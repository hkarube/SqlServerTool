using Microsoft.Data.SqlClient;
using SqlServerTool.Models;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;

namespace SqlServerTool.Services
{
    public class DataMigrationService
    {
        private readonly DatabaseService  _srcDbService;
        private readonly BulkCopyService  _csvService;

        public DataMigrationService(DatabaseService srcDbService)
        {
            _srcDbService = srcDbService;
            _csvService   = new BulkCopyService(srcDbService);
        }

        // ─── 接続ヘルパー ─────────────────────────────────────────────────────

        public static string BuildConnectionString(
            string server, bool windowsAuth, string userId, string password, string database = "master")
        {
            var b = new SqlConnectionStringBuilder
            {
                DataSource              = server,
                InitialCatalog          = database,
                TrustServerCertificate  = true,
                ConnectTimeout          = 15
            };
            if (windowsAuth)
                b.IntegratedSecurity = true;
            else
            {
                b.UserID   = userId;
                b.Password = password;
            }
            return b.ConnectionString;
        }

        public static bool TestConnection(string connStr, out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {
                using var conn = new SqlConnection(connStr);
                conn.Open();
                return true;
            }
            catch (Exception ex) { errorMessage = ex.Message; return false; }
        }

        // ─── ソース情報取得 ───────────────────────────────────────────────────

        public List<MigrationTableItem> GetSourceTables()
        {
            var list = new List<MigrationTableItem>();
            using var conn = _srcDbService.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                SELECT t.name, ISNULL(SUM(p.rows),0)
                FROM sys.tables t
                LEFT JOIN sys.partitions p
                    ON t.object_id = p.object_id AND p.index_id IN (0,1)
                WHERE t.type = 'U'
                GROUP BY t.name
                ORDER BY t.name", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new MigrationTableItem(r.GetString(0), r.GetInt64(1)));
            return list;
        }

        // ─── 宛先 DB 一覧 ─────────────────────────────────────────────────────

        public static List<string> GetDatabaseNames(string connStr)
        {
            var list = new List<string>();
            using var conn = new SqlConnection(connStr);
            conn.Open();
            var cmd = new SqlCommand(
                "SELECT name FROM sys.databases " +
                "WHERE name NOT IN('master','tempdb','model','msdb') ORDER BY name", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
            return list;
        }

        /// <summary>指定したサーバー上に DB が存在するか確認する</summary>
        public static bool DatabaseExists(string serverConnStr, string dbName)
        {
            using var conn = new SqlConnection(serverConnStr) { };
            conn.Open();
            using var cmd = new SqlCommand(
                "SELECT COUNT(1) FROM sys.databases WHERE name = @n", conn);
            cmd.Parameters.AddWithValue("@n", dbName);
            cmd.CommandTimeout = 10;
            return (int)cmd.ExecuteScalar() > 0;
        }

        // ─── FK トポロジカルソート ────────────────────────────────────────────

        public (List<string> Ordered, bool HasCycle) GetTopologicalOrder(List<string> tableNames)
        {
            using var conn = _srcDbService.GetConnection();
            conn.Open();

            // deps[T] = T が参照するテーブルの集合（選択テーブル内のみ）
            var deps = tableNames.ToDictionary(
                t => t,
                _ => new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                StringComparer.OrdinalIgnoreCase);

            var fkCmd = new SqlCommand(@"
                SELECT OBJECT_NAME(fk.parent_object_id)     AS TableName,
                       OBJECT_NAME(fk.referenced_object_id) AS RefTable
                FROM sys.foreign_keys fk
                WHERE OBJECT_NAME(fk.parent_object_id) <>
                      OBJECT_NAME(fk.referenced_object_id)", conn);
            using (var r = fkCmd.ExecuteReader())
                while (r.Read())
                {
                    var tbl = r.GetString(0);
                    var ref_ = r.GetString(1);
                    if (deps.ContainsKey(tbl) && deps.ContainsKey(ref_))
                        deps[tbl].Add(ref_);
                }

            // Kahn's algorithm
            // degree[T]       = 依存先の数（T が待つテーブル数）
            // dependents[R]   = R が終わったら待ち解除するテーブル一覧
            var degree     = tableNames.ToDictionary(t => t, _ => 0, StringComparer.OrdinalIgnoreCase);
            var dependents = tableNames.ToDictionary(
                t => t,
                _ => new List<string>(),
                StringComparer.OrdinalIgnoreCase);

            foreach (var (t, refs) in deps)
                foreach (var r in refs)
                {
                    degree[t]++;
                    dependents[r].Add(t);
                }

            var queue   = new Queue<string>(tableNames.Where(t => degree[t] == 0));
            var ordered = new List<string>();

            while (queue.Count > 0)
            {
                var t = queue.Dequeue();
                ordered.Add(t);
                foreach (var d in dependents[t])
                    if (--degree[d] == 0) queue.Enqueue(d);
            }

            bool hasCycle = ordered.Count < tableNames.Count;
            if (hasCycle)
                foreach (var t in tableNames)
                    if (!ordered.Contains(t, StringComparer.OrdinalIgnoreCase))
                        ordered.Add(t);

            return (ordered, hasCycle);
        }

        // ─── CREATE TABLE スクリプト生成 ──────────────────────────────────────

        public string GetCreateTableScript(string tableName)
        {
            using var conn = _srcDbService.GetConnection();
            conn.Open();

            var colCmd = new SqlCommand(@"
                SELECT c.name, tp.name, c.max_length, c.precision, c.scale,
                       c.is_nullable, c.is_identity,
                       ic.seed_value, ic.increment_value, dc.definition
                FROM sys.columns c
                JOIN sys.types tp ON c.user_type_id = tp.user_type_id
                LEFT JOIN sys.identity_columns ic
                    ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                LEFT JOIN sys.default_constraints dc
                    ON c.object_id = dc.parent_object_id AND c.column_id = dc.parent_column_id
                WHERE c.object_id = OBJECT_ID(@T)
                ORDER BY c.column_id", conn);
            colCmd.Parameters.AddWithValue("@T", tableName);

            var cols = new List<(string Name, string Type, short MaxLen, byte Prec, byte Scale,
                                 bool Nullable, bool IsId, long Seed, long Incr, string? Def)>();
            using (var r = colCmd.ExecuteReader())
                while (r.Read())
                    cols.Add((r.GetString(0), r.GetString(1), r.GetInt16(2), r.GetByte(3), r.GetByte(4),
                              r.GetBoolean(5), r.GetBoolean(6),
                              r.IsDBNull(7) ? 1L : Convert.ToInt64(r[7]),
                              r.IsDBNull(8) ? 1L : Convert.ToInt64(r[8]),
                              r.IsDBNull(9) ? null : r.GetString(9)));

            var pkCmd = new SqlCommand(@"
                SELECT c.name FROM sys.key_constraints kc
                JOIN sys.indexes i
                    ON kc.parent_object_id = i.object_id AND kc.unique_index_id = i.index_id
                JOIN sys.index_columns ic
                    ON i.object_id = ic.object_id AND i.index_id = ic.index_id
                JOIN sys.columns c
                    ON ic.object_id = c.object_id AND ic.column_id = c.column_id
                WHERE kc.type = 'PK' AND kc.parent_object_id = OBJECT_ID(@T)
                ORDER BY ic.key_ordinal", conn);
            pkCmd.Parameters.AddWithValue("@T", tableName);

            var pkCols = new List<string>();
            using (var r = pkCmd.ExecuteReader())
                while (r.Read()) pkCols.Add(r.GetString(0));

            var sb = new StringBuilder();
            sb.AppendLine($"CREATE TABLE [{tableName}] (");
            for (int i = 0; i < cols.Count; i++)
            {
                var (name, type, maxLen, prec, scale, nullable, isId, seed, incr, def) = cols[i];
                bool isLast  = i == cols.Count - 1 && pkCols.Count == 0;
                var  typeDef = FormatDataType(type, maxLen, prec, scale);
                var  idDef   = isId ? $" IDENTITY({seed},{incr})" : "";
                var  nullDef = nullable ? " NULL" : " NOT NULL";
                var  defDef  = def != null ? $" DEFAULT {def}" : "";
                sb.AppendLine($"    [{name}] {typeDef}{idDef}{nullDef}{defDef}{(isLast ? "" : ",")}");
            }
            if (pkCols.Count > 0)
                sb.AppendLine(
                    $"    CONSTRAINT [PK_{tableName}] PRIMARY KEY ({string.Join(", ", pkCols.Select(c => $"[{c}]"))})");
            sb.Append(")");
            return sb.ToString();
        }

        private static string FormatDataType(string typeName, short maxLength, byte precision, byte scale) =>
            typeName.ToLower() switch
            {
                "nvarchar" or "nchar" =>
                    maxLength == -1 ? $"{typeName}(MAX)" : $"{typeName}({maxLength / 2})",
                "varchar" or "char" or "binary" or "varbinary" =>
                    maxLength == -1 ? $"{typeName}(MAX)" : $"{typeName}({maxLength})",
                "decimal" or "numeric" => $"{typeName}({precision},{scale})",
                "datetime2" or "time" or "datetimeoffset" => $"{typeName}({scale})",
                _ => typeName
            };

        // ─── テーブルデータのコピー ───────────────────────────────────────────

        public int CopyTableData(SqlConnection destConn, string tableName, bool preserveIdentity)
        {
            // コピー前に件数を取得
            int rowCount;
            using (var srcConn0 = _srcDbService.GetConnection())
            {
                srcConn0.Open();
                rowCount = (int)(new SqlCommand(
                    $"SELECT COUNT(*) FROM [{tableName}]", srcConn0).ExecuteScalar() ?? 0);
            }
            if (rowCount == 0) return 0;

            using var srcConn = _srcDbService.GetConnection();
            srcConn.Open();
            var cmd = new SqlCommand($"SELECT * FROM [{tableName}]", srcConn)
            {
                CommandTimeout = 600
            };
            using var reader = cmd.ExecuteReader();

            var opts = preserveIdentity
                ? SqlBulkCopyOptions.KeepIdentity
                : SqlBulkCopyOptions.Default;

            using var bulk = new SqlBulkCopy(destConn, opts, null)
            {
                DestinationTableName = $"[{tableName}]",
                BulkCopyTimeout      = 600
            };
            bulk.WriteToServer(reader);
            return rowCount;
        }

        // ─── FK 制約の無効化 / 有効化 ────────────────────────────────────────

        public static void DisableAllForeignKeys(SqlConnection conn)
            => ToggleForeignKeys(conn, enable: false);

        public static void EnableAllForeignKeys(SqlConnection conn)
            => ToggleForeignKeys(conn, enable: true);

        private static void ToggleForeignKeys(SqlConnection conn, bool enable)
        {
            var tables = new List<string>();
            using (var cmd = new SqlCommand(
                "SELECT name FROM sys.tables WHERE type='U'", conn))
            using (var r = cmd.ExecuteReader())
                while (r.Read()) tables.Add(r.GetString(0));

            string sql = enable
                ? "ALTER TABLE [{0}] WITH CHECK CHECK CONSTRAINT ALL"
                : "ALTER TABLE [{0}] NOCHECK CONSTRAINT ALL";

            foreach (var t in tables)
                new SqlCommand(string.Format(sql, t), conn)
                    { CommandTimeout = 60 }.ExecuteNonQuery();
        }

        // ─── 新規 DB 作成 ─────────────────────────────────────────────────────

        public static void CreateDatabase(
            string connStr, string dbName, string? dataFilePath)
        {
            using var conn = new SqlConnection(connStr);
            conn.Open();

            string sql;
            if (string.IsNullOrWhiteSpace(dataFilePath))
            {
                sql = $"CREATE DATABASE [{dbName}]";
            }
            else
            {
                var mdf = Path.Combine(dataFilePath, $"{dbName}.mdf");
                var ldf = Path.Combine(dataFilePath, $"{dbName}_log.ldf");
                sql = $"CREATE DATABASE [{dbName}] " +
                      $"ON PRIMARY (NAME=N'{dbName}', FILENAME=N'{mdf}') " +
                      $"LOG ON (NAME=N'{dbName}_log', FILENAME=N'{ldf}')";
            }
            new SqlCommand(sql, conn) { CommandTimeout = 120 }.ExecuteNonQuery();
        }

        // ─── CSV エクスポート委譲 ─────────────────────────────────────────────

        public int ExportToCsv(string tableName, string filePath,
            Encoding encoding, char delimiter, bool includeHeader)
            => _csvService.ExportToCsv(tableName, filePath, encoding, delimiter, includeHeader);

        // ─── ソース DB のデータファイルフォルダ取得 ───────────────────────────

        public string GetSourceDbDataFolder()
        {
            try
            {
                using var conn = _srcDbService.GetConnection();
                conn.Open();
                var cmd = new SqlCommand(
                    "SELECT TOP 1 physical_name FROM sys.master_files " +
                    "WHERE database_id = DB_ID() AND type = 0", conn);
                var result = cmd.ExecuteScalar() as string;
                return result != null ? Path.GetDirectoryName(result) ?? string.Empty : string.Empty;
            }
            catch { return string.Empty; }
        }

        // ─── 宛先 DB への CSV インポート（静的ヘルパー） ─────────────────────

        /// <summary>
        /// CSV ファイルを指定の SqlConnection のテーブルにインポートする。
        /// ヘッダー行の列名でカラムマッピングを行う。
        /// IDENTITY 列は自動検出し、CSV に値があれば KeepIdentity で保持する。
        /// </summary>
        public static int ImportCsvToTable(
            SqlConnection destConn,
            string tableName,
            string csvFilePath,
            Encoding encoding,
            char delimiter,
            bool hasHeader,
            bool truncateFirst)
        {
            // RFC 4180 対応: 引用符内の改行を正しく処理するため File.ReadAllText を使用
            var records = ParseCsvFileMultiLine(csvFilePath, encoding, delimiter);
            if (records.Count == 0) return 0;

            // ── 宛先テーブルのカラムスキーマを取得 ────────────────────────────
            var destSchema = GetDestColumnSchema(destConn, tableName);

            // ── CSV を DataTable に読み込む ─────────────────────────────────
            var dt       = new DataTable();
            int startRow = 0;

            if (hasHeader && records.Count > 0)
            {
                foreach (var h in records[0])
                    dt.Columns.Add(h, typeof(string));
                startRow = 1;
            }

            for (int i = startRow; i < records.Count; i++)
            {
                var fields = records[i];
                // 全フィールドが空白のレコードはスキップ
                if (fields.All(f => string.IsNullOrWhiteSpace(f))) continue;
                var row    = dt.NewRow();
                for (int j = 0; j < Math.Min(fields.Count, dt.Columns.Count); j++)
                    row[j] = fields[j].Length == 0 ? DBNull.Value : (object)fields[j];
                dt.Rows.Add(row);
            }

            if (dt.Rows.Count == 0) return 0;

            // ── IDENTITY 列の検出と KeepIdentity 判定 ─────────────────────────
            // IDENTITY 列に非 NULL 値が含まれていれば KeepIdentity を使う
            bool useKeepIdentity = false;
            foreach (DataColumn col in dt.Columns)
            {
                if (destSchema.TryGetValue(col.ColumnName, out var info) && info.IsIdentity)
                {
                    bool hasValue = dt.Rows.Cast<DataRow>()
                        .Any(r => r[col] != DBNull.Value);
                    if (hasValue) { useKeepIdentity = true; break; }
                }
            }

            // ── TRUNCATE / DELETE ──────────────────────────────────────────
            if (truncateFirst)
            {
                try
                {
                    new SqlCommand($"TRUNCATE TABLE [{tableName}]", destConn)
                        { CommandTimeout = 120 }.ExecuteNonQuery();
                }
                catch
                {
                    new SqlCommand($"DELETE FROM [{tableName}]", destConn)
                        { CommandTimeout = 120 }.ExecuteNonQuery();
                }
            }

            // ── SqlBulkCopy ───────────────────────────────────────────────
            var opts = useKeepIdentity
                ? SqlBulkCopyOptions.KeepIdentity
                : SqlBulkCopyOptions.Default;

            using var bulk = new SqlBulkCopy(destConn, opts, null)
            {
                DestinationTableName = $"[{tableName}]",
                BulkCopyTimeout      = 300
            };

            // 宛先テーブルに存在する列のみをマッピング対象とする
            // IDENTITY 列が KeepIdentity なし（CSV に値なし）の場合はスキップ
            foreach (DataColumn col in dt.Columns)
            {
                if (!destSchema.TryGetValue(col.ColumnName, out var info)) continue;  // 宛先に存在しない列
                if (info.IsIdentity && !useKeepIdentity) continue;                    // 空 IDENTITY は除外
                bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName);
            }

            bulk.WriteToServer(dt);
            return dt.Rows.Count;
        }

        /// <summary>宛先テーブルのカラム情報（IDENTITY / NULL 許可 / 型名）を取得する</summary>
        private static Dictionary<string, (bool IsIdentity, bool IsNullable, string TypeName)>
            GetDestColumnSchema(SqlConnection conn, string tableName)
        {
            var result = new Dictionary<string, (bool, bool, string)>(StringComparer.OrdinalIgnoreCase);
            using var cmd = new SqlCommand(@"
                SELECT c.name, c.is_identity, c.is_nullable, tp.name
                FROM sys.columns c
                JOIN sys.types tp ON c.user_type_id = tp.user_type_id
                WHERE c.object_id = OBJECT_ID(@T)
                ORDER BY c.column_id", conn);
            cmd.Parameters.AddWithValue("@T", tableName);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                result[r.GetString(0)] = (r.GetBoolean(1), r.GetBoolean(2), r.GetString(3));
            return result;
        }

        /// <summary>
        /// RFC 4180 準拠の CSV パーサー。引用符内の改行（複数行フィールド）を正しく処理する。
        /// File.ReadAllText でファイル全体を読み込み、引用符のネスト状態を行をまたいで追跡する。
        /// </summary>
        private static List<List<string>> ParseCsvFileMultiLine(
            string filePath, Encoding encoding, char delimiter)
        {
            var records       = new List<List<string>>();
            var currentRecord = new List<string>();
            var sb            = new StringBuilder();
            bool inQuote      = false;

            string content = File.ReadAllText(filePath, encoding);
            // BOM が残っている場合は除去
            if (content.Length > 0 && content[0] == '\uFEFF')
                content = content.Substring(1);

            for (int i = 0; i < content.Length; i++)
            {
                char c = content[i];
                if (inQuote)
                {
                    if (c == '"' && i + 1 < content.Length && content[i + 1] == '"')
                    {
                        // エスケープされた引用符 ("") → " として追加
                        sb.Append('"');
                        i++;
                    }
                    else if (c == '"')
                    {
                        inQuote = false;
                    }
                    else
                    {
                        // 引用符内の改行も含めてそのまま追加
                        sb.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuote = true;
                    }
                    else if (c == delimiter)
                    {
                        currentRecord.Add(sb.ToString());
                        sb.Clear();
                    }
                    else if (c == '\r')
                    {
                        // \r\n の \r はスキップ（\n で処理）
                    }
                    else if (c == '\n')
                    {
                        // 引用符外の改行 → レコード終端
                        currentRecord.Add(sb.ToString());
                        sb.Clear();
                        // 空レコード（ファイル末尾の改行など）は追加しない
                        if (currentRecord.Count > 1 ||
                            (currentRecord.Count == 1 && currentRecord[0].Length > 0))
                        {
                            records.Add(currentRecord);
                        }
                        currentRecord = new List<string>();
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
            }

            // ファイル末尾が改行で終わっていない場合の残余処理
            currentRecord.Add(sb.ToString());
            if (currentRecord.Count > 1 ||
                (currentRecord.Count == 1 && currentRecord[0].Length > 0))
            {
                records.Add(currentRecord);
            }

            return records;
        }
    }
}
