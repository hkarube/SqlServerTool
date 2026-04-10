using Microsoft.Data.SqlClient;
using SqlServerTool.Models;
using System.Data;
using System.Text;

namespace SqlServerTool.Services
{
    public class SchemaService
    {
        private readonly DatabaseService _dbService;

        public SchemaService(DatabaseService dbService)
        {
            _dbService = dbService;
        }

        // ─── オブジェクト一覧取得 ───────────────────────────────────────────

        public List<ObjectInfo> GetObjectList(string category) => category switch
        {
            "Tables"      => GetTableList(),
            "Views"       => GetViewList(),
            "StoredProcs" => GetStoredProcList(),
            "Functions"   => GetFunctionList(),
            _             => new()
        };

        private List<ObjectInfo> GetTableList()
        {
            var list = new List<ObjectInfo>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            var sql = @"
                SELECT t.name, s.name AS owner, t.create_date,
                       ISNULL(p.rows, 0) AS row_count,
                       ISNULL(CAST(ep.value AS nvarchar(MAX)), '') AS comment
                FROM sys.tables t
                JOIN sys.schemas s ON t.schema_id = s.schema_id
                LEFT JOIN sys.partitions p ON t.object_id = p.object_id AND p.index_id IN (0,1)
                LEFT JOIN sys.extended_properties ep
                    ON ep.major_id = t.object_id AND ep.minor_id = 0 AND ep.name = 'MS_Description'
                ORDER BY t.name";
            using var cmd = new SqlCommand(sql, conn);
            using var r   = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new ObjectInfo
                {
                    Name       = r.GetString(0),
                    Owner      = r.GetString(1),
                    CreateDate = r.GetDateTime(2),
                    RowCount   = r.GetInt64(3),
                    Comment    = r.GetString(4),
                    ObjectType = "TABLE"
                });
            return list;
        }

        private List<ObjectInfo> GetViewList()
        {
            var list = new List<ObjectInfo>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            var sql = @"
                SELECT v.name, s.name AS owner, v.create_date,
                       ISNULL(CAST(ep.value AS nvarchar(MAX)), '') AS comment
                FROM sys.views v
                JOIN sys.schemas s ON v.schema_id = s.schema_id
                LEFT JOIN sys.extended_properties ep
                    ON ep.major_id = v.object_id AND ep.minor_id = 0 AND ep.name = 'MS_Description'
                ORDER BY v.name";
            using var cmd = new SqlCommand(sql, conn);
            using var r   = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new ObjectInfo
                {
                    Name       = r.GetString(0),
                    Owner      = r.GetString(1),
                    CreateDate = r.GetDateTime(2),
                    Comment    = r.GetString(3),
                    ObjectType = "VIEW"
                });
            return list;
        }

        private List<ObjectInfo> GetStoredProcList()
        {
            var list = new List<ObjectInfo>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            var sql = @"
                SELECT p.name, s.name AS owner, p.create_date,
                       ISNULL(CAST(ep.value AS nvarchar(MAX)), '') AS comment
                FROM sys.procedures p
                JOIN sys.schemas s ON p.schema_id = s.schema_id
                LEFT JOIN sys.extended_properties ep
                    ON ep.major_id = p.object_id AND ep.minor_id = 0 AND ep.name = 'MS_Description'
                ORDER BY p.name";
            using var cmd = new SqlCommand(sql, conn);
            using var r   = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new ObjectInfo
                {
                    Name       = r.GetString(0),
                    Owner      = r.GetString(1),
                    CreateDate = r.GetDateTime(2),
                    Comment    = r.GetString(3),
                    ObjectType = "PROCEDURE"
                });
            return list;
        }

        private List<ObjectInfo> GetFunctionList()
        {
            var list = new List<ObjectInfo>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            var sql = @"
                SELECT o.name, s.name AS owner, o.create_date,
                       ISNULL(CAST(ep.value AS nvarchar(MAX)), '') AS comment
                FROM sys.objects o
                JOIN sys.schemas s ON o.schema_id = s.schema_id
                LEFT JOIN sys.extended_properties ep
                    ON ep.major_id = o.object_id AND ep.minor_id = 0 AND ep.name = 'MS_Description'
                WHERE o.type IN ('FN','IF','TF','FS','FT')
                ORDER BY o.name";
            using var cmd = new SqlCommand(sql, conn);
            using var r   = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new ObjectInfo
                {
                    Name       = r.GetString(0),
                    Owner      = r.GetString(1),
                    CreateDate = r.GetDateTime(2),
                    Comment    = r.GetString(3),
                    ObjectType = "FUNCTION"
                });
            return list;
        }

        // ─── 名前一覧（ツリー用） ──────────────────────────────────────────

        public List<string> GetTableNames()
        {
            var list = new List<string>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            using var cmd = new SqlCommand(
                "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES " +
                "WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
            return list;
        }

        public List<string> GetViewNames()
        {
            var list = new List<string>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            using var cmd = new SqlCommand(
                "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS ORDER BY TABLE_NAME", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
            return list;
        }

        public List<string> GetStoredProcNames()
        {
            var list = new List<string>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            using var cmd = new SqlCommand(
                "SELECT ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES " +
                "WHERE ROUTINE_TYPE = 'PROCEDURE' ORDER BY ROUTINE_NAME", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
            return list;
        }

        public List<string> GetFunctionNames()
        {
            var list = new List<string>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            using var cmd = new SqlCommand(
                "SELECT ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES " +
                "WHERE ROUTINE_TYPE = 'FUNCTION' ORDER BY ROUTINE_NAME", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
            return list;
        }

        // ─── カラム定義取得 ────────────────────────────────────────────────

        public List<ColumnInfo> GetColumns(string tableName)
        {
            var list = new List<ColumnInfo>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            var sql = @"
                SELECT c.COLUMN_NAME, c.DATA_TYPE, c.CHARACTER_MAXIMUM_LENGTH,
                       c.IS_NULLABLE, c.COLUMN_DEFAULT,
                       ISNULL(CAST(ep.value AS nvarchar(MAX)), '') AS LOGICAL_NAME,
                       CASE WHEN pk.COLUMN_NAME IS NOT NULL THEN 1 ELSE 0 END AS IS_PK,
                       CASE WHEN fk.COLUMN_NAME IS NOT NULL THEN 1 ELSE 0 END AS IS_FK,
                       ISNULL(COLUMNPROPERTY(OBJECT_ID(@T), c.COLUMN_NAME, 'IsIdentity'), 0) AS IS_IDENTITY
                FROM INFORMATION_SCHEMA.COLUMNS c
                LEFT JOIN sys.extended_properties ep
                    ON ep.major_id = OBJECT_ID(c.TABLE_NAME)
                    AND ep.minor_id = c.ORDINAL_POSITION
                    AND ep.name = 'MS_Description'
                LEFT JOIN (
                    SELECT col.name AS COLUMN_NAME
                    FROM sys.indexes idx
                    JOIN sys.index_columns ic  ON idx.object_id = ic.object_id AND idx.index_id = ic.index_id
                    JOIN sys.columns     col ON ic.object_id  = col.object_id  AND ic.column_id = col.column_id
                    WHERE idx.object_id = OBJECT_ID(@T) AND idx.is_primary_key = 1
                ) pk ON pk.COLUMN_NAME = c.COLUMN_NAME
                LEFT JOIN (
                    SELECT ku.COLUMN_NAME
                    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
                    JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku ON tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
                    WHERE tc.TABLE_NAME = @T AND tc.CONSTRAINT_TYPE = 'FOREIGN KEY'
                ) fk ON fk.COLUMN_NAME = c.COLUMN_NAME
                WHERE c.TABLE_NAME = @T
                ORDER BY c.ORDINAL_POSITION";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@T", tableName);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var dataType = r.GetString(1);
                var maxLen   = r.IsDBNull(2) ? "" : r.GetInt32(2).ToString();
                if (maxLen == "-1") maxLen = "MAX";
                var typeDisp = maxLen == "" ? dataType : $"{dataType}({maxLen})";
                list.Add(new ColumnInfo
                {
                    ColumnName   = r.GetString(0),
                    DataType     = typeDisp,
                    IsNullable   = r.GetString(3) == "YES",
                    DefaultValue = r.IsDBNull(4) ? "" : r.GetString(4),
                    LogicalName  = r.GetString(5),
                    IsPrimaryKey = r.GetInt32(6) == 1,
                    IsForeignKey = r.GetInt32(7) == 1,
                    IsIdentity   = r.GetInt32(8) == 1,
                });
            }
            return list;
        }

        // ─── テーブルデータ取得 ────────────────────────────────────────────

        public (DataTable display, DataTable original, long totalCount) GetTableData(string tableName, string fullSql)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();

            // 全件数
            long totalCount = 0;
            using (var cntCmd = new SqlCommand($"SELECT COUNT(*) FROM [{tableName}]", conn))
                totalCount = (int)cntCmd.ExecuteScalar()!;

            // ユーザーSQLをそのまま実行（型情報を保持した元データ）
            var adapter = new SqlDataAdapter(fullSql, conn);
            var original = new DataTable();
            adapter.Fill(original);

            // 表示用にすべて文字列変換
            var display = new DataTable();
            foreach (DataColumn col in original.Columns)
                display.Columns.Add(col.ColumnName, typeof(string));
            foreach (DataRow row in original.Rows)
            {
                var nr = display.NewRow();
                foreach (DataColumn col in original.Columns)
                    nr[col.ColumnName] = row[col] == DBNull.Value ? "(NULL)" : row[col]!.ToString();
                display.Rows.Add(nr);
            }
            return (display, original, totalCount);
        }

        // ─── SQL定義取得（SP / ファンクション / ビュー） ──────────────────

        public string GetSqlDefinition(string objectName)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(
                "SELECT sm.definition FROM sys.sql_modules sm " +
                "WHERE sm.object_id = OBJECT_ID(@Name)", conn);
            cmd.Parameters.AddWithValue("@Name", objectName);
            return cmd.ExecuteScalar()?.ToString()
                   ?? $"-- 定義を取得できませんでした: {objectName}";
        }

        // ─── SP / Function 編集 ───────────────────────────────────────────

        /// <summary>SQL の構文チェック（SET PARSEONLY ON）。エラーがあればメッセージを返す。</summary>
        public string? CheckSqlSyntax(string sql)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();
            new SqlCommand("SET PARSEONLY ON", conn).ExecuteNonQuery();
            try
            {
                new SqlCommand(sql, conn).ExecuteNonQuery();
                return null;   // 構文エラーなし
            }
            catch (SqlException ex)
            {
                return ex.Message;
            }
            finally
            {
                try { new SqlCommand("SET PARSEONLY OFF", conn).ExecuteNonQuery(); } catch { }
            }
        }

        /// <summary>ALTER PROCEDURE / FUNCTION などを実行する</summary>
        public void ExecuteAlterObject(string sql)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();
            new SqlCommand(sql, conn) { CommandTimeout = 60 }.ExecuteNonQuery();
        }

        // ─── スクリプト生成用 ──────────────────────────────────────────────

        public List<ScriptColumnInfo> GetScriptColumns(string tableName)
        {
            var list = new List<ScriptColumnInfo>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            var sql = @"
                SELECT c.name, tp.name AS type_name,
                    CASE
                        WHEN tp.name IN ('nvarchar','nchar') THEN
                            CASE WHEN c.max_length = -1 THEN '(MAX)'
                                 ELSE '(' + CAST(c.max_length/2 AS varchar) + ')' END
                        WHEN tp.name IN ('varchar','char','binary','varbinary') THEN
                            CASE WHEN c.max_length = -1 THEN '(MAX)'
                                 ELSE '(' + CAST(c.max_length AS varchar) + ')' END
                        WHEN tp.name IN ('decimal','numeric') THEN
                            '(' + CAST(c.precision AS varchar) + ',' + CAST(c.scale AS varchar) + ')'
                        WHEN tp.name = 'float' THEN '(' + CAST(c.precision AS varchar) + ')'
                        ELSE ''
                    END AS type_suffix,
                    c.is_nullable, c.is_identity,
                    ISNULL(CAST(ic.seed_value      AS bigint), 1) AS seed_val,
                    ISNULL(CAST(ic.increment_value AS bigint), 1) AS inc_val,
                    ISNULL(dc.definition, '') AS default_def
                FROM sys.columns c
                JOIN sys.types tp ON c.user_type_id = tp.user_type_id
                LEFT JOIN sys.default_constraints  dc ON c.default_object_id = dc.object_id
                LEFT JOIN sys.identity_columns     ic ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                WHERE c.object_id = OBJECT_ID(@T)
                ORDER BY c.column_id";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@T", tableName);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new ScriptColumnInfo
                {
                    Name           = r.GetString(0),
                    TypeName       = r.GetString(1),
                    TypeSuffix     = r.GetString(2),
                    IsNullable     = r.GetBoolean(3),
                    IsIdentity     = r.GetBoolean(4),
                    SeedValue      = r.GetInt64(5),
                    IncrementValue = r.GetInt64(6),
                    DefaultDef     = r.GetString(7),
                });
            return list;
        }

        public (string constraintName, List<string> columns) GetPrimaryKey(string tableName)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();
            var sql = @"
                SELECT kc.name, c.name
                FROM sys.key_constraints kc
                JOIN sys.index_columns ic ON kc.parent_object_id = ic.object_id AND kc.unique_index_id = ic.index_id
                JOIN sys.columns c        ON ic.object_id = c.object_id AND ic.column_id = c.column_id
                WHERE kc.parent_object_id = OBJECT_ID(@T) AND kc.type = 'PK'
                ORDER BY ic.key_ordinal";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@T", tableName);
            using var r  = cmd.ExecuteReader();
            var cName    = string.Empty;
            var cols     = new List<string>();
            while (r.Read()) { cName = r.GetString(0); cols.Add(r.GetString(1)); }
            return (cName, cols);
        }

        public string GetFilegroup(string tableName)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                SELECT ds.name FROM sys.indexes i
                JOIN sys.data_spaces ds ON i.data_space_id = ds.data_space_id
                WHERE i.object_id = OBJECT_ID(@T) AND i.index_id <= 1", conn);
            cmd.Parameters.AddWithValue("@T", tableName);
            return cmd.ExecuteScalar()?.ToString() ?? "PRIMARY";
        }

        public List<string> GetViewColumns(string viewName)
        {
            var list = new List<string>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(
                "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS " +
                "WHERE TABLE_NAME = @N ORDER BY ORDINAL_POSITION", conn);
            cmd.Parameters.AddWithValue("@N", viewName);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
            return list;
        }

        // ─── テーブル新規作成 ────────────────────────────────────────────────

        public void CreateTable(string tableName, IReadOnlyList<Models.EditableColumnInfo> columns)
        {
            if (columns.Count == 0)
                throw new ArgumentException("列を1つ以上定義してください。");

            var sb = new StringBuilder();
            sb.AppendLine($"CREATE TABLE [{tableName}] (");

            var pkCols = columns.Where(c => c.IsPrimaryKey).Select(c => c.ColumnName).ToList();

            for (int i = 0; i < columns.Count; i++)
            {
                var col   = columns[i];
                var comma = (i < columns.Count - 1 || pkCols.Count > 0) ? "," : "";
                var nullable = col.IsNullable ? "NULL" : "NOT NULL";
                var def = string.IsNullOrWhiteSpace(col.DefaultValue)
                    ? "" : $" DEFAULT {col.DefaultValue}";
                sb.AppendLine($"    [{col.ColumnName}] {col.DataType} {nullable}{def}{comma}");
            }

            if (pkCols.Count > 0)
            {
                var pkList = string.Join(", ", pkCols.Select(c => $"[{c}]"));
                sb.AppendLine($"    CONSTRAINT [PK_{tableName}] PRIMARY KEY ({pkList})");
            }

            sb.Append(")");

            using var conn = _dbService.GetConnection();
            conn.Open();
            new SqlCommand(sb.ToString(), conn).ExecuteNonQuery();
        }

        // ─── テーブルコピー ───────────────────────────────────────────────

        public void CopyTable(string sourceName, string destName)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();
            var sql = $"SELECT * INTO [{destName}] FROM [{sourceName}]";
            new SqlCommand(sql, conn).ExecuteNonQuery();
        }

        // ─── オブジェクト名変更 ───────────────────────────────────────────

        public void RenameObject(string oldName, string newName)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();
            var cmd = new SqlCommand("EXEC sp_rename @old, @new", conn);
            cmd.Parameters.AddWithValue("@old", oldName);
            cmd.Parameters.AddWithValue("@new", newName);
            cmd.ExecuteNonQuery();
        }

        public bool ObjectNameExists(string name)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(
                "SELECT COUNT(*) FROM sys.objects WHERE name = @N", conn);
            cmd.Parameters.AddWithValue("@N", name);
            return (int)cmd.ExecuteScalar()! > 0;
        }

        // ─── オブジェクト削除 ─────────────────────────────────────────────

        public void DropObject(string name, string objectType)
        {
            var sql = objectType switch
            {
                "TABLE"     => $"DROP TABLE [{name}]",
                "VIEW"      => $"DROP VIEW [{name}]",
                "PROCEDURE" => $"DROP PROCEDURE [{name}]",
                "FUNCTION"  => $"DROP FUNCTION [{name}]",
                _           => throw new ArgumentException($"不明なオブジェクト種別: {objectType}")
            };
            using var conn = _dbService.GetConnection();
            conn.Open();
            new SqlCommand(sql, conn).ExecuteNonQuery();
        }

        // ─── テーブル構造変更 ─────────────────────────────────────────────

        /// <summary>
        /// テーブルの列構造を変更する。
        /// シンプルな ALTER TABLE で対応できない場合は自動マイグレーションを実行する。
        /// </summary>
        public void AlterTableStructure(
            string tableName,
            IReadOnlyList<Models.ColumnInfo> original,
            IReadOnlyList<Models.EditableColumnInfo> edited)
        {
            // 新規列が既存列の間に挿入されている場合は ALTER TABLE では順序を保てないので
            // 直接マイグレーションへ
            if (HasMidInsertion(edited))
            {
                ApplyMigration(tableName, original, edited);
                return;
            }

            try
            {
                ApplySimpleAlters(tableName, original, edited);
            }
            catch
            {
                ApplyMigration(tableName, original, edited);
            }
        }

        /// <summary>新規列が既存列より前に挿入されているか判定</summary>
        private static bool HasMidInsertion(IReadOnlyList<Models.EditableColumnInfo> edited)
        {
            // 末尾から走査して最後の「既存列」インデックスを求める
            int lastExistingIdx = -1;
            for (int i = edited.Count - 1; i >= 0; i--)
            {
                if (!edited[i].IsNew) { lastExistingIdx = i; break; }
            }
            // それより前に新規列があれば中間挿入
            for (int i = 0; i < lastExistingIdx; i++)
                if (edited[i].IsNew) return true;
            return false;
        }

        private void ApplySimpleAlters(
            string tableName,
            IReadOnlyList<Models.ColumnInfo> original,
            IReadOnlyList<Models.EditableColumnInfo> edited)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();
            using var tx = conn.BeginTransaction();

            var origNames = original.Select(c => c.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var editNames = edited.Select(c => c.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);

            // リネーム
            foreach (var e in edited.Where(e => !e.IsNew && !string.Equals(e.OriginalName, e.ColumnName, StringComparison.OrdinalIgnoreCase)))
            {
                var sql = $"EXEC sp_rename '[{tableName}].[{e.OriginalName}]', '{e.ColumnName}', 'COLUMN'";
                new SqlCommand(sql, conn, tx) { CommandTimeout = 300 }.ExecuteNonQuery();
            }

            // 列の追加
            foreach (var e in edited.Where(e => e.IsNew || !origNames.Contains(e.OriginalName)))
            {
                var nullStr = e.IsNullable ? "NULL" : "NOT NULL";
                var defStr  = string.IsNullOrWhiteSpace(e.DefaultValue) ? "" : $" DEFAULT {e.DefaultValue}";
                new SqlCommand($"ALTER TABLE [{tableName}] ADD [{e.ColumnName}] {e.DataType} {nullStr}{defStr}", conn, tx) { CommandTimeout = 300 }.ExecuteNonQuery();
            }

            // 列の削除
            foreach (var o in original.Where(o => !editNames.Contains(o.ColumnName)))
            {
                new SqlCommand($"ALTER TABLE [{tableName}] DROP COLUMN [{o.ColumnName}]", conn, tx) { CommandTimeout = 300 }.ExecuteNonQuery();
            }

            // 型・NULL変更
            foreach (var e in edited.Where(e => !e.IsNew && origNames.Contains(e.OriginalName)))
            {
                var orig = original.First(c => string.Equals(c.ColumnName, e.OriginalName, StringComparison.OrdinalIgnoreCase));
                if (orig.DataType != e.DataType || orig.IsNullable != e.IsNullable)
                {
                    var nullStr = e.IsNullable ? "NULL" : "NOT NULL";
                    new SqlCommand($"ALTER TABLE [{tableName}] ALTER COLUMN [{e.ColumnName}] {e.DataType} {nullStr}", conn, tx) { CommandTimeout = 300 }.ExecuteNonQuery();
                }
            }

            tx.Commit();
        }

        private void ApplyMigration(
            string tableName,
            IReadOnlyList<Models.ColumnInfo> original,
            IReadOnlyList<Models.EditableColumnInfo> edited)
        {
            var bakName = $"{tableName}_bak_{DateTime.Now:yyyyMMddHHmm}";

            using var conn = _dbService.GetConnection();
            conn.Open();
            using var tx = conn.BeginTransaction();

            // 1. データバックアップ
            new SqlCommand($"SELECT * INTO [{bakName}] FROM [{tableName}]", conn, tx) { CommandTimeout = 300 }.ExecuteNonQuery();

            // 2. 元テーブル削除（PK制約などを先に削除）
            DropTableConstraints(tableName, conn, tx);
            new SqlCommand($"DROP TABLE [{tableName}]", conn, tx) { CommandTimeout = 300 }.ExecuteNonQuery();

            // 3. 新テーブル作成
            var createSql = BuildCreateTableSql(tableName, edited);
            new SqlCommand(createSql, conn, tx) { CommandTimeout = 300 }.ExecuteNonQuery();

            // 4. 共通列のデータコピー
            var origColNames = original.Select(c => c.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var commonCols   = edited
                .Where(e => !e.IsNew && origColNames.Contains(e.OriginalName))
                .Select(e => $"[{e.ColumnName}]")
                .ToList();

            if (commonCols.Count > 0)
            {
                var colList = string.Join(", ", commonCols);
                new SqlCommand($"INSERT INTO [{tableName}] ({colList}) SELECT {colList} FROM [{bakName}]", conn, tx) { CommandTimeout = 300 }.ExecuteNonQuery();
            }

            // 5. バックアップテーブル削除
            new SqlCommand($"DROP TABLE [{bakName}]", conn, tx) { CommandTimeout = 300 }.ExecuteNonQuery();

            tx.Commit();
        }

        private static void DropTableConstraints(string tableName, SqlConnection conn, SqlTransaction tx)
        {
            // PK / UQ 制約を取得して削除
            var sql = @"
                SELECT kc.name FROM sys.key_constraints kc
                WHERE kc.parent_object_id = OBJECT_ID(@T)";
            var cmd = new SqlCommand(sql, conn, tx);
            cmd.Parameters.AddWithValue("@T", tableName);
            var constraints = new List<string>();
            using (var r = cmd.ExecuteReader())
                while (r.Read()) constraints.Add(r.GetString(0));

            foreach (var c in constraints)
                new SqlCommand($"ALTER TABLE [{tableName}] DROP CONSTRAINT [{c}]", conn, tx).ExecuteNonQuery();
        }

        private static string BuildCreateTableSql(string tableName, IReadOnlyList<Models.EditableColumnInfo> cols)
        {
            var lines = cols.Select(c =>
            {
                var nullStr = c.IsNullable ? "NULL" : "NOT NULL";
                var defStr  = string.IsNullOrWhiteSpace(c.DefaultValue) ? "" : $" DEFAULT {c.DefaultValue}";
                return $"    [{c.ColumnName}] {c.DataType} {nullStr}{defStr}";
            }).ToList();

            var pkCols = cols.Where(c => c.IsPrimaryKey).Select(c => $"[{c.ColumnName}]").ToList();
            if (pkCols.Count > 0)
                lines.Add($"    CONSTRAINT [PK_{tableName}] PRIMARY KEY ({string.Join(", ", pkCols)})");

            return $"CREATE TABLE [{tableName}] (\n{string.Join(",\n", lines)}\n)";
        }

        // ─── SQL 文字列生成（プレビュー・履歴用）─────────────────────────────

        /// <summary>INSERT SQL を文字列として生成（値はリテラル展開）</summary>
        public string BuildInsertSql(
            string tableName,
            IEnumerable<(string Name, bool IsNull, string Value)> columns)
        {
            var list    = columns.ToList();
            var cols    = string.Join(", ", list.Select(c => $"[{c.Name}]"));
            var vals    = string.Join(",\n        ", list.Select(c =>
                c.IsNull ? "NULL" : $"'{c.Value.Replace("'", "''")}'"));
            return $"INSERT INTO [{tableName}]\n    ({cols})\nVALUES\n    ({vals});";
        }

        /// <summary>UPDATE SQL を文字列として生成</summary>
        public string BuildUpdateSql(
            string tableName,
            IEnumerable<(string Name, bool IsNull, string Value)> setColumns,
            IEnumerable<(string Name, object? Value)> whereColumns,
            bool topOne = false)
        {
            var setList   = setColumns.ToList();
            var whereList = whereColumns.ToList();
            var top       = topOne ? "TOP(1) " : "";
            var sets   = string.Join(",\n    ", setList.Select(f =>
                $"[{f.Name}] = {(f.IsNull ? "NULL" : $"'{f.Value.Replace("'", "''")}'")}" ));
            var wheres = string.Join("\n  AND ", whereList.Select(f =>
                f.Value == null
                    ? $"[{f.Name}] IS NULL"
                    : $"[{f.Name}] = '{f.Value.ToString()?.Replace("'", "''")}'"));
            return $"UPDATE {top}[{tableName}]\nSET\n    {sets}\nWHERE {wheres};";
        }

        /// <summary>DELETE SQL を文字列として生成</summary>
        public string BuildDeleteSql(
            string tableName,
            IEnumerable<(string Name, object? Value)> whereColumns,
            bool topOne = false)
        {
            var list   = whereColumns.ToList();
            var top    = topOne ? "TOP(1) " : "";
            var wheres = string.Join("\n  AND ", list.Select(f =>
                f.Value == null
                    ? $"[{f.Name}] IS NULL"
                    : $"[{f.Name}] = '{f.Value.ToString()?.Replace("'", "''")}'"));
            return $"DELETE {top}FROM [{tableName}]\nWHERE {wheres};";
        }

        /// <summary>
        /// テーブル構造変更 SQL を文字列として生成する（プレビュー用）。
        /// 中間挿入がある場合はマイグレーションスクリプトを返す。
        /// </summary>
        public string BuildAlterTableSql(
            string tableName,
            IReadOnlyList<Models.ColumnInfo> original,
            IReadOnlyList<Models.EditableColumnInfo> edited)
        {
            return HasMidInsertion(edited)
                ? BuildMigrationSqlText(tableName, original, edited)
                : BuildSimpleAlterSqlText(tableName, original, edited);
        }

        private string BuildSimpleAlterSqlText(
            string tableName,
            IReadOnlyList<Models.ColumnInfo> original,
            IReadOnlyList<Models.EditableColumnInfo> edited)
        {
            var sb        = new StringBuilder();
            var origNames = original.Select(c => c.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var editNames = edited.Select(c => c.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var e in edited.Where(e => !e.IsNew
                && !string.Equals(e.OriginalName, e.ColumnName, StringComparison.OrdinalIgnoreCase)))
                sb.AppendLine($"EXEC sp_rename '[{tableName}].[{e.OriginalName}]', '{e.ColumnName}', 'COLUMN';");

            foreach (var e in edited.Where(e => e.IsNew || !origNames.Contains(e.OriginalName)))
            {
                var nullStr = e.IsNullable ? "NULL" : "NOT NULL";
                var defStr  = string.IsNullOrWhiteSpace(e.DefaultValue) ? "" : $" DEFAULT {e.DefaultValue}";
                sb.AppendLine($"ALTER TABLE [{tableName}] ADD [{e.ColumnName}] {e.DataType} {nullStr}{defStr};");
            }

            foreach (var o in original.Where(o => !editNames.Contains(o.ColumnName)))
                sb.AppendLine($"ALTER TABLE [{tableName}] DROP COLUMN [{o.ColumnName}];");

            foreach (var e in edited.Where(e => !e.IsNew && origNames.Contains(e.OriginalName)))
            {
                var orig = original.First(c => string.Equals(c.ColumnName, e.OriginalName, StringComparison.OrdinalIgnoreCase));
                if (orig.DataType != e.DataType || orig.IsNullable != e.IsNullable)
                {
                    var nullStr = e.IsNullable ? "NULL" : "NOT NULL";
                    sb.AppendLine($"ALTER TABLE [{tableName}] ALTER COLUMN [{e.ColumnName}] {e.DataType} {nullStr};");
                }
            }

            return sb.Length > 0 ? sb.ToString().TrimEnd() : "-- 変更なし";
        }

        private string BuildMigrationSqlText(
            string tableName,
            IReadOnlyList<Models.ColumnInfo> original,
            IReadOnlyList<Models.EditableColumnInfo> edited)
        {
            var bakName = $"{tableName}_bak_YYYYMMDDHHMI";
            var sb = new StringBuilder();

            sb.AppendLine($"-- =============================================");
            sb.AppendLine($"-- マイグレーション: [{tableName}]");
            sb.AppendLine($"-- 列の中間挿入があるため一時テーブル経由で再作成します");
            sb.AppendLine($"-- （実際のバックアップテーブル名には実行時の日時が入ります）");
            sb.AppendLine($"-- =============================================");
            sb.AppendLine();
            sb.AppendLine($"-- Step 1: データバックアップ");
            sb.AppendLine($"SELECT * INTO [{bakName}] FROM [{tableName}];");
            sb.AppendLine();
            sb.AppendLine($"-- Step 2: 制約を削除してテーブルを削除");
            sb.AppendLine($"-- (PK / UQ 制約を先に DROP してから)");
            sb.AppendLine($"DROP TABLE [{tableName}];");
            sb.AppendLine();
            sb.AppendLine($"-- Step 3: 新しいテーブルを作成");
            sb.AppendLine(BuildCreateTableSql(tableName, edited) + ";");
            sb.AppendLine();

            var origColNames = original.Select(c => c.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var commonCols   = edited
                .Where(e => !e.IsNew && origColNames.Contains(e.OriginalName))
                .Select(e => $"[{e.ColumnName}]")
                .ToList();

            if (commonCols.Count > 0)
            {
                var colList = string.Join(", ", commonCols);
                sb.AppendLine($"-- Step 4: データコピー（共通列のみ）");
                sb.AppendLine($"INSERT INTO [{tableName}] ({colList})");
                sb.AppendLine($"SELECT {colList} FROM [{bakName}];");
                sb.AppendLine();
            }

            sb.AppendLine($"-- Step 5: バックアップテーブルを削除");
            sb.AppendLine($"DROP TABLE [{bakName}];");
            return sb.ToString();
        }

        // ─── 行更新 ───────────────────────────────────────────────────────

        /// <summary>指定テーブルに 1 行 INSERT する</summary>
        public void InsertRow(
            string tableName,
            IEnumerable<(string Name, bool IsNull, string Value)> columns)
        {
            var list = columns.ToList();
            if (list.Count == 0)
                throw new InvalidOperationException("挿入する列がありません。");

            var colNames   = list.Select(c => $"[{c.Name}]");
            var paramNames = list.Select((_, i) => $"@v{i}");
            var sql = $"INSERT INTO [{tableName}] ({string.Join(", ", colNames)}) " +
                      $"VALUES ({string.Join(", ", paramNames)})";

            using var conn = _dbService.GetConnection();
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            for (int i = 0; i < list.Count; i++)
                cmd.Parameters.AddWithValue($"@v{i}",
                    list[i].IsNull ? (object)DBNull.Value : list[i].Value);
            cmd.ExecuteNonQuery();
        }

        /// <summary>指定テーブルの 1 行を UPDATE する</summary>
        /// <param name="topOne">PK なし時は TOP(1) を付けて重複行の多重更新を防ぐ</param>
        public void UpdateRow(
            string tableName,
            IEnumerable<(string Name, bool IsNull, string Value)> setColumns,
            IEnumerable<(string Name, object? Value)> whereColumns,
            bool topOne = false)
        {
            var setList   = setColumns.ToList();
            var whereList = whereColumns.ToList();

            if (setList.Count == 0)
                throw new InvalidOperationException("更新する列がありません。");
            if (whereList.Count == 0)
                throw new InvalidOperationException("WHERE条件が空です。");

            var setClauses   = setList.Select((f, i) => $"[{f.Name}] = @s{i}");
            var whereClauses = whereList.Select((f, i) =>
                f.Value == null ? $"[{f.Name}] IS NULL" : $"[{f.Name}] = @w{i}");

            var topClause = topOne ? "TOP(1) " : "";
            var sql = $"UPDATE {topClause}[{tableName}] " +
                      $"SET {string.Join(", ", setClauses)} " +
                      $"WHERE {string.Join(" AND ", whereClauses)}";

            using var conn = _dbService.GetConnection();
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);

            for (int i = 0; i < setList.Count; i++)
            {
                var f = setList[i];
                cmd.Parameters.AddWithValue($"@s{i}",
                    f.IsNull ? (object)DBNull.Value : f.Value);
            }
            for (int i = 0; i < whereList.Count; i++)
            {
                if (whereList[i].Value != null)
                    cmd.Parameters.AddWithValue($"@w{i}", whereList[i].Value);
            }

            int affected = cmd.ExecuteNonQuery();
            if (affected == 0)
                throw new InvalidOperationException("更新対象行が見つかりませんでした。");
        }

        /// <summary>指定テーブルの 1 行を DELETE する</summary>
        /// <param name="topOne">PK なし時は TOP(1) を付けて重複行の多重削除を防ぐ</param>
        public void DeleteRow(
            string tableName,
            IEnumerable<(string Name, object? Value)> whereColumns,
            bool topOne = false)
        {
            var list = whereColumns.ToList();
            if (list.Count == 0)
                throw new InvalidOperationException("WHERE条件が空です。");

            var clauses = list.Select((f, i) =>
                f.Value == null ? $"[{f.Name}] IS NULL" : $"[{f.Name}] = @p{i}");

            var topClause = topOne ? "TOP(1) " : "";
            var sql = $"DELETE {topClause}FROM [{tableName}] WHERE {string.Join(" AND ", clauses)}";

            using var conn = _dbService.GetConnection();
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].Value != null)
                    cmd.Parameters.AddWithValue($"@p{i}", list[i].Value);
            }
            int affected = cmd.ExecuteNonQuery();
            if (affected == 0)
                throw new InvalidOperationException($"削除対象行が見つかりませんでした。\nSQL: {sql}");
        }
    }

    public class ScriptColumnInfo
    {
        public string Name           { get; set; } = string.Empty;
        public string TypeName       { get; set; } = string.Empty;
        public string TypeSuffix     { get; set; } = string.Empty;
        public bool   IsNullable     { get; set; }
        public bool   IsIdentity     { get; set; }
        public long   SeedValue      { get; set; }
        public long   IncrementValue { get; set; }
        public string DefaultDef     { get; set; } = string.Empty;
    }
}
