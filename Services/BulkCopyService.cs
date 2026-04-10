using Microsoft.Data.SqlClient;
using SqlServerTool.Models;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;

namespace SqlServerTool.Services
{
    public class BulkCopyService
    {
        private readonly DatabaseService _dbService;

        public BulkCopyService(DatabaseService dbService)
        {
            _dbService = dbService;
        }

        // ─── エクスポート ─────────────────────────────────────────────────────

        /// <summary>テーブルデータを CSV ファイルに出力する</summary>
        public int ExportToCsv(
            string tableName,
            string filePath,
            Encoding encoding,
            char delimiter,
            bool includeHeader)
        {
            using var conn = _dbService.GetConnection();
            conn.Open();

            var cmd = new SqlCommand($"SELECT * FROM [{tableName}]", conn)
            {
                CommandTimeout = 300
            };
            using var reader = cmd.ExecuteReader();

            using var writer = new StreamWriter(filePath, append: false, encoding: encoding);

            if (includeHeader)
            {
                var headers = Enumerable.Range(0, reader.FieldCount)
                    .Select(i => EscapeCsv(reader.GetName(i), delimiter));
                writer.WriteLine(string.Join(delimiter, headers));
            }

            int rowCount = 0;
            while (reader.Read())
            {
                var values = Enumerable.Range(0, reader.FieldCount)
                    .Select(i => EscapeCsv(FormatFieldValue(reader, i), delimiter));
                writer.WriteLine(string.Join(delimiter, values));
                rowCount++;
            }

            return rowCount;
        }

        // ─── インポート ─────────────────────────────────────────────────────

        /// <summary>CSV ファイルからテーブルにデータを一括インポートする</summary>
        public int ImportFromCsv(
            string tableName,
            string filePath,
            Encoding encoding,
            char delimiter,
            bool hasHeader,
            bool truncateFirst)
        {
            // CSV を DataTable に読み込む
            var dt = ReadCsvToDataTable(filePath, encoding, delimiter, hasHeader, tableName);

            using var conn = _dbService.GetConnection();
            conn.Open();

            // TRUNCATE または DELETE
            if (truncateFirst)
            {
                using var truncCmd = new SqlCommand($"TRUNCATE TABLE [{tableName}]", conn)
                {
                    CommandTimeout = 120
                };
                try { truncCmd.ExecuteNonQuery(); }
                catch
                {
                    // TRUNCATE が使えない場合（FK制約など）は DELETE にフォールバック
                    using var delCmd = new SqlCommand($"DELETE FROM [{tableName}]", conn)
                    {
                        CommandTimeout = 120
                    };
                    delCmd.ExecuteNonQuery();
                }
            }

            // SqlBulkCopy で一括挿入
            using var bulk = new SqlBulkCopy(conn)
            {
                DestinationTableName = $"[{tableName}]",
                BulkCopyTimeout      = 300
            };
            foreach (DataColumn col in dt.Columns)
                bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName);

            bulk.WriteToServer(dt);
            return dt.Rows.Count;
        }

        /// <summary>CSVファイルを読み込んで DataTable を返す</summary>
        private DataTable ReadCsvToDataTable(
            string filePath,
            Encoding encoding,
            char delimiter,
            bool hasHeader,
            string tableName)
        {
            var dt = new DataTable();
            var lines = File.ReadAllLines(filePath, encoding);
            if (lines.Length == 0) return dt;

            // テーブルの列情報を取得して DataTable のスキーマを構築
            var columns = GetTableColumns(tableName);

            int startRow = 0;
            if (hasHeader)
            {
                // ヘッダー行の列名でマッピング
                var headerCols = ParseCsvLine(lines[0], delimiter);
                foreach (var h in headerCols)
                {
                    var matched = columns.FirstOrDefault(
                        c => string.Equals(c.ColumnName, h, StringComparison.OrdinalIgnoreCase));
                    var colName = matched == default ? h : matched.ColumnName;
                    dt.Columns.Add(colName, typeof(string));
                }
                startRow = 1;
            }
            else
            {
                // ヘッダーなしの場合はテーブルの列順に従う
                foreach (var col in columns)
                    dt.Columns.Add(col.ColumnName, typeof(string));
            }

            for (int i = startRow; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var fields = ParseCsvLine(lines[i], delimiter);
                var row = dt.NewRow();
                for (int j = 0; j < Math.Min(fields.Count, dt.Columns.Count); j++)
                    row[j] = fields[j].Length == 0 ? DBNull.Value : (object)fields[j];
                dt.Rows.Add(row);
            }

            return dt;
        }

        private List<(string ColumnName, string DataType)> GetTableColumns(string tableName)
        {
            var result = new List<(string, string)>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(
                "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS " +
                "WHERE TABLE_NAME = @T ORDER BY ORDINAL_POSITION", conn);
            cmd.Parameters.AddWithValue("@T", tableName);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add((reader.GetString(0), reader.GetString(1)));
            return result;
        }

        public List<string> GetTableNames()
        {
            var list = new List<string>();
            using var conn = _dbService.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(
                "SELECT name FROM sys.tables WHERE type='U' ORDER BY name", conn);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                list.Add(reader.GetString(0));
            return list;
        }

        // ─── CSV ヘルパー ─────────────────────────────────────────────────────

        /// <summary>
        /// ロケールに依存せず、SQL Server の型を安全な文字列に変換する。
        /// DateTime 等は OS の日付書式ではなく固定 ISO 形式で出力する。
        /// </summary>
        private static string FormatFieldValue(SqlDataReader reader, int i)
        {
            if (reader.IsDBNull(i)) return "";

            var type = reader.GetFieldType(i);

            if (type == typeof(DateTime))
                return reader.GetDateTime(i)
                    .ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture);

            if (type == typeof(DateTimeOffset))
                return ((DateTimeOffset)reader.GetValue(i))
                    .ToString("yyyy-MM-dd HH:mm:ss.fffffffzzz", CultureInfo.InvariantCulture);

            if (type == typeof(TimeSpan))
                return ((TimeSpan)reader.GetValue(i))
                    .ToString(@"hh\:mm\:ss\.fffffff", CultureInfo.InvariantCulture);

            if (type == typeof(float))
                return reader.GetFloat(i).ToString(CultureInfo.InvariantCulture);

            if (type == typeof(double))
                return reader.GetDouble(i).ToString(CultureInfo.InvariantCulture);

            if (type == typeof(decimal))
                return reader.GetDecimal(i).ToString(CultureInfo.InvariantCulture);

            // bool, int, long, byte, string 等はそのまま
            return reader.GetValue(i).ToString()!;
        }

        private static string EscapeCsv(string value, char delimiter)
        {
            if (value.Contains(delimiter) || value.Contains('"') || value.Contains('\n'))
                return $"\"{value.Replace("\"", "\"\"")}\"";
            return value;
        }

        private static List<string> ParseCsvLine(string line, char delimiter)
        {
            var fields = new List<string>();
            var sb = new StringBuilder();
            bool inQuote = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inQuote)
                {
                    if (c == '"' && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else if (c == '"')
                        inQuote = false;
                    else
                        sb.Append(c);
                }
                else
                {
                    if (c == '"')
                        inQuote = true;
                    else if (c == delimiter)
                    {
                        fields.Add(sb.ToString());
                        sb.Clear();
                    }
                    else
                        sb.Append(c);
                }
            }
            fields.Add(sb.ToString());
            return fields;
        }
    }
}
