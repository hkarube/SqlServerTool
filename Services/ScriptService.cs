using Microsoft.Data.SqlClient;
using SqlServerTool.Models;
using System.Data;
using System.IO;
using System.Text;

namespace SqlServerTool.Services
{
    public class ScriptService
    {
        private readonly DatabaseService _dbService;
        private readonly SchemaService   _schemaService;

        public ScriptService(DatabaseService dbService, SchemaService schemaService)
        {
            _dbService     = dbService;
            _schemaService = schemaService;
        }

        /// <summary>
        /// 指定テーブル一覧のスクリプトをファイルに出力する。
        /// テーブルごとに 1 ファイル（tableName.sql）を生成する。
        /// </summary>
        public void GenerateTableScripts(
            IEnumerable<string> tableNames,
            ScriptOptions options,
            IProgress<string>? progress = null)
        {
            Directory.CreateDirectory(options.OutputFolder);

            foreach (var tableName in tableNames)
            {
                progress?.Report($"処理中: {tableName}");
                var filePath = Path.Combine(options.OutputFolder, $"{tableName}.sql");

                if (!options.OverwriteExisting && File.Exists(filePath))
                {
                    progress?.Report($"スキップ（既存）: {tableName}");
                    continue;
                }

                var sb = new StringBuilder();
                sb.AppendLine($"-- ============================================================");
                sb.AppendLine($"-- テーブル: {tableName}");
                sb.AppendLine($"-- 生成日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                sb.AppendLine($"-- ============================================================");
                sb.AppendLine();

                if (options.IncludeDrop)
                {
                    sb.AppendLine($"-- DROP");
                    sb.AppendLine($"IF OBJECT_ID(N'[dbo].[{tableName}]', N'U') IS NOT NULL");
                    sb.AppendLine($"    DROP TABLE [dbo].[{tableName}];");
                    sb.AppendLine("GO");
                    sb.AppendLine();
                }

                sb.Append(BuildCreateTable(tableName, options));
                sb.AppendLine();

                if (options.IncludeData)
                {
                    sb.Append(BuildInsertStatements(tableName));
                    sb.AppendLine();
                }

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                progress?.Report($"完了: {tableName}");
            }
        }

        /// <summary>ビュースクリプトをファイルに出力する</summary>
        public void GenerateViewScripts(
            IEnumerable<string> viewNames,
            ScriptOptions options,
            IProgress<string>? progress = null)
        {
            Directory.CreateDirectory(options.OutputFolder);

            foreach (var viewName in viewNames)
            {
                progress?.Report($"処理中: {viewName}");
                var filePath = Path.Combine(options.OutputFolder, $"{viewName}.sql");

                if (!options.OverwriteExisting && File.Exists(filePath))
                {
                    progress?.Report($"スキップ（既存）: {viewName}");
                    continue;
                }

                var sb = new StringBuilder();
                sb.AppendLine($"-- ============================================================");
                sb.AppendLine($"-- ビュー: {viewName}");
                sb.AppendLine($"-- 生成日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                sb.AppendLine($"-- ============================================================");
                sb.AppendLine();

                if (options.IncludeDrop)
                {
                    sb.AppendLine($"IF OBJECT_ID(N'[dbo].[{viewName}]', N'V') IS NOT NULL");
                    sb.AppendLine($"    DROP VIEW [dbo].[{viewName}];");
                    sb.AppendLine("GO");
                    sb.AppendLine();
                }

                sb.Append(BuildCreateView(viewName, options));

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                progress?.Report($"完了: {viewName}");
            }
        }

        // ─── CREATE TABLE ──────────────────────────────────────────────────

        private string BuildCreateTable(string tableName, ScriptOptions options)
        {
            var sb = new StringBuilder();
            var cols = _schemaService.GetScriptColumns(tableName);
            var (pkName, pkCols) = _schemaService.GetPrimaryKey(tableName);
            var filegroup = options.IncludeFilegroup ? _schemaService.GetFilegroup(tableName) : "PRIMARY";

            sb.AppendLine($"CREATE TABLE [dbo].[{tableName}] (");

            for (int i = 0; i < cols.Count; i++)
            {
                var c = cols[i];
                var line = new StringBuilder();
                line.Append($"    [{c.Name}] {c.TypeName}{c.TypeSuffix}");

                if (c.IsIdentity)
                    line.Append($" IDENTITY({c.SeedValue},{c.IncrementValue})");

                if (!string.IsNullOrEmpty(c.DefaultDef))
                    line.Append($" DEFAULT {c.DefaultDef}");

                line.Append(c.IsNullable ? " NULL" : " NOT NULL");

                bool isLast = (i == cols.Count - 1) && pkCols.Count == 0;
                sb.AppendLine(isLast ? line.ToString() : line + ",");
            }

            if (pkCols.Count > 0)
            {
                var pkColList = string.Join(", ", pkCols.Select(c => $"[{c}]"));
                sb.AppendLine($"    CONSTRAINT [{pkName}] PRIMARY KEY CLUSTERED ({pkColList})");
            }

            if (options.IncludeFilegroup)
                sb.AppendLine($") ON [{filegroup}];");
            else
                sb.AppendLine(");");

            sb.AppendLine("GO");
            return sb.ToString();
        }

        // ─── CREATE VIEW ───────────────────────────────────────────────────

        /// <summary>SP / ファンクションの SQL 定義をファイルに出力する</summary>
        public void GenerateCodeObjectScripts(
            IEnumerable<string> objectNames,
            string objectTypeName,
            ScriptOptions options,
            IProgress<string>? progress = null)
        {
            Directory.CreateDirectory(options.OutputFolder);

            foreach (var name in objectNames)
            {
                progress?.Report($"処理中: {name}");
                var filePath = Path.Combine(options.OutputFolder, $"{name}.sql");

                if (!options.OverwriteExisting && File.Exists(filePath))
                {
                    progress?.Report($"スキップ（既存）: {name}");
                    continue;
                }

                var sb = new StringBuilder();
                sb.AppendLine($"-- ============================================================");
                sb.AppendLine($"-- {objectTypeName}: {name}");
                sb.AppendLine($"-- 生成日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                sb.AppendLine($"-- ============================================================");
                sb.AppendLine();

                if (options.IncludeDrop)
                {
                    var typeCode = objectTypeName == "PROCEDURE" ? "P" : "FN','IF','TF','FS','FT";
                    sb.AppendLine($"IF OBJECT_ID(N'[dbo].[{name}]') IS NOT NULL");
                    sb.AppendLine($"    DROP {objectTypeName} [dbo].[{name}];");
                    sb.AppendLine("GO");
                    sb.AppendLine();
                }

                var definition = _schemaService.GetSqlDefinition(name);
                sb.AppendLine(definition);
                sb.AppendLine("GO");

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                progress?.Report($"完了: {name}");
            }
        }

        private string BuildCreateView(string viewName, ScriptOptions options)
        {
            var sb = new StringBuilder();
            var definition = _schemaService.GetSqlDefinition(viewName);

            if (options.IncludeViewColumns)
            {
                var viewCols = _schemaService.GetViewColumns(viewName);
                if (viewCols.Count > 0)
                {
                    // CREATE VIEW に列名を付加（既存定義にある CREATE VIEW 行を置換）
                    var colList = string.Join(", ", viewCols.Select(c => $"[{c}]"));
                    definition = System.Text.RegularExpressions.Regex.Replace(
                        definition,
                        @"CREATE\s+VIEW\s+(\[?\w+\]?\.?\[?\w+\]?)",
                        $"CREATE VIEW [dbo].[{viewName}] ({colList})",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                }
            }

            sb.AppendLine(definition);
            sb.AppendLine("GO");
            return sb.ToString();
        }

        // ─── INSERT statements ─────────────────────────────────────────────

        private string BuildInsertStatements(string tableName)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"-- INSERT DATA: {tableName}");

            using var conn = _dbService.GetConnection();
            conn.Open();

            // IDENTITY INSERT が必要か確認
            bool hasIdentity = false;
            {
                var identChk = new SqlCommand(
                    "SELECT COUNT(*) FROM sys.columns WHERE object_id = OBJECT_ID(@T) AND is_identity = 1", conn);
                identChk.Parameters.AddWithValue("@T", tableName);
                hasIdentity = (int)identChk.ExecuteScalar()! > 0;
            }

            var dt = new DataTable();
            var adapter = new SqlDataAdapter($"SELECT * FROM [{tableName}]", conn);
            adapter.Fill(dt);

            if (dt.Rows.Count == 0)
            {
                sb.AppendLine($"-- データなし");
                return sb.ToString();
            }

            var colNames = dt.Columns.Cast<DataColumn>()
                             .Select(c => $"[{c.ColumnName}]")
                             .ToList();
            var colList = string.Join(", ", colNames);

            if (hasIdentity)
                sb.AppendLine($"SET IDENTITY_INSERT [dbo].[{tableName}] ON;");

            foreach (DataRow row in dt.Rows)
            {
                var values = dt.Columns.Cast<DataColumn>().Select(col =>
                {
                    if (row[col] == DBNull.Value) return "NULL";
                    var v = row[col].ToString()!.Replace("'", "''");
                    return col.DataType == typeof(bool)
                        ? ((bool)row[col] ? "1" : "0")
                        : $"'{v}'";
                });
                sb.AppendLine($"INSERT INTO [dbo].[{tableName}] ({colList}) VALUES ({string.Join(", ", values)});");
            }

            if (hasIdentity)
                sb.AppendLine($"SET IDENTITY_INSERT [dbo].[{tableName}] OFF;");

            sb.AppendLine("GO");
            return sb.ToString();
        }
    }
}
