using Microsoft.Data.SqlClient;
using System.IO;

namespace SqlServerTool.Services
{
    public class BackupService
    {
        private readonly DatabaseService _dbService;

        public BackupService(DatabaseService dbService)
        {
            _dbService = dbService;
        }

        /// <summary>ファイルパスを直接指定してバックアップする（ダイアログ用）</summary>
        public void BackupToFile(string databaseName, string filePath)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
            RunBackupSql(databaseName, filePath);
        }

        /// <summary>完全バックアップを実行し、生成したファイルパスを返す（従来互換）</summary>
        public string Backup(string databaseName, string outputFolder)
        {
            Directory.CreateDirectory(outputFolder);
            var fileName = $"{databaseName}_{DateTime.Now:yyyyMMdd_HHmmss}.bak";
            var filePath = Path.Combine(outputFolder, fileName);

            RunBackupSql(databaseName, filePath);
            return filePath;
        }

        private void RunBackupSql(string databaseName, string filePath)
        {
            using var conn = _dbService.GetMasterConnection();
            conn.Open();

            var sql = $@"
                BACKUP DATABASE [{databaseName}]
                TO DISK = @FilePath
                WITH FORMAT, COMPRESSION,
                     MEDIANAME = 'SqlServerToolBackup',
                     NAME = 'Full Backup of {databaseName}'";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@FilePath", filePath);
            cmd.CommandTimeout = 600;
            cmd.ExecuteNonQuery();
        }

        /// <summary>バックアップファイルからリストアする</summary>
        public void Restore(string databaseName, string bakFilePath)
        {
            using var conn = _dbService.GetMasterConnection();
            conn.Open();

            // 既存接続を切断してシングルユーザーモードへ
            var killSql = $"ALTER DATABASE [{databaseName}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE";
            using (var cmd = new SqlCommand(killSql, conn))
                cmd.ExecuteNonQuery();

            // リストア実行
            var restoreSql = $@"
                RESTORE DATABASE [{databaseName}]
                FROM DISK = @FilePath
                WITH REPLACE, RECOVERY";
            using (var cmd = new SqlCommand(restoreSql, conn))
            {
                cmd.Parameters.AddWithValue("@FilePath", bakFilePath);
                cmd.CommandTimeout = 600;
                cmd.ExecuteNonQuery();
            }

            // マルチユーザーモードへ戻す
            var multiSql = $"ALTER DATABASE [{databaseName}] SET MULTI_USER";
            using (var cmd = new SqlCommand(multiSql, conn))
                cmd.ExecuteNonQuery();
        }
    }
}
