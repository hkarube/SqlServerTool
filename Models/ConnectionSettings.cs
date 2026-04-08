using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlServerTool.Models
{
    public class ConnectionSettings
    {
        public string ServerName    { get; set; } = string.Empty;
        public string DatabaseName  { get; set; } = string.Empty;
        public string UserId        { get; set; } = "sa";
        public string Password      { get; set; } = string.Empty;
        public bool   UseWindowsAuth { get; set; } = false;

        public string BuildConnectionString()
        {
            if (UseWindowsAuth)
                return $"Server={ServerName};Database={DatabaseName};" +
                       $"Integrated Security=True;TrustServerCertificate=True;";
            return $"Server={ServerName};Database={DatabaseName};" +
                   $"User Id={UserId};Password={Password};" +
                   $"TrustServerCertificate=True;";
        }
    }
}
