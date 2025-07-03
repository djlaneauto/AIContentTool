using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIContentTool
{
    public static class ErrorHandler
    {
        private static readonly string LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PowerPointAddIn", "error.log");

        public static void LogError(string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));
                File.AppendAllText(LogFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                // Silent fail to avoid recursive errors
                System.Diagnostics.Debug.WriteLine($"Failed to log error: {ex.Message}");
            }
        }
    }
}