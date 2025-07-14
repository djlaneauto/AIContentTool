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

        public static void LogError(string message, Exception ex = null)
        {
            string fullMessage = message;
            if (ex != null)
            {
                fullMessage += $"\nException Details: {ex.Message}\nStack Trace: {ex.StackTrace}";
            }
            Log("ERROR", fullMessage);
        }

        public static void LogInfo(string message)
        {
            Log("INFO", message);
        }

        private static void Log(string level, string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));
                File.AppendAllText(LogFilePath, $"{DateTime.Now} [{level}]: {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                // Silent fail to avoid recursive errors
                System.Diagnostics.Debug.WriteLine($"Failed to log {level}: {ex.Message}");
            }
        }
    }
}