using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AIContentTool
{
    public static class KPITracker
    {
        private static readonly string KpiFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PowerPointAddIn", "kpi.csv");

        public static void RecordImport(string format, int placeholderCount)
        {
            AppendKpi($"Import,{format},{DateTime.Now},{placeholderCount},0");
        }

        public static void RecordGeneration(int slideCount, double timeTaken)
        {
            AppendKpi($"Generation,,{DateTime.Now},{slideCount},{timeTaken}");
        }

        private static void AppendKpi(string data)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(KpiFilePath));
                if (!File.Exists(KpiFilePath))
                {
                    File.WriteAllText(KpiFilePath, "Event,Format,Timestamp,Count,TimeTaken\n");
                }
                File.AppendAllText(KpiFilePath, $"{data}\n");
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError($"Failed to record KPI: {ex.Message}");
            }
        }

        public static void ViewKpiReport()
        {
            try
            {
                if (File.Exists(KpiFilePath))
                {
                    Process.Start(new ProcessStartInfo { FileName = KpiFilePath, UseShellExecute = true });
                }
                else
                {
                    MessageBox.Show("No KPI data available yet.");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError($"Failed to view KPI report: {ex.Message}");
                MessageBox.Show("Error opening KPI report.");
            }
        }
    }
}