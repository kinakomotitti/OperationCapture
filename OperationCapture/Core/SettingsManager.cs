using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OperationCapture.Core
{
    public static class SettingsManager
    {
        private static AppSettingsReader render = new AppSettingsReader();

        public static string LocalFolderPath { get; set; }
        public static string LocalFileName { get; set; }
        public static long LocalExcelCellHeight { get; set; } = 24;

        public static String GetFolderPath()
        {
            return String.IsNullOrWhiteSpace(LocalFolderPath) ?
                            render.GetValue("outputFolderPath", typeof(string)).ToString() :
                            LocalFolderPath;
        }

        public static String GetFileName()
        {
            return String.IsNullOrWhiteSpace(LocalFileName) ?
                            render.GetValue("outputExcelFileName", typeof(string)).ToString() :
                            LocalFileName;
        }
    }
}
