namespace OperationCapture.Core
{
    #region using
    using System;
    using System.Configuration;
    #endregion

    public static class SettingsManager
    {
        private static AppSettingsReader render = new AppSettingsReader();
        public static string LocalFolderPath { get; set; } = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        public static string LocalFileName { get; set; } = "output.xlsx";
        public static long LocalExcelCellHeight { get; set; } = 20; //default cell height
        public static bool UseActiveWindowOnly { get; set; } = false;

    }
}
