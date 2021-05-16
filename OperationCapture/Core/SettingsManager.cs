namespace OperationCapture.Core
{
    #region using
    using System;
    using System.Configuration;
    #endregion

    public static class SettingsManager
    {
        #region 変数

        private static AppSettingsReader render = new AppSettingsReader();
        public static string LocalFolderPath { get; set; }
        public static string LocalFileName { get; set; }
        public static long LocalExcelCellHeight { get; set; } = 20; //default cell height

        #endregion

        #region public

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

        #endregion
    }
}
