namespace OperationCapture
{
    #region using
    using OperationCapture.Core;
    using System.Windows;
    #endregion

    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
       public App()
        {
            LogManager.Logger.Info("App Start");
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LogManager.Logger.Fatal(e.Exception.GetBaseException().ToString());
            System.Environment.Exit(9);
        }
    }
}
