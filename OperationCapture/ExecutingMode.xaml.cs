namespace OperationCapture
{
    #region using
    using OperationCapture.Core;
    using System.Windows;
    using System.Windows.Controls;
    #endregion

    /// <summary>
    /// ExecutingMode.xaml の相互作用ロジック
    /// </summary>
    public partial class ExecutingMode : Page
    {
        private double defaultWindowWidth = 500;
        private double defaultWindowHeight=110;
        public ExecutingMode()
        {
            InitializeComponent();
            this.WindowWidth = 50;
            this.WindowHeight = 50;
            this.Opacity = 0.5;
        }

        private void CaptureStartButton_Click(object sender, RoutedEventArgs e)
        {
            GlobalHookManager.FinalizeHooks();
            this.WindowWidth=defaultWindowWidth;
            this.WindowHeight=defaultWindowHeight;
            this.NavigationService.GoBack();
        }
    }
}
