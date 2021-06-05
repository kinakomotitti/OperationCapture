namespace OperationCapture
{
    #region using
    using OperationCapture.Core;
    using System.Windows;
    using System.Windows.Controls;
    #endregion

    /// <summary>
    /// DefaultMode.xaml の相互作用ロジック
    /// </summary>
    public partial class DefaultMode : Page
    {
        public DefaultMode()
        {
            InitializeComponent();
        }

        #region Event

        private void CaptureStartButton_Click(object sender, RoutedEventArgs e)
        {
            GlobalHookManager.InitializeHooks();
            ExecutingMode nextPage = new ExecutingMode();
            this.SaveToExcelButton.IsEnabled  = false;
            this.NavigationService.Navigate(nextPage);
        }

        private void SaveToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            this.SaveToExcelButton.IsEnabled = false;
            EventCore core = new EventCore();
            core.SaveEvidence();
        }

        private void ExitApplicationButton_Click(object sender, RoutedEventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            this.WindowWidth = 500;
            this.WindowHeight = 110;
            this.SaveToExcelButton.IsEnabled= EventCore.Operations.Count > 0;
        }

        private void SettingButton_Click(object sender, RoutedEventArgs e)
        {
            var nextPage = new SettingMode();
            this.NavigationService.Navigate(nextPage);
        }

        #endregion
    }
}
