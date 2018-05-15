using OperationCapture.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OperationCapture
{
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
            GrobalHookManager.InitializeHooks();
            ExecutingMode nextPage = new ExecutingMode();
            this.SaveToExcelButton.IsEnabled  = false;
            this.NavigationService.Navigate(nextPage);
        }

        private void SaveToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            this.SaveToExcelButton.IsEnabled = false;
            EventCore core = new EventCore();
            core.SaveEvidenceToExcelFile();
        }

        private void PrintScreenButton_Click(object sender, RoutedEventArgs e)
        {
            EventCore core = new EventCore();
            core.TakeScreenShot();
        }

        private void ExitApplicationButton_Click(object sender, RoutedEventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            this.SaveToExcelButton.IsEnabled= EventCore.Operations.Count > 0;
        }

        #endregion
    }
}
