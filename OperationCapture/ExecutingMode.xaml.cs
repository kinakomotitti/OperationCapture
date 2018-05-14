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
    /// ExecutingMode.xaml の相互作用ロジック
    /// </summary>
    public partial class ExecutingMode : Page
    {
        public ExecutingMode()
        {
            InitializeComponent();
        }

        private void CaptureStartButton_Click(object sender, RoutedEventArgs e)
        {
            GrobalHookManager.FinalizeHooks();
            this.NavigationService.GoBack();
        }
    }
}
