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
        private double defaultWindowWidth = 400;
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
