namespace OperationCapture
{
    #region using
    using System.Windows;
    using System.Windows.Navigation;
    #endregion

    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : NavigationWindow
    {
        #region MainWindow
        public MainWindow()
        {
            InitializeComponent();
            DefaultMode nextPage = new DefaultMode();
            this.Navigate(nextPage);
        }

        #endregion

        private void NavigationWindow_Loaded(object sender, RoutedEventArgs e)
        {
            var window = Window.GetWindow(this);
            window.Left = 0;
            window.Top = 0;
        }
    }
}
