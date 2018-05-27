namespace OperationCapture
{
    #region using
    using OperationCapture.Core;
    using System.IO;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Forms;
    #endregion

    /// <summary>
    /// Settings.xaml の相互作用ロジック
    /// </summary>
    public partial class SettingMode : Page
    {
        #region SettingMode

        public SettingMode()
        {
            InitializeComponent();
        }

        #endregion

        #region Event

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            this.WindowWidth = 900;
            this.WindowHeight = 200;
        }

        private void FolderPicker_Click(object sender, RoutedEventArgs e)
        {
            // フォルダー参照ダイアログのインスタンスを生成
            var dlg = new FolderBrowserDialog();

            // 説明文を設定
            dlg.Description = "フォルダーを選択してください。";

            // ダイアログを表示
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // 選択されたフォルダーパスをメッセージボックスに表示
                this.FolderPicker_TextBox.Text = dlg.SelectedPath;
            }
        }

        private void Back_Buttton_Click(object sender, RoutedEventArgs e)
        {
            this.NavigationService.GoBack();
        }

        private void Page_Unloaded(object sender, RoutedEventArgs e)
        {
            SettingsManager.LocalFolderPath = Directory.Exists(this.FolderPicker_TextBox.Text) ? 
                                                this.FolderPicker_TextBox.Text : 
                                                SettingsManager.LocalFolderPath;

            SettingsManager.LocalFileName = string.IsNullOrWhiteSpace(this.FileName_TextBox.Text)==false ?
                                                this.FileName_TextBox.Text :
                                                SettingsManager.LocalFileName;
            int heiht = 0;
            if (int.TryParse(this.ExcelCellHeight_TextBox.Text,out heiht))
            {
                SettingsManager.LocalExcelCellHeight = heiht;
            }
        }

        #endregion
    }
}
