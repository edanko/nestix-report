namespace NestixReport
{
    public partial class MainWindow //: Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string GetConnectionString() => $"Data Source=BK-SSK-NESH01.CORP.LOCAL;Initial Catalog={DbComboBox.Text};Integrated Security=SSPI";

        private string GetFilter() => $"%{FilterTextBox.Text}%";
        
        
    }
}
