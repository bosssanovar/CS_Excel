using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelBook book = new();

            book.SetAAAEntity("AAA Entity");
            book.SetBBBEntity("BBB Entity");

            var saveDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var fileName = "ExcelTest.xlsx";
            var savePath = System.IO.Path.Combine(saveDir, fileName);
            book.Save(savePath);

            MessageBox.Show($"{savePath}　を保存完了", "情報");
        }
    }
}