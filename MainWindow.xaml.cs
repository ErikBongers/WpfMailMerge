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

namespace WpfMailMerge;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    MailMerge mailMerge;

    public MainWindow()
    {
        InitializeComponent();
        this.mailMerge = new MailMerge(MailMergeSendTo.TestEmail);
        this.DataContext = this.mailMerge;
        }

    private void btnStart_Click(object sender, RoutedEventArgs e)
        {
        this.mailMerge.Start();
        }
    }