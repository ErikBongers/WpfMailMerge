using System.Windows;

namespace WpfMailMerge;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{

    public MainWindow()
    {
        InitializeComponent();
        }

    private void btnStart_Click(object sender, RoutedEventArgs e)
        {
        MailMergeViewModel? mailMergeViewModel = this.DataContext as MailMergeViewModel;
        mailMergeViewModel?.Start();
        }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
        MailMergeViewModel? mailMerge = this.DataContext as MailMergeViewModel;
        mailMerge?.CloseAll();
        mailMerge?.SaveJsonSettings();
        }
    }