using System.Windows;

namespace Mro.MainApp;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        DispatcherUnhandledException += (_, args) =>
        {
            Mro.Recovered.UI.StartupErrors.Record(args.Exception);
            MessageBox.Show("发生错误：" + args.Exception.Message, "MRO", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };
        base.OnStartup(e);
    }
}
