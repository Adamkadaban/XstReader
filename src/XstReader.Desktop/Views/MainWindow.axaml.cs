using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Platform.Storage;
using XstReader.Desktop.ViewModels;

namespace XstReader.Desktop.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
#if DEBUG
        this.AttachDevTools();
#endif
    }

    private void InitializeComponent()
        => AvaloniaXamlLoader.Load(this);

    private async void OnBrowseClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel viewModel)
            return;

        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel?.StorageProvider == null)
            return;

        var options = new FilePickerOpenOptions
        {
            AllowMultiple = false,
            Title = "Open Outlook Data File",
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Outlook Files") { Patterns = new[] { "*.pst", "*.ost" } },
                FilePickerFileTypes.All
            }
        };

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(options).ConfigureAwait(true);
        var file = files?.FirstOrDefault();
        if (file != null)
        {
            viewModel.PstPath = file.Path.LocalPath;
        }
    }

    private async void OnOpenClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            await viewModel.LoadAsync().ConfigureAwait(true);
        }
    }
}
