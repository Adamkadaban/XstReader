using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ReactiveUI;
using XstReader;

namespace XstReader.Desktop.ViewModels;

public class MainWindowViewModel : ReactiveObject
{
    private string? _pstPath;
    private string? _password;
    private bool _isBusy;
    private string _statusMessage = "Select a PST file to begin.";
    private ObservableCollection<FolderEntry> _folderEntries = new();

    public string? PstPath
    {
        get => _pstPath;
        set => this.RaiseAndSetIfChanged(ref _pstPath, value);
    }

    public string? Password
    {
        get => _password;
        set => this.RaiseAndSetIfChanged(ref _password, value);
    }

    public ObservableCollection<FolderEntry> FolderEntries
    {
        get => _folderEntries;
        private set => this.RaiseAndSetIfChanged(ref _folderEntries, value);
    }

    public bool IsBusy
    {
        get => _isBusy;
        private set => this.RaiseAndSetIfChanged(ref _isBusy, value);
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set => this.RaiseAndSetIfChanged(ref _statusMessage, value);
    }

    public async Task LoadAsync()
    {
        if (IsBusy)
            return;

        var filePath = PstPath?.Trim();
        if (string.IsNullOrWhiteSpace(filePath))
        {
            StatusMessage = "Please select a PST or OST file.";
            return;
        }

        if (!File.Exists(filePath))
        {
            StatusMessage = "The selected file could not be found.";
            return;
        }

        IsBusy = true;
        StatusMessage = "Opening PST...";

        try
        {
            var password = string.IsNullOrWhiteSpace(Password) ? null : Password;
            var entries = await Task.Run(() => LoadFolders(filePath, password)).ConfigureAwait(true);
            FolderEntries = new ObservableCollection<FolderEntry>(entries);

            var totalFolders = entries.Count;
            StatusMessage = totalFolders == 0
                ? "Opened PST successfully." 
                : $"Opened PST successfully: {totalFolders} folders loaded.";
        }
        catch (XstException ex)
        {
            FolderEntries.Clear();
            StatusMessage = ex.Message;
        }
        catch (Exception ex)
        {
            FolderEntries.Clear();
            StatusMessage = $"Unexpected error: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private static List<FolderEntry> LoadFolders(string filePath, string? password)
    {
        using var file = new XstFile(filePath, password);
        var root = file.RootFolder;
        var entries = new List<FolderEntry>();

        foreach (var folder in root.Folders)
            Append(folder, 0);

        return entries;

        void Append(XstFolder folder, int level)
        {
            entries.Add(new FolderEntry(
                folder.DisplayName ?? "(Folder)",
                folder.Path,
                folder.ContentCount,
                folder.ContentUnreadCount,
                level));

            foreach (var child in folder.Folders)
                Append(child, level + 1);
        }
    }
}
