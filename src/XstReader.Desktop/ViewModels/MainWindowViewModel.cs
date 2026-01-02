using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
    private ObservableCollection<MessageEntry> _messages = new();
    private FolderEntry? _selectedFolder;
    private MessageEntry? _selectedMessage;
    private string _messagePreview = "";
    private XstFile? _currentFile;
    private CancellationTokenSource? _messageLoadCts;

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

    public ObservableCollection<MessageEntry> Messages
    {
        get => _messages;
        private set => this.RaiseAndSetIfChanged(ref _messages, value);
    }

    public FolderEntry? SelectedFolder
    {
        get => _selectedFolder;
        set
        {
            if (!ReferenceEquals(_selectedFolder, value))
            {
                this.RaiseAndSetIfChanged(ref _selectedFolder, value);
                _ = LoadMessagesForSelectionAsync(value);
            }
        }
    }

    public MessageEntry? SelectedMessage
    {
        get => _selectedMessage;
        set
        {
            if (!ReferenceEquals(_selectedMessage, value))
            {
                this.RaiseAndSetIfChanged(ref _selectedMessage, value);
                MessagePreview = value is null ? string.Empty : BuildFullBody(value.Message);
            }
        }
    }

    public string MessagePreview
    {
        get => _messagePreview;
        private set => this.RaiseAndSetIfChanged(ref _messagePreview, value);
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
            // Dispose previous file before opening new one
            var newFile = await Task.Run(() => new XstFile(filePath, password)).ConfigureAwait(true);
            _currentFile?.Dispose();
            _currentFile = newFile;
            var entries = await Task.Run(() => LoadFolders(_currentFile)).ConfigureAwait(true);
            FolderEntries = new ObservableCollection<FolderEntry>(entries);
            Messages = new ObservableCollection<MessageEntry>();
            SelectedFolder = null;
            SelectedMessage = null;
            MessagePreview = string.Empty;

            var totalFolders = entries.Count;
            StatusMessage = totalFolders == 0
                ? "Opened PST successfully." 
                : $"Opened PST successfully: {totalFolders} folders loaded.";
        }
        catch (XstException ex)
        {
            FolderEntries.Clear();
            Messages.Clear();
            MessagePreview = string.Empty;
            StatusMessage = ex.Message;
        }
        catch (Exception ex)
        {
            FolderEntries.Clear();
            Messages.Clear();
            MessagePreview = string.Empty;
            StatusMessage = $"Unexpected error: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private List<FolderEntry> LoadFolders(XstFile xstFile)
    {
        var root = xstFile.RootFolder;
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
                level,
                folder));

            foreach (var child in folder.Folders)
                Append(child, level + 1);
        }
    }

    private async Task LoadMessagesForSelectionAsync(FolderEntry? folderEntry)
    {
        _messageLoadCts?.Cancel();
        _messageLoadCts?.Dispose();

        if (folderEntry == null)
        {
            Messages = new ObservableCollection<MessageEntry>();
            SelectedMessage = null;
            MessagePreview = string.Empty;
            return;
        }

        if (_currentFile == null)
            return;

        var cts = new CancellationTokenSource();
        _messageLoadCts = cts;

        try
        {
            StatusMessage = $"Loading messages from \"{folderEntry.Name}\"...";

            var entries = await Task.Run(() =>
            {
                var list = new List<MessageEntry>();
                foreach (var message in folderEntry.Folder.Messages)
                {
                    cts.Token.ThrowIfCancellationRequested();
                    var sender = message.From ?? message.To ?? string.Empty;
                    var preview = BuildPreview(message);
                    list.Add(new MessageEntry(
                        message.Subject ?? string.Empty,
                        message.ReceivedTime ?? message.SubmittedTime,
                        sender,
                        preview,
                        message));
                }
                return list;
            }, cts.Token).ConfigureAwait(true);

            Messages = new ObservableCollection<MessageEntry>(entries);
            SelectedMessage = entries.FirstOrDefault();

            StatusMessage = entries.Count == 0
                ? $"Folder \"{folderEntry.Name}\" is empty."
                : $"Loaded {entries.Count} messages from \"{folderEntry.Name}\".";
        }
        catch (OperationCanceledException)
        {
            // Selection changed again; ignore.
        }
        catch (Exception ex)
        {
            Messages = new ObservableCollection<MessageEntry>();
            SelectedMessage = null;
            MessagePreview = string.Empty;
            StatusMessage = $"Error loading messages: {ex.Message}";
        }
        finally
        {
            if (ReferenceEquals(_messageLoadCts, cts))
            {
                _messageLoadCts = null;
            }
            cts.Dispose();
        }
    }

    private static string BuildPreview(XstMessage message)
    {
        try
        {
            var text = message.Body?.Text ?? string.Empty;
            return Condense(text, 200);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string BuildFullBody(XstMessage message)
    {
        try
        {
            return message.Body?.Text ?? string.Empty;
        }
        catch (Exception ex)
        {
            return $"Unable to load body: {ex.Message}";
        }
    }

    private static string Condense(string text, int maxLength)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        var sb = new StringBuilder(text.Length);
        bool lastWasSpace = false;

        foreach (var ch in text)
        {
            char c = ch;
            if (c == '\r' || c == '\n' || c == '\t')
                c = ' ';

            if (char.IsControl(c))
                continue;

            if (char.IsWhiteSpace(c))
            {
                if (lastWasSpace)
                    continue;
                lastWasSpace = true;
                sb.Append(' ');
            }
            else
            {
                lastWasSpace = false;
                sb.Append(c);
            }

            if (sb.Length >= maxLength)
            {
                sb.Length = maxLength;
                sb.Append('…');
                break;
            }
        }

        return sb.ToString().Trim();
    }
}
