using System;
using XstReader;

namespace XstReader.Desktop.ViewModels;

public sealed class FolderEntry
{
    public FolderEntry(string name, string path, int messageCount, int unreadCount, int level, XstFolder folder)
    {
        Name = string.IsNullOrWhiteSpace(name) ? "(Untitled)" : name.Trim();
        Path = path;
        MessageCount = messageCount;
        UnreadCount = unreadCount;
        Level = level;
        Folder = folder;
    }

    public string Name { get; }
    public string Path { get; }
    public int MessageCount { get; }
    public int UnreadCount { get; }
    public int Level { get; }
    public XstFolder Folder { get; }

    public string Display => UnreadCount > 0
        ? $"{Name} ({MessageCount}, {UnreadCount} unread)"
        : $"{Name} ({MessageCount})";

    public double IndentWidth => Math.Max(0, Level) * 18.0;
}
