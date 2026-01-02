using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace XstReader.Desktop.ViewModels;

public sealed class FolderNode
{
    public FolderNode(string name, string path, int itemCount, int unreadCount, params FolderNode[] children)
        : this(name, path, itemCount, unreadCount, (IEnumerable<FolderNode>)children)
    {
    }

    public FolderNode(string name, string path, int itemCount, int unreadCount, IEnumerable<FolderNode> children)
    {
        Name = string.IsNullOrWhiteSpace(name) ? "(Untitled)" : name.Trim();
        Path = path;
        ItemCount = itemCount;
        UnreadCount = unreadCount;
        Children = new ObservableCollection<FolderNode>(children);
    }

    public string Name { get; }
    public string Path { get; }
    public int ItemCount { get; }
    public int UnreadCount { get; }
    public ObservableCollection<FolderNode> Children { get; }

    public string Title
    {
        get
        {
            var parts = new List<string> { ItemCount.ToString() };
            if (UnreadCount > 0)
                parts.Add($"{UnreadCount} unread");
            return $"{Name} ({string.Join(", ", parts)})";
        }
    }

    public int TotalDescendants => 1 + Children.Sum(child => child.TotalDescendants);
}
