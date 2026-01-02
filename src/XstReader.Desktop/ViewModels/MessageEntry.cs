using System;
using XstReader;

namespace XstReader.Desktop.ViewModels;

public sealed class MessageEntry
{
    public MessageEntry(string subject, DateTime? received, string sender, string preview, XstMessage message)
    {
        Subject = string.IsNullOrWhiteSpace(subject) ? "(No Subject)" : subject.Trim();
        Received = received;
        Sender = sender;
        Preview = preview;
        Message = message;
    }

    public string Subject { get; }
    public DateTime? Received { get; }
    public string Sender { get; }
    public string Preview { get; }
    public XstMessage Message { get; }
}
