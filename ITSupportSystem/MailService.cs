using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Search;
using MailKit.Security;
using MimeKit;

namespace ITSupportSystem;

public sealed class MailService : ISupportMailService
{
    private readonly SupportSettings _settings;

    public MailService(SupportSettings settings)
    {
        _settings = settings;
    }

    public async Task<IReadOnlyList<IncomingMail>> FetchUnseenMailsAsync(CancellationToken cancellationToken)
    {
        using var client = new ImapClient();
        await client.ConnectAsync(_settings.Imap.Host, _settings.Imap.Port, ResolveSecureSocket(_settings.Imap.UseSsl), cancellationToken);
        await client.AuthenticateAsync(_settings.Imap.Username, _settings.Imap.Password, cancellationToken);

        var inbox = client.Inbox;
        await inbox.OpenAsync(MailKit.FolderAccess.ReadOnly, cancellationToken);

        var unseen = await inbox.SearchAsync(SearchQuery.NotSeen, cancellationToken);
        if (unseen.Count == 0)
        {
            await client.DisconnectAsync(true, cancellationToken);
            return Array.Empty<IncomingMail>();
        }

        var result = new List<IncomingMail>(unseen.Count);
        foreach (var uid in unseen)
        {
            var message = await inbox.GetMessageAsync(uid, cancellationToken);
            var sender = message.From.Mailboxes.FirstOrDefault();
            var messageId = string.IsNullOrWhiteSpace(message.MessageId)
                ? $"{uid.Id}@local"
                : message.MessageId;

            result.Add(new IncomingMail(
                uid.Id.ToString(),
                messageId,
                sender?.Address ?? string.Empty,
                sender?.Name ?? string.Empty,
                message.Subject ?? "(uden emne)",
                ExtractBody(message),
                message.Date,
                string.Empty));
        }

        await client.DisconnectAsync(true, cancellationToken);
        return result;
    }

    public Task MarkAsSeenAsync(IEnumerable<string> providerMessageIds, CancellationToken cancellationToken)
    {
        return Task.CompletedTask;
    }

    public Task MoveMessagesToTrashAsync(IEnumerable<string> providerMessageIds, CancellationToken cancellationToken)
    {
        return Task.CompletedTask;
    }

    public Task<SupportThreadAnchor?> SendSupportThreadStartAsync(string subject, string bodyText, string? bodyHtml, CancellationToken cancellationToken)
    {
        return Task.FromResult<SupportThreadAnchor?>(null);
    }

    public Task<SupportThreadAnchor?> ReplyInSupportThreadAsync(string conversationId, string providerMessageId, string bodyText, string? bodyHtml, CancellationToken cancellationToken)
    {
        return Task.FromResult<SupportThreadAnchor?>(null);
    }

    public Task MoveTicketMessagesToTrashAsync(string ticketId, IEnumerable<string> relatedProviderMessageIds, CancellationToken cancellationToken)
    {
        return Task.CompletedTask;
    }

    public async Task SendEmailAsync(
        string toAddress,
        string toName,
        string subject,
        string bodyText,
        string? bodyHtml,
        CancellationToken cancellationToken)
    {
        var message = new MimeMessage();
        message.From.Add(new MailboxAddress(_settings.MailboxDisplayName, _settings.MailboxAddress));
        message.To.Add(new MailboxAddress(toName, toAddress));
        message.Subject = subject;

        var builder = new BodyBuilder
        {
            TextBody = bodyText
        };

        if (!string.IsNullOrWhiteSpace(bodyHtml))
        {
            builder.HtmlBody = bodyHtml;
        }

        message.Body = builder.ToMessageBody();

        using var smtp = new SmtpClient();
        await smtp.ConnectAsync(_settings.Smtp.Host, _settings.Smtp.Port, ResolveSecureSocket(_settings.Smtp.UseSsl), cancellationToken);
        await smtp.AuthenticateAsync(_settings.Smtp.Username, _settings.Smtp.Password, cancellationToken);
        await smtp.SendAsync(message, cancellationToken);
        await smtp.DisconnectAsync(true, cancellationToken);
    }

    private static SecureSocketOptions ResolveSecureSocket(bool useSsl)
    {
        return useSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable;
    }

    private static string ExtractBody(MimeMessage message)
    {
        if (!string.IsNullOrWhiteSpace(message.TextBody))
        {
            return message.TextBody.Trim();
        }

        if (!string.IsNullOrWhiteSpace(message.HtmlBody))
        {
            return StripHtml(message.HtmlBody).Trim();
        }

        return "(ingen tekst i mailen)";
    }

    private static string StripHtml(string html)
    {
        var insideTag = false;
        var buffer = new System.Text.StringBuilder(html.Length);

        foreach (var ch in html)
        {
            if (ch == '<')
            {
                insideTag = true;
                continue;
            }

            if (ch == '>')
            {
                insideTag = false;
                continue;
            }

            if (!insideTag)
            {
                buffer.Append(ch);
            }
        }

        return System.Net.WebUtility.HtmlDecode(buffer.ToString());
    }
}
