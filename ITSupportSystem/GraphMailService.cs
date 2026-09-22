using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.Messages.Item.CreateReply;
using Microsoft.Graph.Users.Item.Messages.Item.Move;
using Microsoft.Graph.Users.Item.SendMail;

namespace ITSupportSystem;

public sealed class GraphMailService : ISupportMailService
{
    private static readonly string[] Scopes = ["https://graph.microsoft.com/.default"];

    private readonly SupportSettings _settings;
    private readonly GraphServiceClient _graphClient;

    public GraphMailService(SupportSettings settings)
    {
        _settings = settings;

        var credential = new ClientSecretCredential(
            _settings.Graph.TenantId,
            _settings.Graph.ClientId,
            _settings.Graph.ClientSecret);

        _graphClient = new GraphServiceClient(credential, Scopes);
    }

    public async Task<IReadOnlyList<IncomingMail>> FetchUnseenMailsAsync(CancellationToken cancellationToken)
    {
        var response = await _graphClient
            .Users[_settings.MailboxAddress]
            .MailFolders["Inbox"]
            .Messages
            .GetAsync(request =>
            {
                request.QueryParameters.Top = _settings.Graph.FetchTop;
                request.QueryParameters.Filter = "isRead eq false";
                request.QueryParameters.Orderby = ["receivedDateTime asc"];
                request.QueryParameters.Select =
                [
                    "id",
                    "internetMessageId",
                    "from",
                    "subject",
                    "body",
                    "bodyPreview",
                    "receivedDateTime"
                ];
            }, cancellationToken);

        var messages = response?.Value;
        if (messages is null || messages.Count == 0)
        {
            return Array.Empty<IncomingMail>();
        }

        var results = new List<IncomingMail>(messages.Count);
        foreach (var message in messages)
        {
            if (string.IsNullOrWhiteSpace(message.Id))
            {
                continue;
            }

            var sender = message.From?.EmailAddress;
            var messageId = string.IsNullOrWhiteSpace(message.InternetMessageId)
                ? message.Id
                : message.InternetMessageId;

            results.Add(new IncomingMail(
                message.Id,
                messageId,
                sender?.Address ?? string.Empty,
                sender?.Name ?? string.Empty,
                message.Subject ?? "(uden emne)",
                ExtractBody(message),
                message.ReceivedDateTime ?? DateTimeOffset.UtcNow,
                message.ConversationId ?? string.Empty));
        }

        return results;
    }

    public async Task MarkAsSeenAsync(IEnumerable<string> providerMessageIds, CancellationToken cancellationToken)
    {
        foreach (var providerMessageId in providerMessageIds)
        {
            if (string.IsNullOrWhiteSpace(providerMessageId))
            {
                continue;
            }

            await _graphClient
                .Users[_settings.MailboxAddress]
                .Messages[providerMessageId]
                .PatchAsync(new Message
                {
                    IsRead = true
                }, cancellationToken: cancellationToken);
        }
    }

    public async Task MoveMessagesToTrashAsync(IEnumerable<string> providerMessageIds, CancellationToken cancellationToken)
    {
        var movedCount = 0;
        var attemptedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var providerMessageId in providerMessageIds)
        {
            if (string.IsNullOrWhiteSpace(providerMessageId) || !attemptedIds.Add(providerMessageId))
            {
                continue;
            }

            movedCount += await TryMoveMessageByIdAsync(providerMessageId, cancellationToken);
        }

        if (movedCount > 0)
        {
            Console.WriteLine($"Papirkurv: flyttede {movedCount} behandlet indgaaende mail(s).");
        }
    }

    public async Task<SupportThreadAnchor?> SendSupportThreadStartAsync(string subject, string bodyText, string? bodyHtml, CancellationToken cancellationToken)
    {
        var useHtml = !string.IsNullOrWhiteSpace(bodyHtml);

        var draft = await _graphClient
            .Users[_settings.MailboxAddress]
            .Messages
            .PostAsync(new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = useHtml ? BodyType.Html : BodyType.Text,
                    Content = useHtml ? bodyHtml : bodyText
                },
                ToRecipients =
                [
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = _settings.MailboxAddress,
                            Name = _settings.MailboxDisplayName
                        }
                    }
                ]
            }, cancellationToken: cancellationToken);

        if (string.IsNullOrWhiteSpace(draft?.Id))
        {
            return null;
        }

        var conversationId = draft.ConversationId ?? string.Empty;

        await _graphClient
            .Users[_settings.MailboxAddress]
            .Messages[draft.Id]
            .Send
            .PostAsync(cancellationToken: cancellationToken);

        var sentMessage = await _graphClient
            .Users[_settings.MailboxAddress]
            .Messages[draft.Id]
            .GetAsync(request =>
            {
                request.QueryParameters.Select = ["id", "conversationId"];
            }, cancellationToken);

        var sentMessageId = string.IsNullOrWhiteSpace(sentMessage?.Id) ? draft.Id : sentMessage.Id;
        var effectiveConversationId = string.IsNullOrWhiteSpace(sentMessage?.ConversationId)
            ? conversationId
            : sentMessage.ConversationId;

        return new SupportThreadAnchor(sentMessageId, effectiveConversationId ?? string.Empty);
    }

    public async Task<SupportThreadAnchor?> ReplyInSupportThreadAsync(string conversationId, string providerMessageId, string bodyText, string? bodyHtml, CancellationToken cancellationToken)
    {
        var anchorMessageId = providerMessageId;
        if (string.IsNullOrWhiteSpace(anchorMessageId))
        {
            anchorMessageId = await FindLatestConversationMessageIdAsync(conversationId, cancellationToken);
        }

        if (string.IsNullOrWhiteSpace(anchorMessageId))
        {
            return null;
        }

        var draft = await _graphClient
            .Users[_settings.MailboxAddress]
            .Messages[anchorMessageId]
            .CreateReply
            .PostAsync(new CreateReplyPostRequestBody
            {
                Comment = string.Empty
            }, cancellationToken: cancellationToken);

        if (string.IsNullOrWhiteSpace(draft?.Id))
        {
            return null;
        }

        var useHtml = !string.IsNullOrWhiteSpace(bodyHtml);
        var effectiveConversationId = string.IsNullOrWhiteSpace(draft.ConversationId)
            ? conversationId
            : draft.ConversationId;

        await _graphClient
            .Users[_settings.MailboxAddress]
            .Messages[draft.Id]
            .PatchAsync(new Message
            {
                Body = new ItemBody
                {
                    ContentType = useHtml ? BodyType.Html : BodyType.Text,
                    Content = useHtml ? bodyHtml : bodyText
                }
            }, cancellationToken: cancellationToken);

        await _graphClient
            .Users[_settings.MailboxAddress]
            .Messages[draft.Id]
            .Send
            .PostAsync(cancellationToken: cancellationToken);

        var sentMessage = await _graphClient
            .Users[_settings.MailboxAddress]
            .Messages[draft.Id]
            .GetAsync(request =>
            {
                request.QueryParameters.Select = ["id", "conversationId"];
            }, cancellationToken);

        var sentMessageId = string.IsNullOrWhiteSpace(sentMessage?.Id) ? draft.Id : sentMessage.Id;
        var updatedConversationId = string.IsNullOrWhiteSpace(sentMessage?.ConversationId)
            ? effectiveConversationId
            : sentMessage.ConversationId;

        return new SupportThreadAnchor(sentMessageId, updatedConversationId ?? string.Empty);
    }

    public async Task MoveTicketMessagesToTrashAsync(string ticketId, IEnumerable<string> relatedProviderMessageIds, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(ticketId))
        {
            return;
        }

        var movedCount = 0;
        var attemptedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var messageId in relatedProviderMessageIds)
        {
            if (string.IsNullOrWhiteSpace(messageId) || !attemptedIds.Add(messageId))
            {
                continue;
            }

            movedCount += await TryMoveMessageByIdAsync(messageId, cancellationToken);
        }

        movedCount += await MoveFolderMessagesToTrashAsync("Inbox", ticketId, cancellationToken);
        movedCount += await MoveFolderMessagesToTrashAsync("SentItems", ticketId, cancellationToken);

        Console.WriteLine($"Papirkurv: flyttede {movedCount} mail(s) for {ticketId}.");
    }

    public async Task SendEmailAsync(
        string toAddress,
        string toName,
        string subject,
        string bodyText,
        string? bodyHtml,
        CancellationToken cancellationToken)
    {
        var useHtml = !string.IsNullOrWhiteSpace(bodyHtml);

        var requestBody = new SendMailPostRequestBody
        {
            Message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = useHtml ? BodyType.Html : BodyType.Text,
                    Content = useHtml ? bodyHtml : bodyText
                },
                ToRecipients =
                [
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = toAddress,
                            Name = string.IsNullOrWhiteSpace(toName) ? null : toName
                        }
                    }
                ]
            },
            SaveToSentItems = true
        };

        await _graphClient
            .Users[_settings.MailboxAddress]
            .SendMail
            .PostAsync(requestBody, cancellationToken: cancellationToken);
    }

    private static string ExtractBody(Message message)
    {
        var raw = message.Body?.Content;
        if (string.IsNullOrWhiteSpace(raw))
        {
            raw = message.BodyPreview;
        }

        if (string.IsNullOrWhiteSpace(raw))
        {
            return "(ingen tekst i mailen)";
        }

        if (message.Body?.ContentType == BodyType.Html)
        {
            return StripHtml(raw).Trim();
        }

        return raw.Trim();
    }

    private async Task<int> MoveFolderMessagesToTrashAsync(string folderId, string ticketId, CancellationToken cancellationToken)
    {
        var response = await _graphClient
            .Users[_settings.MailboxAddress]
            .MailFolders[folderId]
            .Messages
            .GetAsync(request =>
            {
                request.QueryParameters.Top = 100;
                request.QueryParameters.Select = ["id", "subject"];
            }, cancellationToken);

        var messages = response?.Value;
        if (messages is null || messages.Count == 0)
        {
            return 0;
        }

        var moved = 0;

        foreach (var message in messages)
        {
            if (string.IsNullOrWhiteSpace(message.Id))
            {
                continue;
            }

            var subject = message.Subject ?? string.Empty;
            if (!subject.Contains(ticketId, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            await _graphClient
                .Users[_settings.MailboxAddress]
                .Messages[message.Id]
                .Move
                .PostAsync(new MovePostRequestBody
                {
                    DestinationId = "deleteditems"
                }, cancellationToken: cancellationToken);

            moved++;
        }

        return moved;
    }

    private async Task<int> TryMoveMessageByIdAsync(string messageId, CancellationToken cancellationToken)
    {
        try
        {
            await _graphClient
                .Users[_settings.MailboxAddress]
                .Messages[messageId]
                .Move
                .PostAsync(new MovePostRequestBody
                {
                    DestinationId = "deleteditems"
                }, cancellationToken: cancellationToken);

            return 1;
        }
        catch
        {
            return 0;
        }
    }

    private async Task<string?> FindLatestConversationMessageIdAsync(string conversationId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(conversationId))
        {
            return null;
        }

        var response = await _graphClient
            .Users[_settings.MailboxAddress]
            .MailFolders["SentItems"]
            .Messages
            .GetAsync(request =>
            {
                request.QueryParameters.Top = 25;
                request.QueryParameters.Orderby = ["receivedDateTime desc"];
                request.QueryParameters.Select = ["id", "conversationId"];
            }, cancellationToken);

        return response?.Value?
            .FirstOrDefault(message => string.Equals(message.ConversationId, conversationId, StringComparison.OrdinalIgnoreCase))?
            .Id;
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
