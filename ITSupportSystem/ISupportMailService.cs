namespace ITSupportSystem;

public interface ISupportMailService
{
    Task<IReadOnlyList<IncomingMail>> FetchUnseenMailsAsync(CancellationToken cancellationToken);
    Task MarkAsSeenAsync(IEnumerable<string> providerMessageIds, CancellationToken cancellationToken);
    Task MoveMessagesToTrashAsync(IEnumerable<string> providerMessageIds, CancellationToken cancellationToken);
    Task<SupportThreadAnchor?> SendSupportThreadStartAsync(string subject, string bodyText, string? bodyHtml, CancellationToken cancellationToken);
    Task<SupportThreadAnchor?> ReplyInSupportThreadAsync(string conversationId, string providerMessageId, string bodyText, string? bodyHtml, CancellationToken cancellationToken);
    Task MoveTicketMessagesToTrashAsync(string ticketId, IEnumerable<string> relatedProviderMessageIds, CancellationToken cancellationToken);
    Task SendEmailAsync(
        string toAddress,
        string toName,
        string subject,
        string bodyText,
        string? bodyHtml,
        CancellationToken cancellationToken);
}
