namespace ITSupportSystem;

public enum TicketStatus
{
    Open,
    Closed
}

public sealed class Ticket
{
    public string TicketId { get; set; } = string.Empty;
    public string SenderAddress { get; set; } = string.Empty;
    public string SenderName { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public TicketStatus Status { get; set; } = TicketStatus.Open;
    public DateTimeOffset CreatedAt { get; set; }
    public DateTimeOffset UpdatedAt { get; set; }
    public DateTimeOffset? ClosedAt { get; set; }
    public string AssignedToId { get; set; } = string.Empty;
    public string AssignedToName { get; set; } = string.Empty;
    public string AssignedToEmail { get; set; } = string.Empty;
    public DateTimeOffset? AssignedAt { get; set; }
    public List<TicketEntry> Entries { get; set; } = new();
    public List<string> RelatedProviderMessageIds { get; set; } = new();
    public string SupportThreadMessageId { get; set; } = string.Empty;
    public string SupportConversationId { get; set; } = string.Empty;
}

public sealed class TicketEntry
{
    public DateTimeOffset Timestamp { get; set; }
    public string Direction { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
}

public sealed class TicketDatabase
{
    public List<Ticket> Tickets { get; set; } = new();
    public HashSet<string> ProcessedMessageIds { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed record IncomingMail(
    string ProviderMessageId,
    string MessageId,
    string SenderAddress,
    string SenderName,
    string Subject,
    string Body,
    DateTimeOffset ReceivedAt,
    string ConversationId);

public sealed record SupportThreadAnchor(
    string ProviderMessageId,
    string ConversationId);

public sealed class SupportSettings
{
    public string MailboxAddress { get; set; } = "itsupport@blind.dk";
    public string MailboxDisplayName { get; set; } = "IT Support";
    public int PollIntervalSeconds { get; set; } = 30;
    public string DataFilePath { get; set; } = "data/tickets.json";
    public string MailProvider { get; set; } = "Graph";
    public FeatureToggleSettings Features { get; set; } = new();
    public BlobStorageSettings BlobStorage { get; set; } = new();
    public GraphSettings Graph { get; set; } = new();
    public MailServerSettings Imap { get; set; } = new();
    public MailServerSettings Smtp { get; set; } = new();
}

public sealed class BlobStorageSettings
{
    public bool Enabled { get; set; }
    public string AccountUrl { get; set; } = string.Empty;
    public string ConnectionString { get; set; } = string.Empty;
    public string ContainerName { get; set; } = "itsupport";
    public string BlobName { get; set; } = "tickets.json";
}

public sealed class FeatureToggleSettings
{
    public bool EnableTicketPolling { get; set; } = true;
    public bool EnableAutoAcknowledgements { get; set; } = true;
    public bool EnableUpdateAcknowledgements { get; set; } = false;
    public bool EnableOutgoingNotifications { get; set; } = true;
    public bool EnableSupportActionCards { get; set; } = true;
    public bool SendSupportCardOnUpdates { get; set; } = false;
    public bool MoveProcessedIncomingToTrash { get; set; } = true;
    public bool MoveClosedTicketToTrash { get; set; } = true;
    public bool UseTestMailbox { get; set; } = false;
    public string TestMailboxAddress { get; set; } = "teb@blind.dk";
}

public sealed class GraphSettings
{
    public bool Enabled { get; set; } = true;
    public string TenantId { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;
    public int FetchTop { get; set; } = 25;
}

public sealed class MailServerSettings
{
    public string Host { get; set; } = string.Empty;
    public int Port { get; set; }
    public bool UseSsl { get; set; } = true;
    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
}
