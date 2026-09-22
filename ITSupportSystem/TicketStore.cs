using Azure.Identity;
using Azure.Storage.Blobs;
using System.Text.Json;

namespace ITSupportSystem;

public sealed class TicketStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _dataPath;
    private readonly BlobClient? _blobClient;
    private readonly object _syncRoot = new();
    private TicketDatabase _database;

    public TicketStore(string dataPath, BlobStorageSettings? blobStorage = null)
    {
        _dataPath = dataPath;
        _blobClient = CreateBlobClient(blobStorage);
        EnsureDirectoryExists();
        _database = Load();
    }

    public IReadOnlyList<Ticket> GetAllTickets()
    {
        lock (_syncRoot)
        {
            return _database.Tickets
                .OrderByDescending(t => t.UpdatedAt)
                .ToList();
        }
    }

    public Ticket? FindTicket(string ticketId)
    {
        lock (_syncRoot)
        {
            return _database.Tickets
                .FirstOrDefault(t => string.Equals(t.TicketId, ticketId, StringComparison.OrdinalIgnoreCase));
        }
    }

    public Ticket? FindMostRecentClosedTicketBySender(string senderAddress)
    {
        lock (_syncRoot)
        {
            return _database.Tickets
                .Where(t => t.Status == TicketStatus.Closed)
                .Where(t => string.Equals(t.SenderAddress, senderAddress, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(t => t.UpdatedAt)
                .FirstOrDefault();
        }
    }

    public bool HasProcessedMessage(string messageId)
    {
        lock (_syncRoot)
        {
            return _database.ProcessedMessageIds.Contains(messageId);
        }
    }

    public void MarkMessageProcessed(string messageId)
    {
        lock (_syncRoot)
        {
            _database.ProcessedMessageIds.Add(messageId);
        }
    }

    public Ticket CreateTicket(IncomingMail mail)
    {
        lock (_syncRoot)
        {
            var now = DateTimeOffset.UtcNow;
            var ticket = new Ticket
            {
                TicketId = GenerateTicketId(now),
                SenderAddress = mail.SenderAddress,
                SenderName = mail.SenderName,
                Subject = mail.Subject,
                CreatedAt = now,
                UpdatedAt = now
            };

            ticket.Entries.Add(new TicketEntry
            {
                Timestamp = now,
                Direction = "Incoming",
                Message = mail.Body
            });
            AddRelatedProviderMessageId(ticket, mail.ProviderMessageId);

            _database.Tickets.Add(ticket);
            return ticket;
        }
    }

    public void AddIncomingEntry(Ticket ticket, string body, string providerMessageId)
    {
        lock (_syncRoot)
        {
            ticket.Entries.Add(new TicketEntry
            {
                Timestamp = DateTimeOffset.UtcNow,
                Direction = "Incoming",
                Message = body
            });
            AddRelatedProviderMessageId(ticket, providerMessageId);
            ticket.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public IReadOnlyList<string> GetRelatedProviderMessageIds(string ticketId)
    {
        lock (_syncRoot)
        {
            var ticket = _database.Tickets
                .FirstOrDefault(t => string.Equals(t.TicketId, ticketId, StringComparison.OrdinalIgnoreCase));

            if (ticket is null)
            {
                return Array.Empty<string>();
            }

            return ticket.RelatedProviderMessageIds
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }

    public void SetSupportThreadAnchor(Ticket ticket, string providerMessageId, string conversationId)
    {
        lock (_syncRoot)
        {
            if (string.IsNullOrWhiteSpace(providerMessageId) && string.IsNullOrWhiteSpace(conversationId))
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(providerMessageId))
            {
                ticket.SupportThreadMessageId = providerMessageId;
            }

            if (!string.IsNullOrWhiteSpace(conversationId))
            {
                ticket.SupportConversationId = conversationId;
            }

            ticket.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public void AddOutgoingEntry(Ticket ticket, string body)
    {
        lock (_syncRoot)
        {
            ticket.Entries.Add(new TicketEntry
            {
                Timestamp = DateTimeOffset.UtcNow,
                Direction = "Outgoing",
                Message = body
            });
            ticket.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public void CloseTicket(Ticket ticket)
    {
        lock (_syncRoot)
        {
            ticket.Status = TicketStatus.Closed;
            ticket.ClosedAt = DateTimeOffset.UtcNow;
            ticket.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public void ReopenTicket(Ticket ticket)
    {
        lock (_syncRoot)
        {
            ticket.Status = TicketStatus.Open;
            ticket.ClosedAt = null;
            ticket.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public void Save()
    {
        lock (_syncRoot)
        {
            var json = JsonSerializer.Serialize(_database, JsonOptions);
            File.WriteAllText(_dataPath, json);
            SaveToBlobIfConfigured(json);
        }
    }

    private TicketDatabase Load()
    {
        var blobJson = LoadFromBlobIfConfigured();
        if (!string.IsNullOrWhiteSpace(blobJson))
        {
            File.WriteAllText(_dataPath, blobJson);
            return Normalize(JsonSerializer.Deserialize<TicketDatabase>(blobJson, JsonOptions) ?? new TicketDatabase());
        }

        if (!File.Exists(_dataPath))
        {
            return new TicketDatabase();
        }

        var json = File.ReadAllText(_dataPath);
        return Normalize(JsonSerializer.Deserialize<TicketDatabase>(json, JsonOptions) ?? new TicketDatabase());
    }

    private string GenerateTicketId(DateTimeOffset now)
    {
        var datePrefix = now.ToString("yyyyMMdd");
        var countToday = _database.Tickets.Count(t => t.TicketId.StartsWith($"SAG-{datePrefix}-", StringComparison.OrdinalIgnoreCase));
        var sequence = countToday + 1;
        return $"SAG-{datePrefix}-{sequence:0000}";
    }

    private void EnsureDirectoryExists()
    {
        var directory = Path.GetDirectoryName(_dataPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    private static BlobClient? CreateBlobClient(BlobStorageSettings? blobStorage)
    {
        if (blobStorage is null || !blobStorage.Enabled)
        {
            return null;
        }

        BlobContainerClient containerClient;
        if (!string.IsNullOrWhiteSpace(blobStorage.ConnectionString))
        {
            containerClient = new BlobContainerClient(blobStorage.ConnectionString, blobStorage.ContainerName);
        }
        else if (!string.IsNullOrWhiteSpace(blobStorage.AccountUrl))
        {
            var credential = new DefaultAzureCredential();
            var containerUri = new Uri($"{blobStorage.AccountUrl.TrimEnd('/')}/{blobStorage.ContainerName}");
            containerClient = new BlobContainerClient(containerUri, credential);
        }
        else
        {
            return null;
        }

        containerClient.CreateIfNotExists();
        return containerClient.GetBlobClient(blobStorage.BlobName);
    }

    private string? LoadFromBlobIfConfigured()
    {
        if (_blobClient is null)
        {
            return null;
        }

        try
        {
            if (!_blobClient.Exists())
            {
                return null;
            }

            return _blobClient.DownloadContent().Value.Content.ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Advarsel: kunne ikke laese tickets fra Azure Blob Storage. Bruger lokal cache. {ex.Message}");
            return null;
        }
    }

    private void SaveToBlobIfConfigured(string json)
    {
        if (_blobClient is null)
        {
            return;
        }

        try
        {
            using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json));
            _blobClient.Upload(stream, overwrite: true);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Advarsel: kunne ikke gemme tickets til Azure Blob Storage. Lokal fil er opdateret. {ex.Message}");
        }
    }

    private static TicketDatabase Normalize(TicketDatabase data)
    {
        data.ProcessedMessageIds ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        data.Tickets ??= new List<Ticket>();
        foreach (var ticket in data.Tickets)
        {
            ticket.RelatedProviderMessageIds ??= new List<string>();
            ticket.SupportThreadMessageId ??= string.Empty;
            ticket.SupportConversationId ??= string.Empty;
        }

        return data;
    }

    private static void AddRelatedProviderMessageId(Ticket ticket, string providerMessageId)
    {
        if (string.IsNullOrWhiteSpace(providerMessageId))
        {
            return;
        }

        if (!ticket.RelatedProviderMessageIds.Contains(providerMessageId, StringComparer.OrdinalIgnoreCase))
        {
            ticket.RelatedProviderMessageIds.Add(providerMessageId);
        }
    }
}
