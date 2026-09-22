using System.Text.Json;
using System.Text.Json.Nodes;

using ITSupportSystem;

var cancellation = new CancellationTokenSource();
Console.CancelKeyPress += (_, eventArgs) =>
{
	eventArgs.Cancel = true;
	cancellation.Cancel();
};

try
{
	var command = args.Length == 0 ? "run" : args[0].ToLowerInvariant();
	if (command == "help")
	{
		PrintHelp();
		return;
	}

	var settings = LoadSettings(command);
	ApplyMailboxOverride(settings);
	var store = new TicketStore(ResolveDataPath(settings.DataFilePath), settings.BlobStorage);
	ISupportMailService? mailService = null;
	SupportProcessor? processor = null;
	if (CommandNeedsMailProvider(command))
	{
		mailService = CreateMailService(settings);
		processor = new SupportProcessor(store, mailService, settings);
	}

	switch (command)
	{
		case "run":
			await RunContinuouslyAsync(RequireProcessor(processor), settings, cancellation.Token);
			break;

		case "poll-once":
			var processed = await RequireProcessor(processor).PollOnceAsync(cancellation.Token);
			Console.WriteLine($"Behandlet {processed} mail(s).");
			break;

		case "list":
			PrintTickets(store.GetAllTickets());
			break;

		case "reply":
			await ReplyAsync(RequireProcessor(processor), args, cancellation.Token);
			break;

		case "close":
			await CloseAsync(RequireProcessor(processor), args, cancellation.Token);
			break;

		default:
			PrintHelp();
			break;
	}
}
catch (OperationCanceledException)
{
	Console.WriteLine("Afbrudt af bruger.");
}
catch (Exception ex)
{
	Console.Error.WriteLine($"Fejl: {ex.Message}");
	Environment.ExitCode = 1;
}

static SupportSettings LoadSettings(string command)
{
	var appRoot = ResolveAppRoot();
	var configPath = Path.Combine(appRoot, "appsettings.json");
	if (!File.Exists(configPath))
	{
		configPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
	}

	if (!File.Exists(configPath))
	{
		throw new FileNotFoundException(
			"Mangler appsettings.json. Kopier appsettings.example.json til appsettings.json og udfyld loginoplysninger.");
	}

	var json = File.ReadAllText(configPath);
	var rootNode = JsonNode.Parse(json) ?? throw new InvalidOperationException("Kunne ikke laese appsettings.json");
	var localConfigPath = Path.Combine(appRoot, "appsettings.local.json");
	if (File.Exists(localConfigPath))
	{
		var localJson = File.ReadAllText(localConfigPath);
		var localNode = JsonNode.Parse(localJson);
		if (localNode is not null)
		{
			rootNode = MergeJsonNodes(rootNode, localNode);
		}
	}

	var settings = JsonSerializer.Deserialize<SupportSettings>(rootNode.ToJsonString(), new JsonSerializerOptions
	{
		PropertyNameCaseInsensitive = true
	}) ?? throw new InvalidOperationException("Kunne ikke laese appsettings.json");

	ApplyEnvironmentOverrides(settings);

	ValidateSettings(settings, command);
	return settings;
}

static JsonNode MergeJsonNodes(JsonNode baseNode, JsonNode overrideNode)
{
	if (baseNode is JsonObject baseObject && overrideNode is JsonObject overrideObject)
	{
		var merged = (JsonObject)(baseObject.DeepClone() ?? new JsonObject());
		foreach (var property in overrideObject)
		{
			if (property.Value is null)
			{
				merged[property.Key] = null;
				continue;
			}

			if (merged[property.Key] is JsonNode existingNode)
			{
				merged[property.Key] = MergeJsonNodes(existingNode, property.Value);
			}
			else
			{
				merged[property.Key] = property.Value.DeepClone();
			}
		}

		return merged;
	}

	return overrideNode.DeepClone() ?? baseNode.DeepClone() ?? new JsonObject();
}

static void ApplyEnvironmentOverrides(SupportSettings settings)
{
	settings.MailboxAddress = GetOptionalString("SUPPORT_MAILBOX_ADDRESS") ?? settings.MailboxAddress;
	settings.MailboxDisplayName = GetOptionalString("SUPPORT_MAILBOX_DISPLAY_NAME") ?? settings.MailboxDisplayName;
	settings.MailProvider = GetOptionalString("SUPPORT_MAIL_PROVIDER") ?? settings.MailProvider;
	settings.DataFilePath = GetOptionalString("SUPPORT_DATA_FILE_PATH") ?? settings.DataFilePath;
	settings.PollIntervalSeconds = GetOptionalInt("SUPPORT_POLL_INTERVAL_SECONDS") ?? settings.PollIntervalSeconds;

	settings.BlobStorage.Enabled = GetOptionalBool("TICKET_BLOB_ENABLED") ?? settings.BlobStorage.Enabled;
	settings.BlobStorage.AccountUrl = GetOptionalString("TICKET_BLOB_ACCOUNT_URL") ?? settings.BlobStorage.AccountUrl;
	settings.BlobStorage.ConnectionString = GetOptionalString("AZURE_STORAGE_CONNECTION_STRING") ?? settings.BlobStorage.ConnectionString;
	settings.BlobStorage.ContainerName = GetOptionalString("TICKET_BLOB_CONTAINER") ?? settings.BlobStorage.ContainerName;
	settings.BlobStorage.BlobName = GetOptionalString("TICKET_BLOB_NAME") ?? settings.BlobStorage.BlobName;

	settings.Graph.Enabled = GetOptionalBool("GRAPH_ENABLED") ?? settings.Graph.Enabled;
	settings.Graph.TenantId = GetOptionalString("GRAPH_TENANT_ID") ?? settings.Graph.TenantId;
	settings.Graph.ClientId = GetOptionalString("GRAPH_CLIENT_ID") ?? settings.Graph.ClientId;
	settings.Graph.ClientSecret = GetOptionalString("GRAPH_CLIENT_SECRET") ?? settings.Graph.ClientSecret;
	settings.Graph.FetchTop = GetOptionalInt("GRAPH_FETCH_TOP") ?? settings.Graph.FetchTop;

	settings.Features.EnableTicketPolling = GetOptionalBool("FEATURE_ENABLE_TICKET_POLLING") ?? settings.Features.EnableTicketPolling;
	settings.Features.EnableAutoAcknowledgements = GetOptionalBool("FEATURE_ENABLE_AUTO_ACKNOWLEDGEMENTS") ?? settings.Features.EnableAutoAcknowledgements;
	settings.Features.EnableUpdateAcknowledgements = GetOptionalBool("FEATURE_ENABLE_UPDATE_ACKNOWLEDGEMENTS") ?? settings.Features.EnableUpdateAcknowledgements;
	settings.Features.EnableOutgoingNotifications = GetOptionalBool("FEATURE_ENABLE_OUTGOING_NOTIFICATIONS") ?? settings.Features.EnableOutgoingNotifications;
	settings.Features.EnableSupportActionCards = GetOptionalBool("FEATURE_ENABLE_SUPPORT_ACTION_CARDS") ?? settings.Features.EnableSupportActionCards;
	settings.Features.SendSupportCardOnUpdates = GetOptionalBool("FEATURE_SEND_SUPPORT_CARD_ON_UPDATES") ?? settings.Features.SendSupportCardOnUpdates;
	settings.Features.MoveProcessedIncomingToTrash = GetOptionalBool("FEATURE_MOVE_PROCESSED_INCOMING_TO_TRASH") ?? settings.Features.MoveProcessedIncomingToTrash;
	settings.Features.MoveClosedTicketToTrash = GetOptionalBool("FEATURE_MOVE_CLOSED_TICKET_TO_TRASH") ?? settings.Features.MoveClosedTicketToTrash;
	settings.Features.UseTestMailbox = GetOptionalBool("FEATURE_USE_TEST_MAILBOX") ?? settings.Features.UseTestMailbox;
	settings.Features.TestMailboxAddress = GetOptionalString("FEATURE_TEST_MAILBOX_ADDRESS") ?? settings.Features.TestMailboxAddress;
}

static string? GetOptionalString(string variableName)
{
	var value = Environment.GetEnvironmentVariable(variableName);
	return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}

static bool? GetOptionalBool(string variableName)
{
	var value = GetOptionalString(variableName);
	if (value is null)
	{
		return null;
	}

	return bool.TryParse(value, out var parsed) ? parsed : null;
}

static int? GetOptionalInt(string variableName)
{
	var value = GetOptionalString(variableName);
	if (value is null)
	{
		return null;
	}

	return int.TryParse(value, out var parsed) ? parsed : null;
}

static void ValidateSettings(SupportSettings settings, string command)
{
	if (string.IsNullOrWhiteSpace(settings.MailboxAddress))
	{
		throw new InvalidOperationException("MailboxAddress skal vaere udfyldt.");
	}

	if (settings.Features.UseTestMailbox && string.IsNullOrWhiteSpace(settings.Features.TestMailboxAddress))
	{
		throw new InvalidOperationException("TestMailboxAddress skal vaere udfyldt, naar UseTestMailbox er true.");
	}

	if (settings.BlobStorage.Enabled)
	{
		if (string.IsNullOrWhiteSpace(settings.BlobStorage.ConnectionString)
			&& string.IsNullOrWhiteSpace(settings.BlobStorage.AccountUrl))
		{
			throw new InvalidOperationException("BlobStorage.ConnectionString eller BlobStorage.AccountUrl skal vaere udfyldt, naar BlobStorage.Enabled er true.");
		}

		if (string.IsNullOrWhiteSpace(settings.BlobStorage.ContainerName) ||
			string.IsNullOrWhiteSpace(settings.BlobStorage.BlobName))
		{
			throw new InvalidOperationException("BlobStorage.ContainerName og BlobStorage.BlobName skal vaere udfyldt, naar BlobStorage.Enabled er true.");
		}
	}

	if (!CommandNeedsMailProvider(command))
	{
		return;
	}

	if (string.Equals(settings.MailProvider, "Graph", StringComparison.OrdinalIgnoreCase) || settings.Graph.Enabled)
	{
		if (string.IsNullOrWhiteSpace(settings.Graph.TenantId) ||
			string.IsNullOrWhiteSpace(settings.Graph.ClientId) ||
			string.IsNullOrWhiteSpace(settings.Graph.ClientSecret))
		{
			throw new InvalidOperationException("Graph konfiguration er ugyldig. Udfyld TenantId, ClientId og ClientSecret.");
		}

		if (settings.Graph.TenantId.Contains("INDSAET", StringComparison.OrdinalIgnoreCase) ||
			settings.Graph.ClientId.Contains("INDSAET", StringComparison.OrdinalIgnoreCase) ||
			settings.Graph.ClientSecret.Contains("INDSAET", StringComparison.OrdinalIgnoreCase))
		{
			throw new InvalidOperationException("Graph konfiguration indeholder stadig placeholder-vaerdier. Udfyld TenantId, ClientId og ClientSecret.");
		}

		return;
	}

	if (string.IsNullOrWhiteSpace(settings.Imap.Host) || settings.Imap.Port <= 0)
	{
		throw new InvalidOperationException("IMAP konfiguration er ugyldig.");
	}

	if (string.IsNullOrWhiteSpace(settings.Smtp.Host) || settings.Smtp.Port <= 0)
	{
		throw new InvalidOperationException("SMTP konfiguration er ugyldig.");
	}
}

static ISupportMailService CreateMailService(SupportSettings settings)
{
	if (string.Equals(settings.MailProvider, "Graph", StringComparison.OrdinalIgnoreCase) || settings.Graph.Enabled)
	{
		Console.WriteLine($"Mailprovider: Graph ({settings.MailboxAddress})");
		return new GraphMailService(settings);
	}

	Console.WriteLine($"Mailprovider: MailKit ({settings.MailboxAddress})");
	return new MailService(settings);
}

static bool CommandNeedsMailProvider(string command)
{
	return command is "run" or "poll-once" or "reply" or "close";
}

static SupportProcessor RequireProcessor(SupportProcessor? processor)
{
	if (processor is null)
	{
		throw new InvalidOperationException("Mailservice er ikke initialiseret.");
	}

	return processor;
}

static void ApplyMailboxOverride(SupportSettings settings)
{
	if (settings.Features.UseTestMailbox)
	{
		settings.MailboxAddress = settings.Features.TestMailboxAddress;
		Console.WriteLine($"Testmailbox aktiv: {settings.MailboxAddress}");
	}
}

static string ResolveDataPath(string configuredPath)
{
	if (Path.IsPathRooted(configuredPath))
	{
		return configuredPath;
	}

	return Path.GetFullPath(Path.Combine(ResolveAppRoot(), configuredPath));
}

static string ResolveAppRoot()
{
	var current = Directory.GetCurrentDirectory();
	if (File.Exists(Path.Combine(current, "appsettings.json")))
	{
		return current;
	}

	var nestedProjectPath = Path.Combine(current, "ITSupportSystem", "appsettings.json");
	if (File.Exists(nestedProjectPath))
	{
		return Path.GetDirectoryName(nestedProjectPath) ?? current;
	}

	return current;
}

static async Task RunContinuouslyAsync(SupportProcessor processor, SupportSettings settings, CancellationToken cancellationToken)
{
	Console.WriteLine("IT Support overvagning startet. Tryk Ctrl+C for at stoppe.");
	while (!cancellationToken.IsCancellationRequested)
	{
		var handled = await processor.PollOnceAsync(cancellationToken);
		if (handled > 0)
		{
			Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss}: Behandlet {handled} mail(s).");
		}

		await Task.Delay(TimeSpan.FromSeconds(settings.PollIntervalSeconds), cancellationToken);
	}
}

static async Task ReplyAsync(SupportProcessor processor, string[] args, CancellationToken cancellationToken)
{
	if (args.Length < 3)
	{
		throw new ArgumentException("Brug: reply <TicketId> <Besked>");
	}

	var ticketId = args[1];
	var message = string.Join(' ', args.Skip(2));
	await processor.ReplyAsync(ticketId, message, cancellationToken);
	Console.WriteLine($"Svar sendt for {ticketId}.");
}

static async Task CloseAsync(SupportProcessor processor, string[] args, CancellationToken cancellationToken)
{
	if (args.Length < 3)
	{
		throw new ArgumentException("Brug: close <TicketId> <Besked>");
	}

	var ticketId = args[1];
	var message = string.Join(' ', args.Skip(2));
	await processor.CloseAsync(ticketId, message, cancellationToken);
	Console.WriteLine($"Sagen {ticketId} er lukket.");
}

static void PrintTickets(IReadOnlyList<Ticket> tickets)
{
	if (tickets.Count == 0)
	{
		Console.WriteLine("Ingen sager fundet.");
		return;
	}

	foreach (var ticket in tickets)
	{
		Console.WriteLine($"{ticket.TicketId} | {ticket.Status} | {ticket.SenderAddress} | {ticket.Subject}");
	}
}

static void PrintHelp()
{
	Console.WriteLine("Kommandoer:");
	Console.WriteLine("  run                            Start kontinuerlig overvagning");
	Console.WriteLine("  poll-once                      Hent og behandl mails en gang");
	Console.WriteLine("  list                           Vis alle sager");
	Console.WriteLine("  reply <TicketId> <Besked>      Send svar til bruger");
	Console.WriteLine("  close <TicketId> <Besked>      Luk sag og send afslutning");
	Console.WriteLine();
	Console.WriteLine("Feature toggles styres i appsettings.json:");
	Console.WriteLine("  Features.EnableTicketPolling");
	Console.WriteLine("  Features.EnableAutoAcknowledgements");
	Console.WriteLine("  Features.EnableOutgoingNotifications");
	Console.WriteLine("  Features.UseTestMailbox (med Features.TestMailboxAddress)");
}
