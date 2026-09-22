using System.Text.RegularExpressions;

namespace ITSupportSystem;

public sealed class SupportProcessor
{
    private static readonly Regex TicketIdRegex = new(@"SAG-\d{8}-\d{4}", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly TicketStore _store;
    private readonly ISupportMailService _mailService;
    private readonly SupportSettings _settings;

    public SupportProcessor(TicketStore store, ISupportMailService mailService, SupportSettings settings)
    {
        _store = store;
        _mailService = mailService;
        _settings = settings;
    }

    public async Task<int> PollOnceAsync(CancellationToken cancellationToken)
    {
        if (!_settings.Features.EnableTicketPolling)
        {
            Console.WriteLine("Polling er deaktiveret i Features.EnableTicketPolling.");
            return 0;
        }

        var mails = await _mailService.FetchUnseenMailsAsync(cancellationToken);
        if (mails.Count == 0)
        {
            return 0;
        }

        var handledProviderIds = new List<string>();
        var handledSupportCommandProviderIds = new List<string>();
        var createdTicketIncomingProviderIds = new List<string>();
        var ticketIdsToTrash = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var handledCount = 0;

        foreach (var mail in mails)
        {
            if (_store.HasProcessedMessage(mail.MessageId))
            {
                Console.WriteLine($"Springer over allerede behandlet mail: {mail.Subject}");
                handledProviderIds.Add(mail.ProviderMessageId);
                continue;
            }

            var existingTicket = ResolveTicket(mail);
            if (existingTicket is not null && IsSupportThreadAnchorMail(mail))
            {
                _store.SetSupportThreadAnchor(existingTicket, mail.ProviderMessageId, mail.ConversationId);
            }

            if (existingTicket is not null && IsSupportReplyRequest(mail, existingTicket))
            {
                var responseText = ExtractSupportReplyText(mail.Body);
                if (string.IsNullOrWhiteSpace(responseText))
                {
                    Console.WriteLine($"Springer over tom svar-kommando for {existingTicket.TicketId}.");
                }
                else
                {
                    await SendAgentReplyFromSupportInboxAsync(existingTicket, responseText, cancellationToken);
                    _store.AddOutgoingEntry(existingTicket, responseText);
                    await SendSupportTimelineUpdateAsync(
                        existingTicket,
                        "Svar sendt til bruger",
                        responseText,
                        cancellationToken);
                    Console.WriteLine($"Svar sendt til bruger for {existingTicket.TicketId}.");
                }

                _store.MarkMessageProcessed(mail.MessageId);
                handledProviderIds.Add(mail.ProviderMessageId);
                handledSupportCommandProviderIds.Add(mail.ProviderMessageId);
                handledCount++;
                continue;
            }

            if (IsInternalSupportMail(mail))
            {
                Console.WriteLine($"Springer over intern supportmail: {mail.Subject}");
                _store.MarkMessageProcessed(mail.MessageId);
                handledProviderIds.Add(mail.ProviderMessageId);
                continue;
            }

            if (existingTicket is null)
            {
                var newTicket = _store.CreateTicket(mail);
                if (_settings.Features.EnableAutoAcknowledgements)
                {
                    await SendAcknowledgementAsync(newTicket, cancellationToken);
                }
                else
                {
                    Console.WriteLine("Kvittering er deaktiveret i Features.EnableAutoAcknowledgements.");
                }

                if (_settings.Features.EnableSupportActionCards)
                {
                    await SendSupportActionCardAsync(newTicket, "Ny sag modtaget", cancellationToken, isFollowUp: false);
                }
                createdTicketIncomingProviderIds.Add(mail.ProviderMessageId);
                Console.WriteLine($"Ny sag oprettet: {newTicket.TicketId} fra {newTicket.SenderAddress}");
            }
            else
            {
                _store.AddIncomingEntry(existingTicket, mail.Body, mail.ProviderMessageId);
                var updateSnippet = BuildSupportSnippet(mail.Body);
                var reopenedTicket = false;

                if (IsCloseRequest(mail, existingTicket))
                {
                    if (IsSupportCommandMail(mail))
                    {
                        handledSupportCommandProviderIds.Add(mail.ProviderMessageId);
                    }

                    if (existingTicket.Status != TicketStatus.Closed)
                    {
                        _store.CloseTicket(existingTicket);
                        if (_settings.Features.EnableOutgoingNotifications)
                        {
                            await SendClosedByRequesterAsync(existingTicket, cancellationToken);
                        }
                        if (_settings.Features.MoveClosedTicketToTrash)
                        {
                            ticketIdsToTrash.Add(existingTicket.TicketId);
                        }
                        await SendSupportTimelineUpdateAsync(
                            existingTicket,
                            "Sagen er nu lukket",
                            updateSnippet,
                            cancellationToken);
                        Console.WriteLine($"Sagen er nu lukket: {existingTicket.TicketId}");
                    }
                }
                else if (existingTicket.Status == TicketStatus.Closed)
                {
                    _store.ReopenTicket(existingTicket);
                    reopenedTicket = true;
                    Console.WriteLine($"Sagen er genaabnet: {existingTicket.TicketId}");
                }

                if (reopenedTicket)
                {
                    updateSnippet = $"Sagen blev genaabnet af en ny besked fra bruger.\n\n{updateSnippet}";
                }

                if (!IsCloseRequest(mail, existingTicket) && _settings.Features.EnableUpdateAcknowledgements)
                {
                    await SendUpdateAcknowledgementAsync(existingTicket, cancellationToken);
                }

                if (!IsCloseRequest(mail, existingTicket) && !_settings.Features.EnableUpdateAcknowledgements)
                {
                    Console.WriteLine("Kvittering for opdatering er deaktiveret i Features.EnableUpdateAcknowledgements.");
                }

                if (!IsCloseRequest(mail, existingTicket) && _settings.Features.EnableSupportActionCards && _settings.Features.SendSupportCardOnUpdates)
                {
                    await SendSupportActionCardAsync(
                        existingTicket,
                        reopenedTicket ? "Sagen er genaabnet af bruger" : "Opdatering modtaget fra bruger",
                        cancellationToken,
                        isFollowUp: true,
                        details: updateSnippet);
                }
                else if (!IsCloseRequest(mail, existingTicket) && existingTicket.Status != TicketStatus.Closed)
                {
                    await SendSupportTimelineUpdateAsync(
                        existingTicket,
                        reopenedTicket ? "Sagen er genaabnet af bruger" : "Opdatering modtaget fra bruger",
                        updateSnippet,
                        cancellationToken);
                }

                Console.WriteLine($"Opdatering modtaget på {existingTicket.TicketId} fra {mail.SenderAddress}");
            }

            _store.MarkMessageProcessed(mail.MessageId);
            handledProviderIds.Add(mail.ProviderMessageId);
            handledCount++;
        }

        _store.Save();
        await _mailService.MarkAsSeenAsync(handledProviderIds, cancellationToken);

        if (_settings.Features.MoveProcessedIncomingToTrash)
        {
            await _mailService.MoveMessagesToTrashAsync(createdTicketIncomingProviderIds, cancellationToken);
        }

        if (handledSupportCommandProviderIds.Count > 0)
        {
            await _mailService.MoveMessagesToTrashAsync(handledSupportCommandProviderIds, cancellationToken);
        }

        if (_settings.Features.MoveClosedTicketToTrash)
        {
            foreach (var ticketId in ticketIdsToTrash)
            {
                var relatedIds = _store.GetRelatedProviderMessageIds(ticketId);
                await _mailService.MoveTicketMessagesToTrashAsync(ticketId, relatedIds, cancellationToken);
            }
        }

        return handledCount;
    }

    public IReadOnlyList<Ticket> ListTickets()
    {
        return _store.GetAllTickets();
    }

    public async Task ReplyAsync(string ticketId, string message, CancellationToken cancellationToken)
    {
        var ticket = _store.FindTicket(ticketId) ?? throw new InvalidOperationException($"Ticket findes ikke: {ticketId}");

        var subject = $"Re: [{ticket.TicketId}] {ticket.Subject}";
        var body = $"Hej {ResolveName(ticket)},\n\n" +
                   "Der er nyt svar på din IT-support sag.\n\n" +
                   $"Sagsnummer: {ticket.TicketId}\n" +
                   $"Svar:\n{message}\n\n" +
                   "Venlig hilsen\nIT Support";

        if (_settings.Features.EnableOutgoingNotifications)
        {
            await SendSupportEmailAsync(ticket, subject, body, BuildHtmlEmail(ticket, body), cancellationToken);
        }
        _store.AddOutgoingEntry(ticket, message);
        await SendSupportTimelineUpdateAsync(ticket, "Svar sendt til bruger", message, cancellationToken);
        _store.Save();
    }

    public async Task CloseAsync(string ticketId, string closingMessage, CancellationToken cancellationToken)
    {
        var ticket = _store.FindTicket(ticketId) ?? throw new InvalidOperationException($"Ticket findes ikke: {ticketId}");
        if (ticket.Status == TicketStatus.Closed)
        {
            throw new InvalidOperationException($"Ticket er allerede lukket: {ticketId}");
        }

        var subject = $"Sag lukket: [{ticket.TicketId}] {ticket.Subject}";
        var body = $"Hej {ResolveName(ticket)},\n\n" +
                   "Din IT-support sag er nu lukket.\n\n" +
                   $"Sagsnummer: {ticket.TicketId}\n" +
                   $"Afslutning:\n{closingMessage}\n\n" +
                   "Hvis problemet opstår igen, er du velkommen til at svare på denne mail.\n\n" +
                   "Venlig hilsen\nIT Support";

        if (_settings.Features.EnableOutgoingNotifications)
        {
            await SendSupportEmailAsync(ticket, subject, body, BuildHtmlEmail(ticket, body, includeCloseAction: false, replyButtonLabel: "Svar for at genaabne"), cancellationToken);
        }
        _store.AddOutgoingEntry(ticket, closingMessage);
        await SendSupportTimelineUpdateAsync(ticket, "Sagen er nu lukket", closingMessage, cancellationToken);
        _store.CloseTicket(ticket);
        _store.Save();

        if (_settings.Features.MoveClosedTicketToTrash)
        {
            var relatedIds = _store.GetRelatedProviderMessageIds(ticket.TicketId);
            await _mailService.MoveTicketMessagesToTrashAsync(ticket.TicketId, relatedIds, cancellationToken);
        }
    }

    private Ticket? ResolveTicket(IncomingMail mail)
    {
        var subjectMatch = TicketIdRegex.Match(mail.Subject ?? string.Empty);
        if (subjectMatch.Success)
        {
            return _store.FindTicket(subjectMatch.Value);
        }

        var bodyMatch = TicketIdRegex.Match(mail.Body ?? string.Empty);
        if (bodyMatch.Success)
        {
            return _store.FindTicket(bodyMatch.Value);
        }

        if (LooksLikeReply(mail.Subject))
        {
            return _store.FindMostRecentClosedTicketBySender(mail.SenderAddress);
        }

        return null;
    }

    private static bool LooksLikeReply(string? subject)
    {
        if (string.IsNullOrWhiteSpace(subject))
        {
            return false;
        }

        return subject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase)
            || subject.StartsWith("Sv:", StringComparison.OrdinalIgnoreCase)
            || subject.StartsWith("VS:", StringComparison.OrdinalIgnoreCase)
            || subject.StartsWith("FW:", StringComparison.OrdinalIgnoreCase)
            || subject.StartsWith("Fwd:", StringComparison.OrdinalIgnoreCase);
    }

    private Ticket? ResolveTicketBySubject(string subject)
    {
        var match = TicketIdRegex.Match(subject ?? string.Empty);
        if (!match.Success)
        {
            return null;
        }

        return _store.FindTicket(match.Value);
    }

    private async Task SendAcknowledgementAsync(Ticket ticket, CancellationToken cancellationToken)
    {
        var subject = $"Kvittering: [{ticket.TicketId}] {ticket.Subject}";
        var body = $"Hej {ResolveName(ticket)},\n\n" +
                   "Vi har modtaget din henvendelse til IT Support.\n\n" +
                   $"Sagsnummer: {ticket.TicketId}\n" +
                   "Vi vender tilbage hurtigst muligt.\n\n" +
                   "Venlig hilsen\nIT Support";

        await SendSupportEmailAsync(ticket, subject, body, BuildHtmlEmail(ticket, body), cancellationToken);
    }

    private async Task SendUpdateAcknowledgementAsync(Ticket ticket, CancellationToken cancellationToken)
    {
        var subject = $"Kvittering for opdatering: [{ticket.TicketId}] {ticket.Subject}";
        var body = $"Hej {ResolveName(ticket)},\n\n" +
                   "Tak for din opdatering på sagen. Vi har modtaget den.\n\n" +
                   $"Sagsnummer: {ticket.TicketId}\n\n" +
                   "Venlig hilsen\nIT Support";

        await SendSupportEmailAsync(ticket, subject, body, BuildHtmlEmail(ticket, body), cancellationToken);
    }

    private async Task SendClosedByRequesterAsync(Ticket ticket, CancellationToken cancellationToken)
    {
        var subject = $"Sag lukket: [{ticket.TicketId}] {ticket.Subject}";
        var body = $"Hej {ResolveName(ticket)},\n\n" +
                   "Sagen er nu lukket.\n\n" +
                   $"Sagsnummer: {ticket.TicketId}\n\n" +
                   "Venlig hilsen\nIT Support";

        await SendSupportEmailAsync(ticket, subject, body, BuildHtmlEmail(ticket, body, includeCloseAction: false, replyButtonLabel: "Svar for at genaabne"), cancellationToken);
    }

    private async Task SendSupportActionCardAsync(
        Ticket ticket,
        string header,
        CancellationToken cancellationToken,
        bool isFollowUp,
        string? details = null)
    {
        var subjectCore = $"[SUPPORT] [{ticket.TicketId}] {ticket.Subject}";
        var subject = isFollowUp ? $"Re: {subjectCore}" : subjectCore;
        var body = $"{header}\n\n" +
                   $"Sagsnummer: {ticket.TicketId}\n" +
                   $"Bruger: {ticket.SenderAddress}\n" +
                   (string.IsNullOrWhiteSpace(details) ? string.Empty : $"\nSeneste opdatering:\n{details}\n") +
                   "Brug knapperne nedenfor for at svare eller lukke sagen.";

        var html = BuildSupportActionCardHtml(ticket, header);
        if (!string.IsNullOrWhiteSpace(details))
        {
            var encodedDetails = System.Net.WebUtility.HtmlEncode(details).Replace("\n", "<br/>", StringComparison.Ordinal);
            html = html.Replace(
                "</body></html>",
                $"<div style='margin-top:12px;padding:12px;border-top:1px solid #d1d5db;'><p style='margin:0 0 6px 0;'><strong>Seneste opdatering:</strong></p><p style='margin:0;'>{encodedDetails}</p></div></body></html>",
                StringComparison.Ordinal);
        }

        if (!isFollowUp)
        {
            var anchor = await _mailService.SendSupportThreadStartAsync(subject, body, html, cancellationToken);
            if (anchor is not null)
            {
                _store.SetSupportThreadAnchor(ticket, anchor.ProviderMessageId, anchor.ConversationId);
            }
            else
            {
                await _mailService.SendEmailAsync(
                    _settings.MailboxAddress,
                    _settings.MailboxDisplayName,
                    subject,
                    body,
                    html,
                    cancellationToken);
            }
        }
        else
        {
            Console.WriteLine($"Support-kort for opdatering er deaktiveret: {subject}");
        }

        Console.WriteLine($"Support-kort sendt: {subject}");
    }

    private static bool IsCloseRequest(IncomingMail mail, Ticket ticket)
    {
        var subject = mail.Subject ?? string.Empty;
        var body = mail.Body ?? string.Empty;

        return subject.Contains($"LUK [{ticket.TicketId}]", StringComparison.OrdinalIgnoreCase)
            || subject.Contains($"CLOSE [{ticket.TicketId}]", StringComparison.OrdinalIgnoreCase)
            || body.Contains("#LUKSAG", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSupportReplyRequest(IncomingMail mail, Ticket ticket)
    {
        var subject = mail.Subject ?? string.Empty;
        var body = mail.Body ?? string.Empty;

        return subject.Contains($"SVAR [{ticket.TicketId}]", StringComparison.OrdinalIgnoreCase)
            || subject.Contains($"REPLY [{ticket.TicketId}]", StringComparison.OrdinalIgnoreCase)
            || body.Contains("#SVARSAG", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSupportCommandMail(IncomingMail mail)
    {
        var subject = mail.Subject ?? string.Empty;
        var body = mail.Body ?? string.Empty;

        return subject.Contains("SVAR [SAG-", StringComparison.OrdinalIgnoreCase)
            || subject.Contains("REPLY [SAG-", StringComparison.OrdinalIgnoreCase)
            || subject.Contains("LUK [SAG-", StringComparison.OrdinalIgnoreCase)
            || subject.Contains("CLOSE [SAG-", StringComparison.OrdinalIgnoreCase)
            || body.Contains("#SVARSAG", StringComparison.OrdinalIgnoreCase)
            || body.Contains("#LUKSAG", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSupportThreadMail(IncomingMail mail)
    {
        var subject = mail.Subject ?? string.Empty;
        return subject.Contains("[SUPPORT] [SAG-", StringComparison.OrdinalIgnoreCase);
    }

    private bool IsSupportThreadAnchorMail(IncomingMail mail)
    {
        return IsSupportThreadMail(mail) && IsInternalSupportMail(mail) && !IsSupportCommandMail(mail);
    }

    private static string ExtractSupportReplyText(string body)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return string.Empty;
        }

        var normalized = body.Replace("\r\n", "\n", StringComparison.Ordinal);
        var cleaned = normalized
            .Replace("#SVARSAG", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Skriv dit svar her og send mailen.", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();

        return cleaned;
    }

    private async Task SendAgentReplyFromSupportInboxAsync(Ticket ticket, string replyMessage, CancellationToken cancellationToken)
    {
        var subject = $"Re: [{ticket.TicketId}] {ticket.Subject}";
        var body = $"Hej {ResolveName(ticket)},\n\n" +
                   $"{replyMessage}\n\n" +
                   "Venlig hilsen\nIT Support";

        await SendSupportEmailAsync(ticket, subject, body, BuildHtmlEmail(ticket, body), cancellationToken);
    }

    private async Task SendSupportTimelineUpdateAsync(Ticket ticket, string title, string message, CancellationToken cancellationToken)
    {
        _ = message;
        _ = cancellationToken;
        Console.WriteLine($"Support-tidslinje gemt i sagen: {ticket.TicketId} - {title}");
    }

    private static string BuildSupportSnippet(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return "(ingen tekst)";
        }

        var normalized = text.Replace("\r\n", "\n", StringComparison.Ordinal).Trim();
        if (normalized.Length <= 1200)
        {
            return normalized;
        }

        return normalized[..1200] + "...";
    }

    private bool IsInternalSupportMail(IncomingMail mail)
    {
        var sender = mail.SenderAddress ?? string.Empty;
        var subject = mail.Subject ?? string.Empty;
        var body = mail.Body ?? string.Empty;

        if (body.Contains("#SVARSAG", StringComparison.OrdinalIgnoreCase)
            || body.Contains("#LUKSAG", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (sender.Equals(_settings.MailboxAddress, StringComparison.OrdinalIgnoreCase)
            || sender.Equals(_settings.Features.TestMailboxAddress, StringComparison.OrdinalIgnoreCase)
            || sender.Equals(_settings.Imap.Username, StringComparison.OrdinalIgnoreCase)
            || sender.Equals(_settings.Smtp.Username, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (subject.Contains("[SUPPORT] [SAG-", StringComparison.OrdinalIgnoreCase)
            || body.Contains("Brug knapperne nedenfor for at svare eller lukke sagen.", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private async Task SendSupportEmailAsync(
        Ticket ticket,
        string subject,
        string bodyText,
        string bodyHtml,
        CancellationToken cancellationToken)
    {
        await _mailService.SendEmailAsync(
            ticket.SenderAddress,
            ticket.SenderName,
            subject,
            bodyText,
            bodyHtml,
            cancellationToken);
    }

    private string BuildHtmlEmail(Ticket ticket, string bodyText, bool includeCloseAction = true, string replyButtonLabel = "Svar pa sag")
    {
        var encodedBody = System.Net.WebUtility.HtmlEncode(bodyText)
            .Replace("\n", "<br/>", StringComparison.Ordinal);

        var replySubject = Uri.EscapeDataString($"Re: [{ticket.TicketId}] {ticket.Subject}");
        var closeSubject = Uri.EscapeDataString($"LUK [{ticket.TicketId}] {ticket.Subject}");
        var closeBody = Uri.EscapeDataString("Skriv en kort besked (valgfrit) og send mailen.");

        var replyLink = $"mailto:{_settings.MailboxAddress}?subject={replySubject}";
         var closeLink = $"mailto:{_settings.MailboxAddress}?subject={closeSubject}&body={closeBody}";

         var actions = $"<a href='{replyLink}' style='display:inline-block;padding:10px 14px;background:#005a9c;color:#ffffff;text-decoration:none;border-radius:6px;{(includeCloseAction ? "margin-right:10px;" : string.Empty)}'>{System.Net.WebUtility.HtmlEncode(replyButtonLabel)}</a>";
         if (includeCloseAction)
         {
             actions += $"<a href='{closeLink}' style='display:inline-block;padding:10px 14px;background:#b00020;color:#ffffff;text-decoration:none;border-radius:6px;'>Luk sagen</a>";
         }

         return $"<html><body style='font-family:Segoe UI,Arial,sans-serif;'>" +
             $"<p>{encodedBody}</p>" +
             $"<p style='margin-top:20px;'>{actions}</p>" +
             "</body></html>";
    }

    private string BuildSupportActionCardHtml(Ticket ticket, string header)
    {
        var safeHeader = System.Net.WebUtility.HtmlEncode(header);
        var safeTicket = System.Net.WebUtility.HtmlEncode(ticket.TicketId);
        var safeSubject = System.Net.WebUtility.HtmlEncode(ticket.Subject);
        var safeSender = System.Net.WebUtility.HtmlEncode(ticket.SenderAddress);

        var supportThreadSubject = $"Re: [SUPPORT] [{ticket.TicketId}] {ticket.Subject}";
        var replySubject = Uri.EscapeDataString(supportThreadSubject);
        var replyBody = Uri.EscapeDataString("#SVARSAG\n\n");
        var replyLink = $"mailto:{_settings.MailboxAddress}?subject={replySubject}&body={replyBody}";

        var closeSubject = Uri.EscapeDataString(supportThreadSubject);
        var closeBody = Uri.EscapeDataString("#LUKSAG\n\n");
        var closeLink = $"mailto:{_settings.MailboxAddress}?subject={closeSubject}&body={closeBody}";

        return $"<html><body style='font-family:Segoe UI,Arial,sans-serif;'>" +
               $"<h3 style='margin:0 0 8px 0;color:#1f2937;'>{safeHeader}</h3>" +
               $"<p style='margin:0 0 4px 0;'><strong>Sag:</strong> {safeTicket}</p>" +
               $"<p style='margin:0 0 4px 0;'><strong>Emne:</strong> {safeSubject}</p>" +
               $"<p style='margin:0 0 16px 0;'><strong>Bruger:</strong> {safeSender}</p>" +
               "<p style='margin-top:20px;'>" +
               $"<a href='{replyLink}' style='display:inline-block;padding:10px 14px;background:#0b6d3a;color:#ffffff;text-decoration:none;border-radius:6px;margin-right:10px;'>Svar pa sag</a>" +
               $"<a href='{closeLink}' style='display:inline-block;padding:10px 14px;background:#b00020;color:#ffffff;text-decoration:none;border-radius:6px;'>Luk sagen</a>" +
               "</p></body></html>";
    }

    private static string ResolveName(Ticket ticket)
    {
        return string.IsNullOrWhiteSpace(ticket.SenderName)
            ? ""
            : ticket.SenderName;
    }
}
