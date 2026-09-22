document.addEventListener("DOMContentLoaded", () => {
  const tenantId = "7269da2b-d73e-4791-ba91-3675fa4b83f0";
  const clientId = "2ad3b090-e894-456c-8831-5761163ae173";
  const placeholderClientId = "REPLACE_WITH_AZURE_APP_CLIENT_ID";

  const authScreen = document.getElementById("auth-screen");
  const portalShell = document.getElementById("portal-shell");
  const loginButton = document.getElementById("login-button");
  const logoutButton = document.getElementById("logout-button");
  const authStatus = document.getElementById("auth-status");
  const signedInUser = document.getElementById("signed-in-user");

  const dropdowns = document.querySelectorAll(".dropdown");
  const resourceLinks = document.querySelectorAll(".open-resource");
  const formPortalButtons = document.querySelectorAll(".open-form-portal");
  const mailPortalButton = document.querySelector(".open-mail-portal");
  const wordPortalButton = document.querySelector(".open-word-portal");
  const excelPortalButton = document.querySelector(".open-excel-portal");
  const teamsPortalButton = document.querySelector(".open-teams-portal");
  const casesPortalButton = document.querySelector(".open-cases-portal");
  const resourceStatus = document.getElementById("resource-status");
  const sharepointPortalButton = document.querySelector(".open-sharepoint-portal");
  const oneDrivePortalButton = document.querySelector(".open-onedrive-portal");
  const sharepointView = document.getElementById("sharepoint-view");
  const sharepointTitle = document.getElementById("sharepoint-title");
  const mirrorSourceLabel = document.getElementById("mirror-source-label");
  const sharepointRefresh = document.getElementById("sharepoint-refresh");
  const sharepointBack = document.getElementById("sharepoint-back");
  const sharepointPath = document.getElementById("sharepoint-path");
  const sharepointStatus = document.getElementById("sharepoint-status");
  const sharepointList = document.getElementById("sharepoint-list");
  const teamsView = document.getElementById("teams-view");
  const teamsRefreshButton = document.getElementById("teams-refresh");
  const teamsStatus = document.getElementById("teams-status");
  const teamsList = document.getElementById("teams-list");
  const teamsChannelList = document.getElementById("teams-channel-list");
  const teamsChannelsTitle = document.getElementById("teams-channels-title");
  const casesView = document.getElementById("cases-view");
  const casesRefreshButton = document.getElementById("cases-refresh");
  const casesStatus = document.getElementById("cases-status");
  const casesFilterOpenButton = document.getElementById("cases-filter-open");
  const casesFilterClosedButton = document.getElementById("cases-filter-closed");
  const casesFilterAllButton = document.getElementById("cases-filter-all");
  const casesListTitle = document.getElementById("cases-list-title");
  const casesCount = document.getElementById("cases-count");
  const casesList = document.getElementById("cases-list");
  const casesDetail = document.getElementById("cases-detail");
  const formsView = document.getElementById("forms-view");
  const formsTitle = document.getElementById("forms-title");
  const formsDescription = document.getElementById("forms-description");
  const formsStatus = document.getElementById("forms-status");
  const formsReset = document.getElementById("forms-reset");
  const portalForm = document.getElementById("portal-form");
  const mailView = document.getElementById("mail-view");
  const mailForm = document.getElementById("mail-form");
  const mailStatus = document.getElementById("mail-status");
  const mailReset = document.getElementById("mail-reset");
  const mailToInput = document.getElementById("mail-to");
  const mailCcInput = document.getElementById("mail-cc");
  const mailBccInput = document.getElementById("mail-bcc");
  const mailSubjectInput = document.getElementById("mail-subject");
  const mailBodyInput = document.getElementById("mail-body");
  const mailContentTypeSelect = document.getElementById("mail-content-type");
  const mailHtmlPreviewWrapper = document.getElementById("mail-html-preview-wrapper");
  const mailHtmlPreview = document.getElementById("mail-html-preview");
  const mailAttachmentsInput = document.getElementById("mail-attachments");
  const mailPreviewHtmlButton = document.getElementById("mail-preview-html");
  const mailRefreshInboxButton = document.getElementById("mail-refresh-inbox");
  const mailCalendarRefreshButton = document.getElementById("mail-calendar-refresh");
  const mailCalendarPrevButton = document.getElementById("mail-calendar-prev");
  const mailCalendarTodayButton = document.getElementById("mail-calendar-today");
  const mailCalendarNextButton = document.getElementById("mail-calendar-next");
  const mailCalendarMonthLabel = document.getElementById("mail-calendar-month-label");
  const mailCalendarStatus = document.getElementById("mail-calendar-status");
  const mailCalendarModeMonthButton = document.getElementById("mail-calendar-mode-month");
  const mailCalendarModeWeekButton = document.getElementById("mail-calendar-mode-week");
  const mailCalendarModeWorkweekButton = document.getElementById("mail-calendar-mode-workweek");
  const mailCalendarGrid = document.getElementById("mail-calendar-grid");
  const mailCalendarAgenda = document.getElementById("mail-calendar-agenda");
  const mailCalendarDraftSubject = document.getElementById("mail-calendar-draft-subject");
  const mailCalendarDraftDate = document.getElementById("mail-calendar-draft-date");
  const mailCalendarDraftStart = document.getElementById("mail-calendar-draft-start");
  const mailCalendarDraftEnd = document.getElementById("mail-calendar-draft-end");
  const mailCalendarDraftLocation = document.getElementById("mail-calendar-draft-location");
  const mailCalendarDraftCopyButton = document.getElementById("mail-calendar-draft-copy");
  const mailCalendarDraftIcsButton = document.getElementById("mail-calendar-draft-ics");
  const mailCalendarDraftClearButton = document.getElementById("mail-calendar-draft-clear");
  const mailComposeTitle = document.getElementById("mail-compose-title");
  const mailInboxStatus = document.getElementById("mail-inbox-status");
  const mailFolderList = document.getElementById("mail-folder-list");
  const mailInboxList = document.getElementById("mail-inbox-list");
  const mailInboxDetail = document.getElementById("mail-inbox-detail");
  const mailReaderLive = document.getElementById("mail-reader-live");
  const wordView = document.getElementById("word-view");
  const wordStatus = document.getElementById("word-status");
  const wordEditor = document.getElementById("word-editor");
  const wordClearButton = document.getElementById("word-clear");
  const wordCopyButton = document.getElementById("word-copy");
  const wordOpenButton = document.getElementById("word-open");
  const wordOpenFileInput = document.getElementById("word-open-file");
  const wordSaveOneDriveButton = document.getElementById("word-save-onedrive");
  const wordSaveSharePointButton = document.getElementById("word-save-sharepoint");
  const wordBlockStyleSelect = document.getElementById("word-block-style");
  const wordFontFamilySelect = document.getElementById("word-font-family");
  const wordFontSizeSelect = document.getElementById("word-font-size");
  const wordLineHeightSelect = document.getElementById("word-line-height");
  const wordTextColorInput = document.getElementById("word-text-color");
  const wordHighlightColorInput = document.getElementById("word-highlight-color");
  const excelView = document.getElementById("excel-view");
  const excelStatus = document.getElementById("excel-status");
  const excelActiveCell = document.getElementById("excel-active-cell");
  const excelFormulaInput = document.getElementById("excel-formula");
  const excelGrid = document.getElementById("excel-grid");
  const excelAddRowButton = document.getElementById("excel-add-row");
  const excelAddColButton = document.getElementById("excel-add-col");
  const excelOpenButton = document.getElementById("excel-open");
  const excelOpenFileInput = document.getElementById("excel-open-file");
  const excelSaveOneDriveButton = document.getElementById("excel-save-onedrive");
  const excelSaveSharePointButton = document.getElementById("excel-save-sharepoint");
  const excelClearButton = document.getElementById("excel-clear");
  const excelFontFamilySelect = document.getElementById("excel-font-family");
  const excelFontSizeSelect = document.getElementById("excel-font-size");
  const excelBoldButton = document.getElementById("excel-bold");
  const excelTextColorInput = document.getElementById("excel-text-color");
  const excelFillColorInput = document.getElementById("excel-fill-color");
  const excelBorderStyleSelect = document.getElementById("excel-border-style");
  const excelColWidthInput = document.getElementById("excel-col-width");
  const excelRowHeightInput = document.getElementById("excel-row-height");
  const excelApplyColWidthButton = document.getElementById("excel-apply-col-width");
  const excelApplyRowHeightButton = document.getElementById("excel-apply-row-height");
  const excelAutoFitColButton = document.getElementById("excel-autofit-col");
  const excelAutoFitAllButton = document.getElementById("excel-autofit-all");

  const sharepointHost = "blindesamfund.sharepoint.com";
  const sharepointSitePath = "/sites/Faelles";
  const sharepointLibraryName = "Delte dokumenter";
  const powerAutomateFlowUrl = "REPLACE_WITH_POWER_AUTOMATE_HTTP_TRIGGER_URL";
  const portalLoginScopes = ["User.Read"];
  const mailGraphScopes = ["User.Read", "Mail.Send", "Mail.Read"];
  const calendarReadScopes = ["User.Read", "Calendars.ReadBasic"];
  const sharepointReadScopes = ["User.Read", "Sites.Read.All", "Files.Read.All"];
  const mailReadWriteScopes = ["User.Read", "Mail.ReadWrite"];
  const graphWriteScopes = ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"];
  const teamsReadScopes = ["User.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All"];
  const casesAutoRefreshMs = 2 * 60 * 1000;
  let sharepointSiteId = null;
  let sharepointDriveId = null;
  let sharepointDriveName = sharepointLibraryName;
  let sharepointCurrentPath = "";
  let mirrorMode = "sharepoint";
  let activeFormType = null;
  let joinedTeams = [];
  let selectedTeamId = "";
  let selectedTeamName = "";
  let tickets = [];
  let caseViewFilter = "open";
  let selectedTicketId = "";
  let casesAutoRefreshHandle = null;
  let inboxMessages = [];
  let mailboxFolders = [];
  let selectedMailFolderId = "__all__";
  let selectedMailFolderName = "Alle mails";
  let selectedInboxMessageId = "";
  let focusReadPaneAfterLoad = false;
  let markMessageAsReadAfterLoadId = "";
  let calendarVisibleMonth = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
  let calendarSelectedDate = new Date();
  let calendarViewMode = "month";
  let calendarEvents = [];
  let calendarLoading = false;
  let calendarDraggedEventId = "";
  let calendarDraft = {
    subject: "",
    date: new Date(),
    start: "09:00",
    end: "10:00",
    location: "",
  };

  const hiddenMailFolderNames = new Set([
    "alle mails",
    "arkiv",
    "rss-kilder2",
    "samtalehistorik",
    "synkroniseringsproblemer",
  ]);
  let excelRows = 20;
  let excelCols = 10;
  let excelData = [];
  let excelActiveRow = 0;
  let excelActiveCol = 0;
  let excelCellFormats = {};
  let excelColWidths = {};
  let excelRowHeights = {};
  let excelMeasureCanvas = null;

  const excelDefaultFontFamily = "Calibri";
  const excelDefaultFontSize = "12px";
  const excelDefaultTextColor = "#1f1f1f";
  const excelDefaultFillColor = "#ffffff";
  const excelDefaultBorderStyle = "none";
  const excelDefaultColWidth = 110;
  const excelDefaultRowHeight = 36;

  function hasXlsxRuntime() {
    return typeof window !== "undefined" && !!window.XLSX;
  }

  function hasMammothRuntime() {
    return typeof window !== "undefined" && !!window.mammoth && typeof window.mammoth.convertToHtml === "function";
  }

  function isDocxFile(file) {
    const lowerName = String(file?.name || "").toLowerCase();
    const mime = String(file?.type || "").toLowerCase();
    return lowerName.endsWith(".docx") || mime === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  }

  async function decodeTextFile(file) {
    const arrayBuffer = await file.arrayBuffer();
    const uint8 = new Uint8Array(arrayBuffer);

    try {
      return new TextDecoder("utf-8", { fatal: false }).decode(uint8).replace(/^\uFEFF/, "");
    } catch {
      // Fallback improves compatibility with Office exports saved in legacy encodings.
      return new TextDecoder("windows-1252", { fatal: false }).decode(uint8).replace(/^\uFEFF/, "");
    }
  }

  async function parseWordDocumentFile(file) {
    const lowerName = String(file?.name || "").toLowerCase();
    const mime = String(file?.type || "").toLowerCase();
    const isHtml = lowerName.endsWith(".html") || lowerName.endsWith(".htm") || mime.includes("html");

    if (isDocxFile(file)) {
      if (!hasMammothRuntime()) {
        throw new Error("DOCX bibliotek mangler. Prøv at genindlæse siden.");
      }

      const arrayBuffer = await file.arrayBuffer();
      const result = await window.mammoth.convertToHtml({ arrayBuffer });
      const html = String(result?.value || "").trim();
      if (!html) {
        throw new Error("DOCX-filen indeholdt ingen læsbar tekst.");
      }

      return {
        mode: "html",
        content: html,
        status: "Word-dokument (.docx) åbnet.",
      };
    }

    if (isHtml) {
      return {
        mode: "html",
        content: await decodeTextFile(file),
        status: "HTML-dokument åbnet.",
      };
    }

    return {
      mode: "text",
      content: await decodeTextFile(file),
      status: "Tekstdokument åbnet.",
    };
  }

  const formDefinitions = {
    hire: {
      title: "Ansættelse i portal",
      description: "Udfyld oplysningerne for ny ansættelse. Formularen gemmes lokalt i portalen som kvittering.",
      successMessage: "Ansættelsesformularen er registreret i portalen.",
      fields: [
        { name: "fuldeNavn", label: "Fulde navn", type: "text", required: true },
        { name: "email", label: "E-mail", type: "email", required: true },
        { name: "afdeling", label: "Afdeling", type: "text", required: true },
        { name: "stilling", label: "Stilling", type: "text", required: true },
        { name: "startdato", label: "Startdato", type: "date", required: true },
        { name: "leder", label: "Nærmeste leder", type: "text", required: true },
        { name: "noter", label: "Noter", type: "textarea", required: false },
      ],
    },
    termination: {
      title: "Afskedigelse i portal",
      description: "Udfyld oplysningerne ved afslutning af ansættelse. Formularen gemmes lokalt i portalen som kvittering.",
      successMessage: "Afskedigelsesformularen er registreret i portalen.",
      fields: [
        { name: "fuldeNavn", label: "Fulde navn", type: "text", required: true },
        { name: "email", label: "E-mail", type: "email", required: true },
        { name: "afdeling", label: "Afdeling", type: "text", required: true },
        { name: "slutdato", label: "Slutdato", type: "date", required: true },
        {
          name: "aarsag",
          label: "Årsag",
          type: "select",
          required: true,
          options: [
            { value: "", text: "Vælg årsag" },
            { value: "opsigelse-medarbejder", text: "Opsigelse fra medarbejder" },
            { value: "opsigelse-arbejdsgiver", text: "Opsigelse fra arbejdsgiver" },
            { value: "kontraktudloeb", text: "Kontraktudløb" },
            { value: "andet", text: "Andet" },
          ],
        },
        { name: "udstyr", label: "Udstyr til aflevering", type: "textarea", required: false },
        { name: "noter", label: "Noter", type: "textarea", required: false },
      ],
    },
  };

  function setAuthUi(isAuthenticated, account) {
    if (portalShell) {
      portalShell.classList.toggle("is-hidden", !isAuthenticated);
    }

    if (authScreen) {
      authScreen.classList.toggle("is-hidden", isAuthenticated);
    }

    document.body.classList.toggle("auth-locked", !isAuthenticated);

    if (signedInUser) {
      const displayName = account?.name || account?.username || "Ukendt bruger";
      signedInUser.textContent = displayName;
    }
  }

  function setAuthStatus(message) {
    if (authStatus) {
      authStatus.textContent = message;
    }
  }

  function setResourceStatus(message) {
    if (!resourceStatus) {
      return;
    }

    if (!message) {
      resourceStatus.classList.add("is-hidden");
      resourceStatus.textContent = "";
      return;
    }

    resourceStatus.classList.remove("is-hidden");
    resourceStatus.textContent = message;
  }

  function hidePortalViews() {
    if (sharepointView) {
      sharepointView.classList.add("is-hidden");
    }

    if (formsView) {
      formsView.classList.add("is-hidden");
    }

    if (mailView) {
      mailView.classList.add("is-hidden");
    }

    if (teamsView) {
      teamsView.classList.add("is-hidden");
    }

    if (casesView) {
      casesView.classList.add("is-hidden");
    }

    stopCasesAutoRefresh();

    if (wordView) {
      wordView.classList.add("is-hidden");
    }

    if (excelView) {
      excelView.classList.add("is-hidden");
    }
  }

  function setMailStatus(message) {
    if (mailStatus) {
      mailStatus.textContent = message;
    }
  }

  function setMailInboxStatus(message) {
    if (mailInboxStatus) {
      mailInboxStatus.textContent = message;
    }
  }

  function setTeamsStatus(message) {
    if (teamsStatus) {
      teamsStatus.textContent = message;
    }
  }

  function setCasesStatus(message) {
    if (casesStatus) {
      casesStatus.textContent = message;
    }
  }

  function getCasesAutoRefreshLabel() {
    return `Automatisk opdatering hvert ${Math.round(casesAutoRefreshMs / 60000)}. minut.`;
  }

  function stopCasesAutoRefresh() {
    if (casesAutoRefreshHandle) {
      window.clearInterval(casesAutoRefreshHandle);
      casesAutoRefreshHandle = null;
    }
  }

  function startCasesAutoRefresh() {
    stopCasesAutoRefresh();
    casesAutoRefreshHandle = window.setInterval(async () => {
      if (!casesView || casesView.classList.contains("is-hidden")) {
        stopCasesAutoRefresh();
        return;
      }

      await loadCases({ silent: true });
    }, casesAutoRefreshMs);
  }

  function clearTeamsLists() {
    if (teamsList) {
      teamsList.textContent = "";
    }

    if (teamsChannelList) {
      teamsChannelList.textContent = "";
    }
  }

  function renderTeamsList(items) {
    if (!teamsList) {
      return;
    }

    teamsList.textContent = "";
    if (!items || items.length === 0) {
      const emptyItem = document.createElement("li");
      emptyItem.className = "teams-item";
      emptyItem.textContent = "Ingen Teams fundet.";
      teamsList.appendChild(emptyItem);
      return;
    }

    items.forEach((team) => {
      const item = document.createElement("li");
      item.className = "teams-item";

      const button = document.createElement("button");
      button.type = "button";
      button.className = "teams-button";
      if ((team.id || "") === selectedTeamId) {
        button.classList.add("is-active");
      }

      const name = document.createElement("span");
      name.textContent = team.displayName || "Team uden navn";

      const meta = document.createElement("span");
      meta.className = "teams-item-meta";
      meta.textContent = team.description || "Ingen beskrivelse.";

      button.appendChild(name);
      button.appendChild(meta);
      button.addEventListener("click", async () => {
        selectedTeamId = team.id || "";
        selectedTeamName = team.displayName || "Team";
        renderTeamsList(joinedTeams);
        await loadTeamChannels(false, selectedTeamId, selectedTeamName);
      });

      item.appendChild(button);
      teamsList.appendChild(item);
    });
  }

  function buildTeamsChannelWebUrl(teamId, channel) {
    if (channel?.webUrl) {
      return channel.webUrl;
    }

    const encodedChannelId = encodeURIComponent(channel?.id || "");
    const encodedName = encodeURIComponent(channel?.displayName || "Kanal");
    const encodedTeamId = encodeURIComponent(teamId || "");
    const encodedTenantId = encodeURIComponent(tenantId);
    return `https://teams.microsoft.com/l/channel/${encodedChannelId}/${encodedName}?groupId=${encodedTeamId}&tenantId=${encodedTenantId}`;
  }

  function renderTeamChannels(channels, teamName, teamId) {
    if (teamsChannelsTitle) {
      teamsChannelsTitle.textContent = teamName ? `Kanaler i ${teamName}` : "Kanaler";
    }

    if (!teamsChannelList) {
      return;
    }

    teamsChannelList.textContent = "";
    if (!channels || channels.length === 0) {
      const emptyItem = document.createElement("li");
      emptyItem.className = "teams-channel-item";
      emptyItem.textContent = "Ingen kanaler fundet i dette Team.";
      teamsChannelList.appendChild(emptyItem);
      return;
    }

    channels.forEach((channel) => {
      const item = document.createElement("li");
      item.className = "teams-channel-item";

      const link = document.createElement("a");
      link.className = "teams-channel-link";
      link.href = buildTeamsChannelWebUrl(teamId, channel);
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = channel.displayName || "Kanal uden navn";

      const meta = document.createElement("span");
      meta.className = "teams-channel-meta";
      meta.textContent = channel.description || "Åbner kanalen i Teams-web.";

      link.appendChild(meta);
      item.appendChild(link);
      teamsChannelList.appendChild(item);
    });
  }

  function getTicketId(ticket) {
    return ticket?.TicketId || ticket?.ticketId || "SAG-?";
  }

  function getTicketSubject(ticket) {
    return ticket?.Subject || ticket?.subject || "(intet emne)";
  }

  function getTicketEntries(ticket) {
    if (Array.isArray(ticket?.Entries)) {
      return ticket.Entries;
    }

    if (Array.isArray(ticket?.entries)) {
      return ticket.entries;
    }

    return [];
  }

  function getTicketStatusLabel(ticket) {
    return ticket?.Status === 1 || ticket?.status === 1 ? "Lukket" : "Åben";
  }

  function isTicketClosed(ticket) {
    return getTicketStatusLabel(ticket) === "Lukket";
  }

  function getTicketAssignee(ticket) {
    return {
      id: String(ticket?.AssignedToId || ticket?.assignedToId || ticket?.AssignedTo?.id || ticket?.assignedTo?.id || "").trim(),
      name: String(ticket?.AssignedToName || ticket?.assignedToName || ticket?.AssignedTo?.name || ticket?.assignedTo?.name || "").trim(),
      email: String(ticket?.AssignedToEmail || ticket?.assignedToEmail || ticket?.AssignedTo?.email || ticket?.assignedTo?.email || "").trim(),
      assignedAt: ticket?.AssignedAt || ticket?.assignedAt || ticket?.AssignedTo?.assignedAt || ticket?.assignedTo?.assignedAt || null,
    };
  }

  function normalizeUserEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function getCaseFilterLabel(filter) {
    if (filter === "closed") {
      return "Alle lukkede sager";
    }

    if (filter === "all") {
      return "Alle sager";
    }

    return "Alle åbne sager";
  }

  function getFilteredTickets(items) {
    if (caseViewFilter === "closed") {
      return items.filter((ticket) => isTicketClosed(ticket));
    }

    if (caseViewFilter === "all") {
      return items;
    }

    return items.filter((ticket) => !isTicketClosed(ticket));
  }

  function updateCaseFilterButtons() {
    const mapping = [
      { button: casesFilterOpenButton, filter: "open" },
      { button: casesFilterClosedButton, filter: "closed" },
      { button: casesFilterAllButton, filter: "all" },
    ];

    mapping.forEach(({ button, filter }) => {
      if (!button) {
        return;
      }

      const isActive = caseViewFilter === filter;
      button.classList.toggle("is-active", isActive);
      button.setAttribute("aria-selected", isActive ? "true" : "false");
    });

    if (casesListTitle) {
      casesListTitle.textContent = getCaseFilterLabel(caseViewFilter);
    }
  }

  function setCaseViewFilter(nextFilter) {
    const allowedFilters = new Set(["open", "closed", "all"]);
    if (!allowedFilters.has(nextFilter)) {
      return;
    }

    caseViewFilter = nextFilter;
    updateCaseFilterButtons();
    const filteredTickets = getFilteredTickets(tickets);
    if (!filteredTickets.some((ticket) => getTicketId(ticket) === selectedTicketId)) {
      selectedTicketId = getTicketId(filteredTickets[0]);
    }
    renderCasesList(filteredTickets);
    const currentEmail = getCurrentUserEmailForUi();
    if (filteredTickets.length === 0) {
      setCasesStatus(currentEmail
        ? `Ingen sager fundet for ${currentEmail} i visningen "${getCaseFilterLabel(caseViewFilter)}". ${getCasesAutoRefreshLabel()}`
        : `Ingen sager fundet for den aktuelle bruger. ${getCasesAutoRefreshLabel()}`);
      return;
    }

    setCasesStatus(`Viser ${filteredTickets.length} sager i visningen "${getCaseFilterLabel(caseViewFilter)}". ${getCasesAutoRefreshLabel()}`);
  }

  function getCurrentUserIdentity() {
    const serverName = String(serverAuthenticatedUser?.name || "").trim();
    const serverEmail = String(serverAuthenticatedUser?.email || "").trim();
    if (serverEmail) {
      return {
        id: String(serverAuthenticatedUser?.id || "").trim(),
        name: serverName || serverEmail,
        email: serverEmail,
      };
    }

    const account = getActiveOrFirstAccount();
    const claims = account?.idTokenClaims || {};
    const claimEmailCandidates = [
      claims.preferred_username,
      claims.upn,
      Array.isArray(claims.emails) ? claims.emails[0] : claims.emails,
      account?.username,
    ];
    const email = String(claimEmailCandidates.find((value) => String(value || "").trim()) || "").trim();

    return {
      id: String(claims.oid || claims.objectId || "").trim(),
      name: String(account?.name || claims.name || email || "").trim(),
      email,
    };
  }

  async function runCaseAction(ticketId, message, action) {
    const response = await fetch(`/api/tickets/${action}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...getLocalIdentityHeaders(),
      },
      body: JSON.stringify({ ticketId, message }),
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload?.error || `Kunne ikke ${action === "reply" ? "sende svar" : "lukke sag"}.`);
    }

    return payload;
  }

  async function assignTicketToCurrentUser(ticketId) {
    const response = await fetch("/api/tickets/assign-self", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...getLocalIdentityHeaders(),
      },
      body: JSON.stringify({ ticketId }),
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload?.error || "Kunne ikke tildele sag.");
    }

    return payload;
  }

  function renderCaseDetail(ticket) {
    if (!casesDetail) {
      return;
    }

    casesDetail.textContent = "";

    if (ticket) {
      casesDetail.setAttribute("aria-busy", "true");
    }

    if (!ticket) {
      const heading = document.createElement("h3");
      heading.textContent = "Vælg en sag";
      const info = document.createElement("p");
      info.textContent = "Vælg en sag i listen for at se historikken.";
      casesDetail.appendChild(heading);
      casesDetail.appendChild(info);
      casesDetail.removeAttribute("aria-busy");
      return;
    }

    const header = document.createElement("div");
    header.className = "case-detail-header";

    const heading = document.createElement("h3");
    heading.textContent = `${getTicketId(ticket)} - ${getTicketSubject(ticket)}`;

    const meta = document.createElement("p");
    meta.className = "case-detail-meta";
    const senderName = ticket?.SenderName || ticket?.senderName || "Ukendt";
    const senderAddress = ticket?.SenderAddress || ticket?.senderAddress || "ukendt";
    const assignee = getTicketAssignee(ticket);
    const assigneeLabel = assignee.email
      ? `Tildelt: ${assignee.name || assignee.email} <${assignee.email}>`
      : "Tildelt: Ingen";
    meta.textContent = `${getTicketStatusLabel(ticket)} | ${senderName} <${senderAddress}> | ${assigneeLabel} | Oprettet ${formatDate(ticket?.CreatedAt || ticket?.createdAt)}`;

    header.appendChild(heading);
    header.appendChild(meta);

    const entries = document.createElement("ol");
    entries.className = "case-entry-list";

    const actions = document.createElement("section");
    actions.className = "case-actions";
    actions.setAttribute("aria-label", "Handlinger for sag");

    const actionHeading = document.createElement("h4");
    actionHeading.textContent = "Svar eller luk sag";

    const actionHelp = document.createElement("p");
    actionHelp.className = "case-detail-meta";
    actionHelp.textContent = "Skriv dit svar eller afslutningsbesked her. Svaret bliver sendt til brugeren og gemt i sagen.";

    const actionLabel = document.createElement("label");
    const actionInputId = `case-message-${getTicketId(ticket)}`;
    actionLabel.className = "visually-hidden";
    actionLabel.setAttribute("for", actionInputId);
    actionLabel.textContent = `Besked til ${getTicketId(ticket)}`;

    const actionInput = document.createElement("textarea");
    actionInput.id = actionInputId;
    actionInput.className = "case-action-input";
    const canReply = hasPortalPermission("portal.cases.reply");
    const canClose = hasPortalPermission("portal.cases.close");
    const isClosed = getTicketStatusLabel(ticket) === "Lukket";
    actionInput.placeholder = getTicketStatusLabel(ticket) === "Lukket"
      ? "Sagen er lukket."
      : "Skriv dit svar eller din afslutningsbesked her";
    actionInput.disabled = isClosed || (!canReply && !canClose);

    const actionStatus = document.createElement("p");
    actionStatus.className = "cases-status";
    actionStatus.setAttribute("role", "status");
    actionStatus.setAttribute("aria-live", "polite");
    actionStatus.textContent = isClosed
      ? "Sagen er allerede lukket."
      : (!canReply && !canClose)
        ? "Du har ikke adgang til at svare på eller lukke denne sag."
        : "Klar til at sende svar eller lukke sagen.";

    const actionButtons = document.createElement("div");
    actionButtons.className = "portal-form-actions";

    const assignButton = document.createElement("button");
    assignButton.type = "button";
    assignButton.className = "secondary-button";
    assignButton.textContent = "Sæt mig på sagen";

    const replyButton = document.createElement("button");
    replyButton.type = "button";
    replyButton.className = "primary-button";
    replyButton.textContent = "Send svar";
    replyButton.disabled = isClosed || !canReply;

    const closeButton = document.createElement("button");
    closeButton.type = "button";
    closeButton.className = "secondary-button";
    closeButton.textContent = "Luk sag";
    closeButton.disabled = isClosed || !canClose;

    if (!canReply) {
      replyButton.title = "Din rolle har ikke adgang til at svare på sager.";
      replyButton.setAttribute("aria-disabled", "true");
    }

    if (!canClose) {
      closeButton.title = "Din rolle har ikke adgang til at lukke sager.";
      closeButton.setAttribute("aria-disabled", "true");
    }

    const currentUser = getCurrentUserIdentity();
    const canAssign = hasPortalPermission("portal.cases.assign") || hasPortalPermission("*");
    const isAssignedToCurrentUser = assignee.email && normalizeUserEmail(assignee.email) === normalizeUserEmail(currentUser.email);
    assignButton.disabled = !canAssign || !currentUser.email || isAssignedToCurrentUser;

    if (!canAssign) {
      assignButton.title = "Din rolle har ikke adgang til at tildele sager.";
      assignButton.setAttribute("aria-disabled", "true");
    } else if (!currentUser.email) {
      assignButton.title = "Kunne ikke bestemme din brugeridentitet.";
      assignButton.setAttribute("aria-disabled", "true");
    } else if (isAssignedToCurrentUser) {
      assignButton.title = "Sagen er allerede tildelt dig.";
      assignButton.setAttribute("aria-disabled", "true");
    }

    const setBusyState = (isBusy) => {
      assignButton.disabled = isBusy || !canAssign || !currentUser.email || isAssignedToCurrentUser;
      replyButton.disabled = isBusy || isClosed || !canReply;
      closeButton.disabled = isBusy || isClosed || !canClose;
      actionInput.disabled = isBusy || isClosed || (!canReply && !canClose);
    };

    assignButton.addEventListener("click", async () => {
      actionStatus.textContent = "Tildeler sag...";
      setBusyState(true);

      try {
        const result = await assignTicketToCurrentUser(getTicketId(ticket));
        actionStatus.textContent = result?.warning
          ? `Sagen er tildelt dig. Advarsel: ${result.warning}`
          : "Sagen er tildelt dig. Brugeren er informeret.";
        await loadCases();
      } catch (error) {
        actionStatus.textContent = error?.message || "Kunne ikke tildele sag.";
      } finally {
        setBusyState(false);
      }
    });

    replyButton.addEventListener("click", async () => {
      const messageText = String(actionInput.value || "").trim();
      if (!messageText) {
        actionStatus.textContent = "Skriv en besked før du sender svar.";
        actionInput.focus();
        return;
      }

      actionStatus.textContent = "Sender svar...";
      setBusyState(true);

      try {
        await runCaseAction(getTicketId(ticket), messageText, "reply");
        actionInput.value = "";
        actionStatus.textContent = "Svar er sendt.";
        await loadCases();
      } catch (error) {
        actionStatus.textContent = error?.message || "Kunne ikke sende svar.";
      } finally {
        setBusyState(false);
      }
    });

    closeButton.addEventListener("click", async () => {
      const messageText = String(actionInput.value || "").trim();
      if (!messageText) {
        actionStatus.textContent = "Skriv en afslutningsbesked før du lukker sagen.";
        actionInput.focus();
        return;
      }

      actionStatus.textContent = "Lukker sagen...";
      setBusyState(true);

      try {
        await runCaseAction(getTicketId(ticket), messageText, "close");
        actionInput.value = "";
        actionStatus.textContent = "Sagen er lukket.";
        await loadCases();
      } catch (error) {
        actionStatus.textContent = error?.message || "Kunne ikke lukke sagen.";
      } finally {
        setBusyState(false);
      }
    });

    actionButtons.appendChild(assignButton);
    actionButtons.appendChild(replyButton);
    actionButtons.appendChild(closeButton);
    actions.appendChild(actionHeading);
    actions.appendChild(actionHelp);
    actions.appendChild(actionLabel);
    actions.appendChild(actionInput);
    actions.appendChild(actionButtons);
    actions.appendChild(actionStatus);

    const ticketEntries = getTicketEntries(ticket);
    if (ticketEntries.length === 0) {
      const empty = document.createElement("li");
      empty.className = "case-entry";
      empty.textContent = "Ingen historik i denne sag endnu.";
      entries.appendChild(empty);
    } else {
      ticketEntries.forEach((entry) => {
        const item = document.createElement("li");
        item.className = `case-entry ${String(entry?.Direction || entry?.direction || "").toLowerCase()}`;

        const entryMeta = document.createElement("p");
        entryMeta.className = "case-entry-meta";
        entryMeta.textContent = `${entry?.Direction || entry?.direction || "Ukendt"} | ${formatDate(entry?.Timestamp || entry?.timestamp)}`;

        const message = document.createElement("p");
        message.className = "case-entry-message";
        message.textContent = entry?.Message || entry?.message || "(tom besked)";

        item.appendChild(entryMeta);
        item.appendChild(message);
        entries.appendChild(item);
      });
    }

    casesDetail.appendChild(header);
    casesDetail.appendChild(actions);
    casesDetail.appendChild(entries);
    casesDetail.removeAttribute("aria-busy");
  }

  function renderCasesList(items) {
    if (!casesList) {
      return;
    }

    casesList.textContent = "";

    if (casesCount) {
      if (tickets.length !== items.length) {
        casesCount.textContent = `${items.length} af ${tickets.length} sager`;
      } else {
        casesCount.textContent = `${items.length} sager`;
      }
    }

    if (!items || items.length === 0) {
      const emptyItem = document.createElement("li");
      emptyItem.className = "case-item";
      emptyItem.textContent = "Ingen sager fundet.";
      casesList.appendChild(emptyItem);
      renderCaseDetail(null);
      return;
    }

    items.forEach((ticket) => {
      const item = document.createElement("li");
      const button = document.createElement("button");
      button.type = "button";
      button.className = "case-item-button";
      const ticketId = getTicketId(ticket);
      if (ticketId === selectedTicketId) {
        button.classList.add("is-active");
        button.setAttribute("aria-current", "true");
      } else {
        button.removeAttribute("aria-current");
      }
      button.setAttribute("aria-controls", "cases-detail");

      const title = document.createElement("p");
      title.className = "case-item-title";
      title.textContent = `${ticketId} - ${getTicketSubject(ticket)}`;

      const meta = document.createElement("p");
      meta.className = "case-item-meta";
      const assignee = getTicketAssignee(ticket);
      const assigneeLabel = assignee.email
        ? `Tildelt: ${assignee.name || assignee.email}`
        : "Tildelt: Ingen";
      meta.textContent = `${getTicketStatusLabel(ticket)} | ${assigneeLabel} | ${formatDate(ticket?.UpdatedAt || ticket?.updatedAt)}`;

      button.appendChild(title);
      button.appendChild(meta);
      button.addEventListener("click", () => {
        selectedTicketId = ticketId;
        renderCasesList(getFilteredTickets(tickets));
        if (casesDetail) {
          casesDetail.focus();
        }
      });

      item.appendChild(button);
      casesList.appendChild(item);
    });

    const activeTicket = items.find((ticket) => getTicketId(ticket) === selectedTicketId) || items[0];
    selectedTicketId = getTicketId(activeTicket);
    renderCaseDetail(activeTicket);
  }

  async function loadCases(options = {}) {
    if (!casesView) {
      return;
    }

    const silent = options?.silent === true;

    if (!silent) {
      setCasesStatus("Henter sager...");
    }

    try {
      const response = await fetch("/api/tickets", {
        cache: "no-store",
        headers: {
          ...getLocalIdentityHeaders(),
        },
      });
      if (!response.ok) {
        throw new Error(`Kunne ikke hente sager (${response.status})`);
      }

      const payload = await response.json();
      tickets = Array.isArray(payload?.Tickets) ? payload.Tickets : [];
      tickets.sort((a, b) => new Date(b.UpdatedAt || b.updatedAt || 0) - new Date(a.UpdatedAt || a.updatedAt || 0));
      const filteredTickets = getFilteredTickets(tickets);
      if (!selectedTicketId || !filteredTickets.some((ticket) => getTicketId(ticket) === selectedTicketId)) {
        selectedTicketId = getTicketId(filteredTickets[0]);
      }

      renderCasesList(filteredTickets);
      if (filteredTickets.length === 0) {
        const currentEmail = getCurrentUserEmailForUi();
        setCasesStatus(currentEmail
          ? `Ingen sager fundet for ${currentEmail} i visningen "${getCaseFilterLabel(caseViewFilter)}". ${getCasesAutoRefreshLabel()}`
          : `Ingen sager fundet for den aktuelle bruger. ${getCasesAutoRefreshLabel()}`);
      } else {
        const nowLabel = new Date().toLocaleTimeString("da-DK", { hour: "2-digit", minute: "2-digit" });
        setCasesStatus(
          silent
            ? `Sager opdateret automatisk kl. ${nowLabel}. ${getCasesAutoRefreshLabel()}`
            : `Viser ${filteredTickets.length} sager i visningen "${getCaseFilterLabel(caseViewFilter)}". ${getCasesAutoRefreshLabel()}`,
        );
      }
      setResourceStatus("Sagsvisning er åbnet i portalen.");
    } catch (error) {
      tickets = [];
      renderCasesList(tickets);
      const message = error?.message || "Ukendt fejl";
      setCasesStatus(`Kunne ikke hente sager: ${message}`);
    }
  }

  async function loadTeamsList(interactiveAllowed) {
    if (!teamsView) {
      return;
    }

    setTeamsStatus("Henter Teams...");
    if (teamsList) {
      teamsList.textContent = "";
    }

    try {
      let accessToken = await acquireGraphToken(interactiveAllowed, false, teamsReadScopes, {
        allowRedirect: !isLoopback,
        pendingAction: {
          action: "teams",
          payload: {},
        },
      });
      let response;

      try {
        response = await fetchGraph(
          "https://graph.microsoft.com/v1.0/me/joinedTeams?$select=id,displayName,description,webUrl",
          accessToken,
        );
      } catch (error) {
        if (!interactiveAllowed || !isAccessDeniedMessage(error?.message)) {
          throw error;
        }

        accessToken = await acquireGraphToken(true, true, teamsReadScopes);
        response = await fetchGraph(
          "https://graph.microsoft.com/v1.0/me/joinedTeams?$select=id,displayName,description,webUrl",
          accessToken,
        );
      }

      joinedTeams = response.value || [];
      if (!selectedTeamId || !joinedTeams.some((team) => (team.id || "") === selectedTeamId)) {
        selectedTeamId = joinedTeams[0]?.id || "";
        selectedTeamName = joinedTeams[0]?.displayName || "";
      }

      renderTeamsList(joinedTeams);
      setTeamsStatus(`Viser ${joinedTeams.length} Teams.`);
    } catch (error) {
      if (isGraphRedirectStartedError(error)) {
        setTeamsStatus("Fortsætter login i samme fane. Du sendes tilbage til Teams automatisk.");
        return;
      }

      const message = error?.message || "Ukendt fejl";
      setTeamsStatus(`Kunne ikke hente Teams: ${message}`);
      clearTeamsLists();
      throw error;
    }
  }

  async function loadTeamChannels(interactiveAllowed, teamId = selectedTeamId, teamName = selectedTeamName) {
    if (!teamsView) {
      return;
    }

    if (!teamId) {
      renderTeamChannels([], "", "");
      return;
    }

    setTeamsStatus(`Henter kanaler i ${teamName || "valgt Team"}...`);
    if (teamsChannelList) {
      teamsChannelList.textContent = "";
    }

    try {
      let accessToken = await acquireGraphToken(interactiveAllowed, false, teamsReadScopes);
      let response;

      try {
        response = await fetchGraph(
          `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels?$select=id,displayName,description,webUrl,membershipType`,
          accessToken,
        );
      } catch (error) {
        if (!interactiveAllowed || !isAccessDeniedMessage(error?.message)) {
          throw error;
        }

        accessToken = await acquireGraphToken(true, true, teamsReadScopes);
        response = await fetchGraph(
          `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels?$select=id,displayName,description,webUrl,membershipType`,
          accessToken,
        );
      }

      const channels = response.value || [];
      renderTeamChannels(channels, teamName || "Team", teamId);
      setTeamsStatus(`Viser ${channels.length} kanaler i ${teamName || "valgt Team"}.`);
    } catch (error) {
      const message = error?.message || "Ukendt fejl";
      setTeamsStatus(`Kunne ikke hente kanaler: ${message}`);
      renderTeamChannels([], teamName || "Team", teamId);
    }
  }

  async function openTeamsPortal(interactiveAllowed) {
    if (!teamsView) {
      return;
    }

    if (!hasPortalPermission("portal.teams")) {
      setResourceStatus("Du har ikke adgang til Teams i portalen.");
      return;
    }

    hidePortalViews();
    teamsView.classList.remove("is-hidden");
    setResourceStatus("Teams mirror er åbnet i portalen.");

    try {
      await loadTeamsList(interactiveAllowed);
      await loadTeamChannels(false, selectedTeamId, selectedTeamName);
    } catch {
      // Status messages are handled in load functions.
    }

    teamsView.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function getMailDetailTextElement() {
    return document.getElementById("mail-detail-text");
  }

  function announceMailForScreenReader(text) {
    if (!mailReaderLive) {
      return;
    }

    mailReaderLive.textContent = "";
    window.setTimeout(() => {
      mailReaderLive.textContent = text;
    }, 20);
  }

  function focusMailReadPane(announce = false) {
    const target = getMailDetailTextElement() || mailInboxDetail;
    if (!target) {
      return;
    }

    target.focus();
    if (announce) {
      const detailText = (mailInboxDetail?.innerText || "").trim();
      if (detailText) {
        announceMailForScreenReader(detailText);
      }
    }
  }

  async function ensureMessageDetailContent(message, interactiveAllowed) {
    if (!message || !message.id) {
      return;
    }

    if (message._detailLoaded) {
      return;
    }

    let accessToken = await acquireGraphToken(interactiveAllowed, false, mailGraphScopes);
    const detailUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(message.id)}?$select=id,subject,from,receivedDateTime,body,bodyPreview`;

    let detail;
    try {
      detail = await fetchGraph(detailUrl, accessToken, {
        Prefer: 'outlook.body-content-type="text"',
      });
    } catch (error) {
      if (!interactiveAllowed || !isAccessDeniedMessage(error?.message)) {
        throw error;
      }

      accessToken = await acquireGraphToken(true, true, mailGraphScopes);
      detail = await fetchGraph(detailUrl, accessToken, {
        Prefer: 'outlook.body-content-type="text"',
      });
    }

    message.subject = detail.subject || message.subject;
    message.from = detail.from || message.from;
    message.receivedDateTime = detail.receivedDateTime || message.receivedDateTime;
    message.bodyPreview = detail.bodyPreview || message.bodyPreview;
    message.body = detail.body || message.body;
    message._detailBodyText = String(detail.body?.content || detail.bodyPreview || "");
    message._detailLoaded = true;
  }

  async function markMessageAsRead(message, interactiveAllowed) {
    if (!message || !message.id || message.isRead) {
      return;
    }

    const patchUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(message.id)}`;
    const patchBody = JSON.stringify({ isRead: true });

    async function patchReadState(accessToken) {
      const response = await fetch(patchUrl, {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: patchBody,
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph fejl (${response.status}): ${errorText || "ukendt fejl"}`);
      }
    }

    let accessToken = await acquireGraphToken(interactiveAllowed, false, mailReadWriteScopes);

    try {
      await patchReadState(accessToken);
    } catch (error) {
      if (!interactiveAllowed || !isAccessDeniedMessage(error?.message)) {
        throw error;
      }

      accessToken = await acquireGraphToken(true, true, mailReadWriteScopes);
      await patchReadState(accessToken);
    }

    message.isRead = true;
  }

  async function openMessageInDetail(message, interactiveAllowed, moveFocusToReadPane, markAsRead = false) {
    if (!message) {
      return;
    }

    selectedInboxMessageId = message.id || "";

    const detailTextElement = getMailDetailTextElement();
    if (detailTextElement) {
      detailTextElement.textContent = "";
    }
    setMailInboxStatus("Henter valgt mails indhold...");

    try {
      await ensureMessageDetailContent(message, interactiveAllowed);
      if (markAsRead) {
        try {
          await markMessageAsRead(message, true);
        } catch (error) {
          const messageText = error?.message || "Ukendt fejl";
          setMailStatus(`Kunne ikke markere mail som læst: ${messageText}`);
        }
      }
      renderMailDetail(message);
      if (moveFocusToReadPane) {
        focusMailReadPane(true);
      }
    } catch (error) {
      const messageText = error?.message || "Ukendt fejl";
      setMailInboxStatus(`Kunne ikke hente mailindhold: ${messageText}`);
      renderMailDetail(message);
      if (moveFocusToReadPane) {
        focusMailReadPane(true);
      }
    }
  }

  function setWordStatus(message) {
    if (wordStatus) {
      wordStatus.textContent = message;
    }
  }

  function openWordPortal() {
    if (!wordView) {
      return;
    }

    if (!hasPortalPermission("portal.word")) {
      setResourceStatus("Du har ikke adgang til Word i portalen.");
      return;
    }

    hidePortalViews();
    wordView.classList.remove("is-hidden");
    setResourceStatus("Word editor er åbnet i portalen.");
    setWordStatus("Editor klar.");
    wordView.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function runWordCommand(command, value = null) {
    if (!wordEditor) {
      return;
    }

    wordEditor.focus();
    document.execCommand("styleWithCSS", false, true);
    document.execCommand(command, false, value);
  }

  function handleWordToolCommand(command) {
    if (!command) {
      return;
    }

    if (command === "createLink") {
      const url = window.prompt("Indsæt link (fx https://blind.dk):", "https://");
      if (!url) {
        setWordStatus("Link-indsættelse annulleret.");
        return;
      }

      runWordCommand("createLink", url);
      setWordStatus("Link indsat.");
      return;
    }

    runWordCommand(command);

    const commandLabels = {
      bold: "Fed",
      italic: "Kursiv",
      underline: "Understreg",
      superscript: "Hævet",
      subscript: "Sænket",
      strikeThrough: "Gennemstreget",
      insertUnorderedList: "Punktliste",
      insertOrderedList: "Nummereret liste",
      justifyLeft: "Venstre",
      justifyCenter: "Centrer",
      justifyRight: "Højre",
      justifyFull: "Juster",
      unlink: "Fjern link",
      undo: "Fortryd",
      redo: "Gendan",
      removeFormat: "Fjern formatering",
    };

    setWordStatus(`Værktøj anvendt: ${commandLabels[command] || command}.`);
  }

  function applyWordFontSize(fontSize) {
    if (!wordEditor) {
      return;
    }

    const normalizedSize = String(fontSize || "12pt").trim();
    runWordCommand("fontSize", "7");

    const fontElements = wordEditor.querySelectorAll('font[size="7"]');
    fontElements.forEach((fontElement) => {
      fontElement.removeAttribute("size");
      fontElement.style.fontSize = normalizedSize;
    });
  }

  function setExcelStatus(message) {
    if (excelStatus) {
      excelStatus.textContent = message;
    }
  }

  function createExcelData(rows, cols) {
    return Array.from({ length: rows }, () => Array.from({ length: cols }, () => ""));
  }

  function ensureExcelDataShape() {
    while (excelData.length < excelRows) {
      excelData.push(Array.from({ length: excelCols }, () => ""));
    }

    excelData = excelData.slice(0, excelRows).map((row) => {
      const resizedRow = Array.from({ length: excelCols }, (_, idx) => row[idx] ?? "");
      return resizedRow;
    });
  }

  function clampNumber(value, min, max, fallback) {
    const parsed = Number.parseInt(String(value), 10);
    if (!Number.isFinite(parsed)) {
      return fallback;
    }
    return Math.max(min, Math.min(max, parsed));
  }

  function normalizeHexColor(value, fallback) {
    const normalized = String(value || "").trim().toLowerCase();
    return /^#[0-9a-f]{6}$/i.test(normalized) ? normalized : fallback;
  }

  function getExcelCellFormatKey(row, col) {
    return `${row}:${col}`;
  }

  function getExcelCellFormat(row, col) {
    const stored = excelCellFormats[getExcelCellFormatKey(row, col)] || {};
    return {
      fontFamily: stored.fontFamily || excelDefaultFontFamily,
      fontSize: stored.fontSize || excelDefaultFontSize,
      isBold: Boolean(stored.isBold),
      textColor: normalizeHexColor(stored.textColor, excelDefaultTextColor),
      fillColor: normalizeHexColor(stored.fillColor, excelDefaultFillColor),
      borderStyle: stored.borderStyle === "all" ? "all" : excelDefaultBorderStyle,
    };
  }

  function setExcelCellFormat(row, col, partialFormat) {
    const key = getExcelCellFormatKey(row, col);
    const previous = getExcelCellFormat(row, col);
    const next = {
      ...previous,
      ...(partialFormat || {}),
    };

    next.fontFamily = String(next.fontFamily || excelDefaultFontFamily);
    next.fontSize = String(next.fontSize || excelDefaultFontSize);
    next.isBold = Boolean(next.isBold);
    next.textColor = normalizeHexColor(next.textColor, excelDefaultTextColor);
    next.fillColor = normalizeHexColor(next.fillColor, excelDefaultFillColor);
    next.borderStyle = next.borderStyle === "all" ? "all" : "none";

    const isDefault = next.fontFamily === excelDefaultFontFamily
      && next.fontSize === excelDefaultFontSize
      && !next.isBold
      && next.textColor === excelDefaultTextColor
      && next.fillColor === excelDefaultFillColor
      && next.borderStyle === excelDefaultBorderStyle;

    if (isDefault) {
      delete excelCellFormats[key];
      return;
    }

    excelCellFormats[key] = next;
  }

  function getExcelColumnWidth(col) {
    return clampNumber(excelColWidths[col], 60, 720, excelDefaultColWidth);
  }

  function setExcelColumnWidth(col, width) {
    excelColWidths[col] = clampNumber(width, 60, 720, excelDefaultColWidth);
  }

  function getExcelRowHeight(row) {
    return clampNumber(excelRowHeights[row], 28, 120, excelDefaultRowHeight);
  }

  function setExcelRowHeight(row, height) {
    excelRowHeights[row] = clampNumber(height, 28, 120, excelDefaultRowHeight);
  }

  function getExcelMeasureContext() {
    if (!excelMeasureCanvas) {
      excelMeasureCanvas = document.createElement("canvas");
    }

    return excelMeasureCanvas.getContext("2d");
  }

  function measureExcelTextWidth(text, row, col) {
    const context = getExcelMeasureContext();
    if (!context) {
      return 60;
    }

    const format = getExcelCellFormat(row, col);
    const fontSizePx = clampNumber(String(format.fontSize).replace(/[^0-9]/g, ""), 8, 72, 12);
    context.font = `${format.isBold ? "700" : "400"} ${fontSizePx}px ${format.fontFamily}, sans-serif`;

    const lines = String(text || "").split(/\r?\n/);
    let maxLineWidth = 0;
    lines.forEach((line) => {
      const measured = context.measureText(line || "").width;
      maxLineWidth = Math.max(maxLineWidth, measured);
    });

    return Math.ceil(maxLineWidth);
  }

  function calculateAutoFitColumnWidth(col) {
    let maxWidth = 0;

    for (let row = 0; row < excelRows; row += 1) {
      const raw = getExcelRaw(row, col);
      const display = raw.startsWith("=") ? getExcelDisplay(row, col) : raw;
      const text = String(display || raw || "");
      maxWidth = Math.max(maxWidth, measureExcelTextWidth(text, row, col));
    }

    // Extra spacing keeps content readable and avoids clipping at the right edge.
    return clampNumber(Math.ceil(maxWidth + 28), 60, 720, excelDefaultColWidth);
  }

  function autoFitColumn(col) {
    const width = calculateAutoFitColumnWidth(col);
    setExcelColumnWidth(col, width);
    return width;
  }

  function ensureColumnFitsCellContent(row, col) {
    const raw = getExcelRaw(row, col);
    const display = raw.startsWith("=") ? getExcelDisplay(row, col) : raw;
    const required = clampNumber(Math.ceil(measureExcelTextWidth(display || raw, row, col) + 28), 60, 720, excelDefaultColWidth);
    const current = getExcelColumnWidth(col);

    if (required <= current) {
      return false;
    }

    setExcelColumnWidth(col, required);
    return true;
  }

  function applyExcelCellVisualStyle(input, row, col) {
    if (!input) {
      return;
    }

    const format = getExcelCellFormat(row, col);
    const rowHeight = getExcelRowHeight(row);

    input.style.fontFamily = format.fontFamily;
    input.style.fontSize = format.fontSize;
    input.style.fontWeight = format.isBold ? "700" : "400";
    input.style.color = format.textColor;
    input.style.backgroundColor = format.fillColor;
    input.style.height = `${Math.max(24, rowHeight - 2)}px`;
    input.classList.toggle("has-border", format.borderStyle === "all");
  }

  function syncExcelFormattingControls() {
    const activeFormat = getExcelCellFormat(excelActiveRow, excelActiveCol);

    if (excelFontFamilySelect) {
      excelFontFamilySelect.value = activeFormat.fontFamily;
    }
    if (excelFontSizeSelect) {
      excelFontSizeSelect.value = activeFormat.fontSize;
    }
    if (excelBoldButton) {
      excelBoldButton.setAttribute("aria-pressed", String(activeFormat.isBold));
    }
    if (excelTextColorInput) {
      excelTextColorInput.value = activeFormat.textColor;
    }
    if (excelFillColorInput) {
      excelFillColorInput.value = activeFormat.fillColor;
    }
    if (excelBorderStyleSelect) {
      excelBorderStyleSelect.value = activeFormat.borderStyle;
    }
    if (excelColWidthInput) {
      excelColWidthInput.value = String(getExcelColumnWidth(excelActiveCol));
    }
    if (excelRowHeightInput) {
      excelRowHeightInput.value = String(getExcelRowHeight(excelActiveRow));
    }
  }

  function colToLabel(colIndex) {
    let value = colIndex + 1;
    let label = "";
    while (value > 0) {
      const mod = (value - 1) % 26;
      label = String.fromCharCode(65 + mod) + label;
      value = Math.floor((value - 1) / 26);
    }
    return label;
  }

  function labelToCol(label) {
    const normalized = String(label || "").toUpperCase();
    let total = 0;
    for (let i = 0; i < normalized.length; i += 1) {
      const code = normalized.charCodeAt(i);
      if (code < 65 || code > 90) {
        return -1;
      }
      total = total * 26 + (code - 64);
    }
    return total - 1;
  }

  function parseCellReference(cellRef) {
    const match = String(cellRef || "").toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!match) {
      return null;
    }

    const col = labelToCol(match[1]);
    const row = Number.parseInt(match[2], 10) - 1;
    if (col < 0 || row < 0) {
      return null;
    }

    return { row, col };
  }

  function getExcelRaw(row, col) {
    if (row < 0 || col < 0 || row >= excelRows || col >= excelCols) {
      return "";
    }
    return String(excelData[row]?.[col] ?? "");
  }

  function getExcelCellKey(row, col) {
    return `${row}:${col}`;
  }

  function getExcelNumericValueFromRef(cellRef, visited) {
    const parsed = parseCellReference(cellRef);
    if (!parsed) {
      return 0;
    }
    const value = evaluateExcelCell(parsed.row, parsed.col, visited);
    const number = Number.parseFloat(String(value).replace(",", "."));
    return Number.isFinite(number) ? number : 0;
  }

  function getExcelRangeRefs(rangeText) {
    const [startRef, endRef] = String(rangeText || "").split(":");
    const start = parseCellReference(startRef);
    const end = parseCellReference(endRef);
    if (!start || !end) {
      return [];
    }

    const rowStart = Math.min(start.row, end.row);
    const rowEnd = Math.max(start.row, end.row);
    const colStart = Math.min(start.col, end.col);
    const colEnd = Math.max(start.col, end.col);

    const refs = [];
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        refs.push(`${colToLabel(col)}${row + 1}`);
      }
    }
    return refs;
  }

  function getExcelFunctionValues(argsText, visited) {
    const tokens = String(argsText || "")
      .split(",")
      .map((token) => token.trim())
      .filter(Boolean);

    const values = [];
    tokens.forEach((token) => {
      if (token.includes(":")) {
        const refs = getExcelRangeRefs(token);
        refs.forEach((ref) => {
          values.push(getExcelNumericValueFromRef(ref, visited));
        });
        return;
      }

      if (/^[A-Z]+\d+$/i.test(token)) {
        values.push(getExcelNumericValueFromRef(token, visited));
        return;
      }

      const parsed = Number.parseFloat(token.replace(",", "."));
      values.push(Number.isFinite(parsed) ? parsed : 0);
    });

    return values;
  }

  function evaluateExcelFormula(formula, visited) {
    let expression = String(formula || "").toUpperCase();

    const functionPattern = /(SUM|AVERAGE|MIN|MAX)\(([^()]*)\)/g;
    let functionMatch = functionPattern.exec(expression);
    while (functionMatch) {
      const fnName = functionMatch[1];
      const values = getExcelFunctionValues(functionMatch[2], visited);
      let result = 0;
      if (fnName === "SUM") {
        result = values.reduce((acc, value) => acc + value, 0);
      } else if (fnName === "AVERAGE") {
        result = values.length === 0 ? 0 : values.reduce((acc, value) => acc + value, 0) / values.length;
      } else if (fnName === "MIN") {
        result = values.length === 0 ? 0 : Math.min(...values);
      } else if (fnName === "MAX") {
        result = values.length === 0 ? 0 : Math.max(...values);
      }

      expression = `${expression.slice(0, functionMatch.index)}${String(result)}${expression.slice(functionMatch.index + functionMatch[0].length)}`;
      functionPattern.lastIndex = 0;
      functionMatch = functionPattern.exec(expression);
    }

    expression = expression.replace(/\b([A-Z]+\d+)\b/g, (cellRef) => String(getExcelNumericValueFromRef(cellRef, visited)));

    if (!/^[0-9+\-*/().\s]+$/.test(expression)) {
      throw new Error("Ugyldig formel");
    }

    const result = Function(`"use strict"; return (${expression});`)();
    if (result === Infinity || result === -Infinity || Number.isNaN(result)) {
      throw new Error("Ugyldig beregning");
    }

    return result;
  }

  function evaluateExcelCell(row, col, visited = new Set()) {
    const key = getExcelCellKey(row, col);
    if (visited.has(key)) {
      throw new Error("Cirkulær reference");
    }

    const raw = getExcelRaw(row, col);
    if (!raw.startsWith("=")) {
      return raw;
    }

    const nextVisited = new Set(visited);
    nextVisited.add(key);
    const formula = raw.slice(1);
    const evaluated = evaluateExcelFormula(formula, nextVisited);
    if (typeof evaluated === "number") {
      return Number.isInteger(evaluated) ? String(evaluated) : String(Number(evaluated.toFixed(6)));
    }
    return String(evaluated);
  }

  function getExcelDisplay(row, col) {
    const raw = getExcelRaw(row, col);
    if (!raw.startsWith("=")) {
      return raw;
    }

    try {
      return evaluateExcelCell(row, col);
    } catch {
      return "#FEJL";
    }
  }

  function setExcelActiveCell(row, col) {
    excelActiveRow = Math.max(0, Math.min(excelRows - 1, row));
    excelActiveCol = Math.max(0, Math.min(excelCols - 1, col));
    if (excelActiveCell) {
      excelActiveCell.textContent = `Aktiv celle: ${colToLabel(excelActiveCol)}${excelActiveRow + 1}`;
    }
    if (excelFormulaInput) {
      excelFormulaInput.value = getExcelRaw(excelActiveRow, excelActiveCol);
    }

    syncExcelFormattingControls();
  }

  function refreshExcelGridDisplay() {
    if (!excelGrid) {
      return;
    }

    const inputs = excelGrid.querySelectorAll(".excel-cell-input");
    inputs.forEach((input) => {
      const row = Number.parseInt(input.getAttribute("data-row") || "0", 10);
      const col = Number.parseInt(input.getAttribute("data-col") || "0", 10);
      applyExcelCellVisualStyle(input, row, col);
      const isFocused = document.activeElement === input;
      if (!isFocused) {
        input.value = getExcelDisplay(row, col);
      }
    });
  }

  function renderExcelGrid() {
    if (!excelGrid) {
      return;
    }

    ensureExcelDataShape();
    const headerRow = document.createElement("tr");
    const corner = document.createElement("th");
    corner.className = "excel-row-head";
    corner.textContent = "#";
    headerRow.appendChild(corner);

    for (let col = 0; col < excelCols; col += 1) {
      const th = document.createElement("th");
      const colWidth = getExcelColumnWidth(col);
      th.scope = "col";
      th.textContent = colToLabel(col);
      th.style.width = `${colWidth}px`;
      th.style.minWidth = `${colWidth}px`;
      headerRow.appendChild(th);
    }

    const thead = document.createElement("thead");
    thead.appendChild(headerRow);

    const tbody = document.createElement("tbody");
    for (let row = 0; row < excelRows; row += 1) {
      const tr = document.createElement("tr");
      const rowHeight = getExcelRowHeight(row);
      tr.style.height = `${rowHeight}px`;
      const rowHead = document.createElement("th");
      rowHead.scope = "row";
      rowHead.className = "excel-row-head";
      rowHead.textContent = String(row + 1);
      tr.appendChild(rowHead);

      for (let col = 0; col < excelCols; col += 1) {
        const colWidth = getExcelColumnWidth(col);
        const td = document.createElement("td");
        td.style.width = `${colWidth}px`;
        td.style.minWidth = `${colWidth}px`;
        const input = document.createElement("input");
        input.type = "text";
        input.className = "excel-cell-input";
        input.setAttribute("data-row", String(row));
        input.setAttribute("data-col", String(col));
        input.setAttribute("aria-label", `Celle ${colToLabel(col)}${row + 1}`);
        input.style.minWidth = `${Math.max(60, colWidth - 4)}px`;
        input.value = getExcelDisplay(row, col);
        applyExcelCellVisualStyle(input, row, col);
        td.appendChild(input);
        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    }

    excelGrid.textContent = "";
    excelGrid.appendChild(thead);
    excelGrid.appendChild(tbody);
    setExcelActiveCell(excelActiveRow, excelActiveCol);
  }

  function openExcelPortal() {
    if (!excelView) {
      return;
    }

    if (!hasPortalPermission("portal.excel")) {
      setResourceStatus("Du har ikke adgang til Excel i portalen.");
      return;
    }

    hidePortalViews();
    excelView.classList.remove("is-hidden");
    setResourceStatus("Excel er åbnet i portalen.");
    setExcelStatus("Regneark klar.");

    if (!Array.isArray(excelData) || excelData.length === 0) {
      excelData = createExcelData(excelRows, excelCols);
    }

    renderExcelGrid();
    excelView.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function updateExcelCellFromInput(inputElement) {
    const row = Number.parseInt(inputElement.getAttribute("data-row") || "0", 10);
    const col = Number.parseInt(inputElement.getAttribute("data-col") || "0", 10);
    excelData[row][col] = inputElement.value;
    setExcelActiveCell(row, col);
  }

  function detectDelimitedTextDelimiter(text) {
    const sampleLines = String(text || "")
      .split(/\r?\n/)
      .slice(0, 5)
      .filter((line) => line.trim().length > 0);

    const delimiters = [",", ";", "\t"];
    let bestDelimiter = ",";
    let bestScore = -1;

    delimiters.forEach((delimiter) => {
      const score = sampleLines.reduce((total, line) => total + (line.split(delimiter).length - 1), 0);
      if (score > bestScore) {
        bestScore = score;
        bestDelimiter = delimiter;
      }
    });

    return bestDelimiter;
  }

  function parseCsv(text, delimiter = ",") {
    const rows = [];
    let current = "";
    let row = [];
    let inQuotes = false;

    for (let i = 0; i < text.length; i += 1) {
      const char = text[i];
      const next = text[i + 1];

      if (char === '"') {
        if (inQuotes && next === '"') {
          current += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === delimiter && !inQuotes) {
        row.push(current);
        current = "";
      } else if ((char === "\n" || char === "\r") && !inQuotes) {
        if (char === "\r" && next === "\n") {
          i += 1;
        }
        row.push(current);
        rows.push(row);
        row = [];
        current = "";
      } else {
        current += char;
      }
    }

    if (current.length > 0 || row.length > 0) {
      row.push(current);
      rows.push(row);
    }

    return rows;
  }

  function getExcelRowsFromWorksheet(workbook, worksheetName) {
    if (!hasXlsxRuntime()) {
      throw new Error("XLSX bibliotek er ikke indlæst.");
    }

    const sheetName = worksheetName || workbook?.SheetNames?.[0];
    if (!sheetName) {
      return [];
    }

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet || !worksheet["!ref"]) {
      return [];
    }

    const range = window.XLSX.utils.decode_range(worksheet["!ref"]);
    const rows = [];

    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
      const row = [];
      for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
        const address = window.XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        const cell = worksheet[address];

        if (!cell) {
          row.push("");
          continue;
        }

        if (cell.f) {
          row.push(`=${cell.f}`);
          continue;
        }

        const formatted = window.XLSX.utils.format_cell(cell);
        row.push(formatted == null ? "" : String(formatted));
      }

      rows.push(row);
    }

    return rows;
  }

  function applyExcelRows(parsedRows) {
    if (!Array.isArray(parsedRows) || parsedRows.length === 0) {
      setExcelStatus("Filen var tom.");
      return;
    }

    excelRows = Math.max(20, parsedRows.length);
    excelCols = Math.max(10, parsedRows.reduce((max, row) => Math.max(max, row.length), 0));
    excelData = createExcelData(excelRows, excelCols);
    excelCellFormats = {};
    excelColWidths = {};
    excelRowHeights = {};

    parsedRows.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        if (rowIndex < excelRows && colIndex < excelCols) {
          excelData[rowIndex][colIndex] = String(value ?? "");
        }
      });
    });

    excelActiveRow = 0;
    excelActiveCol = 0;
    renderExcelGrid();
  }

  function excelHexToArgb(hexColor, fallback = "#000000") {
    const normalized = normalizeHexColor(hexColor, fallback).slice(1).toUpperCase();
    return `FF${normalized}`;
  }

  function hasCustomExcelCellFormat(row, col) {
    return Object.prototype.hasOwnProperty.call(excelCellFormats, getExcelCellFormatKey(row, col));
  }

  function buildExcelCellStyle(row, col) {
    const format = getExcelCellFormat(row, col);
    const fontSizePx = clampNumber(String(format.fontSize).replace(/[^0-9]/g, ""), 8, 72, 12);

    const style = {
      font: {
        name: format.fontFamily,
        sz: fontSizePx,
        bold: format.isBold,
        color: { rgb: excelHexToArgb(format.textColor, excelDefaultTextColor) },
      },
      fill: {
        patternType: "solid",
        fgColor: { rgb: excelHexToArgb(format.fillColor, excelDefaultFillColor) },
        bgColor: { rgb: excelHexToArgb(format.fillColor, excelDefaultFillColor) },
      },
    };

    if (format.borderStyle === "all") {
      style.border = {
        top: { style: "thin", color: { rgb: "FF8A8A8A" } },
        right: { style: "thin", color: { rgb: "FF8A8A8A" } },
        bottom: { style: "thin", color: { rgb: "FF8A8A8A" } },
        left: { style: "thin", color: { rgb: "FF8A8A8A" } },
      };
    }

    return style;
  }

  function exportExcelWorkbookBinary() {
    if (!hasXlsxRuntime()) {
      throw new Error("XLSX bibliotek er ikke indlæst.");
    }

    ensureExcelDataShape();
    const worksheet = {};

    for (let row = 0; row < excelRows; row += 1) {
      for (let col = 0; col < excelCols; col += 1) {
        const raw = String(excelData[row]?.[col] ?? "");
        const hasCustomFormat = hasCustomExcelCellFormat(row, col);
        if (!raw && !hasCustomFormat) {
          continue;
        }

        const address = window.XLSX.utils.encode_cell({ r: row, c: col });
        const cell = {};

        if (raw.startsWith("=")) {
          cell.f = raw.slice(1);
        } else if (/^-?\d+(?:[.,]\d+)?$/.test(raw)) {
          cell.t = "n";
          cell.v = Number.parseFloat(raw.replace(",", "."));
        } else {
          cell.t = "s";
          cell.v = raw;
        }

        if (!raw) {
          cell.t = "s";
          cell.v = "";
        }

        if (hasCustomFormat) {
          cell.s = buildExcelCellStyle(row, col);
        }

        worksheet[address] = cell;
      }
    }

    worksheet["!cols"] = Array.from({ length: excelCols }, (_, col) => ({
      wpx: getExcelColumnWidth(col),
    }));

    worksheet["!rows"] = Array.from({ length: excelRows }, (_, row) => ({
      hpx: getExcelRowHeight(row),
    }));

    worksheet["!ref"] = window.XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(0, excelRows - 1), c: Math.max(0, excelCols - 1) },
    });

    const workbook = {
      SheetNames: ["Ark1"],
      Sheets: {
        Ark1: worksheet,
      },
    };

    return window.XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
      cellStyles: true,
    });
  }

  function splitRecipients(value) {
    return String(value || "")
      .split(/[;,]/)
      .map((item) => item.trim())
      .filter(Boolean);
  }

  function isValidEmail(address) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(address || "").trim());
  }

  function toGraphRecipients(addresses) {
    return addresses.map((address) => ({
      emailAddress: {
        address,
      },
    }));
  }

  function buildReplySubject(subject) {
    const normalized = String(subject || "").trim();
    if (!normalized) {
      return "Re: (intet emne)";
    }

    if (/^re:/i.test(normalized)) {
      return normalized;
    }

    return `Re: ${normalized}`;
  }

  function buildReplyBodyText(message) {
    const senderName = message?.from?.emailAddress?.name || "Ukendt afsender";
    const senderAddress = message?.from?.emailAddress?.address || "ukendt";
    const received = formatDate(message?.receivedDateTime);
    const originalPreview = String(message?.bodyPreview || "").trim();
    const quoted = originalPreview
      ? originalPreview.split("\n").map((line) => `> ${line}`).join("\n")
      : "> (ingen forhåndsvisning)";

    return [
      "",
      "",
      `----- Oprindelig meddelelse -----`,
      `Fra: ${senderName} <${senderAddress}>`,
      `Sendt: ${received}`,
      "",
      quoted,
    ].join("\n");
  }

  function prefillReplyForm(message) {
    if (!message || !mailToInput || !mailSubjectInput || !mailBodyInput) {
      return;
    }

    const senderAddress = message.from?.emailAddress?.address || "";
    mailToInput.value = senderAddress;
    if (mailCcInput) {
      mailCcInput.value = "";
    }
    if (mailBccInput) {
      mailBccInput.value = "";
    }

    mailSubjectInput.value = buildReplySubject(message.subject);
    if (mailContentTypeSelect) {
      mailContentTypeSelect.value = "Text";
    }
    updateMailPreviewVisibility();

    const prefix = "Hej,\n\n";
    mailBodyInput.value = `${prefix}${buildReplyBodyText(message)}`;
    mailBodyInput.focus();
    if (typeof mailBodyInput.setSelectionRange === "function") {
      mailBodyInput.setSelectionRange(prefix.length, prefix.length);
    }

    if (mailComposeTitle) {
      mailComposeTitle.scrollIntoView({ behavior: "smooth", block: "start" });
    } else if (mailForm) {
      mailForm.scrollIntoView({ behavior: "smooth", block: "start" });
    }

    setMailStatus("Svar er klargjort i formularen. Tilpas teksten og tryk Send mail.");
    setResourceStatus("Svarfunktion er åbnet for valgt mail.");
  }

  function isAccessDeniedMessage(message) {
    const text = String(message || "").toLowerCase();
    return text.includes("403")
      || text.includes("access denied")
      || text.includes("erroraccessdenied")
      || text.includes("invalid_grant")
      || text.includes("aadsts65001")
      || text.includes("consent");
  }

  function isConsentRequiredMessage(message) {
    const text = String(message || "").toLowerCase();
    return text.includes("invalid_grant")
      || text.includes("aadsts65001")
      || text.includes("consent");
  }

  function updateMailPreviewVisibility() {
    if (!mailHtmlPreviewWrapper || !mailContentTypeSelect || !mailPreviewHtmlButton) {
      return;
    }

    const isHtml = mailContentTypeSelect.value === "HTML";
    mailHtmlPreviewWrapper.classList.toggle("is-hidden", !isHtml);
    mailPreviewHtmlButton.disabled = !isHtml;
  }

  function clearMailInboxList() {
    if (mailInboxList) {
      mailInboxList.textContent = "";
    }
  }

  function clearMailFolderList() {
    if (mailFolderList) {
      mailFolderList.textContent = "";
    }
  }

  function sortMailFoldersByName(folders) {
    return [...folders].sort((a, b) => {
      const aName = String(a?.displayName || "");
      const bName = String(b?.displayName || "");
      return aName.localeCompare(bName, "da-DK", { sensitivity: "base" });
    });
  }

  function normalizeMailFolderName(name) {
    return String(name || "")
      .trim()
      .toLowerCase();
  }

  function shouldHideMailFolder(folder) {
    return hiddenMailFolderNames.has(normalizeMailFolderName(folder?.displayName));
  }

  function renderMailFolders(folders) {
    clearMailFolderList();

    if (!mailFolderList) {
      return;
    }

    const virtualAllFolder = {
      id: "__all__",
      displayName: "Alle mails",
      parentFolderId: null,
      totalItemCount: 0,
      unreadItemCount: 0,
    };

    const allFolders = [virtualAllFolder, ...(folders || [])].filter((folder) => !shouldHideMailFolder(folder));
    if (allFolders.length === 0) {
      const emptyItem = document.createElement("li");
      emptyItem.className = "mail-folder-item";
      emptyItem.textContent = "Ingen mapper fundet.";
      mailFolderList.appendChild(emptyItem);
      return;
    }

    const childrenByParent = new Map();
    const idSet = new Set(allFolders.map((folder) => String(folder.id || "")));

    allFolders.forEach((folder) => {
      const parentRaw = folder.parentFolderId == null ? "" : String(folder.parentFolderId);
      const parentKey = idSet.has(parentRaw) ? parentRaw : "";
      if (!childrenByParent.has(parentKey)) {
        childrenByParent.set(parentKey, []);
      }
      childrenByParent.get(parentKey).push(folder);
    });

    const appendFolder = (folder, depth) => {
      const item = document.createElement("li");
      item.className = "mail-folder-item";

      const button = document.createElement("button");
      button.type = "button";
      button.className = "mail-folder-button";
      if (folder.id === selectedMailFolderId) {
        button.classList.add("is-active");
      }

      const unread = Number(folder.unreadItemCount || 0);
      const total = Number(folder.totalItemCount || 0);
      const unreadText = unread > 0 ? ` (${unread} ulæste)` : "";
      button.textContent = `${folder.displayName || "Mappe"} - ${total}${unreadText}`;
      button.style.paddingLeft = `${12 + depth * 16}px`;
      button.addEventListener("click", () => {
        selectedMailFolderId = String(folder.id || "");
        selectedMailFolderName = folder.displayName || "Mappe";
        renderMailFolders(mailboxFolders);
        loadMessagesForSelectedFolder(false);
      });

      item.appendChild(button);
      mailFolderList.appendChild(item);

      const children = sortMailFoldersByName(childrenByParent.get(String(folder.id || "")) || []);
      children.forEach((child) => appendFolder(child, depth + 1));
    };

    const roots = sortMailFoldersByName(childrenByParent.get("") || []);
    roots.forEach((folder) => appendFolder(folder, 0));
  }

  async function loadMailFolders(interactiveAllowed) {
    setMailInboxStatus("Henter mailmapper...");

    try {
      let accessToken = await acquireGraphToken(interactiveAllowed, false, mailGraphScopes, {
        allowRedirect: !isLoopback,
        pendingAction: {
          action: "mail",
          payload: {},
        },
      });
      const allFolders = [];
      let nextUrl = "https://graph.microsoft.com/v1.0/me/mailFolders?$top=200&$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount";

      while (nextUrl) {
        let folderResponse;

        try {
          folderResponse = await fetchGraph(nextUrl, accessToken);
        } catch (error) {
          if (!interactiveAllowed || !isAccessDeniedMessage(error?.message)) {
            throw error;
          }

          accessToken = await acquireGraphToken(true, true);
          folderResponse = await fetchGraph(nextUrl, accessToken);
        }

        allFolders.push(...(folderResponse.value || []));
        nextUrl = folderResponse["@odata.nextLink"] || "";
      }

      mailboxFolders = allFolders;

      mailboxFolders = mailboxFolders.filter((folder) => !shouldHideMailFolder(folder));

      if (selectedMailFolderId === "__all__") {
        selectedMailFolderId = mailboxFolders.length > 0 ? String(mailboxFolders[0].id || "") : "";
        selectedMailFolderName = mailboxFolders[0]?.displayName || "Mappe";
      } else {
        const selected = mailboxFolders.find((folder) => folder.id === selectedMailFolderId);
        if (!selected) {
          selectedMailFolderId = mailboxFolders.length > 0 ? String(mailboxFolders[0].id || "") : "";
          selectedMailFolderName = mailboxFolders[0]?.displayName || "Mappe";
        }
      }

      renderMailFolders(mailboxFolders);
    } catch (error) {
      if (isGraphRedirectStartedError(error)) {
        setMailInboxStatus("Fortsætter login i samme fane. Du sendes tilbage til Mail automatisk.");
        return;
      }

      const message = error?.message || "Ukendt fejl";
      setMailInboxStatus(`Kunne ikke hente mapper: ${message}`);
      clearMailFolderList();
      throw error;
    }
  }

  async function loadMessagesForSelectedFolder(interactiveAllowed) {
    if (!mailView || !mailInboxList) {
      return;
    }

    if (!selectedMailFolderId) {
      inboxMessages = [];
      clearMailInboxList();
      setMailInboxStatus("Ingen synlige mapper at vise mails fra.");
      renderMailDetail(null);
      return;
    }

    const folderName = selectedMailFolderName || "valgt mappe";
    setMailInboxStatus(`Henter mails fra ${folderName}...`);
    clearMailInboxList();

    const baseUrl = selectedMailFolderId === "__all__"
      ? "https://graph.microsoft.com/v1.0/me/messages"
      : `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(selectedMailFolderId)}/messages`;

    let nextUrl = `${baseUrl}?$top=200&$orderby=receivedDateTime%20DESC&$select=id,subject,from,receivedDateTime,bodyPreview,isRead`;
    const loadedMessages = [];

    try {
      let accessToken = await acquireGraphToken(interactiveAllowed, false, mailGraphScopes);

      while (nextUrl) {
        let messagesResponse;

        try {
          messagesResponse = await fetchGraph(nextUrl, accessToken);
        } catch (error) {
          if (!interactiveAllowed || !isAccessDeniedMessage(error?.message)) {
            throw error;
          }

          accessToken = await acquireGraphToken(true, true, mailGraphScopes);
          messagesResponse = await fetchGraph(nextUrl, accessToken);
        }

        loadedMessages.push(...(messagesResponse.value || []));
        nextUrl = messagesResponse["@odata.nextLink"] || "";
      }

      inboxMessages = loadedMessages;
      setMailInboxStatus(`Viser ${inboxMessages.length} mails fra ${folderName}.`);
      renderInboxMessages(inboxMessages);
    } catch (error) {
      const message = error?.message || "Ukendt fejl";
      setMailInboxStatus(`Kunne ikke hente mails: ${message}`);
      renderMailDetail(null);
    }
  }

  function renderMailDetail(message) {
    if (!mailInboxDetail) {
      return;
    }

    mailInboxDetail.textContent = "";

    if (!message) {
      const heading = document.createElement("h4");
      heading.textContent = "Vælg en mail";
      const info = document.createElement("p");
      info.textContent = "Vælg en mail fra listen for at se detaljer.";
      mailInboxDetail.appendChild(heading);
      mailInboxDetail.appendChild(info);
      announceMailForScreenReader("Ingen mail valgt.");
      return;
    }

    const heading = document.createElement("h4");
    heading.textContent = message.subject || "(intet emne)";

    const from = document.createElement("p");
    from.className = "mail-detail-meta";
    const senderName = message.from?.emailAddress?.name || "Ukendt";
    const senderAddress = message.from?.emailAddress?.address || "ukendt";
    from.textContent = `Fra: ${senderName} <${senderAddress}>`;

    const received = document.createElement("p");
    received.className = "mail-detail-meta";
    received.textContent = `Modtaget: ${formatDate(message.receivedDateTime)}`;

    const textBody = document.createElement("div");
    textBody.className = "mail-detail-text";
    textBody.id = "mail-detail-text";
    textBody.tabIndex = 0;
    textBody.setAttribute("aria-label", "Mailtekst");
    textBody.setAttribute("aria-readonly", "true");
    textBody.textContent = message._detailBodyText || message.bodyPreview || "Ingen indhold tilgængelig.";

    const actions = document.createElement("div");
    actions.className = "mail-detail-actions";

    const replyButton = document.createElement("button");
    replyButton.type = "button";
    replyButton.className = "secondary-button";
    replyButton.textContent = "Besvar i formular";
    replyButton.addEventListener("click", () => {
      prefillReplyForm(message);
    });
    actions.appendChild(replyButton);

    const quickReplyButton = document.createElement("button");
    quickReplyButton.type = "button";
    quickReplyButton.className = "primary-button";
    quickReplyButton.textContent = "Et-klik svar";
    quickReplyButton.addEventListener("click", async () => {
      quickReplyButton.disabled = true;
      setMailStatus("Sender et-klik svar...");

      try {
        await sendQuickReply(message, true);
        setMailStatus("Et-klik svar er sendt.");
        setResourceStatus("Svar blev sendt direkte fra indbakken.");
        await loadMessagesForSelectedFolder(false);
      } catch (error) {
        const messageText = error?.message || "Ukendt fejl";
        setMailStatus(`Et-klik svar fejlede: ${messageText}`);
        setResourceStatus("Kunne ikke sende et-klik svar.");
      } finally {
        quickReplyButton.disabled = false;
      }
    });
    actions.appendChild(quickReplyButton);

    mailInboxDetail.appendChild(heading);
    mailInboxDetail.appendChild(from);
    mailInboxDetail.appendChild(received);
    mailInboxDetail.appendChild(textBody);
    mailInboxDetail.appendChild(actions);

    const summaryForReader = [
      `Emne: ${message.subject || "intet emne"}.`,
      `Fra: ${senderName}.`,
      `Modtaget: ${formatDate(message.receivedDateTime)}.`,
      `Indhold: ${message.bodyPreview || "Ingen forhåndsvisning."}`,
    ].join(" ");
    announceMailForScreenReader(summaryForReader);
  }


  function setCalendarStatus(message) {
    if (mailCalendarStatus) {
      mailCalendarStatus.textContent = message;
    }
  }

  function getCalendarStartOfMonth(date) {
    const source = date instanceof Date ? date : new Date(date || Date.now());
    return new Date(source.getFullYear(), source.getMonth(), 1);
  }

  function getCalendarStartOfWeek(date) {
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const dayIndex = (start.getDay() + 6) % 7;
    start.setDate(start.getDate() - dayIndex);
    start.setHours(0, 0, 0, 0);
    return start;
  }

  function getCalendarEndOfWeek(date) {
    const end = getCalendarStartOfWeek(date);
    end.setDate(end.getDate() + 6);
    end.setHours(23, 59, 59, 999);
    return end;
  }

  function getCalendarDateKey(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function getCalendarDayLabel(date) {
    return date.toLocaleDateString("da-DK", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
  }

  function getCalendarMonthLabel(date) {
    return date.toLocaleDateString("da-DK", { month: "long", year: "numeric" });
  }

  function getCalendarTimeLabel(date) {
    return date.toLocaleTimeString("da-DK", { hour: "2-digit", minute: "2-digit" });
  }

  function getCalendarRangeDates(anchorDate, includeWeekends = true) {
    const start = getCalendarStartOfWeek(anchorDate);
    const end = new Date(start);
    end.setDate(end.getDate() + (includeWeekends ? 6 : 4));

    const dates = [];
    for (let current = new Date(start); current <= end; current.setDate(current.getDate() + 1)) {
      dates.push(new Date(current));
    }

    return dates;
  }

  function getCalendarEventsInRange(startDate, endDate) {
    return calendarEvents
      .filter((event) => event.start < endDate && event.end > startDate)
      .sort((left, right) => left.start - right.start);
  }

  function setCalendarDraftFromDate(date, sourceEvent = null) {
    const resolvedDate = date instanceof Date ? new Date(date) : new Date(date || Date.now());
    calendarDraft.date = resolvedDate;
    calendarDraft.subject = String(sourceEvent?.subject || calendarDraft.subject || "").trim();
    calendarDraft.location = String(sourceEvent?.location || calendarDraft.location || "").trim();

    if (!mailCalendarDraftStart?.value || mailCalendarDraftStart.value === "09:00") {
      calendarDraft.start = sourceEvent?.start ? getCalendarTimeLabel(sourceEvent.start) : "09:00";
    }

    if (!mailCalendarDraftEnd?.value || mailCalendarDraftEnd.value === "10:00") {
      calendarDraft.end = sourceEvent?.end ? getCalendarTimeLabel(sourceEvent.end) : "10:00";
    }

    if (mailCalendarDraftDate) {
      mailCalendarDraftDate.value = getCalendarDateKey(resolvedDate);
    }
    if (mailCalendarDraftSubject) {
      mailCalendarDraftSubject.value = calendarDraft.subject;
    }
    if (mailCalendarDraftStart) {
      mailCalendarDraftStart.value = sourceEvent?.start ? sourceEvent.start.toTimeString().slice(0, 5) : calendarDraft.start;
    }
    if (mailCalendarDraftEnd) {
      mailCalendarDraftEnd.value = sourceEvent?.end ? sourceEvent.end.toTimeString().slice(0, 5) : calendarDraft.end;
    }
    if (mailCalendarDraftLocation) {
      mailCalendarDraftLocation.value = calendarDraft.location;
    }
  }

  function readCalendarDraft() {
    const dateValue = mailCalendarDraftDate?.value || getCalendarDateKey(calendarDraft.date);
    const startValue = mailCalendarDraftStart?.value || calendarDraft.start || "09:00";
    const endValue = mailCalendarDraftEnd?.value || calendarDraft.end || "10:00";
    const subject = String(mailCalendarDraftSubject?.value || calendarDraft.subject || "").trim() || "Nyt møde";
    const location = String(mailCalendarDraftLocation?.value || calendarDraft.location || "").trim();
    const date = new Date(`${dateValue}T00:00:00`);

    return {
      subject,
      location,
      start: new Date(`${dateValue}T${startValue}:00`),
      end: new Date(`${dateValue}T${endValue}:00`),
      date,
    };
  }

  function formatCalendarIcsDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    const seconds = String(date.getSeconds()).padStart(2, "0");
    return `${year}${month}${day}T${hours}${minutes}${seconds}`;
  }

  function buildCalendarIcsContent(draft) {
    const uid = `dbs-portal-${Date.now()}@blind.dk`;
    const lines = [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//DBS Portal//Kalenderkladde//DA",
      "CALSCALE:GREGORIAN",
      "METHOD:PUBLISH",
      "BEGIN:VEVENT",
      `UID:${uid}`,
      `DTSTAMP:${formatCalendarIcsDate(new Date())}`,
      `DTSTART:${formatCalendarIcsDate(draft.start)}`,
      `DTEND:${formatCalendarIcsDate(draft.end)}`,
      `SUMMARY:${String(draft.subject || "Nyt møde").replace(/\n/g, " ")}`,
    ];

    if (draft.location) {
      lines.push(`LOCATION:${String(draft.location).replace(/\n/g, " ")}`);
    }

    lines.push(
      "DESCRIPTION:Oprettet i DBS Portal",
      "END:VEVENT",
      "END:VCALENDAR",
    );

    return lines.join("\r\n");
  }

  function downloadCalendarDraftIcs() {
    const draft = readCalendarDraft();
    const icsContent = buildCalendarIcsContent(draft);
    const blob = new Blob([icsContent], { type: "text/calendar;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `dbs-mode-${getCalendarDateKey(draft.date)}.ics`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function copyCalendarDraftText() {
    const draft = readCalendarDraft();
    const text = [
      `Emne: ${draft.subject}`,
      `Dato: ${getCalendarDayLabel(draft.date)}`,
      `Tid: ${draft.start.toLocaleTimeString("da-DK", { hour: "2-digit", minute: "2-digit" })} - ${draft.end.toLocaleTimeString("da-DK", { hour: "2-digit", minute: "2-digit" })}`,
      draft.location ? `Sted: ${draft.location}` : "",
    ].filter(Boolean).join("\n");

    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(text).catch(() => {});
    }

    return text;
  }

  function setCalendarViewMode(nextMode) {
    const allowedModes = new Set(["month", "week", "workweek"]);
    calendarViewMode = allowedModes.has(nextMode) ? nextMode : "month";

    [
      { button: mailCalendarModeMonthButton, mode: "month" },
      { button: mailCalendarModeWeekButton, mode: "week" },
      { button: mailCalendarModeWorkweekButton, mode: "workweek" },
    ].forEach(({ button, mode }) => {
      if (!button) {
        return;
      }

      const isActive = calendarViewMode === mode;
      button.classList.toggle("is-active", isActive);
      button.setAttribute("aria-selected", isActive ? "true" : "false");
    });

    renderCalendarGrid();
    renderCalendarAgenda(calendarSelectedDate);
  }

  function openCalendarDraftForDate(date, sourceEvent = null) {
    calendarSelectedDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    setCalendarDraftFromDate(calendarSelectedDate, sourceEvent);
    renderCalendarGrid();
    renderCalendarAgenda(calendarSelectedDate);
    mailCalendarDraftSubject?.focus();
  }

  function parseCalendarDateTime(rawValue) {
    const value = rawValue?.dateTime || rawValue || null;
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function normalizeCalendarEvent(rawEvent) {
    const start = parseCalendarDateTime(rawEvent?.start);
    const end = parseCalendarDateTime(rawEvent?.end);

    if (!start || !end) {
      return null;
    }

    return {
      id: String(rawEvent?.id || "").trim(),
      subject: String(rawEvent?.subject || "(intet emne)").trim(),
      start,
      end,
      location: String(rawEvent?.location?.displayName || rawEvent?.location || "").trim(),
      organizer: String(rawEvent?.organizer?.emailAddress?.name || rawEvent?.organizer?.emailAddress?.address || "").trim(),
      isAllDay: Boolean(rawEvent?.isAllDay),
      webLink: String(rawEvent?.webLink || "").trim(),
    };
  }

  function getCalendarEventsForDate(date) {
    const startOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const endOfDay = new Date(startOfDay);
    endOfDay.setDate(endOfDay.getDate() + 1);

    return calendarEvents
      .filter((event) => event.start < endOfDay && event.end > startOfDay)
      .sort((left, right) => left.start - right.start);
  }

  function clearCalendarPanel() {
    if (mailCalendarGrid) {
      mailCalendarGrid.textContent = "";
    }

    if (mailCalendarAgenda) {
      mailCalendarAgenda.textContent = "";
    }
  }

  function renderCalendarAgenda(date) {
    if (!mailCalendarAgenda) {
      return;
    }

    mailCalendarAgenda.textContent = "";
    const heading = document.createElement("h4");
    heading.textContent = calendarViewMode === "month"
      ? `Agenda for ${getCalendarDayLabel(date)}`
      : calendarViewMode === "workweek"
        ? `Arbejdsuge for ${getCalendarDayLabel(getCalendarStartOfWeek(date))}`
        : `Uge for ${getCalendarDayLabel(getCalendarStartOfWeek(date))}`;
    mailCalendarAgenda.appendChild(heading);

    const includeWeekends = calendarViewMode !== "workweek";
    const rangeDates = getCalendarRangeDates(date, includeWeekends);
    const rangeStart = rangeDates[0];
    const rangeEnd = new Date(rangeDates[rangeDates.length - 1]);
    rangeEnd.setHours(23, 59, 59, 999);
    const events = calendarViewMode === "month"
      ? getCalendarEventsForDate(date)
      : getCalendarEventsInRange(rangeStart, rangeEnd);

    if (calendarViewMode === "month") {
      if (events.length === 0) {
        const empty = document.createElement("p");
        empty.className = "mail-calendar-empty";
        empty.textContent = "Ingen møder eller aftaler på denne dag.";
        mailCalendarAgenda.appendChild(empty);
        return;
      }

      const list = document.createElement("ul");
      list.className = "mail-calendar-agenda-list";

      events.forEach((event) => {
        const item = document.createElement("li");
        item.className = "mail-calendar-agenda-item";
        item.draggable = true;
        item.addEventListener("dragstart", () => {
          calendarDraggedEventId = event.id;
        });

        const topRow = document.createElement("div");
        topRow.className = "mail-calendar-agenda-top";

        const title = document.createElement("strong");
        title.textContent = event.subject;

        const time = document.createElement("span");
        time.className = "mail-calendar-agenda-time";
        time.textContent = event.isAllDay
          ? "Hele dagen"
          : `${getCalendarTimeLabel(event.start)} - ${getCalendarTimeLabel(event.end)}`;

        topRow.appendChild(title);
        topRow.appendChild(time);
        item.appendChild(topRow);

        const meta = document.createElement("p");
        meta.className = "mail-calendar-agenda-meta";
        meta.textContent = [event.location, event.organizer].filter(Boolean).join(" | ") || "Ingen yderligere detaljer.";
        item.appendChild(meta);

        const actionRow = document.createElement("div");
        actionRow.className = "mail-calendar-agenda-actions";

        if (event.webLink) {
          const link = document.createElement("a");
          link.href = event.webLink;
          link.target = "_blank";
          link.rel = "noopener noreferrer";
          link.textContent = "Åbn i Outlook";
          link.className = "mail-calendar-link";
          actionRow.appendChild(link);
        }

        const createButton = document.createElement("button");
        createButton.type = "button";
        createButton.className = "secondary-button mail-calendar-create-button";
        createButton.textContent = "Opret møde ud fra dag";
        createButton.addEventListener("click", () => openCalendarDraftForDate(date, event));
        actionRow.appendChild(createButton);

        item.appendChild(actionRow);
        list.appendChild(item);
      });

      mailCalendarAgenda.appendChild(list);
      return;
    }

    const list = document.createElement("div");
    list.className = "mail-calendar-week-list";

    rangeDates.forEach((day) => {
      const dayEvents = getCalendarEventsForDate(day);
      const section = document.createElement("section");
      section.className = "mail-calendar-week-day";

      const sectionHeader = document.createElement("button");
      sectionHeader.type = "button";
      sectionHeader.className = "mail-calendar-week-day-header";
      sectionHeader.textContent = getCalendarDayLabel(day);
      sectionHeader.addEventListener("click", () => openCalendarDraftForDate(day));

      section.addEventListener("dragover", (event) => event.preventDefault());
      section.addEventListener("drop", (event) => {
        event.preventDefault();
        const draggedEvent = calendarEvents.find((entry) => entry.id === calendarDraggedEventId) || null;
        openCalendarDraftForDate(day, draggedEvent);
        calendarDraggedEventId = "";
      });

      const eventList = document.createElement("div");
      eventList.className = "mail-calendar-week-events";

      if (dayEvents.length === 0) {
        const empty = document.createElement("p");
        empty.className = "mail-calendar-empty";
        empty.textContent = "Ingen aftaler.";
        eventList.appendChild(empty);
      } else {
        dayEvents.forEach((event) => {
          const item = document.createElement("button");
          item.type = "button";
          item.className = "mail-calendar-week-event";
          item.draggable = true;
          item.addEventListener("dragstart", () => {
            calendarDraggedEventId = event.id;
          });
          item.addEventListener("click", () => openCalendarDraftForDate(day, event));

          const time = document.createElement("span");
          time.className = "mail-calendar-week-event-time";
          time.textContent = event.isAllDay ? "Hele dagen" : `${getCalendarTimeLabel(event.start)} - ${getCalendarTimeLabel(event.end)}`;

          const title = document.createElement("span");
          title.className = "mail-calendar-week-event-title";
          title.textContent = event.subject;

          item.appendChild(time);
          item.appendChild(title);
          eventList.appendChild(item);
        });
      }

      const createButton = document.createElement("button");
      createButton.type = "button";
      createButton.className = "secondary-button mail-calendar-create-button";
      createButton.textContent = "Opret møde";
      createButton.addEventListener("click", () => openCalendarDraftForDate(day));

      section.appendChild(sectionHeader);
      section.appendChild(eventList);
      section.appendChild(createButton);
      list.appendChild(section);
    });

    mailCalendarAgenda.appendChild(list);
  }

  function renderCalendarGrid() {
    if (!mailCalendarGrid || !mailCalendarMonthLabel) {
      return;
    }

    mailCalendarGrid.textContent = "";
    mailCalendarMonthLabel.textContent = getCalendarMonthLabel(calendarVisibleMonth);

    const rangeStart = getCalendarStartOfWeek(getCalendarStartOfMonth(calendarVisibleMonth));
    const rangeEnd = getCalendarEndOfWeek(new Date(calendarVisibleMonth.getFullYear(), calendarVisibleMonth.getMonth() + 1, 0));
    const todayKey = getCalendarDateKey(new Date());
    const selectedKey = getCalendarDateKey(calendarSelectedDate);

    for (let current = new Date(rangeStart); current <= rangeEnd; current.setDate(current.getDate() + 1)) {
      const dayDate = new Date(current);
      const dayEvents = getCalendarEventsForDate(dayDate);
      const dayKey = getCalendarDateKey(dayDate);
      const cell = document.createElement("button");
      cell.type = "button";
      cell.className = "mail-calendar-day";
      cell.draggable = true;
      if (dayDate.getMonth() !== calendarVisibleMonth.getMonth()) {
        cell.classList.add("is-outside");
      }
      if (dayKey === todayKey) {
        cell.classList.add("is-today");
      }
      if (dayKey === selectedKey) {
        cell.classList.add("is-selected");
      }

      const dayNumber = document.createElement("span");
      dayNumber.className = "mail-calendar-day-number";
      dayNumber.textContent = String(dayDate.getDate());

      const eventCount = document.createElement("span");
      eventCount.className = "mail-calendar-day-count";
      eventCount.textContent = dayEvents.length > 0 ? `${dayEvents.length} møde${dayEvents.length === 1 ? "" : "r"}` : "";

      const chips = document.createElement("div");
      chips.className = "mail-calendar-day-chips";
      dayEvents.slice(0, 2).forEach((event) => {
        const chip = document.createElement("span");
        chip.className = "mail-calendar-chip";
        chip.textContent = event.subject;
        chips.appendChild(chip);
      });

      const createButton = document.createElement("button");
      createButton.type = "button";
      createButton.className = "mail-calendar-day-create";
      createButton.textContent = "+ Møde";
      createButton.addEventListener("click", (event) => {
        event.stopPropagation();
        openCalendarDraftForDate(dayDate);
      });

      cell.appendChild(dayNumber);
      cell.appendChild(eventCount);
      cell.appendChild(chips);
      cell.appendChild(createButton);
      cell.addEventListener("click", () => {
        calendarSelectedDate = dayDate;
        setCalendarDraftFromDate(dayDate);
        renderCalendarGrid();
        renderCalendarAgenda(calendarSelectedDate);
      });
      cell.addEventListener("dragover", (event) => event.preventDefault());
      cell.addEventListener("drop", (event) => {
        event.preventDefault();
        const draggedEvent = calendarEvents.find((entry) => entry.id === calendarDraggedEventId) || null;
        openCalendarDraftForDate(dayDate, draggedEvent);
        calendarDraggedEventId = "";
      });

      mailCalendarGrid.appendChild(cell);
    }

    if (mailCalendarStatus && !calendarLoading) {
      const modeLabel = calendarViewMode === "month" ? "måned" : calendarViewMode === "week" ? "uge" : "arbejdsuge";
      mailCalendarStatus.textContent = `Viser ${calendarEvents.length} kalenderaftaler i ${getCalendarMonthLabel(calendarVisibleMonth)} (${modeLabel}).`;
    }
  }

  async function loadCalendarEvents(interactiveAllowed) {
    if (!mailCalendarGrid || !mailCalendarAgenda) {
      return;
    }

    if (calendarLoading) {
      return;
    }

    calendarLoading = true;
    setCalendarStatus("Henter kalender...");
    clearCalendarPanel();
    renderCalendarGrid();

    const rangeStart = getCalendarStartOfWeek(getCalendarStartOfMonth(calendarVisibleMonth));
    const rangeEnd = getCalendarEndOfWeek(new Date(calendarVisibleMonth.getFullYear(), calendarVisibleMonth.getMonth() + 1, 0));
    const startDateTime = rangeStart.toISOString();
    const endDateTime = rangeEnd.toISOString();
    const queryUrl = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}&$orderby=start/dateTime&$select=id,subject,start,end,organizer,location,isAllDay,webLink`;

    try {
      let accessToken = await acquireGraphToken(interactiveAllowed, false, calendarReadScopes);
      let nextUrl = queryUrl;
      const loadedEvents = [];

      while (nextUrl) {
        let response;
        try {
          response = await fetchGraph(nextUrl, accessToken);
        } catch (error) {
          if (!interactiveAllowed || isConsentRequiredMessage(error?.message) || !isAccessDeniedMessage(error?.message)) {
            throw error;
          }

          accessToken = await acquireGraphToken(true, false, calendarReadScopes);
          response = await fetchGraph(nextUrl, accessToken);
        }

        loadedEvents.push(...(response.value || []));
        nextUrl = response["@odata.nextLink"] || "";
      }

      calendarEvents = loadedEvents
        .map((event) => normalizeCalendarEvent(event))
        .filter(Boolean);

      setCalendarStatus(`Viser ${calendarEvents.length} aftaler i ${getCalendarMonthLabel(calendarVisibleMonth)}.`);
      renderCalendarGrid();
      renderCalendarAgenda(calendarSelectedDate);
    } catch (error) {
      const message = error?.message || "Ukendt fejl";
      if (isConsentRequiredMessage(message)) {
        setCalendarStatus("Kalenderdata kræver admin-godkendelse i blind.dk. Kalenderen kan stadig bruges lokalt (dag/uge/måned), men Outlook-aftaler kan ikke hentes endnu.");
      } else {
        setCalendarStatus(`Kunne ikke hente kalender: ${message}`);
      }
      calendarEvents = [];
      renderCalendarGrid();
      renderCalendarAgenda(calendarSelectedDate);
    } finally {
      calendarLoading = false;
    }
  }

  async function moveCalendarMonth(deltaMonths) {
    calendarVisibleMonth = new Date(calendarVisibleMonth.getFullYear(), calendarVisibleMonth.getMonth() + deltaMonths, 1);
    calendarSelectedDate = new Date(calendarVisibleMonth.getFullYear(), calendarVisibleMonth.getMonth(), 1);
    await loadCalendarEvents(true);
  }

  function renderInboxMessages(messages) {
    clearMailInboxList();

    if (!mailInboxList) {
      return;
    }

    if (!messages || messages.length === 0) {
      const empty = document.createElement("li");
      empty.className = "mail-inbox-item";
      empty.textContent = "Ingen mails fundet i den valgte mappe.";
      mailInboxList.appendChild(empty);
      renderMailDetail(null);
      return;
    }

    messages.forEach((message) => {
      const item = document.createElement("li");
      item.className = "mail-inbox-item";

      const button = document.createElement("button");
      button.type = "button";
      button.className = "mail-inbox-button";
      if (!message.isRead) {
        button.classList.add("is-unread");
      }
      if (message.id === selectedInboxMessageId) {
        button.classList.add("is-active");
      }

      const subject = document.createElement("p");
      subject.className = "mail-inbox-subject";
      subject.textContent = message.subject || "(intet emne)";

      const stateBadge = document.createElement("span");
      stateBadge.className = `mail-read-badge ${message.isRead ? "is-read" : "is-unread"}`;
      stateBadge.textContent = message.isRead ? "Læst" : "Ulæst";

      const heading = document.createElement("div");
      heading.className = "mail-inbox-heading";
      heading.appendChild(subject);
      heading.appendChild(stateBadge);

      const senderName = message.from?.emailAddress?.name || "Ukendt";
      const meta = document.createElement("p");
      meta.className = "mail-inbox-meta";
      meta.textContent = `${senderName} | ${formatDate(message.receivedDateTime)}`;

      button.appendChild(heading);
      button.appendChild(meta);
      button.addEventListener("click", () => {
        selectedInboxMessageId = message.id || "";
        focusReadPaneAfterLoad = false;
        renderInboxMessages(inboxMessages);
      });
      button.addEventListener("dblclick", () => {
        selectedInboxMessageId = message.id || "";
        focusReadPaneAfterLoad = true;
        markMessageAsReadAfterLoadId = message.id || "";
        message.isRead = true;
        renderInboxMessages(inboxMessages);
      });

      item.appendChild(button);
      mailInboxList.appendChild(item);
    });

    const activeMessage = messages.find((message) => message.id === selectedInboxMessageId) || messages[0];
    selectedInboxMessageId = activeMessage.id || "";
    const moveFocusToReadPane = focusReadPaneAfterLoad;
    const shouldMarkAsRead = activeMessage.id === markMessageAsReadAfterLoadId;
    focusReadPaneAfterLoad = false;
    markMessageAsReadAfterLoadId = "";
    openMessageInDetail(activeMessage, false, moveFocusToReadPane, shouldMarkAsRead);
  }

  function arrayBufferToBase64(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    const chunkSize = 0x8000;
    let binary = "";

    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode(...chunk);
    }

    return window.btoa(binary);
  }

  async function buildGraphAttachments(fileList) {
    const files = Array.from(fileList || []);
    if (files.length === 0) {
      return [];
    }

    const maxFileSizeBytes = 3 * 1024 * 1024;
    for (const file of files) {
      if (file.size > maxFileSizeBytes) {
        throw new Error(`Filen ${file.name} er for stor. Maks størrelse er 3 MB per fil.`);
      }
    }

    const attachments = await Promise.all(files.map(async (file) => {
      const arrayBuffer = await file.arrayBuffer();
      return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: file.name,
        contentType: file.type || "application/octet-stream",
        contentBytes: arrayBufferToBase64(arrayBuffer),
      };
    }));

    return attachments;
  }

  async function openMailPortal() {
    if (!mailView) {
      return;
    }

    if (!hasPortalPermission("portal.mail")) {
      setResourceStatus("Du har ikke adgang til Mail i portalen.");
      return;
    }

    hidePortalViews();
    mailView.classList.remove("is-hidden");
    setMailStatus("Udfyld mailfelterne og tryk Send mail.");
    setResourceStatus("Mailformular er åbnet i portalen.");
    updateMailPreviewVisibility();

    // Initialize calendar UI immediately so mode/day controls are usable
    // even while Graph data is still loading.
    setCalendarViewMode(calendarViewMode);
    setCalendarDraftFromDate(calendarSelectedDate, null);
    setCalendarStatus("Initialiserer kalender...");

    try {
      await loadMailFolders(true);
      await loadMessagesForSelectedFolder(false);
      void loadCalendarEvents(false);
      setCalendarDraftFromDate(calendarSelectedDate, null);
    } catch {
      // Statustekster sættes i de underliggende funktioner.
    }
    mailView.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function openCasesPortal() {
    if (!casesView) {
      return;
    }

    if (!hasPortalPermission("portal.cases.view")) {
      setResourceStatus("Du har ikke adgang til sagsvisning.");
      return;
    }

    hidePortalViews();
    casesView.classList.remove("is-hidden");
    setResourceStatus("Sagsvisning er åbnet i portalen.");
    await loadCases();
    startCasesAutoRefresh();
    casesView.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function createFormField(field) {
    const fieldWrapper = document.createElement("div");
    fieldWrapper.className = "portal-form-field";

    const inputId = `portal-field-${field.name}`;
    const label = document.createElement("label");
    label.setAttribute("for", inputId);
    label.textContent = field.required ? `${field.label} *` : field.label;
    fieldWrapper.appendChild(label);

    let input;
    if (field.type === "textarea") {
      input = document.createElement("textarea");
    } else if (field.type === "select") {
      input = document.createElement("select");
      (field.options || []).forEach((option) => {
        const optionElement = document.createElement("option");
        optionElement.value = option.value;
        optionElement.textContent = option.text;
        input.appendChild(optionElement);
      });
    } else {
      input = document.createElement("input");
      input.type = field.type || "text";
    }

    input.id = inputId;
    input.name = field.name;
    input.required = Boolean(field.required);
    input.autocomplete = "on";
    fieldWrapper.appendChild(input);

    return fieldWrapper;
  }

  function setFormsStatus(message) {
    if (formsStatus) {
      formsStatus.textContent = message;
    }
  }

  function renderPortalForm(formType) {
    if (!portalForm || !formsView || !formsTitle || !formsDescription) {
      return;
    }

    if (!hasPortalPermission("portal.forms")) {
      setResourceStatus("Du har ikke adgang til formularer.");
      return;
    }

    const formDefinition = formDefinitions[formType];
    if (!formDefinition) {
      return;
    }

    hidePortalViews();
    setResourceStatus(`${formDefinition.title} er åbnet i portalen.`);
    formsView.classList.remove("is-hidden");
    formsTitle.textContent = formDefinition.title;
    formsDescription.textContent = formDefinition.description;
    setFormsStatus("Udfyld felterne og send formularen.");

    portalForm.textContent = "";

    formDefinition.fields.forEach((field) => {
      portalForm.appendChild(createFormField(field));
    });

    const actions = document.createElement("div");
    actions.className = "portal-form-actions";

    const submitButton = document.createElement("button");
    submitButton.type = "submit";
    submitButton.className = "primary-button";
    submitButton.textContent = "Send formular";

    const clearButton = document.createElement("button");
    clearButton.type = "reset";
    clearButton.className = "secondary-button";
    clearButton.textContent = "Ryd felter";

    actions.appendChild(submitButton);
    actions.appendChild(clearButton);
    portalForm.appendChild(actions);

    activeFormType = formType;
    portalForm.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function initPortalInteractions() {
    function closeAll(except) {
      dropdowns.forEach((dd) => {
        if (dd !== except) {
          dd.classList.remove("open");
          dd.querySelector(".dropbtn").setAttribute("aria-expanded", "false");
        }
      });
    }

    dropdowns.forEach((dropdown) => {
      const button = dropdown.querySelector(".dropbtn");

      button.addEventListener("click", (e) => {
        e.stopPropagation();
        const isOpen = dropdown.classList.contains("open");
        closeAll(dropdown);
        dropdown.classList.toggle("open", !isOpen);
        button.setAttribute("aria-expanded", String(!isOpen));
      });

      // Escape closes the menu and returns focus to the trigger button (WCAG 2.1.1, 2.4.3)
      dropdown.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
          e.stopPropagation();
          dropdown.classList.remove("open");
          button.setAttribute("aria-expanded", "false");
          button.focus();
        }
      });

      // Close the menu once focus moves outside it (e.g. via Tab)
      dropdown.addEventListener("focusout", (e) => {
        if (!dropdown.contains(e.relatedTarget)) {
          dropdown.classList.remove("open");
          button.setAttribute("aria-expanded", "false");
        }
      });
    });

    document.addEventListener("click", () => closeAll());
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        closeAll();
      }
    });

    if (resourceLinks.length > 0) {
      resourceLinks.forEach((link) => {
        link.addEventListener("click", (e) => {
          e.preventDefault();

          const url = link.getAttribute("href");
          const resourceName = link.getAttribute("data-resource-name") || "Indhold";
          const openMode = link.getAttribute("data-open-mode") || "same-tab";

          if (!url) {
            return;
          }

          hidePortalViews();
          const statusText = openMode === "new-window"
            ? `${resourceName} åbnes nu i nyt vindue.`
            : `${resourceName} åbnes nu i samme vindue.`;
          setResourceStatus(statusText);

          if (openMode === "new-window") {
            window.open(url, "_blank", "noopener,noreferrer");
            return;
          }

          window.location.assign(url);
        });
      });
    }
  }

  let portalAuthorization = {
    role: "requester",
    permissions: [],
  };

  setAuthUi(false, null);
  applyPortalAuthorizationToUi();
  initPortalInteractions();

  const isLoopback = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";

  if (window.location.protocol === "file:") {
    setAuthStatus("Login kræver hosting på https:// eller http://localhost. Åbn ikke via file://.");
    if (loginButton) {
      loginButton.disabled = true;
    }
    return;
  }

  if (isLoopback && (!window.msal || !window.msal.PublicClientApplication)) {
    setAuthStatus("MSAL bibliotek kunne ikke indlæses. Kontrollér internetforbindelse.");
    if (loginButton) {
      loginButton.disabled = true;
    }
    return;
  }

  if (isLoopback && clientId === placeholderClientId) {
    setAuthStatus("Konfiguration mangler: indsæt jeres App (client) ID i script.js.");
    if (loginButton) {
      loginButton.disabled = true;
    }
    return;
  }

  const redirectUri = window.location.origin + window.location.pathname;
  const serverLoginUrl = `/.auth/login/aad?post_login_redirect_uri=${encodeURIComponent(window.location.pathname || "/")}&domain_hint=blind.dk&prompt=login`;
  const serverLogoutUrl = `/.auth/logout?post_logout_redirect_uri=${encodeURIComponent(window.location.origin + (window.location.pathname || "/"))}`;
  const serverLoginAttemptStorageKey = "dbsPortalServerLoginAttemptAt";
  const serverLoginCooldownMs = 30000;

  const msalConfig = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri,
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
  };

  const loginRequest = {
    scopes: portalLoginScopes,
  };

  const msalInstance = window.msal && window.msal.PublicClientApplication
    ? new window.msal.PublicClientApplication(msalConfig)
    : null;
  let isMsalReady = false;
  let serverAuthenticatedUser = null;
  let graphInteractionPromise = null;
  const graphPendingActionStorageKey = "dbsPortalPendingGraphAction";
  const graphRedirectStartedCode = "graph_redirect_started";
  const localMsalDefaultAuthorization = {
    role: "requester",
    permissions: ["portal.forms", "portal.cases.view", "portal.cases.read.own"],
  };
  const authGroupClaimTypes = new Set([
    "groups",
    "http://schemas.microsoft.com/ws/2008/06/identity/claims/groups",
    "http://schemas.microsoft.com/identity/claims/groups",
  ]);
  const graphProtectionNotice = "Graph-funktioner i browseren er slået fra her, fordi tenantens tokenbeskyttelse blokerer SPA-adgang. Sager, formularer og lokale editorer virker stadig.";

  function isServerAuthenticated() {
    return !!serverAuthenticatedUser;
  }

  function getActiveOrFirstAccount() {
    if (!msalInstance) {
      return null;
    }

    return msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
  }

  function getLocalIdentityHeaders() {
    const account = getActiveOrFirstAccount();
    if (!account) {
      return {};
    }

    const claims = account.idTokenClaims || {};
    const claimEmailCandidates = [
      claims.preferred_username,
      claims.upn,
      Array.isArray(claims.emails) ? claims.emails[0] : claims.emails,
      account.username,
    ];
    const email = String(claimEmailCandidates.find((value) => String(value || "").trim()) || "").trim();
    if (!email) {
      return {};
    }

    const name = String(account.name || claims.name || email).trim();
    const objectId = String(claims.oid || claims.objectId || "").trim();
    const groups = Array.isArray(claims.groups)
      ? claims.groups.map((entry) => String(entry || "").trim()).filter(Boolean)
      : [];

    const domain = email.includes("@") ? email.split("@")[1] : "";
    const nameParts = String(name || "")
      .trim()
      .split(/\s+/)
      .map((part) => part.replace(/[^a-zA-Z0-9.-]/g, "").toLowerCase())
      .filter(Boolean);
    const derivedNameEmail = (domain && nameParts.length >= 2)
      ? `${nameParts[0]}.${nameParts[nameParts.length - 1]}@${domain}`
      : "";
    const emailCandidates = [...new Set([
      ...claimEmailCandidates.map((value) => String(value || "").trim()).filter(Boolean),
      derivedNameEmail,
    ])];

    const headers = {
      "x-dbs-user-email": email,
      "x-dbs-user-name": name,
    };

    if (objectId) {
      headers["x-dbs-user-id"] = objectId;
    }

    if (groups.length > 0) {
      headers["x-dbs-user-groups"] = groups.join(",");
    }

    if (emailCandidates.length > 0) {
      headers["x-dbs-user-email-candidates"] = emailCandidates.join(",");
    }

    return headers;
  }

  function getCurrentUserEmailForUi() {
    const serverEmail = String(serverAuthenticatedUser?.email || "").trim();
    if (serverEmail) {
      return serverEmail;
    }

    const account = getActiveOrFirstAccount();
    const claims = account?.idTokenClaims || {};
    const claimEmailCandidates = [
      claims.preferred_username,
      claims.upn,
      Array.isArray(claims.emails) ? claims.emails[0] : claims.emails,
      account?.username,
    ];
    return String(claimEmailCandidates.find((value) => String(value || "").trim()) || "").trim();
  }

  function normalizePermissionList(permissions) {
    if (!Array.isArray(permissions)) {
      return [];
    }

    return permissions
      .map((permission) => String(permission || "").trim())
      .filter(Boolean);
  }

  function hasPortalPermission(permission) {
    const permissionList = normalizePermissionList(portalAuthorization?.permissions);
    if (permissionList.includes("*")) {
      return true;
    }

    return permissionList.includes(permission);
  }

  function mapAuthMeEntries(payload) {
    if (Array.isArray(payload)) {
      return payload;
    }

    if (Array.isArray(payload?.clientPrincipal)) {
      return payload.clientPrincipal;
    }

    if (payload?.clientPrincipal && typeof payload.clientPrincipal === "object") {
      return [payload.clientPrincipal];
    }

    if (payload?.user_claims) {
      return [payload];
    }

    if (payload?.claims) {
      return [payload];
    }

    return [];
  }

  function normalizeAuthMeResponse(payload) {
    const entries = mapAuthMeEntries(payload);
    const firstEntry = entries[0] || null;
    if (!firstEntry) {
      return null;
    }

    const claims = Array.isArray(firstEntry.user_claims)
      ? firstEntry.user_claims
      : Array.isArray(firstEntry.claims)
        ? firstEntry.claims
        : [];

    const findClaim = (type) => {
      const match = claims.find((claim) => String(claim?.typ || claim?.type || '').toLowerCase() === type.toLowerCase());
      return String(match?.val || match?.value || '').trim();
    };

    const groups = claims
      .filter((claim) => authGroupClaimTypes.has(String(claim?.typ || claim?.type || '').trim().toLowerCase()))
      .flatMap((claim) => String(claim?.val || claim?.value || '').split(','))
      .map((entry) => entry.trim())
      .filter(Boolean);

    const email = String(firstEntry.user_details || firstEntry.userDetails || findClaim('preferred_username') || findClaim('emails') || '').trim();
    const name = String(findClaim('name') || firstEntry.user_name || firstEntry.userName || firstEntry.user_details || firstEntry.userDetails || email || '').trim();

    return {
      user: {
        id: String(firstEntry.user_id || firstEntry.userId || findClaim('http://schemas.microsoft.com/identity/claims/objectidentifier') || '').trim(),
        name: name || email || 'Ukendt bruger',
        email: email || name,
        provider: String(firstEntry.identity_provider || firstEntry.identityProvider || firstEntry.auth_typ || firstEntry.authenticationType || 'aad').trim() || 'aad',
        groups,
      },
    };
  }

  function setPortalAuthorization(authorization) {
    const normalized = {
      role: String(authorization?.role || "requester").trim().toLowerCase() || "requester",
      permissions: normalizePermissionList(authorization?.permissions),
    };

    portalAuthorization = normalized;
  }

  function clearServerLoginAttempt() {
    try {
      window.sessionStorage.removeItem(serverLoginAttemptStorageKey);
    } catch {
      // Ignore session storage failures.
    }
  }

  function getLastServerLoginAttemptAt() {
    try {
      const rawValue = window.sessionStorage.getItem(serverLoginAttemptStorageKey);
      const parsedValue = Number(rawValue);
      return Number.isFinite(parsedValue) ? parsedValue : 0;
    } catch {
      return 0;
    }
  }

  function shouldThrottleServerLoginRedirect() {
    const lastAttemptAt = getLastServerLoginAttemptAt();
    if (!lastAttemptAt) {
      return false;
    }

    return Date.now() - lastAttemptAt < serverLoginCooldownMs;
  }

  function redirectToServerLogin(statusMessage) {
    try {
      window.sessionStorage.setItem(serverLoginAttemptStorageKey, String(Date.now()));
    } catch {
      // Ignore session storage failures.
    }

    if (statusMessage) {
      setAuthStatus(statusMessage);
    }

    window.location.assign(serverLoginUrl);
  }

  function setPendingGraphAction(action, payload = {}) {
    if (!action) {
      return;
    }

    try {
      const value = JSON.stringify({
        action: String(action).trim(),
        payload: payload && typeof payload === "object" ? payload : {},
      });
      window.sessionStorage.setItem(graphPendingActionStorageKey, value);
    } catch {
      // Ignore session storage failures.
    }
  }

  function getPendingGraphAction() {
    try {
      const rawValue = window.sessionStorage.getItem(graphPendingActionStorageKey);
      if (!rawValue) {
        return null;
      }

      const parsed = JSON.parse(rawValue);
      const action = String(parsed?.action || "").trim();
      if (!action) {
        return null;
      }

      return {
        action,
        payload: parsed?.payload && typeof parsed.payload === "object" ? parsed.payload : {},
      };
    } catch {
      return null;
    }
  }

  function clearPendingGraphAction() {
    try {
      window.sessionStorage.removeItem(graphPendingActionStorageKey);
    } catch {
      // Ignore session storage failures.
    }
  }

  function createGraphRedirectStartedError() {
    const error = new Error("Videresender til Microsoft-login for at fuldfoere adgangen.");
    error.errorCode = graphRedirectStartedCode;
    return error;
  }

  function isGraphRedirectStartedError(error) {
    const code = String(error?.errorCode || "").toLowerCase();
    return code === graphRedirectStartedCode;
  }

  function setMenuActionAvailability(button, allowed, reason) {
    if (!button) {
      return;
    }

    button.disabled = !allowed;

    if (allowed) {
      button.removeAttribute("aria-disabled");
      button.removeAttribute("title");
      return;
    }

    button.setAttribute("aria-disabled", "true");
    if (reason) {
      button.title = reason;
    }
  }

  function applyPortalAuthorizationToUi() {
    const deniedReason = "Din rolle har ikke adgang til denne funktion.";

    setMenuActionAvailability(mailPortalButton, hasPortalPermission("portal.mail"), deniedReason);
    setMenuActionAvailability(wordPortalButton, hasPortalPermission("portal.word"), deniedReason);
    setMenuActionAvailability(excelPortalButton, hasPortalPermission("portal.excel"), deniedReason);
    setMenuActionAvailability(sharepointPortalButton, hasPortalPermission("portal.sharepoint"), deniedReason);
    setMenuActionAvailability(oneDrivePortalButton, hasPortalPermission("portal.onedrive"), deniedReason);
    setMenuActionAvailability(teamsPortalButton, hasPortalPermission("portal.teams"), deniedReason);
    setMenuActionAvailability(casesPortalButton, hasPortalPermission("portal.cases.view"), deniedReason);

    formPortalButtons.forEach((button) => {
      setMenuActionAvailability(button, hasPortalPermission("portal.forms"), deniedReason);
    });
  }

  function setGraphFeatureAvailability(enabled, reason = "") {
    const graphButtons = [
      { button: mailPortalButton, permission: "portal.mail" },
      { button: sharepointPortalButton, permission: "portal.sharepoint" },
      { button: oneDrivePortalButton, permission: "portal.onedrive" },
      { button: teamsPortalButton, permission: "portal.teams" },
      { button: wordSaveOneDriveButton, permission: "portal.onedrive" },
      { button: wordSaveSharePointButton, permission: "portal.sharepoint" },
      { button: excelSaveOneDriveButton, permission: "portal.onedrive" },
      { button: excelSaveSharePointButton, permission: "portal.sharepoint" },
    ];

    graphButtons.forEach(({ button, permission }) => {
      if (!button) {
        return;
      }

      const allowedByRole = hasPortalPermission(permission);
      const allowed = enabled && allowedByRole;
      button.disabled = !allowed;

      if (!allowed) {
        button.setAttribute("aria-disabled", "true");
        button.title = allowedByRole ? reason : "Din rolle har ikke adgang til denne funktion.";
      } else {
        button.removeAttribute("title");
        button.removeAttribute("aria-disabled");
      }
    });
  }

  async function detectServerAuthorization() {
    try {
      const authzResponse = await fetch("/api/authz/me", {
        cache: "no-store",
        credentials: "include",
        headers: {
          ...getLocalIdentityHeaders(),
        },
      });

      if (authzResponse.ok) {
        const authzPayload = await authzResponse.json();
        if (authzPayload?.authenticated && authzPayload?.user) {
          return {
            user: authzPayload.user,
            authorization: authzPayload.authorization || null,
          };
        }
      }

      const response = await fetch("/.auth/me", { cache: "no-store", credentials: "include" });
      if (!response.ok) {
        return null;
      }

      const payload = await response.json();
      const normalized = normalizeAuthMeResponse(payload);
      if (!normalized?.user) {
        return null;
      }

      return {
        user: normalized.user,
        authorization: null,
      };
    } catch {
      return null;
    }
  }

  async function detectServerAuthenticatedUser() {
    try {
      const response = await fetch("/.auth/me", { cache: "no-store", credentials: "include" });
      if (!response.ok) {
        return null;
      }

      const payload = await response.json();
      return normalizeAuthMeResponse(payload)?.user || null;
    } catch {
      return null;
    }
  }

  function isTokenProtectionError(error) {
    const text = String(error?.message || error || "").toLowerCase();
    return text.includes("tokenbeskyttelse")
      || text.includes("token protection")
      || text.includes("forhindrer dette program i at få adgang til ressourcen")
      || text.includes("prevents this program from accessing the resource");
  }

  function isInteractionInProgressError(error) {
    const code = String(error?.errorCode || "").toLowerCase();
    const message = String(error?.message || error || "").toLowerCase();
    return code === "interaction_in_progress"
      || message.includes("interaction_in_progress")
      || message.includes("interaction is currently in progress");
  }

  async function runGraphInteraction(interactionFactory) {
    if (graphInteractionPromise) {
      await graphInteractionPromise;
    }

    graphInteractionPromise = (async () => interactionFactory())();
    try {
      return await graphInteractionPromise;
    } finally {
      graphInteractionPromise = null;
    }
  }

  function setSharepointStatus(message) {
    if (sharepointStatus) {
      sharepointStatus.textContent = message;
    }
  }

  function clearSharepointList() {
    if (sharepointList) {
      sharepointList.textContent = "";
    }
  }

  function setSharepointPath(path) {
    sharepointCurrentPath = path || "";
    if (sharepointPath) {
      sharepointPath.textContent = `Placering: ${sharepointCurrentPath || "Rod"}`;
    }
  }

  function setMirrorMode(mode) {
    mirrorMode = mode === "onedrive" ? "onedrive" : "sharepoint";

    if (sharepointTitle) {
      sharepointTitle.textContent = mirrorMode === "onedrive" ? "OneDrive i portal" : "SharePoint i portal";
    }

    if (mirrorSourceLabel) {
      mirrorSourceLabel.textContent = mirrorMode === "onedrive" ? "Kilde: OneDrive" : "Kilde: SharePoint";
    }
  }

  function getParentPath(path) {
    const segments = (path || "").split("/").filter(Boolean);
    if (segments.length === 0) {
      return "";
    }

    return segments.slice(0, -1).join("/");
  }

  function buildFolderPath(folderName) {
    if (!folderName) {
      return sharepointCurrentPath;
    }

    return sharepointCurrentPath ? `${sharepointCurrentPath}/${folderName}` : folderName;
  }

  function encodeGraphPath(path) {
    return (path || "")
      .split("/")
      .filter(Boolean)
      .map((segment) => encodeURIComponent(segment))
      .join("/");
  }

  function formatDate(isoDate) {
    if (!isoDate) {
      return "Ukendt dato";
    }

    const date = new Date(isoDate);
    if (Number.isNaN(date.getTime())) {
      return "Ukendt dato";
    }

    return date.toLocaleString("da-DK");
  }

  function formatSize(bytes) {
    if (typeof bytes !== "number" || bytes < 0) {
      return "Ukendt størrelse";
    }

    if (bytes < 1024) {
      return `${bytes} B`;
    }

    const kb = bytes / 1024;
    if (kb < 1024) {
      return `${kb.toFixed(1)} KB`;
    }

    const mb = kb / 1024;
    if (mb < 1024) {
      return `${mb.toFixed(1)} MB`;
    }

    const gb = mb / 1024;
    return `${gb.toFixed(1)} GB`;
  }

  function buildPortalFileName(prefix, extension) {
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    return `${prefix}-${stamp}.${extension}`;
  }

  function encodeGraphFilePath(path) {
    return String(path || "")
      .split("/")
      .filter(Boolean)
      .map((segment) => encodeURIComponent(segment))
      .join("/");
  }

  async function uploadContentToOnedrive(accessToken, filePath, content, mimeType) {
    const encodedPath = encodeGraphFilePath(filePath);
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/content`;
    const response = await fetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": mimeType,
      },
      body: content,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`OneDrive upload fejl (${response.status}): ${errorText || "ukendt fejl"}`);
    }

    return response.json();
  }

  async function uploadContentToSharepoint(accessToken, filePath, content, mimeType) {
    await ensureSharepointDriveContext(accessToken);
    const encodedPath = encodeGraphFilePath(filePath);
    const url = `https://graph.microsoft.com/v1.0/drives/${sharepointDriveId}/root:/${encodedPath}:/content`;
    const response = await fetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": mimeType,
      },
      body: content,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`SharePoint upload fejl (${response.status}): ${errorText || "ukendt fejl"}`);
    }

    return response.json();
  }

  async function saveWordToCloud(target) {
    const html = wordEditor?.innerHTML || "";
    if (!html.trim()) {
      setWordStatus("Der er intet indhold at gemme.");
      return;
    }

    const htmlDocument = [
      "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>Dokument</title></head><body>",
      html,
      "</body></html>",
    ].join("");

    const fileName = buildPortalFileName("dbs-word-dokument", "html");
    const targetPath = fileName;

    try {
      setWordStatus(target === "sharepoint" ? "Gemmer til SharePoint..." : "Gemmer til OneDrive...");
      const accessToken = await acquireGraphToken(true, false, graphWriteScopes);
      const result = target === "sharepoint"
        ? await uploadContentToSharepoint(accessToken, targetPath, htmlDocument, "text/html; charset=utf-8")
        : await uploadContentToOnedrive(accessToken, targetPath, htmlDocument, "text/html; charset=utf-8");

      const location = target === "sharepoint" ? "SharePoint" : "OneDrive";
      setWordStatus(`Dokument gemt i ${location}: ${result?.name || fileName}.`);
    } catch (error) {
      const message = error?.message || "Ukendt fejl";
      setWordStatus(`Kunne ikke gemme dokumentet: ${message}`);
    }
  }

  async function saveExcelToCloud(target) {
    try {
      const workbookData = exportExcelWorkbookBinary();
      const fileName = buildPortalFileName("dbs-excel-ark", "xlsx");
      const targetPath = fileName;

      setExcelStatus(target === "sharepoint" ? "Gemmer til SharePoint..." : "Gemmer til OneDrive...");
      const accessToken = await acquireGraphToken(true, false, graphWriteScopes);
      const result = target === "sharepoint"
        ? await uploadContentToSharepoint(
          accessToken,
          targetPath,
          workbookData,
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        : await uploadContentToOnedrive(
          accessToken,
          targetPath,
          workbookData,
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        );

      const location = target === "sharepoint" ? "SharePoint" : "OneDrive";
      setExcelStatus(`Regneark gemt i ${location}: ${result?.name || fileName}.`);
    } catch (error) {
      const message = error?.message || "Ukendt fejl";
      setExcelStatus(`Kunne ikke gemme regnearket: ${message}`);
    }
  }

  async function fetchGraph(url, accessToken, extraHeaders = {}) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        ...extraHeaders,
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Graph fejl (${response.status}): ${errorText || "ukendt fejl"}`);
    }

    return response.json();
  }

  function isMsalPopupTimeoutError(error) {
    const code = String(error?.errorCode || "").toLowerCase();
    const message = String(error?.message || "").toLowerCase();
    return code === "timed_out"
      || code === "popup_window_error"
      || message.includes("timed_out")
      || message.includes("popup")
      || message.includes("window.open");
  }

  async function acquireGraphToken(interactiveAllowed, forceConsent = false, requestedScopes = mailGraphScopes, options = {}) {
    if (!msalInstance) {
      throw new Error("Graph-login er ikke tilgængelig i denne session.");
    }

    const allowRedirect = options?.allowRedirect === true;
    const pendingAction = options?.pendingAction || null;
    const requestRedirect = async (request) => {
      if (pendingAction?.action) {
        setPendingGraphAction(pendingAction.action, pendingAction.payload || {});
      }

      await msalInstance.acquireTokenRedirect(request);
      throw createGraphRedirectStartedError();
    };

    let account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
    if (!account && isServerAuthenticated()) {
      const loginHint = String(serverAuthenticatedUser?.email || "").trim();

      if (!interactiveAllowed) {
        try {
          const silentResult = await msalInstance.ssoSilent({
            scopes: requestedScopes,
            loginHint: loginHint || undefined,
          });
          if (silentResult?.account) {
            msalInstance.setActiveAccount(silentResult.account);
            account = silentResult.account;
          }
        } catch {
          // Silent bootstrap can fail if there is no browser session yet.
        }
      } else {
        if (allowRedirect) {
          if (pendingAction?.action) {
            setPendingGraphAction(pendingAction.action, pendingAction.payload || {});
          }

          await msalInstance.loginRedirect({
            scopes: requestedScopes,
            loginHint: loginHint || undefined,
            prompt: forceConsent ? "consent" : "select_account",
          });
          throw createGraphRedirectStartedError();
        }

        try {
          const loginResult = await runGraphInteraction(() => msalInstance.loginPopup({
            scopes: requestedScopes,
            loginHint: loginHint || undefined,
            prompt: forceConsent ? "consent" : "select_account",
          }));
          if (loginResult?.account) {
            msalInstance.setActiveAccount(loginResult.account);
            account = loginResult.account;
          }
        } catch (error) {
          if (isInteractionInProgressError(error)) {
            if (graphInteractionPromise) {
              await graphInteractionPromise;
              account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
            }
          } else {
            throw error;
          }
        }
      }
    }

    if (!account) {
      if (isServerAuthenticated()) {
        throw new Error("Du er logget ind til portalen, men Graph-login i browseren er ikke aktiv. Prøv handlingen igen og fuldfør Microsoft-popup'en.");
      }

      throw new Error("Ingen aktiv login-session fundet. Log ind igen.");
    }

    const tokenRequest = {
      account,
      scopes: requestedScopes,
    };

    try {
      const token = await msalInstance.acquireTokenSilent(tokenRequest);
      return token.accessToken;
    } catch (error) {
      if (interactiveAllowed) {
        if (allowRedirect) {
          await requestRedirect({
            scopes: tokenRequest.scopes,
            prompt: forceConsent ? "consent" : undefined,
          });
        }

        try {
          const popupToken = await runGraphInteraction(() => msalInstance.acquireTokenPopup({
            scopes: tokenRequest.scopes,
            prompt: forceConsent ? "consent" : undefined,
          }));
          return popupToken.accessToken;
        } catch (interactiveError) {
          if (isTokenProtectionError(interactiveError)) {
            throw new Error("Tokenbeskyttelse i Microsoft Entra blokerer denne browserapp fra at hente Graph-token. Portalen kan godt beskyttes med Microsoft-login, men Graph-funktioner kræver en understøttet klient eller en ændring i Conditional Access-politikken.");
          }

          if (isInteractionInProgressError(interactiveError)) {
            if (graphInteractionPromise) {
              await graphInteractionPromise;
            }

            const retryToken = await msalInstance.acquireTokenSilent(tokenRequest);
            return retryToken.accessToken;
          }

          if (!isMsalPopupTimeoutError(interactiveError)) {
            throw interactiveError;
          }

          await requestRedirect({
            scopes: tokenRequest.scopes,
            prompt: forceConsent ? "consent" : undefined,
          });
        }
      }

      throw error;
    }
  }

  async function resumePendingGraphActionIfNeeded() {
    if (!isMsalReady || !msalInstance) {
      return;
    }

    const pending = getPendingGraphAction();
    if (!pending?.action) {
      return;
    }

    if (!getActiveOrFirstAccount()) {
      return;
    }

    try {
      if (pending.action === "sharepoint") {
        const path = String(pending.payload?.path || "");
        await loadSharepointPortal(false, path);
      } else if (pending.action === "onedrive") {
        const path = String(pending.payload?.path || "");
        await loadOnedrivePortal(false, path);
      } else if (pending.action === "teams") {
        await openTeamsPortal(false);
      } else if (pending.action === "mail") {
        await openMailPortal();
      }

      clearPendingGraphAction();
    } catch (error) {
      if (isGraphRedirectStartedError(error)) {
        return;
      }

      clearPendingGraphAction();
    }
  }

  async function ensureSharepointDriveContext(accessToken) {
    if (sharepointSiteId && sharepointDriveId) {
      return;
    }

    const site = await fetchGraph(
      `https://graph.microsoft.com/v1.0/sites/${sharepointHost}:${sharepointSitePath}?$select=id,displayName,webUrl`,
      accessToken,
    );

    const drivesResponse = await fetchGraph(
      `https://graph.microsoft.com/v1.0/sites/${site.id}/drives?$select=id,name,webUrl`,
      accessToken,
    );

    const drives = drivesResponse.value || [];
    const targetDrive = drives.find((drive) =>
      (drive.name || "").toLowerCase() === sharepointLibraryName.toLowerCase()
    ) || drives[0];

    if (!targetDrive) {
      throw new Error("Ingen dokumentbiblioteker fundet på SharePoint-sitet.");
    }

    sharepointSiteId = site.id;
    sharepointDriveId = targetDrive.id;
    sharepointDriveName = targetDrive.name || sharepointLibraryName;
  }

  async function invokePowerAutomateFlow(formType, payload) {
    if (!powerAutomateFlowUrl || powerAutomateFlowUrl === "REPLACE_WITH_POWER_AUTOMATE_HTTP_TRIGGER_URL") {
      throw new Error("Power Automate URL mangler. Indsæt jeres HTTP trigger URL i script.js.");
    }

    const flowPayload = {
      formType,
      submittedAt: new Date().toISOString(),
      submittedBy: msalInstance.getActiveAccount()?.username || "ukendt",
      data: payload,
    };

    const response = await fetch(powerAutomateFlowUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json; charset=utf-8",
      },
      body: JSON.stringify(flowPayload),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Flow fejl (${response.status}): ${errorText || "ukendt fejl"}`);
    }

    return true;
  }

  async function sendMailViaGraph(messagePayload, accessToken) {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message: messagePayload,
        saveToSentItems: true,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Mail fejl (${response.status}): ${errorText || "ukendt fejl"}`);
    }
  }

  async function sendQuickReply(message, interactiveAllowed) {
    const messageId = message?.id || "";
    if (!messageId) {
      throw new Error("Kan ikke svare: mail-id mangler.");
    }

    const replyUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/reply`;
    const payload = {
      comment: "Hej\n\nTak for din mail. Jeg vender tilbage hurtigst muligt.\n\nVenlig hilsen",
    };

    let accessToken = await acquireGraphToken(interactiveAllowed, false, mailGraphScopes);

    const sendReply = async (token) => {
      const response = await fetch(replyUrl, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Svar fejl (${response.status}): ${errorText || "ukendt fejl"}`);
      }
    };

    try {
      await sendReply(accessToken);
    } catch (error) {
      if (!interactiveAllowed || !isAccessDeniedMessage(error?.message)) {
        throw error;
      }

      accessToken = await acquireGraphToken(true, true, mailGraphScopes);
      await sendReply(accessToken);
    }
  }

  function renderSharepointItems(items, onOpenFolder) {
    clearSharepointList();

    if (!sharepointList) {
      return;
    }

    if (!items || items.length === 0) {
      const emptyItem = document.createElement("li");
      emptyItem.className = "sharepoint-item";
      emptyItem.textContent = "Ingen elementer fundet i biblioteket.";
      sharepointList.appendChild(emptyItem);
      return;
    }

    items.forEach((item) => {
      const listItem = document.createElement("li");
      listItem.className = "sharepoint-item";

      const isFolder = Boolean(item.folder);
      const label = item.name || "Uden navn";

      if (isFolder) {
        const folderButton = document.createElement("button");
        folderButton.type = "button";
        folderButton.className = "sharepoint-folder-button";
        folderButton.textContent = label;
        folderButton.addEventListener("click", () => {
          onOpenFolder(buildFolderPath(label));
        });
        listItem.appendChild(folderButton);
      } else {
        const link = document.createElement("a");
        link.className = "sharepoint-item-link";
        link.href = item.webUrl || "#";
        link.target = "_blank";
        link.rel = "noopener noreferrer";
        link.textContent = label;
        listItem.appendChild(link);
      }

      const meta = document.createElement("div");
      meta.className = "sharepoint-meta";
      const itemType = isFolder ? "Mappe" : "Fil";
      const sizeText = isFolder ? "-" : formatSize(item.size);
      meta.textContent = `${itemType} | Opdateret: ${formatDate(item.lastModifiedDateTime)} | Størrelse: ${sizeText}`;

      listItem.appendChild(meta);
      sharepointList.appendChild(listItem);
    });
  }

  async function loadSharepointPortal(interactiveAllowed, requestedPath = sharepointCurrentPath) {
    if (!hasPortalPermission("portal.sharepoint")) {
      setResourceStatus("Du har ikke adgang til SharePoint i portalen.");
      return;
    }

    if (!isMsalReady) {
      setSharepointStatus("Login initialiseres stadig. Prøv igen om et øjeblik.");
      return;
    }

    if (!sharepointView) {
      return;
    }

    hidePortalViews();
    sharepointView.classList.remove("is-hidden");
    setMirrorMode("sharepoint");
    setResourceStatus("SharePoint vises i portalvisningen.");
    clearSharepointList();
    setSharepointStatus("Henter SharePoint-indhold...");

    try {
      const accessToken = await acquireGraphToken(interactiveAllowed, false, sharepointReadScopes, {
        allowRedirect: !isLoopback,
        pendingAction: {
          action: "sharepoint",
          payload: {
            path: requestedPath,
          },
        },
      });

      await ensureSharepointDriveContext(accessToken);

      const encodedPath = encodeGraphPath(requestedPath);
      const childrenUrl = encodedPath
        ? `https://graph.microsoft.com/v1.0/drives/${sharepointDriveId}/root:/${encodedPath}:/children?$top=200&$select=id,name,webUrl,size,lastModifiedDateTime,folder,file`
        : `https://graph.microsoft.com/v1.0/drives/${sharepointDriveId}/root/children?$top=200&$select=id,name,webUrl,size,lastModifiedDateTime,folder,file`;

      const childrenResponse = await fetchGraph(childrenUrl, accessToken);

      const items = childrenResponse.value || [];
      setSharepointPath(requestedPath);
      renderSharepointItems(items, (folderPath) => {
        loadSharepointPortal(false, folderPath);
      });
      setSharepointStatus(`Viser ${items.length} elementer fra ${sharepointDriveName}.`);
      sharepointView.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (error) {
      if (isGraphRedirectStartedError(error)) {
        setSharepointStatus("Fortsætter login i samme fane. Du sendes tilbage til SharePoint automatisk.");
        return;
      }

      const message = error?.message || "Ukendt fejl";
      setSharepointStatus(`Kunne ikke hente SharePoint-indhold: ${message}`);
      clearSharepointList();
    }
  }

  async function loadOnedrivePortal(interactiveAllowed, requestedPath = sharepointCurrentPath) {
    if (!hasPortalPermission("portal.onedrive")) {
      setResourceStatus("Du har ikke adgang til OneDrive i portalen.");
      return;
    }

    if (!isMsalReady) {
      setSharepointStatus("Login initialiseres stadig. Prøv igen om et øjeblik.");
      return;
    }

    if (!sharepointView) {
      return;
    }

    hidePortalViews();
    sharepointView.classList.remove("is-hidden");
    setMirrorMode("onedrive");
    setResourceStatus("OneDrive vises i portalvisningen.");
    clearSharepointList();
    setSharepointStatus("Henter OneDrive-indhold...");

    try {
      const accessToken = await acquireGraphToken(interactiveAllowed, false, sharepointReadScopes, {
        allowRedirect: !isLoopback,
        pendingAction: {
          action: "onedrive",
          payload: {
            path: requestedPath,
          },
        },
      });

      const encodedPath = encodeGraphPath(requestedPath);
      const childrenUrl = encodedPath
        ? `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/children?$top=200&$select=id,name,webUrl,size,lastModifiedDateTime,folder,file`
        : "https://graph.microsoft.com/v1.0/me/drive/root/children?$top=200&$select=id,name,webUrl,size,lastModifiedDateTime,folder,file";

      const childrenResponse = await fetchGraph(childrenUrl, accessToken);

      const items = childrenResponse.value || [];
      setSharepointPath(requestedPath);
      renderSharepointItems(items, (folderPath) => {
        loadOnedrivePortal(false, folderPath);
      });
      setSharepointStatus(`Viser ${items.length} elementer fra din OneDrive.`);
      sharepointView.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (error) {
      if (isGraphRedirectStartedError(error)) {
        setSharepointStatus("Fortsætter login i samme fane. Du sendes tilbage til OneDrive automatisk.");
        return;
      }

      const message = error?.message || "Ukendt fejl";
      setSharepointStatus(`Kunne ikke hente OneDrive-indhold: ${message}`);
      clearSharepointList();
    }
  }

  if (loginButton) {
    loginButton.disabled = true;
  }

  async function handleAuthentication() {
    if (!isLoopback) {
      try {
        setAuthStatus("Kontrollerer login...");
        const serverAuthContext = await detectServerAuthorization();
        serverAuthenticatedUser = serverAuthContext?.user || null;

        if (!serverAuthenticatedUser) {
          setPortalAuthorization({ role: "requester", permissions: [] });
          applyPortalAuthorizationToUi();
          setAuthUi(false, null);
          setAuthStatus("Du skal logge ind for at bruge portalen.");
          return;
        }

        setPortalAuthorization(serverAuthContext?.authorization || localMsalDefaultAuthorization);
        applyPortalAuthorizationToUi();
        setAuthUi(true, {
          name: serverAuthenticatedUser?.name || serverAuthenticatedUser?.email || "Bruger",
          username: serverAuthenticatedUser?.email || "",
        });
        setAuthStatus("Logget ind.");
        setGraphFeatureAvailability(true);
        setResourceStatus("");
        return;
      } catch (error) {
        setPortalAuthorization({ role: "requester", permissions: [] });
        applyPortalAuthorizationToUi();
        setAuthUi(false, null);
        setAuthStatus(`Loginfejl: ${error?.message || "Ukendt fejl"}`);
        return;
      }
    }

    try {
      setAuthStatus("Kontrollerer login...");
      const redirectResult = await msalInstance.handleRedirectPromise();

      if (redirectResult?.account) {
        msalInstance.setActiveAccount(redirectResult.account);
      }

      const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;

      if (account) {
        clearServerLoginAttempt();
        const serverAuthContext = await detectServerAuthorization();
        serverAuthenticatedUser = serverAuthContext?.user || null;

        if (!serverAuthenticatedUser) {
          const fallbackUser = await detectServerAuthenticatedUser();
          if (fallbackUser) {
            serverAuthenticatedUser = fallbackUser;
          }
        }

        setPortalAuthorization(serverAuthContext?.authorization || localMsalDefaultAuthorization);
        applyPortalAuthorizationToUi();
        setAuthUi(true, account);
        setAuthStatus("Logget ind.");
        setGraphFeatureAvailability(true);
        setResourceStatus("");
      } else {
        setPortalAuthorization({ role: "requester", permissions: [] });
        applyPortalAuthorizationToUi();
        setAuthUi(false, null);
        setAuthStatus("Du skal logge ind for at bruge portalen.");
      }
    } catch (error) {
      setPortalAuthorization({ role: "requester", permissions: [] });
      applyPortalAuthorizationToUi();
      setAuthUi(false, null);
      if (isTokenProtectionError(error)) {
        setAuthStatus("Loginfejl: Microsoft Entra tokenbeskyttelse blokerer denne browserapp. Jeg anbefaler serverbaseret login for portalen og en separat løsning for Graph-funktioner.");
        return;
      }

      setAuthStatus(`Loginfejl: ${error?.message || "Ukendt fejl"}`);
    }
  }

  if (loginButton) {
    loginButton.addEventListener("click", () => {
      if (!isLoopback) {
        if (shouldThrottleServerLoginRedirect()) {
          setAuthStatus("Microsoft-login blev ikke fuldført. Vent et øjeblik og prøv igen.");
          return;
        }

        redirectToServerLogin("Sender dig til Microsoft login...");
        return;
      }

      if (!isMsalReady) {
        setAuthStatus("Login initialiseres, prøv igen om et øjeblik.");
        return;
      }

      if (!isLoopback && shouldThrottleServerLoginRedirect()) {
        setAuthStatus("Microsoft-login blev ikke fuldført. Vent et øjeblik og prøv igen.");
        return;
      }

      clearServerLoginAttempt();
      setAuthStatus("Sender dig til Microsoft login...");
      msalInstance.loginRedirect(loginRequest);
    });
  }

  if (logoutButton) {
    logoutButton.addEventListener("click", () => {
      if (!isLoopback) {
        window.location.assign(serverLogoutUrl);
        return;
      }

      if (!isMsalReady) {
        return;
      }
      msalInstance.logoutRedirect({
        account: msalInstance.getActiveAccount() || undefined,
        postLogoutRedirectUri: redirectUri,
      });
    });
  }

  if (sharepointPortalButton) {
    sharepointPortalButton.addEventListener("click", () => {
      loadSharepointPortal(true, "");
    });
  }

  if (oneDrivePortalButton) {
    oneDrivePortalButton.addEventListener("click", () => {
      loadOnedrivePortal(true, "");
    });
  }

  if (teamsPortalButton) {
    teamsPortalButton.addEventListener("click", () => {
      openTeamsPortal(true);
    });
  }

  if (formPortalButtons.length > 0) {
    formPortalButtons.forEach((button) => {
      button.addEventListener("click", () => {
        const formType = button.getAttribute("data-form-type") || "";
        renderPortalForm(formType);
      });
    });
  }

  if (mailPortalButton) {
    mailPortalButton.addEventListener("click", () => {
      openMailPortal();
    });
  }

  if (mailCalendarModeMonthButton) {
    mailCalendarModeMonthButton.addEventListener("click", () => setCalendarViewMode("month"));
  }

  if (mailCalendarModeWeekButton) {
    mailCalendarModeWeekButton.addEventListener("click", () => setCalendarViewMode("week"));
  }

  if (mailCalendarModeWorkweekButton) {
    mailCalendarModeWorkweekButton.addEventListener("click", () => setCalendarViewMode("workweek"));
  }

  if (mailCalendarDraftCopyButton) {
    mailCalendarDraftCopyButton.addEventListener("click", () => {
      copyCalendarDraftText();
      setCalendarStatus("Mødekladden er kopieret til udklipsholderen.");
    });
  }

  if (mailCalendarDraftIcsButton) {
    mailCalendarDraftIcsButton.addEventListener("click", () => {
      downloadCalendarDraftIcs();
      setCalendarStatus("Mødeinvitationen er hentet som .ics-fil.");
    });
  }

  if (mailCalendarDraftClearButton) {
    mailCalendarDraftClearButton.addEventListener("click", () => {
      calendarDraft = {
        subject: "",
        date: new Date(),
        start: "09:00",
        end: "10:00",
        location: "",
      };
      setCalendarDraftFromDate(calendarSelectedDate, null);
      if (mailCalendarDraftSubject) {
        mailCalendarDraftSubject.value = "";
      }
      if (mailCalendarDraftLocation) {
        mailCalendarDraftLocation.value = "";
      }
      if (mailCalendarDraftStart) {
        mailCalendarDraftStart.value = "09:00";
      }
      if (mailCalendarDraftEnd) {
        mailCalendarDraftEnd.value = "10:00";
      }
      setCalendarStatus("Mødekladden er ryddet.");
    });
  }

  if (casesPortalButton) {
    casesPortalButton.addEventListener("click", () => {
      openCasesPortal();
    });
  }

  if (casesFilterOpenButton) {
    casesFilterOpenButton.addEventListener("click", () => {
      setCaseViewFilter("open");
    });
  }

  if (casesFilterClosedButton) {
    casesFilterClosedButton.addEventListener("click", () => {
      setCaseViewFilter("closed");
    });
  }

  if (casesFilterAllButton) {
    casesFilterAllButton.addEventListener("click", () => {
      setCaseViewFilter("all");
    });
  }

  if (wordPortalButton) {
    wordPortalButton.addEventListener("click", () => {
      openWordPortal();
    });
  }

  if (excelPortalButton) {
    excelPortalButton.addEventListener("click", () => {
      openExcelPortal();
    });
  }

  if (portalForm) {
    portalForm.addEventListener("submit", async (event) => {
      event.preventDefault();

      if (!activeFormType || !formDefinitions[activeFormType]) {
        setFormsStatus("Vælg en formular i menuen før afsendelse.");
        return;
      }

      if (!portalForm.reportValidity()) {
        setFormsStatus("Udfyld de obligatoriske felter markeret med *.");
        return;
      }

      const formData = new FormData(portalForm);
      const submittedData = {};
      formData.forEach((value, key) => {
        submittedData[key] = value;
      });

      const formDefinition = formDefinitions[activeFormType];
      const timestamp = new Date().toLocaleString("da-DK");
      setFormsStatus("Sender til Power Automate...");

      try {
        await invokePowerAutomateFlow(activeFormType, submittedData);

        setFormsStatus(`${formDefinition.successMessage} Tidspunkt: ${timestamp}.`);
        setResourceStatus(`${formDefinition.title} er sendt til Power Automate flow.`);
      } catch (error) {
        const message = error?.message || "Ukendt fejl";
        const isNetworkOrCors = error instanceof TypeError;
        if (isNetworkOrCors) {
          setFormsStatus("Afsendelse fejlede: Netværk/CORS blokerede kaldet til Power Automate. Kontrollér flowets URL og CORS-politik.");
        } else {
          setFormsStatus(`Afsendelse fejlede: ${message}`);
        }
        setResourceStatus("Formularen kunne ikke sendes til Power Automate.");
        return;
      }

      // Keep a local trace for debugging and support if users report field mapping issues.
      console.info("Portal form submission", {
        formType: activeFormType,
        submittedAt: timestamp,
        data: submittedData,
      });
    });
  }

  if (formsReset) {
    formsReset.addEventListener("click", () => {
      if (!activeFormType) {
        setFormsStatus("Ingen formular er valgt endnu.");
        return;
      }

      renderPortalForm(activeFormType);
      setFormsStatus("Formularen er nulstillet.");
    });
  }

  if (mailForm) {
    mailForm.addEventListener("submit", async (event) => {
      event.preventDefault();

      if (!mailForm.reportValidity()) {
        setMailStatus("Udfyld de obligatoriske mailfelter markeret med *.");
        return;
      }

      const toAddresses = splitRecipients(mailToInput?.value || "");
      const ccAddresses = splitRecipients(mailCcInput?.value || "");
      const bccAddresses = splitRecipients(mailBccInput?.value || "");
      const invalidAddresses = [...toAddresses, ...ccAddresses, ...bccAddresses].filter((address) => !isValidEmail(address));

      if (toAddresses.length === 0) {
        setMailStatus("Angiv mindst én modtager i Til-feltet.");
        return;
      }

      if (invalidAddresses.length > 0) {
        setMailStatus(`Ugyldig e-mailadresse fundet: ${invalidAddresses[0]}`);
        return;
      }

      const bodyContentType = mailContentTypeSelect?.value === "HTML" ? "HTML" : "Text";
      setMailStatus("Sender mail via Microsoft Graph...");

      try {
        const attachments = await buildGraphAttachments(mailAttachmentsInput?.files);
        const messagePayload = {
          subject: mailSubjectInput?.value?.trim() || "(intet emne)",
          body: {
            contentType: bodyContentType,
            content: mailBodyInput?.value || "",
          },
          toRecipients: toGraphRecipients(toAddresses),
          ccRecipients: toGraphRecipients(ccAddresses),
          bccRecipients: toGraphRecipients(bccAddresses),
          attachments,
        };

        let accessToken = await acquireGraphToken(true);

        try {
          await sendMailViaGraph(messagePayload, accessToken);
        } catch (sendError) {
          const sendMessage = sendError?.message || "";
          const accessDenied = sendMessage.includes("Mail fejl (403)") || sendMessage.toLowerCase().includes("access denied");

          if (!accessDenied) {
            throw sendError;
          }

          accessToken = await acquireGraphToken(true, true);
          await sendMailViaGraph(messagePayload, accessToken);
        }

        setMailStatus("Mail er sendt og gemt i Sendt post.");
        setResourceStatus("Mail blev sendt fra portalen.");
        mailForm.reset();
        if (mailHtmlPreview) {
          mailHtmlPreview.textContent = "";
        }
        updateMailPreviewVisibility();
        loadMessagesForSelectedFolder(false);
      } catch (error) {
        const message = error?.message || "Ukendt fejl";
        const isNetworkOrCors = error instanceof TypeError;
        if (isNetworkOrCors) {
          setMailStatus("Afsendelse fejlede: Netværk/CORS blokerede kaldet til Graph.");
        } else {
          setMailStatus(`Afsendelse fejlede: ${message}`);
        }
        setResourceStatus("Mail kunne ikke sendes fra portalen.");
      }
    });
  }

  if (mailContentTypeSelect) {
    mailContentTypeSelect.addEventListener("change", () => {
      updateMailPreviewVisibility();
    });
  }

  if (mailPreviewHtmlButton) {
    mailPreviewHtmlButton.addEventListener("click", () => {
      if (!mailHtmlPreview || !mailBodyInput) {
        return;
      }

      const isHtml = mailContentTypeSelect?.value === "HTML";
      if (!isHtml) {
        setMailStatus("Skift mailformat til HTML før forhåndsvisning.");
        return;
      }

      mailHtmlPreview.innerHTML = mailBodyInput.value || "";
      setMailStatus("HTML-forhåndsvisning opdateret.");
    });
  }

  if (mailRefreshInboxButton) {
    mailRefreshInboxButton.addEventListener("click", async () => {
      await loadMailFolders(true);
      await loadMessagesForSelectedFolder(false);
    });
  }

  if (mailCalendarRefreshButton) {
    mailCalendarRefreshButton.addEventListener("click", async () => {
      await loadCalendarEvents(true);
    });
  }

  if (mailCalendarPrevButton) {
    mailCalendarPrevButton.addEventListener("click", async () => {
      await moveCalendarMonth(-1);
    });
  }

  if (mailCalendarTodayButton) {
    mailCalendarTodayButton.addEventListener("click", async () => {
      calendarVisibleMonth = getCalendarStartOfMonth(new Date());
      calendarSelectedDate = new Date();
      await loadCalendarEvents(true);
    });
  }

  if (mailCalendarNextButton) {
    mailCalendarNextButton.addEventListener("click", async () => {
      await moveCalendarMonth(1);
    });
  }

  if (casesRefreshButton) {
    casesRefreshButton.addEventListener("click", async () => {
      await loadCases();
    });
  }

  updateCaseFilterButtons();

  if (teamsRefreshButton) {
    teamsRefreshButton.addEventListener("click", async () => {
      await loadTeamsList(true);
      await loadTeamChannels(false, selectedTeamId, selectedTeamName);
    });
  }

  document.addEventListener("keydown", (event) => {
    if (!mailView || mailView.classList.contains("is-hidden")) {
      return;
    }

    const isReadPaneShortcut = event.altKey && !event.shiftKey && !event.ctrlKey && !event.metaKey
      && String(event.key || "").toLowerCase() === "r";
    if (!isReadPaneShortcut) {
      return;
    }

    event.preventDefault();
    focusMailReadPane(true);
    setMailStatus("Fokus flyttet til læsefeltet (Alt+R).");
  });

  if (wordView) {
    wordView.addEventListener("click", (event) => {
      const toolButton = event.target.closest(".word-tool");
      if (!toolButton) {
        return;
      }

      const command = toolButton.getAttribute("data-word-cmd");
      if (!command) {
        return;
      }

      handleWordToolCommand(command);
    });
  }

  if (wordBlockStyleSelect) {
    wordBlockStyleSelect.addEventListener("change", () => {
      const selected = wordBlockStyleSelect.value || "P";
      runWordCommand("formatBlock", selected);
      setWordStatus(`Afsnitstype sat til ${selected}.`);
      if (wordEditor) {
        wordEditor.focus();
      }
    });
  }

  if (wordFontFamilySelect) {
    wordFontFamilySelect.addEventListener("change", () => {
      const selected = wordFontFamilySelect.value || "Calibri";
      runWordCommand("fontName", selected);
      setWordStatus(`Skrifttype sat til ${selected}.`);
      if (wordEditor) {
        wordEditor.focus();
      }
    });
  }

  if (wordFontSizeSelect) {
    wordFontSizeSelect.addEventListener("change", () => {
      const selected = wordFontSizeSelect.value || "12pt";
      applyWordFontSize(selected);
      setWordStatus("Skriftstørrelse opdateret.");
      if (wordEditor) {
        wordEditor.focus();
      }
    });
  }

  if (wordLineHeightSelect) {
    wordLineHeightSelect.addEventListener("change", () => {
      if (!wordEditor) {
        return;
      }

      const selected = wordLineHeightSelect.value || "1.5";
      wordEditor.style.lineHeight = selected;
      wordEditor.focus();
      setWordStatus(`Linjeafstand sat til ${selected}.`);
    });
  }

  if (wordTextColorInput) {
    wordTextColorInput.addEventListener("input", () => {
      const color = wordTextColorInput.value || "#1f1f1f";
      runWordCommand("foreColor", color);
      setWordStatus("Tekstfarve opdateret.");
    });
  }

  if (wordHighlightColorInput) {
    wordHighlightColorInput.addEventListener("input", () => {
      const color = wordHighlightColorInput.value || "#fff59d";
      runWordCommand("hiliteColor", color);
      setWordStatus("Markeringsfarve opdateret.");
    });
  }

  if (wordClearButton) {
    wordClearButton.addEventListener("click", () => {
      if (wordEditor) {
        wordEditor.innerHTML = "";
        wordEditor.focus();
      }
      setWordStatus("Dokument ryddet.");
    });
  }

  if (wordCopyButton) {
    wordCopyButton.addEventListener("click", async () => {
      const text = wordEditor?.innerText || "";
      if (!text.trim()) {
        setWordStatus("Der er ingen tekst at kopiere.");
        return;
      }

      try {
        await navigator.clipboard.writeText(text);
        setWordStatus("Tekst kopieret til udklipsholder.");
      } catch {
        setWordStatus("Kunne ikke kopiere automatisk. Markér teksten manuelt.");
      }
    });
  }

  if (wordSaveOneDriveButton) {
    wordSaveOneDriveButton.addEventListener("click", async () => {
      await saveWordToCloud("onedrive");
    });
  }

  if (wordSaveSharePointButton) {
    wordSaveSharePointButton.addEventListener("click", async () => {
      await saveWordToCloud("sharepoint");
    });
  }

  if (wordOpenButton && wordOpenFileInput) {
    wordOpenButton.addEventListener("click", () => {
      wordOpenFileInput.click();
    });

    wordOpenFileInput.addEventListener("change", async () => {
      const file = wordOpenFileInput.files?.[0];
      if (!file || !wordEditor) {
        return;
      }

      try {
        const parsedDocument = await parseWordDocumentFile(file);
        if (parsedDocument.mode === "html") {
          wordEditor.innerHTML = parsedDocument.content;
        } else {
          wordEditor.textContent = parsedDocument.content;
        }

        setWordStatus(parsedDocument.status);
        wordEditor.focus();
      } catch (error) {
        const message = error?.message || "Ukendt fejl";
        setWordStatus(`Kunne ikke åbne filen: ${message}`);
      } finally {
        wordOpenFileInput.value = "";
      }
    });
  }

  if (excelGrid) {
    excelGrid.addEventListener("focusin", (event) => {
      const input = event.target.closest(".excel-cell-input");
      if (!input) {
        return;
      }

      const row = Number.parseInt(input.getAttribute("data-row") || "0", 10);
      const col = Number.parseInt(input.getAttribute("data-col") || "0", 10);
      setExcelActiveCell(row, col);
      input.value = getExcelRaw(row, col);
    });

    excelGrid.addEventListener("input", (event) => {
      const input = event.target.closest(".excel-cell-input");
      if (!input) {
        return;
      }

      updateExcelCellFromInput(input);
      if (excelFormulaInput) {
        excelFormulaInput.value = input.value;
      }
    });

    excelGrid.addEventListener("blur", (event) => {
      const input = event.target.closest(".excel-cell-input");
      if (!input) {
        return;
      }

      const row = Number.parseInt(input.getAttribute("data-row") || "0", 10);
      const col = Number.parseInt(input.getAttribute("data-col") || "0", 10);
      updateExcelCellFromInput(input);

      if (ensureColumnFitsCellContent(row, col)) {
        renderExcelGrid();
      } else {
        refreshExcelGridDisplay();
      }
    }, true);

    excelGrid.addEventListener("keydown", (event) => {
      const input = event.target.closest(".excel-cell-input");
      if (!input) {
        return;
      }

      const row = Number.parseInt(input.getAttribute("data-row") || "0", 10);
      const col = Number.parseInt(input.getAttribute("data-col") || "0", 10);

      if (event.key === "Enter") {
        event.preventDefault();
        input.blur();
        const nextRow = Math.min(excelRows - 1, row + 1);
        const nextInput = excelGrid.querySelector(`.excel-cell-input[data-row=\"${nextRow}\"][data-col=\"${col}\"]`);
        if (nextInput) {
          nextInput.focus();
        }
      }
    });
  }

  if (excelFormulaInput) {
    excelFormulaInput.addEventListener("input", () => {
      const raw = excelFormulaInput.value;
      excelData[excelActiveRow][excelActiveCol] = raw;
      const activeInput = excelGrid?.querySelector(`.excel-cell-input[data-row=\"${excelActiveRow}\"][data-col=\"${excelActiveCol}\"]`);
      if (activeInput && document.activeElement === activeInput) {
        activeInput.value = raw;
      }
    });

    excelFormulaInput.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") {
        return;
      }

      event.preventDefault();
      excelData[excelActiveRow][excelActiveCol] = excelFormulaInput.value;
      refreshExcelGridDisplay();
      const activeInput = excelGrid?.querySelector(`.excel-cell-input[data-row=\"${excelActiveRow}\"][data-col=\"${excelActiveCol}\"]`);
      if (activeInput) {
        activeInput.focus();
      }
      setExcelStatus("Formel opdateret.");
    });
  }

  if (excelFontFamilySelect) {
    excelFontFamilySelect.addEventListener("change", () => {
      setExcelCellFormat(excelActiveRow, excelActiveCol, {
        fontFamily: excelFontFamilySelect.value || excelDefaultFontFamily,
      });
      refreshExcelGridDisplay();
      setExcelStatus("Skrifttype opdateret for aktiv celle.");
    });
  }

  if (excelFontSizeSelect) {
    excelFontSizeSelect.addEventListener("change", () => {
      setExcelCellFormat(excelActiveRow, excelActiveCol, {
        fontSize: excelFontSizeSelect.value || excelDefaultFontSize,
      });
      refreshExcelGridDisplay();
      setExcelStatus("Skriftstoerrelse opdateret for aktiv celle.");
    });
  }

  if (excelBoldButton) {
    excelBoldButton.addEventListener("click", () => {
      const current = getExcelCellFormat(excelActiveRow, excelActiveCol);
      setExcelCellFormat(excelActiveRow, excelActiveCol, {
        isBold: !current.isBold,
      });
      refreshExcelGridDisplay();
      syncExcelFormattingControls();
      setExcelStatus("Fed skrift er opdateret for aktiv celle.");
    });
  }

  if (excelTextColorInput) {
    excelTextColorInput.addEventListener("input", () => {
      setExcelCellFormat(excelActiveRow, excelActiveCol, {
        textColor: excelTextColorInput.value || excelDefaultTextColor,
      });
      refreshExcelGridDisplay();
      setExcelStatus("Tekstfarve opdateret for aktiv celle.");
    });
  }

  if (excelFillColorInput) {
    excelFillColorInput.addEventListener("input", () => {
      setExcelCellFormat(excelActiveRow, excelActiveCol, {
        fillColor: excelFillColorInput.value || excelDefaultFillColor,
      });
      refreshExcelGridDisplay();
      setExcelStatus("Baggrundsfarve opdateret for aktiv celle.");
    });
  }

  if (excelBorderStyleSelect) {
    excelBorderStyleSelect.addEventListener("change", () => {
      setExcelCellFormat(excelActiveRow, excelActiveCol, {
        borderStyle: excelBorderStyleSelect.value === "all" ? "all" : "none",
      });
      refreshExcelGridDisplay();
      setExcelStatus("Kant opdateret for aktiv celle.");
    });
  }

  function applyActiveColumnWidth() {
    if (!excelColWidthInput) {
      return;
    }

    const width = clampNumber(excelColWidthInput.value, 60, 720, excelDefaultColWidth);
    setExcelColumnWidth(excelActiveCol, width);
    renderExcelGrid();
    setExcelStatus(`Kolonnebredde sat til ${width}px for kolonne ${colToLabel(excelActiveCol)}.`);
  }

  function applyAutoFitActiveColumn() {
    const width = autoFitColumn(excelActiveCol);
    renderExcelGrid();
    setExcelStatus(`Kolonne ${colToLabel(excelActiveCol)} auto-tilpasset til ${width}px.`);
  }

  function applyAutoFitAllColumns() {
    for (let col = 0; col < excelCols; col += 1) {
      autoFitColumn(col);
    }

    renderExcelGrid();
    setExcelStatus("Alle kolonner er auto-tilpasset til indhold.");
  }

  function applyActiveRowHeight() {
    if (!excelRowHeightInput) {
      return;
    }

    const height = clampNumber(excelRowHeightInput.value, 28, 120, excelDefaultRowHeight);
    setExcelRowHeight(excelActiveRow, height);
    renderExcelGrid();
    setExcelStatus(`Raekkehoejde sat til ${height}px for raekke ${excelActiveRow + 1}.`);
  }

  if (excelApplyColWidthButton) {
    excelApplyColWidthButton.addEventListener("click", () => {
      applyActiveColumnWidth();
    });
  }

  if (excelAutoFitColButton) {
    excelAutoFitColButton.addEventListener("click", () => {
      applyAutoFitActiveColumn();
    });
  }

  if (excelAutoFitAllButton) {
    excelAutoFitAllButton.addEventListener("click", () => {
      applyAutoFitAllColumns();
    });
  }

  if (excelApplyRowHeightButton) {
    excelApplyRowHeightButton.addEventListener("click", () => {
      applyActiveRowHeight();
    });
  }

  if (excelColWidthInput) {
    excelColWidthInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        applyActiveColumnWidth();
      }
    });
  }

  if (excelRowHeightInput) {
    excelRowHeightInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        applyActiveRowHeight();
      }
    });
  }

  if (excelAddRowButton) {
    excelAddRowButton.addEventListener("click", () => {
      excelRows += 1;
      ensureExcelDataShape();
      renderExcelGrid();
      setExcelStatus(`Række tilføjet. Arkstørrelse: ${excelRows} x ${excelCols}.`);
    });
  }

  if (excelAddColButton) {
    excelAddColButton.addEventListener("click", () => {
      excelCols += 1;
      ensureExcelDataShape();
      renderExcelGrid();
      setExcelStatus(`Kolonne tilføjet. Arkstørrelse: ${excelRows} x ${excelCols}.`);
    });
  }

  if (excelSaveOneDriveButton) {
    excelSaveOneDriveButton.addEventListener("click", async () => {
      await saveExcelToCloud("onedrive");
    });
  }

  if (excelSaveSharePointButton) {
    excelSaveSharePointButton.addEventListener("click", async () => {
      await saveExcelToCloud("sharepoint");
    });
  }

  if (excelClearButton) {
    excelClearButton.addEventListener("click", () => {
      excelData = createExcelData(excelRows, excelCols);
      excelCellFormats = {};
      excelColWidths = {};
      excelRowHeights = {};
      excelActiveRow = 0;
      excelActiveCol = 0;
      renderExcelGrid();
      setExcelStatus("Arket er ryddet.");
    });
  }

  if (excelOpenButton && excelOpenFileInput) {
    excelOpenButton.addEventListener("click", () => {
      excelOpenFileInput.click();
    });

    excelOpenFileInput.addEventListener("change", async () => {
      const file = excelOpenFileInput.files?.[0];
      if (!file) {
        return;
      }

      try {
        const lowerName = String(file.name || "").toLowerCase();
        const mime = String(file.type || "").toLowerCase();
        const isDelimitedText = lowerName.endsWith(".csv")
          || lowerName.endsWith(".tsv")
          || mime === "text/csv"
          || mime === "text/tab-separated-values";

        if (isDelimitedText) {
          const text = await decodeTextFile(file);
          const delimiter = lowerName.endsWith(".tsv") ? "\t" : detectDelimitedTextDelimiter(text);
          const parsedRows = parseCsv(text, delimiter);
          applyExcelRows(parsedRows);
          setExcelStatus("CSV/TSV åbnet i regnearket.");
        } else {
          if (!hasXlsxRuntime()) {
            throw new Error("XLSX bibliotek mangler i portalen.");
          }

          const arrayBuffer = await file.arrayBuffer();
          const workbook = window.XLSX.read(arrayBuffer, {
            type: "array",
            cellFormula: true,
            cellDates: true,
            cellNF: true,
            raw: false,
            sheetStubs: true,
          });
          const sheetName = workbook.SheetNames.find((name) => workbook.Sheets?.[name]?.["!ref"]) || workbook.SheetNames[0];
          const parsedRows = getExcelRowsFromWorksheet(workbook, sheetName);
          applyExcelRows(parsedRows);
          setExcelStatus(`Excel-fil åbnet i regnearket (${sheetName || "Ark1"}).`);
        }
      } catch (error) {
        const message = error?.message || "Ukendt fejl";
        setExcelStatus(`Kunne ikke åbne filen: ${message}`);
      } finally {
        excelOpenFileInput.value = "";
      }
    });
  }

  if (mailReset) {
    mailReset.addEventListener("click", () => {
      if (mailForm) {
        mailForm.reset();
      }
      if (mailHtmlPreview) {
        mailHtmlPreview.textContent = "";
      }
      updateMailPreviewVisibility();
      setMailStatus("Mailformularen er nulstillet.");
    });
  }

  if (sharepointRefresh) {
    sharepointRefresh.addEventListener("click", () => {
      if (mirrorMode === "onedrive") {
        loadOnedrivePortal(true, sharepointCurrentPath);
        return;
      }

      loadSharepointPortal(true, sharepointCurrentPath);
    });
  }

  if (sharepointBack) {
    sharepointBack.addEventListener("click", () => {
      const parentPath = getParentPath(sharepointCurrentPath);
      if (parentPath === sharepointCurrentPath) {
        setSharepointStatus("Du er allerede i rodmappen.");
        return;
      }

      if (mirrorMode === "onedrive") {
        loadOnedrivePortal(false, parentPath);
        return;
      }

      loadSharepointPortal(false, parentPath);
    });
  }

  function initializeCalendarUi() {
    if (!mailCalendarGrid || !mailCalendarAgenda) {
      return;
    }

    setCalendarViewMode(calendarViewMode);
    setCalendarDraftFromDate(calendarSelectedDate, null);
    if (!calendarLoading) {
      setCalendarStatus("Kalender klar. Tryk Opdater kalender for at hente aftaler.");
    }
  }

  async function bootstrapAuth() {
    if (!msalInstance && isLoopback) {
      setAuthStatus("MSAL bibliotek kunne ikke indlæses. Kontrollér internetforbindelse.");
      if (loginButton) {
        loginButton.disabled = true;
      }
      return;
    }

    if (!isLoopback) {
      if (loginButton) {
        loginButton.disabled = false;
      }

      if (msalInstance) {
        try {
          await msalInstance.initialize();
          isMsalReady = true;
          const redirectResult = await msalInstance.handleRedirectPromise();
          if (redirectResult?.account) {
            msalInstance.setActiveAccount(redirectResult.account);
          }
        } catch {
          // Cloud portal can continue with server-side auth if MSAL init fails.
          isMsalReady = false;
        }
      }

      await handleAuthentication();
      await resumePendingGraphActionIfNeeded();
      return;
    }

    try {
      setAuthStatus("Initialiserer login...");
      await msalInstance.initialize();
      isMsalReady = true;

      if (loginButton) {
        loginButton.disabled = false;
      }

      await handleAuthentication();
      await resumePendingGraphActionIfNeeded();
    } catch (error) {
      setAuthStatus(`Loginfejl: ${error?.message || "Ukendt fejl"}`);
      if (loginButton) {
        loginButton.disabled = true;
      }
    }
  }

  initializeCalendarUi();
  bootstrapAuth();
});
