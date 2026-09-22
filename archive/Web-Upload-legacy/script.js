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
  const sharepointReadScopes = ["User.Read", "Sites.Read.All", "Files.Read.All"];
  const mailReadWriteScopes = ["User.Read", "Mail.ReadWrite"];
  const graphWriteScopes = ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"];
  const teamsReadScopes = ["User.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All"];
  let sharepointSiteId = null;
  let sharepointDriveId = null;
  let sharepointDriveName = sharepointLibraryName;
  let sharepointCurrentPath = "";
  let mirrorMode = "sharepoint";
  let activeFormType = null;
  let joinedTeams = [];
  let selectedTeamId = "";
  let selectedTeamName = "";
  let inboxMessages = [];
  let mailboxFolders = [];
  let selectedMailFolderId = "__all__";
  let selectedMailFolderName = "Alle mails";
  let selectedInboxMessageId = "";
  let focusReadPaneAfterLoad = false;
  let markMessageAsReadAfterLoadId = "";

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

  async function loadTeamsList(interactiveAllowed) {
    if (!teamsView) {
      return;
    }

    setTeamsStatus("Henter Teams...");
    if (teamsList) {
      teamsList.textContent = "";
    }

    try {
      let accessToken = await acquireGraphToken(interactiveAllowed, false, teamsReadScopes);
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
      detailTextElement.value = "";
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
    return text.includes("403") || text.includes("access denied") || text.includes("erroraccessdenied");
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
      let accessToken = await acquireGraphToken(interactiveAllowed, false, mailGraphScopes);
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

          accessToken = await acquireGraphToken(true, true, mailGraphScopes);
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

    const preview = document.createElement("p");
    preview.className = "mail-detail-preview";
    preview.textContent = message.bodyPreview || "Ingen forhåndsvisning tilgængelig.";

    const textBody = document.createElement("textarea");
    textBody.className = "mail-detail-text";
    textBody.id = "mail-detail-text";
    textBody.readOnly = true;
    textBody.setAttribute("aria-label", "Mailtekst");
    textBody.value = message._detailBodyText || message.bodyPreview || "Ingen indhold tilgængelig.";

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
    mailInboxDetail.appendChild(preview);
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

    hidePortalViews();
    mailView.classList.remove("is-hidden");
    setMailStatus("Udfyld mailfelterne og tryk Send mail.");
    setResourceStatus("Mailformular er åbnet i portalen.");
    updateMailPreviewVisibility();
    loadMailFolders(true).catch(() => {});
    loadMessagesForSelectedFolder(false).catch(() => {});
    mailView.scrollIntoView({ behavior: "smooth", block: "start" });
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

  setAuthUi(false, null);
  initPortalInteractions();

  if (window.location.protocol === "file:") {
    setAuthStatus("Login kræver hosting på https:// eller http://localhost. Åbn ikke via file://.");
    if (loginButton) {
      loginButton.disabled = true;
    }
    return;
  }

  if (!window.msal || !window.msal.PublicClientApplication) {
    setAuthStatus("MSAL bibliotek kunne ikke indlæses. Kontrollér internetforbindelse.");
    if (loginButton) {
      loginButton.disabled = true;
    }
    return;
  }

  if (clientId === placeholderClientId) {
    setAuthStatus("Konfiguration mangler: indsæt jeres App (client) ID i script.js.");
    if (loginButton) {
      loginButton.disabled = true;
    }
    return;
  }

  const isLoopback = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";
  const canonicalLocalRedirectUri = "http://localhost:5500/index.html";
  const redirectUri = isLoopback ? canonicalLocalRedirectUri : window.location.origin + window.location.pathname;

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

  const msalInstance = new window.msal.PublicClientApplication(msalConfig);
  let isMsalReady = false;

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

  async function acquireGraphToken(interactiveAllowed, forceConsent = false, requestedScopes = mailGraphScopes) {
    const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
    if (!account) {
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
        try {
          const popupToken = await msalInstance.acquireTokenPopup({
            scopes: tokenRequest.scopes,
            prompt: forceConsent ? "consent" : undefined,
          });
          return popupToken.accessToken;
        } catch (interactiveError) {
          if (!isMsalPopupTimeoutError(interactiveError)) {
            throw interactiveError;
          }

          await msalInstance.acquireTokenRedirect({
            scopes: tokenRequest.scopes,
            prompt: forceConsent ? "consent" : undefined,
          });

          throw new Error("Omdirigerer til Microsoft for at fuldføre Teams-adgang...");
        }
      }

      throw error;
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
      const accessToken = await acquireGraphToken(interactiveAllowed, false, sharepointReadScopes);

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
      const message = error?.message || "Ukendt fejl";
      setSharepointStatus(`Kunne ikke hente SharePoint-indhold: ${message}`);
      clearSharepointList();
    }
  }

  async function loadOnedrivePortal(interactiveAllowed, requestedPath = sharepointCurrentPath) {
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
      const accessToken = await acquireGraphToken(interactiveAllowed, false, sharepointReadScopes);

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
      const message = error?.message || "Ukendt fejl";
      setSharepointStatus(`Kunne ikke hente OneDrive-indhold: ${message}`);
      clearSharepointList();
    }
  }

  if (loginButton) {
    loginButton.disabled = true;
  }

  async function handleAuthentication() {
    try {
      setAuthStatus("Kontrollerer login...");
      const redirectResult = await msalInstance.handleRedirectPromise();

      if (redirectResult?.account) {
        msalInstance.setActiveAccount(redirectResult.account);
      }

      const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;

      if (account) {
        setAuthUi(true, account);
        setAuthStatus("Logget ind.");
      } else {
        setAuthUi(false, null);
        setAuthStatus("Du skal logge ind for at bruge portalen.");
      }
    } catch (error) {
      setAuthUi(false, null);
      setAuthStatus(`Loginfejl: ${error?.message || "Ukendt fejl"}`);
    }
  }

  if (loginButton) {
    loginButton.addEventListener("click", () => {
      if (!isMsalReady) {
        setAuthStatus("Login initialiseres, prøv igen om et øjeblik.");
        return;
      }
      setAuthStatus("Sender dig til Microsoft login...");
      msalInstance.loginRedirect(loginRequest);
    });
  }

  if (logoutButton) {
    logoutButton.addEventListener("click", () => {
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

  async function bootstrapAuth() {
    try {
      setAuthStatus("Initialiserer login...");
      await msalInstance.initialize();
      isMsalReady = true;

      if (loginButton) {
        loginButton.disabled = false;
      }

      await handleAuthentication();
    } catch (error) {
      setAuthStatus(`Loginfejl: ${error?.message || "Ukendt fejl"}`);
      if (loginButton) {
        loginButton.disabled = true;
      }
    }
  }

  bootstrapAuth();
});
