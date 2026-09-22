#!/usr/bin/env node
const http = require('http');
const { spawn } = require('child_process');
const { DefaultAzureCredential } = require('@azure/identity');
const { BlobServiceClient } = require('@azure/storage-blob');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT || 5500;
const ticketsFilePath = path.join(__dirname, '..', 'ITSupportSystem', 'data', 'tickets.json');
const supportSystemPath = path.join(__dirname, '..', 'ITSupportSystem');
const supportSettingsPath = process.env.SUPPORT_SETTINGS_PATH || path.join(supportSystemPath, 'appsettings.json');
const isProduction = String(process.env.NODE_ENV || '').toLowerCase() === 'production';
const authzDevBypass = String(process.env.AUTHZ_DEV_BYPASS || 'false').toLowerCase() === 'true';

const groupClaimTypes = new Set([
  'groups',
  'http://schemas.microsoft.com/ws/2008/06/identity/claims/groups',
  'http://schemas.microsoft.com/identity/claims/groups',
]);

function decodeClientPrincipal(req) {
  const encodedPrincipal = String(req.headers['x-ms-client-principal'] || '').trim();
  if (!encodedPrincipal) {
    return null;
  }

  try {
    const json = Buffer.from(encodedPrincipal, 'base64').toString('utf8');
    return JSON.parse(json);
  } catch {
    return null;
  }
}

function getAuthenticatedUser(req) {
  const principal = decodeClientPrincipal(req);
  const principalName = String(req.headers['x-ms-client-principal-name'] || '').trim();
  const principalId = String(req.headers['x-ms-client-principal-id'] || '').trim();

  if (!principal && !principalName && !principalId) {
    return null;
  }

  const claims = Array.isArray(principal?.claims) ? principal.claims : [];
  const findClaim = (type) => {
    const match = claims.find((claim) => String(claim?.typ || '').toLowerCase() === type.toLowerCase());
    return String(match?.val || '').trim();
  };
  const findClaimAny = (types) => {
    for (const type of types) {
      const value = findClaim(type);
      if (value) {
        return value;
      }
    }

    return '';
  };
  const collectEmailLikeClaims = () => {
    const emailClaimTypes = new Set([
      'preferred_username',
      'upn',
      'email',
      'emails',
      'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress',
      'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn',
    ]);

    const emailSet = new Set();
    for (const claim of claims) {
      const claimType = String(claim?.typ || '').trim().toLowerCase();
      if (!emailClaimTypes.has(claimType)) {
        continue;
      }

      const rawValue = String(claim?.val || '').trim();
      if (!rawValue) {
        continue;
      }

      rawValue
        .split(/[;,]/)
        .map((entry) => entry.trim())
        .filter(Boolean)
        .forEach((entry) => emailSet.add(entry));
    }

    if (principalName) {
      emailSet.add(principalName);
    }

    return Array.from(emailSet);
  };

  const groups = [];
  for (const claim of claims) {
    const claimType = String(claim?.typ || '').trim().toLowerCase();
    if (!groupClaimTypes.has(claimType)) {
      continue;
    }

    const raw = String(claim?.val || '').trim();
    if (!raw) {
      continue;
    }

    raw
      .split(',')
      .map((entry) => entry.trim())
      .filter(Boolean)
      .forEach((entry) => groups.push(entry));
  }

  const givenName = findClaimAny([
    'given_name',
    'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/givenname',
  ]);
  const familyName = findClaimAny([
    'family_name',
    'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/surname',
  ]);
  const fullNameFromParts = `${givenName} ${familyName}`.trim();
  const emailCandidates = collectEmailLikeClaims();
  const primaryEmail = findClaimAny([
    'preferred_username',
    'upn',
    'email',
    'emails',
    'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress',
    'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn',
  ]) || principalName;

  return {
    id: principalId || findClaim('http://schemas.microsoft.com/identity/claims/objectidentifier'),
    name: findClaimAny([
      'name',
      'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name',
    ]) || fullNameFromParts || principalName,
    email: primaryEmail,
    emailCandidates,
    provider: String(principal?.auth_typ || '').trim() || 'aad',
    groups,
  };
}

function getLocalForwardedUser(req) {
  const email = String(req.headers['x-dbs-user-email'] || '').trim();
  if (!email) {
    return null;
  }

  const name = String(req.headers['x-dbs-user-name'] || '').trim();
  const groupsHeader = String(req.headers['x-dbs-user-groups'] || '').trim();
  const groups = groupsHeader
    ? groupsHeader.split(',').map((entry) => entry.trim()).filter(Boolean)
    : [];
  const emailCandidatesHeader = String(req.headers['x-dbs-user-email-candidates'] || '').trim();
  const emailCandidates = emailCandidatesHeader
    ? emailCandidatesHeader.split(',').map((entry) => String(entry || '').trim()).filter(Boolean)
    : [];

  return {
    id: String(req.headers['x-dbs-user-id'] || '').trim(),
    name: name || email,
    email,
    emailCandidates,
    provider: 'local-msal',
    groups,
  };
}

function parseCsvSet(value) {
  return new Set(
    String(value || '')
      .split(',')
      .map((entry) => entry.trim().toLowerCase())
      .filter(Boolean)
  );
}

function parseRolePermissions(raw) {
  if (!raw || typeof raw !== 'object') {
    return null;
  }

  const result = {};
  for (const [role, permissions] of Object.entries(raw)) {
    if (!Array.isArray(permissions)) {
      continue;
    }

    result[String(role).toLowerCase()] = permissions
      .map((permission) => String(permission || '').trim())
      .filter(Boolean);
  }

  return result;
}

function parseRoleMappings(raw) {
  if (!raw || typeof raw !== 'object') {
    return null;
  }

  const result = {};
  for (const [role, entries] of Object.entries(raw)) {
    if (!Array.isArray(entries)) {
      continue;
    }

    result[String(role).toLowerCase()] = new Set(
      entries
        .map((entry) => String(entry || '').trim().toLowerCase())
        .filter(Boolean)
    );
  }

  return result;
}

function loadAuthorizationConfig() {
  const defaultMasterGroupIdFallback = 'a9de0c67-6e00-4321-afbe-751547f67f58';
  const defaultAgentEmailFallback = isProduction ? '' : 'teb@blind.dk';

  const defaultConfig = {
    defaultRole: 'requester',
    masterGroupNames: parseCsvSet(process.env.AUTHZ_MASTER_GROUP_NAMES || 'sg_itsupportsystem'),
    masterGroupIds: parseCsvSet(process.env.AUTHZ_MASTER_GROUP_IDS || defaultMasterGroupIdFallback),
    rolePermissions: {
      admin: ['*'],
      agent: [
        'portal.mail',
        'portal.word',
        'portal.excel',
        'portal.teams',
        'portal.sharepoint',
        'portal.onedrive',
        'portal.forms',
        'portal.cases.view',
        'portal.cases.reply',
        'portal.cases.close',
        'portal.cases.assign',
        'portal.cases.read.all',
      ],
      viewer: [
        'portal.forms',
        'portal.cases.view',
        'portal.cases.read.all',
      ],
      requester: [
        'portal.forms',
        'portal.cases.view',
        'portal.cases.read.own',
      ],
    },
    roleMappings: {
      admin: parseCsvSet(process.env.AUTHZ_ADMIN_EMAILS),
      agent: parseCsvSet(process.env.AUTHZ_AGENT_EMAILS || defaultAgentEmailFallback),
      viewer: parseCsvSet(process.env.AUTHZ_VIEWER_EMAILS),
      requester: parseCsvSet(process.env.AUTHZ_REQUESTER_EMAILS),
    },
  };

  const rawJson = String(process.env.AUTHZ_RULES_JSON || '').trim();
  if (!rawJson) {
    return defaultConfig;
  }

  try {
    const parsed = JSON.parse(rawJson);
    const parsedRolePermissions = parseRolePermissions(parsed?.rolePermissions);
    const parsedRoleMappings = parseRoleMappings(parsed?.roleMappings);
    const defaultRole = String(parsed?.defaultRole || defaultConfig.defaultRole).trim().toLowerCase() || 'requester';
    const masterGroupNames = Array.isArray(parsed?.masterGroupNames)
      ? new Set(parsed.masterGroupNames.map((entry) => String(entry || '').trim().toLowerCase()).filter(Boolean))
      : defaultConfig.masterGroupNames;
    const masterGroupIds = Array.isArray(parsed?.masterGroupIds)
      ? new Set(parsed.masterGroupIds.map((entry) => String(entry || '').trim().toLowerCase()).filter(Boolean))
      : defaultConfig.masterGroupIds;

    return {
      defaultRole,
      masterGroupNames,
      masterGroupIds,
      rolePermissions: parsedRolePermissions || defaultConfig.rolePermissions,
      roleMappings: parsedRoleMappings || defaultConfig.roleMappings,
    };
  } catch (error) {
    console.warn(`AUTHZ_RULES_JSON kunne ikke parses. Bruger fallback-konfiguration. ${error?.message || error}`);
    return defaultConfig;
  }
}

const authorizationConfig = loadAuthorizationConfig();

function resolveUserRole(user) {
  const email = String(user?.email || '').trim().toLowerCase();
  const groups = Array.isArray(user?.groups)
    ? user.groups.map((entry) => String(entry || '').trim().toLowerCase()).filter(Boolean)
    : [];

  const hasMasterGroup = groups.some((groupValue) =>
    authorizationConfig.masterGroupNames.has(groupValue)
    || authorizationConfig.masterGroupIds.has(groupValue)
  );

  if (hasMasterGroup) {
    return 'admin';
  }

  const roleMappings = authorizationConfig.roleMappings;

  if (email) {
    for (const role of Object.keys(roleMappings)) {
      if (roleMappings[role]?.has(email)) {
        return role;
      }
    }
  }

  return authorizationConfig.defaultRole;
}

function buildAuthorizationForUser(user) {
  const role = resolveUserRole(user);
  const permissions = authorizationConfig.rolePermissions[role] || authorizationConfig.rolePermissions.requester || [];
  return {
    role,
    permissions,
  };
}

function hasPermission(authorization, permission) {
  const permissionList = Array.isArray(authorization?.permissions) ? authorization.permissions : [];
  if (permissionList.includes('*')) {
    return true;
  }

  return permissionList.includes(permission);
}

function getRequestSecurityContext(req) {
  const user = getAuthenticatedUser(req);
  if (user) {
    return {
      user,
      authorization: buildAuthorizationForUser(user),
    };
  }

  const forwardedUser = getLocalForwardedUser(req);
  if (forwardedUser) {
    return {
      user: forwardedUser,
      authorization: buildAuthorizationForUser(forwardedUser),
    };
  }

  if (authzDevBypass) {
    const devUser = {
      id: 'local-dev',
      name: 'Local Developer',
      email: 'local@dev',
      provider: 'dev-bypass',
    };

    return {
      user: devUser,
      authorization: {
        role: 'admin',
        permissions: ['*'],
      },
    };
  }

  return null;
}

function denyForbidden(res, message = 'Du har ikke adgang til denne handling.') {
  res.writeHead(403, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify({ error: message }));
}

function denyUnauthorized(res) {
  res.writeHead(401, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify({ error: 'Ikke logget ind.' }));
}

function parseTicketDatabase(data) {
  const parsed = JSON.parse(String(data || '{}'));
  const tickets = Array.isArray(parsed?.Tickets) ? parsed.Tickets : [];
  const processedMessageIds = Array.isArray(parsed?.ProcessedMessageIds)
    ? parsed.ProcessedMessageIds
    : [];

  return {
    ...parsed,
    Tickets: tickets,
    ProcessedMessageIds: processedMessageIds,
  };
}

function normalizeEmail(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) {
    return '';
  }

  const mailtoRemoved = raw.startsWith('mailto:') ? raw.slice('mailto:'.length) : raw;
  const angleMatch = mailtoRemoved.match(/<\s*([^>\s]+@[^>\s]+)\s*>/);
  if (angleMatch?.[1]) {
    return angleMatch[1].trim().toLowerCase();
  }

  const directMatch = mailtoRemoved.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
  if (directMatch?.[0]) {
    return directMatch[0].trim().toLowerCase();
  }

  return mailtoRemoved;
}

function normalizeTicketSenderAddress(ticket) {
  return normalizeEmail(ticket?.SenderAddress || ticket?.senderAddress || '');
}

function tryBuildNameBasedEmailCandidate(userName, emailDomain) {
  const normalizedDomain = normalizeEmail(emailDomain);
  if (!normalizedDomain) {
    return '';
  }

  const parts = String(userName || '')
    .trim()
    .split(/\s+/)
    .map((part) => part.replace(/[^a-zA-Z0-9.-]/g, '').toLowerCase())
    .filter(Boolean);

  if (parts.length < 2) {
    return '';
  }

  const first = parts[0];
  const last = parts[parts.length - 1];
  if (!first || !last) {
    return '';
  }

  return `${first}.${last}@${normalizedDomain}`;
}

function getUserEmailCandidates(user) {
  const candidateSet = new Set();
  const add = (value) => {
    const normalized = normalizeEmail(value);
    if (normalized) {
      candidateSet.add(normalized);
    }
  };

  add(user?.email);
  if (Array.isArray(user?.emailCandidates)) {
    user.emailCandidates.forEach((candidate) => add(candidate));
  }

  const primaryEmail = normalizeEmail(user?.email);
  const domain = primaryEmail.includes('@') ? primaryEmail.split('@')[1] : '';
  const nameBasedCandidate = tryBuildNameBasedEmailCandidate(user?.name, domain);
  add(nameBasedCandidate);

  return candidateSet;
}

function canReadTicket(securityContext, ticket) {
  if (!securityContext?.authorization) {
    return false;
  }

  if (hasPermission(securityContext.authorization, 'portal.cases.read.all')) {
    return true;
  }

  if (!hasPermission(securityContext.authorization, 'portal.cases.read.own')) {
    return false;
  }

  const userEmailCandidates = getUserEmailCandidates(securityContext.user);
  if (userEmailCandidates.size === 0) {
    return false;
  }

  return userEmailCandidates.has(normalizeTicketSenderAddress(ticket));
}

function getTicketId(ticket) {
  return String(ticket?.TicketId || ticket?.ticketId || '').trim();
}

function findReadableTicketById(tickets, ticketId, securityContext) {
  const normalizedTarget = String(ticketId || '').trim();
  if (!normalizedTarget) {
    return null;
  }

  return tickets.find((ticket) => getTicketId(ticket) === normalizedTarget && canReadTicket(securityContext, ticket)) || null;
}

function buildTicketAssignmentIdentity(user) {
  const userEmailCandidates = Array.from(getUserEmailCandidates(user));
  const primaryEmail = userEmailCandidates[0] || normalizeEmail(user?.email);
  const name = String(user?.name || primaryEmail || '').trim();
  return {
    id: String(user?.id || '').trim(),
    name,
    email: primaryEmail,
  };
}

function resolveSupportSystemDllPath() {
  const fromEnv = String(process.env.SUPPORT_SYSTEM_DLL_PATH || '').trim();
  if (fromEnv) {
    return fromEnv;
  }

  const candidates = [
    path.join(supportSystemPath, 'bin', 'Debug', 'net8.0', 'ITSupportSystem.dll'),
    path.join(supportSystemPath, 'bin', 'Release', 'net8.0', 'ITSupportSystem.dll'),
    path.join(supportSystemPath, 'publish', 'ITSupportSystem.dll'),
    path.join(__dirname, '..', 'ITSupportSystem', 'publish', 'ITSupportSystem.dll')
  ];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return candidates[0];
}

const supportSystemDllPath = resolveSupportSystemDllPath();

function loadBlobStorageSettings() {
  const envAccountUrl = String(process.env.TICKET_BLOB_ACCOUNT_URL || '').trim();
  const envConnectionString = String(process.env.AZURE_STORAGE_CONNECTION_STRING || '').trim();
  const envContainerName = String(process.env.TICKET_BLOB_CONTAINER || '').trim();
  const envBlobName = String(process.env.TICKET_BLOB_NAME || '').trim();

  if (envConnectionString || envAccountUrl) {
    return {
      enabled: true,
      accountUrl: envAccountUrl,
      connectionString: envConnectionString,
      containerName: envContainerName || 'itsupport',
      blobName: envBlobName || 'tickets.json',
    };
  }

  if (!fs.existsSync(supportSettingsPath)) {
    return null;
  }

  try {
    const settings = JSON.parse(fs.readFileSync(supportSettingsPath, 'utf8'));
    const blob = settings?.BlobStorage;
    if (!blob?.Enabled || (!blob?.ConnectionString && !blob?.AccountUrl)) {
      return null;
    }

    return {
      enabled: true,
      accountUrl: String(blob.AccountUrl || '').trim(),
      connectionString: String(blob.ConnectionString || '').trim(),
      containerName: String(blob.ContainerName || 'itsupport').trim(),
      blobName: String(blob.BlobName || 'tickets.json').trim(),
    };
  } catch {
    return null;
  }
}

function createBlobServiceClient(blobSettings) {
  if (blobSettings.connectionString) {
    return BlobServiceClient.fromConnectionString(blobSettings.connectionString);
  }

  if (blobSettings.accountUrl) {
    return new BlobServiceClient(blobSettings.accountUrl, new DefaultAzureCredential());
  }

  return null;
}

async function loadTicketsJson() {
  const blobSettings = loadBlobStorageSettings();
  if (blobSettings?.enabled) {
    try {
      const blobServiceClient = createBlobServiceClient(blobSettings);
      if (!blobServiceClient) {
        throw new Error('Blob-klient kunne ikke oprettes.');
      }
      const containerClient = blobServiceClient.getContainerClient(blobSettings.containerName);
      const blobClient = containerClient.getBlobClient(blobSettings.blobName);

      if (await blobClient.exists()) {
        const download = await blobClient.download();
        return await streamToString(download.readableStreamBody);
      }
    } catch (error) {
      console.warn(`Kunne ikke læse tickets fra Azure Blob Storage. Fallback til lokal fil. ${error?.message || error}`);
    }
  }

  return fs.readFileSync(ticketsFilePath, 'utf8');
}

async function saveTicketsJson(content) {
  const blobSettings = loadBlobStorageSettings();
  const payload = String(content || '');

  if (blobSettings?.enabled) {
    const blobServiceClient = createBlobServiceClient(blobSettings);
    if (!blobServiceClient) {
      throw new Error('Blob-klient kunne ikke oprettes.');
    }

    const containerClient = blobServiceClient.getContainerClient(blobSettings.containerName);
    await containerClient.createIfNotExists();
    const blobClient = containerClient.getBlockBlobClient(blobSettings.blobName);
    await blobClient.upload(payload, Buffer.byteLength(payload), {
      blobHTTPHeaders: {
        blobContentType: 'application/json; charset=utf-8',
      },
    });
    return;
  }

  fs.mkdirSync(path.dirname(ticketsFilePath), { recursive: true });
  fs.writeFileSync(ticketsFilePath, payload, 'utf8');
}

function streamToString(readableStream) {
  return new Promise((resolve, reject) => {
    if (!readableStream) {
      resolve('');
      return;
    }

    const chunks = [];
    readableStream.on('data', (chunk) => {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    });
    readableStream.on('end', () => {
      resolve(Buffer.concat(chunks).toString('utf8'));
    });
    readableStream.on('error', reject);
  });
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';

    req.on('data', (chunk) => {
      body += chunk;
      if (body.length > 1024 * 1024) {
        reject(new Error('Request body er for stor.'));
      }
    });

    req.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        reject(new Error('Ugyldig JSON i request body.'));
      }
    });

    req.on('error', reject);
  });
}

function runSupportCommand(command, ticketId, message) {
  return new Promise((resolve, reject) => {
    const child = spawn('dotnet', [supportSystemDllPath, command, ticketId, message], {
      cwd: supportSystemPath,
      windowsHide: true,
    });

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (chunk) => {
      stdout += chunk.toString();
    });

    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString();
    });

    child.on('error', (error) => {
      reject(error);
    });

    child.on('close', (code) => {
      if (code === 0) {
        resolve({ stdout: stdout.trim() });
        return;
      }

      reject(new Error((stderr || stdout || `Kommando fejlede med exit code ${code}.`).trim()));
    });
  });
}

const server = http.createServer((req, res) => {
  const parsedUrl = url.parse(req.url);
  let pathname = parsedUrl.pathname;

  if (pathname === '/healthz') {
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: true }));
    return;
  }

  if (pathname === '/api/auth/me') {
    const user = getAuthenticatedUser(req);
    if (!user) {
      res.writeHead(401, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ authenticated: false }));
      return;
    }

    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ authenticated: true, user }));
    return;
  }

  if (pathname === '/api/authz/me') {
    const securityContext = getRequestSecurityContext(req);
    if (!securityContext) {
      res.writeHead(401, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ authenticated: false }));
      return;
    }

    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({
      authenticated: true,
      user: securityContext.user,
      authorization: securityContext.authorization,
    }));
    return;
  }

  if (pathname === '/api/tickets') {
    const securityContext = getRequestSecurityContext(req);
    if (!securityContext) {
      denyUnauthorized(res);
      return;
    }

    if (!hasPermission(securityContext.authorization, 'portal.cases.view')) {
      denyForbidden(res, 'Du har ikke adgang til at se sager.');
      return;
    }

    loadTicketsJson()
      .then((data) => {
        const db = parseTicketDatabase(data);
        const visibleTickets = db.Tickets.filter((ticket) => canReadTicket(securityContext, ticket));

        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({
          ...db,
          Tickets: visibleTickets,
        }));
      })
      .catch(() => {
        res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ error: 'Kunne ikke læse tickets' }));
      });

    return;
  }

  if (pathname === '/api/tickets/reply' && req.method === 'POST') {
    const securityContext = getRequestSecurityContext(req);
    if (!securityContext) {
      denyUnauthorized(res);
      return;
    }

    if (!hasPermission(securityContext.authorization, 'portal.cases.reply')) {
      denyForbidden(res, 'Du har ikke adgang til at svare på sager.');
      return;
    }

    readJsonBody(req)
      .then(async (payload) => {
        const ticketId = String(payload.ticketId || '').trim();
        const message = String(payload.message || '').trim();

        if (!ticketId || !message) {
          res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ error: 'ticketId og message skal udfyldes.' }));
          return;
        }

        const db = parseTicketDatabase(await loadTicketsJson());
        const ticket = findReadableTicketById(db.Tickets, ticketId, securityContext);
        if (!ticket) {
          denyForbidden(res, 'Du har ikke adgang til den valgte sag.');
          return;
        }

        const result = await runSupportCommand('reply', ticketId, message);
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ ok: true, output: result.stdout }));
      })
      .catch((error) => {
        res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ error: error.message || 'Kunne ikke sende svar.' }));
      });

    return;
  }

  if (pathname === '/api/tickets/close' && req.method === 'POST') {
    const securityContext = getRequestSecurityContext(req);
    if (!securityContext) {
      denyUnauthorized(res);
      return;
    }

    if (!hasPermission(securityContext.authorization, 'portal.cases.close')) {
      denyForbidden(res, 'Du har ikke adgang til at lukke sager.');
      return;
    }

    readJsonBody(req)
      .then(async (payload) => {
        const ticketId = String(payload.ticketId || '').trim();
        const message = String(payload.message || '').trim();

        if (!ticketId || !message) {
          res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ error: 'ticketId og message skal udfyldes.' }));
          return;
        }

        const db = parseTicketDatabase(await loadTicketsJson());
        const ticket = findReadableTicketById(db.Tickets, ticketId, securityContext);
        if (!ticket) {
          denyForbidden(res, 'Du har ikke adgang til den valgte sag.');
          return;
        }

        const result = await runSupportCommand('close', ticketId, message);
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ ok: true, output: result.stdout }));
      })
      .catch((error) => {
        res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ error: error.message || 'Kunne ikke lukke sag.' }));
      });

    return;
  }

  if (pathname === '/api/tickets/assign-self' && req.method === 'POST') {
    const securityContext = getRequestSecurityContext(req);
    if (!securityContext) {
      denyUnauthorized(res);
      return;
    }

    if (!hasPermission(securityContext.authorization, 'portal.cases.assign')) {
      denyForbidden(res, 'Du har ikke adgang til at tildele sager.');
      return;
    }

    readJsonBody(req)
      .then(async (payload) => {
        const ticketId = String(payload.ticketId || '').trim();
        if (!ticketId) {
          res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ error: 'ticketId skal udfyldes.' }));
          return;
        }

        const db = parseTicketDatabase(await loadTicketsJson());
        const ticket = findReadableTicketById(db.Tickets, ticketId, securityContext);
        if (!ticket) {
          denyForbidden(res, 'Du har ikke adgang til den valgte sag.');
          return;
        }

        const assignment = buildTicketAssignmentIdentity(securityContext.user);
        if (!assignment.email) {
          res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ error: 'Kunne ikke bestemme din e-mail til tildeling.' }));
          return;
        }

        const nowIso = new Date().toISOString();
        ticket.AssignedToId = assignment.id;
        ticket.AssignedToName = assignment.name;
        ticket.AssignedToEmail = assignment.email;
        ticket.AssignedAt = nowIso;
        ticket.UpdatedAt = nowIso;

        await saveTicketsJson(JSON.stringify(db, null, 2));

        const assigneeDisplayName = String(assignment.name || assignment.email || 'En supportmedarbejder').trim();
        const assignmentNotificationMessage = `${assigneeDisplayName} har taget din sag og arbejder pa den nu.`;
        let notifyWarning = '';

        try {
          await runSupportCommand('reply', ticketId, assignmentNotificationMessage);
        } catch (notifyError) {
          notifyWarning = notifyError?.message || 'Sagen blev tildelt, men notifikation til brugeren kunne ikke sendes.';
        }

        // Re-apply assignment after support command to guard against any downstream save path
        // that might not preserve assignment metadata yet.
        const refreshedDb = parseTicketDatabase(await loadTicketsJson());
        const refreshedTicket = Array.isArray(refreshedDb?.Tickets)
          ? refreshedDb.Tickets.find((entry) => getTicketId(entry) === ticketId)
          : null;

        if (refreshedTicket) {
          refreshedTicket.AssignedToId = assignment.id;
          refreshedTicket.AssignedToName = assignment.name;
          refreshedTicket.AssignedToEmail = assignment.email;
          refreshedTicket.AssignedAt = nowIso;
          refreshedTicket.UpdatedAt = nowIso;
          await saveTicketsJson(JSON.stringify(refreshedDb, null, 2));
        }

        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({
          ok: true,
          warning: notifyWarning || undefined,
          ticket: {
            ticketId: getTicketId(ticket),
            assignedToId: ticket.AssignedToId,
            assignedToName: ticket.AssignedToName,
            assignedToEmail: ticket.AssignedToEmail,
            assignedAt: ticket.AssignedAt,
          },
        }));
      })
      .catch((error) => {
        res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ error: error.message || 'Kunne ikke tildele sag.' }));
      });

    return;
  }
  
  if (pathname === '/') {
    pathname = '/index.html';
  }
  
  const filePath = path.join(__dirname, pathname);
  
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('404 Not Found');
      return;
    }
    
    let contentType = 'text/plain';
    if (filePath.endsWith('.html')) contentType = 'text/html';
    else if (filePath.endsWith('.css')) contentType = 'text/css';
    else if (filePath.endsWith('.js')) contentType = 'application/javascript';
    else if (filePath.endsWith('.json')) contentType = 'application/json';
    else if (filePath.endsWith('.png')) contentType = 'image/png';
    else if (filePath.endsWith('.jpg') || filePath.endsWith('.jpeg')) contentType = 'image/jpeg';
    else if (filePath.endsWith('.gif')) contentType = 'image/gif';
    
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
