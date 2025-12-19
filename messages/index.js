"use strict";

// We now also import 'TeamsInfo' to get user details
const { CloudAdapter, ConfigurationBotFrameworkAuthentication, TeamsInfo } = require("botbuilder");
const { DefaultAzureCredential } = require("@azure/identity");
const { LogsQueryClient, LogsQueryResultStatus } = require("@azure/monitor-query");
const fetch = require("node-fetch"); // Ensure fetch is available in Node 18 environments if not native

// ---------------- Bot auth (Single Tenant) ----------------
const botAuth = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId: process.env.MicrosoftAppId,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword,
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId, 
  MicrosoftAppType: "SingleTenant",
});
const adapter = new CloudAdapter(botAuth);

// ---------------- Helpers ----------------
// Formatting fix for Teams (adds two spaces for line breaks)
function stringifyLogsTable(result, tableIndex = 0, maxRows = 5, maxChars = 3500) {
  if (!result?.tables || result.tables.length <= tableIndex) return "(no tables)";
  const table = result.tables[tableIndex];
  if (!table?.rows?.length) return "(no rows)";
  // Add "  " (two spaces) to the end of the header
  const header = table.columns.map((c) => c.name).join(" | ") + "  ";
  // Add "  " (two spaces) to the end of each row
  const rows = table.rows.slice(0, maxRows).map((r) => r.map(String).join(" | ") + "  ");
  let text = [header, ...rows].join("\n"); 
  if (text.length > maxChars) text = text.slice(0, maxChars) + "\n…truncated…";
  return text;
}


// ====================================================================
// Sentinel Write-Back Helpers
// ====================================================================

/**
 * Safely gets the full Azure Resource Manager (ARM) ID for an incident number
 * by using the Sentinel REST API, requiring specific environment variables.
 */
async function getIncidentArmIdByRestApi(incidentNumber, context) {
    const subscriptionId = process.env.AZURE_SUBSCRIPTION_ID || process.env.SubscriptionId; // Support both naming conventions
    // Try to derive resource group/workspace if not explicitly set, or error out
    // Note: For best results, add WORKSPACE_ID to App Settings, but for ARM ID lookup we ideally need RG name.
    // If you haven't set AZURE_RESOURCE_GROUP in App Settings, this might fail unless we hardcode or discover it.
    // For this Lite version, we will assume the user sets these in App Settings or we use KQL to find it (harder).
    // Let's rely on the KQL query result in 'summarize' usually, but here we need it for 'comment'.
    // SIMPLIFICATION: We will try to find the incident via KQL first to get its ARM ID if possible, 
    // OR we ask the user to set "AZURE_RESOURCE_GROUP" and "WORKSPACE_NAME" in Configuration.
    
    // Fallback: If variables are missing, we can't build the URL.
    // Check main.json parameters? The template asks for WorkspaceID but not Name/RG.
    // FIX for "Deploy to Azure" users: We will skip the complex REST lookup if vars are missing and warn the user.
    const resourceGroupName = process.env.AZURE_RESOURCE_GROUP;
    const workspaceName = process.env.WORKSPACE_NAME;

    if (!subscriptionId || !resourceGroupName || !workspaceName) {
        throw new Error("Missing App Settings: AZURE_SUBSCRIPTION_ID, AZURE_RESOURCE_GROUP, WORKSPACE_NAME are required for Write-Back.");
    }

    const credential = new DefaultAzureCredential();
    const token = await credential.getToken("https://management.azure.com/.default");

    const apiVersion = "2024-03-01"; 
    
    const apiUrl = `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.OperationalInsights/workspaces/${workspaceName}/providers/Microsoft.SecurityInsights/incidents?api-version=${apiVersion}`;
    
    const filter = `$filter=properties/incidentNumber eq ${incidentNumber}`;
    const url = `${apiUrl}&${filter}`;

    context.log(`[arm] Looking up ARM ID for incident ${incidentNumber} at: ${url}`);

    const res = await fetch(url, {
        method: "GET",
        headers: { "Authorization": `Bearer ${token.token}` },
    });

    if (!res.ok) {
        const text = await res.text().catch(() => "");
        context.log.error("[arm] HTTP error:", res.status, res.statusText);
        throw new Error(`Failed to retrieve incident list for ARM ID lookup: ${res.status} ${res.statusText}`);
    }

    const json = await res.json();
    const incident = json.value?.[0]; 

    if (!incident || !incident.id) {
        throw new Error(`Incident ${incidentNumber} not found via REST API.`);
    }

    return incident.id;
}


/**
 * Posts a comment to a Sentinel incident using the REST API and the Function's Managed Identity.
 */
async function addSentinelCommentWithMsi(incidentArmId, commentText) {
  const commentId = `c-${Date.now()}`;
  const apiVersion = "2024-03-01"; 

  const baseUrl = "https://management.azure.com";
  const url = `${baseUrl}${incidentArmId}/comments/${commentId}?api-version=${apiVersion}`;

  const cred = new DefaultAzureCredential();
  const token = await cred.getToken("https://management.azure.com/.default");

  const body = {
    properties: {
      message: commentText
    }
  };

  const res = await fetch(url, {
    method: "PUT",
    headers: {
      "Authorization": `Bearer ${token.token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Sentinel comment API failed: ${res.status} ${res.statusText} - Body: ${txt.slice(0, 500)}`);
  }

  return true;
}

/**
 * Gets the full incident object and its etag.
 */
async function getIncidentObject(incidentArmId, context) {
  const apiVersion = "2024-03-01";
  const url = `https://management.azure.com${incidentArmId}?api-version=${apiVersion}`;

  const cred = new DefaultAzureCredential();
  const token = await cred.getToken("https://management.azure.com/.default");

  context.log(`[assign] GET ${url}`);
  const res = await fetch(url, {
    method: "GET",
    headers: { "Authorization": `Bearer ${token.token}` },
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Failed to get incident for update: ${res.status} ${res.statusText} - ${txt.slice(0, 200)}`);
  }

  const etag = res.headers.get("etag");
  const incidentJson = await res.json();
  
  if (!etag) {
    context.log.warn("[assign] No etag found on incident object. PUT may fail.");
  }

  return { incidentJson, etag };
}

/**
 * Assigns a Sentinel incident to a user using the REST API.
 */
async function assignIncidentToUser(incidentArmId, ownerObject, context) {
  // 1. Get the incident object and its etag
  const { incidentJson, etag } = await getIncidentObject(incidentArmId, context);

  // 2. Modify the owner
  incidentJson.properties.owner = ownerObject;
  // We must also update the status if it's 'New'
  if (incidentJson.properties.status === "New") {
    incidentJson.properties.status = "Active";
    context.log("[assign] Incident status was 'New', auto-setting to 'Active'.");
  }

  // 3. PUT the modified object back
  const apiVersion = "2024-03-01";
  const url = `https://management.azure.com${incidentArmId}?api-version=${apiVersion}`;
  
  const cred = new DefaultAzureCredential();
  const token = await cred.getToken("https://management.azure.com/.default");

  const headers = {
    "Authorization": `Bearer ${token.token}`,
    "Content-Type": "application/json"
  };

  // Add etag for concurrency control if it exists
  if (etag) {
    headers["If-Match"] = etag;
  }
  
  context.log(`[assign] PUT ${url}`);
  const res = await fetch(url, {
    method: "PUT",
    headers: headers,
    body: JSON.stringify(incidentJson)
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Sentinel assignment API failed: ${res.status} ${res.statusText} - ${txt.slice(0, 500)}`);
  }

  return true;
}

// ====================================================================
// END Sentinel Write-Back Helpers
// ====================================================================


// ---------------- KQL Query Function (Used by entities/summarize) ----------------
async function querySentinelByIncident(context, workspaceId, incidentNumber) {
  const credential = new DefaultAzureCredential();
  const client = new LogsQueryClient(credential);

  const kql = `
    let incidentNumber = toint('${incidentNumber}');
    SecurityIncident
    | where IncidentNumber == incidentNumber
    | top 1 by TimeGenerated desc
    | mv-expand AlertIds to typeof(string)
    | join kind=leftouter (
        SecurityAlert
        | extend AlertEntities = parse_json(Entities)
        | mv-expand AlertEntities
    ) on $left.AlertIds == $right.SystemAlertId
    | project
        IncidentNumber,
        Title,
        Description,
        Severity,
        Status,
        Owner,
        FirstActivityTime,
        LastActivityTime,
        TimeGenerated,
        AlertId = SystemAlertId,
        AlertName,
        ProductName,
        AlertEntities,
        AlertSeverity = column_ifexists('Severity1', 'Unknown')
  `;

  context.log("[sentinel] running KQL for incident", incidentNumber);
  const result = await client.queryWorkspace(workspaceId, kql, { duration: "P7D" });
  if (result.status !== LogsQueryResultStatus.Success) {
    throw new Error(`Logs query failed: ${result.status}`);
  }
  return result;
}

/**
 * Extracts and formats entities (Markdown Table)
 */
function extractAndFormatEntities(logsResult) {
  const allEntities = new Map();
  const table = logsResult?.tables?.[0];
  if (!table || !table.rows?.length) {
    return "No entities found for this incident.";
  }

  for (const row of table.rows) {
    const entity = row[12]; // AlertEntities column (Index 12)
    if (entity && entity.Type) { 
      const entityType = entity.Type;
      let entityName = '';
      if (entity.Address) {
        entityName = entity.Address;
      } else if (entity.HostName) {
        entityName = entity.HostName;
      } else if (entity.Name) {
        entityName = entity.Name;
      } else if (entity.ProcessId && entity.CommandLine) {
        entityName = `PID ${entity.ProcessId} (${entity.CommandLine})`;
      } else {
        entityName = JSON.stringify(entity);
      }
      
      if (entityName) {
        if (!allEntities.has(entityType)) {
          allEntities.set(entityType, new Set());
        }
        const shortName = entityName.length > 80 ? entityName.substring(0, 77) + '...' : entityName;
        allEntities.get(entityType).add(shortName);
      }
    }
  }

  if (allEntities.size === 0) {
    return "No entities found for this incident.";
  }

  let output = "**Entities found:**\n\n";

  for (const [type, entities] of allEntities.entries()) {
    output += `### ${type}\n`;
    output += `| Name | Type |\n`;
    output += `| :--- | :--- |\n`;
    for (const entityName of entities) {
      const safeEntityName = entityName.replace(/\|/g, '\\|');
      output += `| ${safeEntityName} | ${type} |\n`;
    }
    output += `\n`;
  }

  return output;
}

async function callAOAIChat(context, { endpoint, deployment, apiVersion, apiKey, systemPrompt, userPrompt }) {
  systemPrompt = (systemPrompt ?? "").toString();
  userPrompt = (userPrompt ?? "").toString();

  const url =
    `${endpoint.replace(/\/+$/, "")}` +
    `/openai/deployments/${encodeURIComponent(deployment)}/chat/completions?api-version=${encodeURIComponent(apiVersion)}`;

  const payload = {
    temperature: 0.2,
    max_tokens: 500,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
  };

  context.log("[aoai] POST", url);

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json", "api-key": apiKey },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    context.log.error("[aoai] HTTP error:", res.status, res.statusText);
    throw new Error(`AOAI ${res.status} ${res.statusText}`);
  }

  const json = await res.json();
  const content = json?.choices?.[0]?.message?.content;
  if (!content) {
    return "(no content)";
  }
  return content;
}

// ---------------- Azure Function entry ----------------
module.exports = function (context, req) {
  adapter.process(req, context.res, async (turn) => {
    if (turn.activity.type !== "message") {
      await turn.sendActivity(`Event received: ${turn.activity.type}`);
      return;
    }

    const raw = (turn.activity.text || "").trim();
    const text = raw.toLowerCase();
    context.log("[msg] raw:", raw);

    // ---- ping? ----
    if (text === "ping?" || text === "ping") {
      await turn.sendActivity("pong");
      return;
    }

    // ---- aoai? ----
    if (text === "aoai?" || text === "aoai") {
      const ok =
        !!process.env.AI_ENDPOINT &&
        !!process.env.AI_DEPLOYMENT &&
        !!process.env.AI_API_VERSION &&
        !!process.env.AI_API_KEY;
      const msg = ok
        ? `AOAI config OK\nendpoint: ${process.env.AI_ENDPOINT}\n` +
          `deployment: ${process.env.AI_DEPLOYMENT}\nversion: ${process.env.AI_API_VERSION}`
        : "AOAI config MISSING (set AI_ENDPOINT, AI_DEPLOYMENT, AI_API_VERSION, AI_API_KEY)";
      await turn.sendActivity(msg);
      return;
    }

    // ---- whoami ----
    if (text === "whoami") {
      const userName = turn.activity.from.name;
      const userId = turn.activity.from.id;
      await turn.sendActivity(`You are: ${userName}\nTeams ID: ${userId}`);
      return;
    }

    // ---- sentinel? (last 1h) ----
    if (text === "sentinel?" || text === "sentinel") {
      try {
        const workspaceId = process.env.WORKSPACE_ID;
        if (!workspaceId) throw new Error("WORKSPACE_ID is not set.");
        const client = new LogsQueryClient(new DefaultAzureCredential());
        const kql =
          "SecurityIncident | top 5 by TimeGenerated desc | project TimeGenerated, IncidentNumber, Title, Severity, Status";
        const result = await client.queryWorkspace(workspaceId, kql, { duration: "PT1H" });

        if (result.status === LogsQueryResultStatus.Success) {
          const table = stringifyLogsTable(result, 0, 5);
          await turn.sendActivity(`Sentinel incidents (last 1h):\n${table}`);
        } else {
          await turn.sendActivity(`Query failed: ${result.status}`);
        }
      } catch (err) {
        context.log.error("[sentinel] error:", err);
        await turn.sendActivity(`Error querying Sentinel: ${err.message}`);
      }
      return;
    }

    // ---- summarize <incidentId> [question] ----
    if (text.startsWith("summarize ")) {
      const parts = raw.split(/\s+/);
      const incidentId = parts[1];
      const question = raw.slice(raw.indexOf(incidentId) + incidentId.length).trim();

      if (!/^\d+$/.test(incidentId)) {
        await turn.sendActivity("Usage: summarize <incidentId> [optional question]");
        return;
      }

      try {
        const workspaceId = process.env.WORKSPACE_ID;
        if (!workspaceId) throw new Error("WORKSPACE_ID is not set.");

        const logsResult = await querySentinelByIncident(context, workspaceId, incidentId);
        const incidentAndAlertsTxt = stringifyLogsTable(logsResult, 0, 11, 4000);
        const entitiesTxt = extractAndFormatEntities(logsResult); 

        const aoaiCfg = {
          endpoint: process.env.AI_ENDPOINT,
          deployment: process.env.AI_DEPLOYMENT,
          apiVersion: process.env.AI_API_VERSION,
          apiKey: process.env.AI_API_KEY,
          systemPrompt:
            "You are a senior SOC analyst. Produce concise, actionable guidance. " +
            "Prefer bullet points. Keep to <= 12 lines. If information is missing, state clear assumptions.",
          userPrompt:
            `Summarize and advise on Sentinel Incident ${incidentId}.\n\n` +
            `=== INCIDENT AND ALERTS ===\n${incidentAndAlertsTxt}\n\n` +
            `=== KEY ENTITIES ===\n${entitiesTxt}\n\n` + 
            (question ? `Analyst question:\n${question}\n\n` : "") +
            "Return: 1) Brief summary, 2) Indicators, 3) Triage steps, 4) Remediation, 5) Next actions.",
        };

        const answer = await callAOAIChat(context, aoaiCfg);
        await turn.sendActivity(answer);
      } catch (err) {
        context.log.error("[summarize] error:", err);
        await turn.sendActivity(`Error in summarize: ${err.message}`);
      }
      return;
    }

    // ---- kql <incidentId> ----
    if (text.startsWith("kql ")) {
      const parts = raw.split(/\s+/);
      const incidentId = parts[1];

      if (!/^\d+$/.test(incidentId)) {
        await turn.sendActivity("Usage: kql <incidentId>");
        return;
      }

      try {
        const workspaceId = process.env.WORKSPACE_ID;
        if (!workspaceId) throw new Error("WORKSPACE_ID is not set.");
        
        await turn.sendActivity(`Getting incident ${incidentId} details for KQL generation...`);

        const logsResult = await querySentinelByIncident(context, workspaceId, incidentId);
        const entitiesTxt = extractAndFormatEntities(logsResult); 

        const aoaiCfg = {
          endpoint: process.env.AI_ENDPOINT,
          deployment: process.env.AI_DEPLOYMENT,
          apiVersion: process.env.AI_API_VERSION,
          apiKey: process.env.AI_API_KEY,
          systemPrompt:
            "You are a Kusto Query Language (KQL) expert for Microsoft Sentinel. " +
            "Your job is to provide a single, simple KQL query for further investigation. " +
            "You MUST NOT query the 'SecurityEvent' table. " +
            "Focus on tables like 'SecurityAlert', 'SigninLogs', 'AADNonInteractiveUserSignInLogs', 'CommonSecurityLog', or 'DeviceProcessEvents'. " +
            "You ONLY return the KQL query, inside a code block. Do not add any explanation.",
          userPrompt:
            "Based on the key entities from the incident below, write one simple KQL query to find *more* activity related to them. " +
            "For example, search for the users in 'SigninLogs' or the hosts in 'DeviceProcessEvents'.\n\n" +
            `=== KEY ENTITIES ===\n${entitiesTxt}\n\n` + 
            "Return ONLY the KQL query, inside a Markdown code block."
        };

        const answer = await callAOAIChat(context, aoaiCfg);
        await turn.sendActivity(answer);
      } catch (err) {
        context.log.error("[kql] error:", err);
        await turn.sendActivity(`Error in kql: ${err.message}`);
      }
      return;
    }

    // ---- assign-me <incidentId> ----
    if (text.startsWith("assign-me ")) {
      const parts = raw.split(/\s+/);
      const incidentId = parts[1];

      if (!/^\d+$/.test(incidentId)) {
        await turn.sendActivity("Usage: assign-me <incidentId>");
        return;
      }

      try {
        await turn.sendActivity(`Attempting to assign incident ${incidentId} to you...`);

        let member;
        try {
            member = await TeamsInfo.getMember(turn, turn.activity.from.id);
        } catch (e) {
            context.log.error("[assign-me] Failed to get Teams member details.", e.message);
            throw new Error("Could not get your AAD details from Teams. Make sure the bot has 'TeamsActivity.Read.User' API permissions in Azure.");
        }
        
        if (!member.aadObjectId) {
             throw new Error("Your AAD Object ID could not be found. Cannot assign incident.");
        }
        
        const ownerObject = {
            objectId: member.aadObjectId,
            email: member.email,
            name: member.name,
            userPrincipalName: member.userPrincipalName || member.email // Fallback
        };
        
        context.log(`[assign-me] Assigning to: ${ownerObject.name} (${ownerObject.objectId})`);

        const armId = await getIncidentArmIdByRestApi(incidentId, context);
        await assignIncidentToUser(armId, ownerObject, context);

        await turn.sendActivity(`✅ Incident **${incidentId}** is now assigned to you!`);
      } catch (err) {
        context.log.error("[assign-me] error:", err);
        await turn.sendActivity(`❌ Assignment error: ${err.message}`);
      }
      return;
    }

    // ---- entities <incidentId> ----
    if (text.startsWith("entities ")) {
      const parts = raw.split(/\s+/);
      const incidentId = parts[1];

      if (!/^\d+$/.test(incidentId)) {
        await turn.sendActivity("Usage: entities <incidentId>");
        return;
      }

      try {
        const workspaceId = process.env.WORKSPACE_ID;
        if (!workspaceId) throw new Error("WORKSPACE_ID is not set.");

        await turn.sendActivity(`Fetching entities for incident ${incidentId}...`);
        const logsResult = await querySentinelByIncident(context, workspaceId, incidentId);
        const entityList = extractAndFormatEntities(logsResult);

        await turn.sendActivity(entityList);
      } catch (err) {
        context.log.error("[entities] error:", err);
        await turn.sendActivity(`Error fetching entities: ${err.message}`);
      }
      return;
    }
    
    // ---- comment <incidentNumber> <text...> ----
    if (text.startsWith("comment ")) {
      const parts = raw.split(/\s+/);
      const inc = parts[1];
      const message = raw.slice(raw.indexOf(inc) + inc.length).trim();

      if (!/^\d+$/.test(inc) || message.length < 5) {
        await turn.sendActivity("Usage: comment <incidentNumber> <Your comment text (min 5 chars)>");
        return;
      }

      try {
        await turn.sendActivity(`Attempting to post comment to incident ${inc}...`);

        const armId = await getIncidentArmIdByRestApi(inc, context); 
        const fullComment = `[SOC Copilot Lite] ${message}`;
        await addSentinelCommentWithMsi(armId, fullComment);

        await turn.sendActivity(`✅ Comment successfully posted to Sentinel Incident **${inc}**!`);
      } catch (err) {
        context.log.error("[comment] error:", err);
        const trimmedError = err.message.substring(0, 300); 
        await turn.sendActivity(`❌ Comment error: ${trimmedError}`);
      }
      return;
    }

    // ---- default echo ----
    await turn.sendActivity(`Echo: ${raw}`);
  });
};