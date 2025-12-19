# SOC Copilot Lite 🤖

**SOC Copilot Lite** is a serverless, AI-powered security assistant that integrates **Microsoft Sentinel** with **Microsoft Teams**. It allows SOC analysts to triage incidents, summarize alerts, and generate KQL queries directly from chat using Azure OpenAI (GPT-4o).

---

## 🚀 Features

* **⚡ Serverless:** Built on Azure Functions (Node.js) & Azure OpenAI.
* **🛡️ Sentinel Integrated:** Fetch incidents, get details, and run triage commands.
* **💬 Teams Native:** Chat directly with your SOC data.
* **🧠 AI Powered:** Uses GPT-4o to summarize complex alerts and generate KQL.

### 🤖 Available Commands
| Command | Description |
| :--- | :--- |
| `sentinel` | Lists the top 5 most recent incidents. |
| `summarize <IncidentID>` | analyzing the incident using AI and providing a summary. |
| `entities <IncidentID>` | Extracts IPs, Users, and URLs from the incident. |
| `kql <IncidentID>` | Generates a specific KQL query to investigate the incident. |
| `assign-me <IncidentID>` | Assigns the incident to you (the user chatting). |
| `comment <IncidentID> <msg>` | Posts a comment to the Sentinel incident. |
| `ping` | Checks if the bot is online. |

---

## 🛠️ Prerequisites

Before you deploy, make sure you have:
1.  **Azure Subscription** (with access to create Resource Groups and OpenAI resources).
2.  **Microsoft Sentinel Workspace** (already set up).
3.  **Microsoft Teams** (with permission to upload custom apps).
4.  **GitHub Account** (to fork or view this repo).

---

## 📦 Installation Guide

### Step 1: Create an App Registration (Identity)
The bot needs an identity to communicate with Teams.
1.  Go to **[Azure Portal > Microsoft Entra ID > App registrations](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps)**.
2.  Click **+ New registration**.
    * **Name:** `SOC Copilot Lite`
    * **Supported account types:** "Accounts in this organizational directory only (Single tenant)".
    * Click **Register**.
3.  Copy the **Application (client) ID** and save it for later.
4.  Go to **Certificates & secrets** -> **+ New client secret**.
    * Copy the **Secret Value** immediately (you won't see it again).

### Step 2: Deploy to Azure
Click the button below to deploy the entire infrastructure (Azure Function, OpenAI, Storage, etc.) automatically.

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Ffkh090%2FSOC3Lite%2Fmain%2Finfrastructure%2Fmain.json)

* **Resource Group:** Create a NEW one (e.g., `rg-soc-copilot`).
* **Region:** Select **Sweden Central** (Recommended for OpenAI quota availability).
* **App Name:** Enter a unique name (e.g., `soc-lite-yourname`).
* **Microsoft App ID / Password:** Paste the values from Step 1.
* **Sentinel Details:** Enter your existing Sentinel Resource Group and Workspace Name.

### Step 3: Grant Permissions (Important!)
After deployment finishes, you must grant the bot permission to access Sentinel and Teams.

**A. Sentinel Permissions (Managed Identity)**
1.  Go to your **Microsoft Sentinel Workspace**.
2.  Click **Access control (IAM)** -> **+ Add** -> **Add role assignment**.
3.  Select Role: **Microsoft Sentinel Responder** (allows reading and editing incidents).
4.  Assign to: **Managed Identity** -> **Function App** -> Select your new bot (`soc-lite-...`).
5.  Click **Review + assign**.

**B. Teams API Permissions (App Registration)**
1.  Go back to **Entra ID > App registrations > SOC Copilot Lite**.
2.  Click **API permissions** -> **+ Add a permission** -> **Microsoft Graph** -> **Delegated**.
3.  Add these permissions:
    * `User.Read`
    * `offline_access`
    * `openid`
    * `Chat.ReadWrite`
    * `TeamsActivity.Send`
4.  **Crucial:** Click **Grant admin consent for [Your Org]**.

### Step 4: Install in Teams
1.  Download the `app-package` folder from this repository to your computer.
2.  Open `manifest.json` in a text editor.
3.  Replace `{{BOT_ID}}` (in two places) with your **Application (client) ID** from Step 1.
4.  Select the 3 files inside the folder (`manifest.json`, `color.png`, `outline.png`) and **Zip them** into a file named `soc-copilot.zip`.
5.  Go to the **[Microsoft Teams Developer Portal](https://dev.teams.microsoft.com/apps)**.
6.  Click **Apps** -> **Manage your apps** -> **Import app**.
7.  Upload your `soc-copilot.zip`.

---

## 🚀 Usage
Once installed in Teams, simply start a chat with the bot and type `sentinel` to see your latest incidents!
