# SOC Copilot Lite (Teams Edition)

A lightweight Security Operations Center (SOC) chatbot for Microsoft Teams, powered by Azure OpenAI and Microsoft Sentinel.

## 🚀 Quick Deployment

Click the button below to deploy the infrastructure (Function App, OpenAI, Storage) directly to your Azure subscription.

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Ffkh090%2FSOC3Lite%2Fmain%2Finfrastructure%2Fmain.json)

## 📋 Prerequisites

1.  **Azure OpenAI Access:** You must have an active Azure subscription with access to OpenAI models.
2.  **Microsoft App ID:** You need to create an App Registration in Entra ID (formerly Azure AD) to get a `MicrosoftAppId` and `ClientSecret`.
3.  **Sentinel Workspace:** You need the `Workspace ID` of your Sentinel instance.

## 📦 Project Structure

* `infrastructure/`: Contains the Azure Resource Manager (ARM) template.
* `src/`: The Node.js source code for the bot.
* `teams-app/`: The manifest package to install inside Microsoft Teams.