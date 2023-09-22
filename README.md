# Create Azure Bot service

- Go to the https://portal.azure.com and create an Azure Bot service
- Open **Configuration** tab
- Put your https Messaging endpoint
- Save Microsoft App ID
- Go to Manage Password after create and save Microsoft App Password
- Go to Channels and connect **Microsoft Teams**

# MS Teams App

- open manifest.json from ./manifest folder
- replace id to randrom guid
- (optional) update packageName, developer, name, description, icons(files in the same folder)
- replace botId to your MicrosoftAppId
- replace webApplicationInfo.id to your MicrosoftAppId
- complress to zip archive these files (manifest.json, color.png, outline.png)
- go to [MS teams](https://teams.microsoft.com)
- click three dots new the team and open **Manage team** otion
- go to apps section and click **upload a custom app** choose zip archive

# MS Teams App

- open ./src/index.ts
- replace **MicrosoftAppId** and **MicrosoftAppPassword**
- replace **userTeamsId** and **userTeamsName** (write smth ti chat, bot will catch it and using debugger save this data from **"context.activity.from"**)
- replace **conversationId** from **"context.activity.conversation.id"**
- -replace **channelId** from **"context.activity"** or from teams url
  teams.microsoft.com/\_?culture=en&country=ua#/conversations/**channel-name**?threadId=**channelId**&ctx=channel
