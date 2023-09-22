import {
  Channels,
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  ConfigurationBotFrameworkAuthenticationOptions,
  ConversationParameters,
  ConversationReference,
  ConversationState,
  MemoryStorage,
  Mention,
  MessageFactory,
  TeamsChannelData,
  UserState,
} from "botbuilder";
import express from "express";
import { DialogBot } from "./bot";

const expressApp = express();

expressApp.use(express.json());
expressApp.use(express.urlencoded({ extended: true }));

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId: "MicrosoftAppId",
  MicrosoftAppPassword: "MicrosoftAppPassword",
} as ConfigurationBotFrameworkAuthenticationOptions);

const adapter = new CloudAdapter(botFrameworkAuthentication);

const onTurnErrorHandler = async (context: any, error: any) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity("The bot encountered an error or bug.");
  await context.sendActivity(
    "To continue to run this bot, please fix the bot source code."
  );
};

adapter.onTurnError = onTurnErrorHandler;

const conversationState = new ConversationState(new MemoryStorage());
const userState = new UserState(new MemoryStorage());
const bot = new DialogBot(conversationState, userState);

expressApp.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await bot.run(context);
  });
});

expressApp.get("/create", async (req, res) => {
  const mention = {
    mentioned: {
      id: "userTeamsId",
      name: "userTeamsName",
    },
    text: `<at>userTeamsName</at>`,
    type: "mention",
  } as Mention;

  // Returns a simple text message.
  const replyActivity = MessageFactory.text(
    `Hello ${mention.text} you have an issue`
  );
  replyActivity.entities = [mention];

  const parameters = {
    isGroup: true,
    channelData: {
      channel: {
        id: "channelId",
      },
    } as TeamsChannelData,
    activity: replyActivity,
  } as ConversationParameters;

  await adapter.createConversationAsync(
    "MicrosoftAppId",
    Channels.Msteams,
    "https://smba.trafficmanager.net/emea/",
    "https://api.botframework.com",
    parameters,
    async () => {}
  );

  res.setHeader("Content-Type", "text/html");
  res.send(
    "<html><body><h1>Proactive messages have been sent.</h1></body></html>"
  );
});

expressApp.get("/continue", async (req, res) => {
  const conversationReference = {
    channelId: "msteams",
    serviceUrl: "https://smba.trafficmanager.net/emea/",
    conversation: {
      isGroup: true,
      conversationType: "channel",
      id: "conversationId",
    },
  } as Partial<ConversationReference>;
  await adapter.continueConversationAsync(
    "MicrosoftAppId",
    conversationReference,
    async (context) => {
      await context.sendActivity(
        `send message to thread${new Date().toISOString()}`
      );
    }
  );
  res.send();
});

// Start the webserver
expressApp.listen(3000, () => {
  console.log(`Server running on 3000`);
});
