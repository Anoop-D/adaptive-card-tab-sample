import * as Express from "express";
import * as http from "http";
import * as path from "path";
import * as morgan from "morgan";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import {
  BotFrameworkAdapter,
  ConversationState,
  MemoryStorage,
} from "botbuilder";
import * as debug from "debug";
import { renderFile } from "ejs";
import { AcPrototypeBot } from "./acPrototypeBot/AcPrototypeBot";

// Initialize debug logging module
const log = debug("msteams");

log(`Initializing Microsoft Teams Express hosted App...`);

// Initialize dotenv, to use .env file settings if existing
// tslint:disable-next-line:no-var-requires
require("dotenv").config();

// The import of components has to be done AFTER the dotenv config
import * as allComponents from "./TeamsAppsComponents";

// Create the Express webserver
const express = Express();
const port = process.env.port || process.env.PORT || 3007;

const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights. See https://aka.ms/bottelemetry for telemetry
  //       configuration instructions.
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
  // Clear out state
  await conversationState.delete(context);
};

// Define a state store for your bot. See https://aka.ms/about-bot-state to learn more about using MemoryStorage.
// A bot requires a state store to persist the dialog and user state between messages.

// For local development, in-memory storage is used.
// CAUTION: The Memory Storage used here is for local bot debugging only. When the bot
// is restarted, anything stored in memory will be gone.
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

const bot = new AcPrototypeBot(memoryStorage, conversationState);

// Inject the raw request body onto the request object
express.use(
  Express.json({
    verify: (req, res, buf: Buffer, encoding: string): void => {
      (req as any).rawBody = buf.toString();
    },
  })
);
express.use(Express.urlencoded({ extended: true }));

express.engine("html", renderFile);
express.set("view engine", "ejs");

// Express configuration
express.set("views", path.join(__dirname, "/"));

// Add simple logging
express.use(morgan("tiny"));

// Listen for incoming requests.
express.post("/api/messages", (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// Add /scripts and /assets as static folders
express.use("/scripts", Express.static(path.join(__dirname, "web/scripts")));
express.use("/assets", Express.static(path.join(__dirname, "web/assets")));

express.post("/acPrototypeTab/redirect.html", (req, res) => {
  const authResponse = req.body;
  const state = Math.random().toString(); // _guid() is a helper function in the sample
  memoryStorage.write({
    [state]: {
      idToken: authResponse.id_token,
      accessToken: authResponse.access_token,
      tokenType: authResponse.token_type,
      expiresIn: authResponse.expires_in,
      scope: authResponse.scope,
    },
  });
  res.render(path.join(__dirname, "web/acPrototypeTab/redirect.ejs"), {
    state,
    error: authResponse.error,
  });
});

// routing for bots, connectors and incoming web hooks - based on the decorators
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsApiRouter(allComponents));

// routing for pages for tabs and connector configuration
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(
  MsTeamsPageRouter({
    root: path.join(__dirname, "web/"),
    components: allComponents,
  })
);

// Set default web page
express.use(
  "/",
  Express.static(path.join(__dirname, "web/"), {
    index: "index.html",
  })
);

// Set the port
express.set("port", port);

// Start the webserver
http.createServer(express).listen(port, () => {
  log(`Server running on ${port}`);
});
