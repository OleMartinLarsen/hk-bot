const path = require('path');
const restify = require('restify');
const { BotFrameworkAdapter } = require('botbuilder');
const { ActivityTypes } = require('botbuilder-core');

// Bot
const { QnABot } = require('./bots/QnABot');

// Env
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

// Adapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
});

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
  // Create a trace activity that contains the error object
  const traceActivity = {
    type: ActivityTypes.Trace,
    timestamp: new Date(),
    name: 'onTurnError Trace',
    label: 'TurnError',
    value: `${error}`,
    valueType: 'https://www.botframework.com/schemas/error',
  };
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendActivity(traceActivity);

  // Send a message to the user
  await context.sendActivity(`The bot encountered an error or bug.`);
  await context.sendActivity(
    `To continue to run this bot, please fix the bot source code.`
  );
};

const bot = new QnABot();

// HTTP server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
  console.log(`\n${server.name} listening to ${server.url}`);
});

// Listen for incoming activities and route them to your bot main dialog.
server.post('/api/messages', (req, res) => {
  // Route received a request to adapter for processing
  adapter.processActivity(req, res, async (turnContext) => {
    // route to bot activity handler.
    await bot.run(turnContext);
  });
});
