const path = require('path');
const restify = require('restify');
const { BotFrameworkAdapter } = require('botbuilder');
const { ActivityTypes } = require('botbuilder-core');

const { QnABot } = require('./bots/QnABot');

const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

// Adapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
});

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
  const traceActivity = {
    type: ActivityTypes.Trace,
    timestamp: new Date(),
    name: 'onTurnError Trace',
    label: 'TurnError',
    value: `${error}`,
    valueType: 'https://www.botframework.com/schemas/error',
  };

  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendActivity(traceActivity);

  // Error message to user
  await context.sendActivity(`The bot encountered an error or bug.`);
  await context.sendActivity(
    `To continue to run this bot, please fix the bot source code.`
  );
};

const bot = new QnABot();

// HTTP server setup
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
  console.log(`\n${server.name} listening to ${server.url}`);
});

server.post('/api/messages', (req, res) => {
  // Route received a request to adapter for processing
  adapter.processActivity(req, res, async (turnContext) => {
    // route to bot activity handler.
    await bot.run(turnContext);
  });
});
