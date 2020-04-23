const { ActivityHandler, CardFactory } = require('botbuilder');
const { QnAMaker } = require('botbuilder-ai');
const WelcomeCard = require('./resources/welcomeCard.json');

class QnABot extends ActivityHandler {
  constructor() {
    super();

    try {
      this.qnaMaker = new QnAMaker({
        knowledgeBaseId: process.env.QnAKnowledgebaseId,
        endpointKey: process.env.QnAEndpointKey,
        host: process.env.QnAEndpointHostName,
      });
    } catch (err) {
      console.warn(
        `QnAMaker Exception: ${err} Check your QnAMaker configuration in .env`
      );
    }

    // Method for when a new user gets added, displays the WelcomeCard
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id !== context.activity.recipient.id) {
          const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
          await context.sendActivity({ attachments: [welcomeCard] });
          await context.sendActivity(
            'Hi! I am a chatbot from Kristiania University College, i will try to help you to the best of my ability.'
          );
          await context.sendActivity('What can i help you with?');
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // When a user sends a message, perform a call to the QnA Maker service to retrieve matching Question and Answer pairs.
    this.onMessage(async (context, next) => {
      if (
        !process.env.QnAKnowledgebaseId ||
        !process.env.QnAEndpointKey ||
        !process.env.QnAEndpointHostName
      ) {
        const unconfiguredQnaMessage =
          'NOTE: \r\n' +
          'QnA Maker is not configured. To enable all capabilities, add `QnAKnowledgebaseId`, `QnAEndpointKey` and `QnAEndpointHostName` to the .env file. \r\n' +
          'You may visit www.qnamaker.ai to create a QnA Maker knowledge base.';

        await context.sendActivity(unconfiguredQnaMessage);
      } else {
        console.log('Calling QnA Maker');

        const qnaResults = await this.qnaMaker.getAnswers(context);

        // If an answer was received from QnA Maker, send the answer back to the user.
        if (qnaResults[0]) {
          const {
            answer,
            context: { prompts },
          } = qnaResults[0];

          let reply;
          if (prompts.length) {
            const card = {
              type: 'AdaptiveCard',
              body: [
                {
                  type: 'TextBlock',
                  text: answer,
                  wrap: true,
                },
              ],
              actions: prompts.map(({ displayText }) => ({
                type: 'Action.Submit',
                title: displayText,
                data: displayText,
              })),
              $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
              version: '1.1',
            };

            reply = { attachments: [CardFactory.adaptiveCard(card)] };
          } else {
            reply = answer;
          }

          await context.sendActivity(reply);

          // If no answers were returned from QnA Maker, reply with help.
        } else {
          await context.sendActivity('No QnA Maker answers were found.');
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });
  }
}

module.exports.QnABot = QnABot;
