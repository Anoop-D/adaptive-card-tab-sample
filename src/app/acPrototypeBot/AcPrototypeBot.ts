import {
  BotDeclaration,
  MessageExtensionDeclaration,
  IBot,
  PreventIframe,
} from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import {
  StatePropertyAccessor,
  CardFactory,
  TurnContext,
  MemoryStorage,
  ConversationState,
  ActivityTypes,
  InvokeResponse,
} from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import FlightItineraryCard from "./dialogs/FlightItineraryDialog";
import { TeamsContext, TeamsActivityProcessor } from "botbuilder-teams";
import AdminCard from "./dialogs/AdminCard";
import QuickActionCard from "./dialogs/QuickActionsCard";
import ManagerDashboardCard from "./dialogs/ManagerDashboard";

// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for acPrototype Bot
 */
@BotDeclaration(
  "/api/messages",
  new MemoryStorage(),
  process.env.MICROSOFT_APP_ID,
  process.env.MICROSOFT_APP_PASSWORD
)
@PreventIframe("/acPrototypeBot/acProtoBotTab.html")
export class AcPrototypeBot implements IBot {
  private readonly conversationState: ConversationState;
  private readonly dialogs: DialogSet;
  private dialogState: StatePropertyAccessor<DialogState>;
  private readonly activityProc = new TeamsActivityProcessor();

  /**
   * The constructor
   * @param conversationState
   */
  public constructor(conversationState: ConversationState) {
    this.conversationState = conversationState;
    this.dialogState = conversationState.createProperty("dialogState");
    this.dialogs = new DialogSet(this.dialogState);
    this.dialogs.add(new HelpDialog("help"));

    // Set up the Activity processing

    this.activityProc.messageActivityHandler = {
      // Incoming messages
      onMessage: async (context: TurnContext): Promise<void> => {
        // get the Microsoft Teams context, will be undefined if not in Microsoft Teams
        const teamsContext: TeamsContext = TeamsContext.from(context);

        // TODO: add your own bot logic in here
        switch (context.activity.type) {
          case ActivityTypes.Message:
            const text = teamsContext
              ? teamsContext.getActivityTextWithoutMentions().toLowerCase()
              : context.activity.text;

            if (text.startsWith("hello")) {
              await context.sendActivity("Oh, hello to you as well!");
              return;
            } else if (text.startsWith("help")) {
              const dc = await this.dialogs.createContext(context);
              await dc.beginDialog("help");
            } else if (text.includes("card")) {
              try {
                const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                await context.sendActivity({ attachments: [welcomeCard] });
              } catch (err) {
                console.log(err);
              }
            } else {
              console.log(text);
              await context.sendActivity(
                `I\'m terribly sorry, but my master hasn\'t trained me to do anything yet...`
              );
            }
            break;
          case ActivityTypes.Invoke:
          default:
            break;
        }

        // Save state changes
        return this.conversationState.saveChanges(context);
      },
    };

    this.activityProc.conversationUpdateActivityHandler = {
      onConversationUpdateActivity: async (
        context: TurnContext
      ): Promise<void> => {
        if (
          context.activity.membersAdded &&
          context.activity.membersAdded.length !== 0
        ) {
          for (const idx in context.activity.membersAdded) {
            if (
              context.activity.membersAdded[idx].id ===
              context.activity.recipient.id
            ) {
              const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
              await context.sendActivity({ attachments: [welcomeCard] });
            }
          }
        }
      },
    };

    // Message reactions in Microsoft Teams
    this.activityProc.messageReactionActivityHandler = {
      onMessageReaction: async (context: TurnContext): Promise<void> => {
        const added = context.activity.reactionsAdded;
        if (added && added[0]) {
          await context.sendActivity({
            textFormat: "xml",
            text: `That was an interesting reaction (<b>${added[0].type}</b>)`,
          });
        }
      },
    };

    this.activityProc.invokeActivityHandler = {
      onInvoke: async (context: TurnContext): Promise<InvokeResponse> => {
        // console.dir(context);
        const ctx: any = context;
        // console.log("entities", ctx.entities);
        // console.log("value", ctx.value);
        // console.log("channelData", ctx.channelData);

        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
        const adminCard = CardFactory.adaptiveCard(AdminCard);
        const quickActionsCard = CardFactory.adaptiveCard(QuickActionCard);
        const managerCard = CardFactory.adaptiveCard(ManagerDashboardCard);
        const taskModuleCard = CardFactory.adaptiveCard(FlightItineraryCard);
        // Return the specified task module response to the bot

        // tslint:disable-next-line: no-string-literal
        managerCard.content["$data"] = {
          creator: {
            name: ctx.activity.name,
            profileImage: "https://randomuser.me/api/portraits/women/32.jpg",
          },
        };
        let responseBody: any;

        const tabResponse: any = {
          tab: {
            type: "continue",
            value: {
              cards: [
                { card: quickActionsCard.content },
                { card: managerCard.content },
                { card: adminCard.content },
                { card: welcomeCard.content },
              ],
            },
          },
        };

        const tabSubmitResponse: any = {
          tab: {
            type: "continue",
            value: {
              cards: [{ card: welcomeCard.content }],
            },
          },
        };

        switch (ctx.activity.name) {
          case "task/fetch":
            responseBody = {
              task: {
                type: "continue",
                value: {
                  height: "medium",
                  width: "medium",
                  title: "task",
                  card: taskModuleCard,
                },
              },
            };
            break;
          case "task/submit":
            responseBody = {
              task: {
                type: "continue",
                value: tabResponse,
              },
            };
            break;
          case "tab/submit":
            responseBody = tabSubmitResponse;
            break;
          case "tab/fetch":
          default:
            responseBody = tabResponse;
            break;
        }
        return { status: 200, body: responseBody };
      },
    };
  }

  /**
   * The Bot Framework `onTurn` handler.
   * The Microsoft Teams middleware for Bot Framework uses a custom activity processor (`TeamsActivityProcessor`)
   * which is configured in the constructor of this sample
   */
  public async onTurn(context: TurnContext): Promise<any> {
    // transfer the activity to the TeamsActivityProcessor
    await this.activityProc.processIncomingActivity(context);
  }
}
