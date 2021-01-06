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
import VideoPlayerCard from "./dialogs/VideoPlayerCard";
import { TeamsContext, TeamsActivityProcessor } from "botbuilder-teams";
import AdminCard from "./dialogs/AdminCard";
import QuickActionCard from "./dialogs/QuickActionsCard";
import ManagerDashboardCard from "./dialogs/ManagerDashboard";
import InterviewCandidatesCard from "./dialogs/interviewCandidates";
import SuccessCard from "./dialogs/SuccessCard";

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
  private loggedInMemberOIDs: Map<string, object> = new Map();
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
        const ctx: any = context;
        // console.log(ctx.activity);
        // console.log(ctx.activity.value);

        // If not logged in
        if (ctx.activity.value.state != null) {
          this.loggedInMemberOIDs.set(
            ctx.activity.from.aadObjectId,
            ctx.activity.value.state
          );
        }
        const profile = await this.getUserProfile(
          ctx.activity.from.aadObjectId
        );

        if (
          ctx.activity.value.tabContext.tabEntityId === "workday" &&
          !profile
        ) {
          return {
            status: 200,
            body: {
              tab: {
                type: "auth",
                suggestedActions: {
                  actions: [
                    {
                      type: "openUrl",
                      value:
                        "https://andhillo-relay.servicebus.windows.net/MININT-S5EDEDH/acPrototypeTab/login.html",
                      title: "Sign in to this app!",
                    },
                  ],
                },
              },
            },
          };
        }

        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
        const adminCard = CardFactory.adaptiveCard(AdminCard);
        const quickActionsCard = CardFactory.adaptiveCard(QuickActionCard);
        const managerCard = CardFactory.adaptiveCard(
          ManagerDashboardCard(profile)
        );
        const videoPlayerCard = CardFactory.adaptiveCard(VideoPlayerCard);
        const interviewCard = CardFactory.adaptiveCard(InterviewCandidatesCard);
        const successCard = CardFactory.adaptiveCard(SuccessCard);
        let responseBody: any;

        const workdayTabResponse: any = {
          tab: {
            type: "continue",
            value: {
              cards: [
                { card: quickActionsCard.content },
                { card: managerCard.content },
                { card: adminCard.content },
              ],
            },
          },
        };

        const sampleTabResponse: any = {
          tab: {
            type: "continue",
            value: {
              cards: [
                { card: welcomeCard.content },
                { card: interviewCard.content },
                { card: videoPlayerCard.content },
              ],
            },
          },
        };

        const tabSubmitResponse: any = {
          tab: {
            type: "continue",
            value: {
              cards: [
                { card: successCard.content },
                { card: quickActionsCard.content },
                { card: managerCard.content },
                { card: adminCard.content },
              ],
            },
          },
        };

        const sampleSubmitTabResponse: any = {
          tab: {
            type: "continue",
            value: {
              cards: [
                { card: successCard.content },
                { card: welcomeCard.content },
                { card: interviewCard.content },
                { card: videoPlayerCard.content },
              ],
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
                  card: videoPlayerCard,
                },
              },
            };
            break;
          case "task/submit":
            if (ctx.activity.value.tabContext.tabEntityId === "workday") {
              responseBody = {
                task: {
                  type: "continue",
                  value: tabSubmitResponse,
                },
              };
            } else {
              responseBody = {
                task: {
                  type: "continue",
                  value: sampleSubmitTabResponse,
                },
              };
            }
            break;
          case "tab/submit":
            responseBody = tabSubmitResponse;
            break;
          case "tab/fetch":
          default:
            if (ctx.activity.value.tabContext.tabEntityId === "workday") {
              responseBody = workdayTabResponse;
            } else {
              responseBody = sampleTabResponse;
            }
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

  private async getUserProfile(aadObjectId: string): Promise<any> {
    const authState: any = this.loggedInMemberOIDs.get(aadObjectId);
    if (!authState) {
      return false;
    }

    const response = await fetch("https://graph.microsoft.com/v1.0/me/", {
      headers: {
        Authorization: "Bearer " + authState.accessToken,
      },
    });
    const profile = await response.json();
    return profile.error == null ? profile : undefined;
  }
}
