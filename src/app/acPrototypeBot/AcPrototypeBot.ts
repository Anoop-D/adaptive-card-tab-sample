import { BotDeclaration, IBot, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import {
  CardFactory,
  TurnContext,
  MemoryStorage,
  ConversationState,
  InvokeResponse,
} from "botbuilder";
import fetch from "node-fetch";
import WelcomeCard from "./dialogs/WelcomeDialog";
import VideoPlayerCard from "./dialogs/VideoPlayerCard";
import { TeamsActivityProcessor } from "botbuilder-teams";
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
  private readonly activityProc = new TeamsActivityProcessor();
  private loggedInMemberOIDs: Map<string, object> = new Map();
  /**
   * The constructor
   * @param conversationState
   */
  public constructor(conversationState: ConversationState) {
    this.conversationState = conversationState;

    // Set up the Activity processing
    this.activityProc.invokeActivityHandler = {
      onInvoke: async (context: TurnContext): Promise<InvokeResponse> => {
        const ctx: any = context;
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
                        "https://acprototype.azurewebsites.net/acPrototypeTab/login.html",
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

        const primaryTabResponse: any = {
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

        const secondaryTabResponse: any = {
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

        const primaryTabSubmitResponse: any = {
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

        const secondaryTabSubmitResponse: any = {
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
                  value: primaryTabSubmitResponse,
                },
              };
            } else {
              responseBody = {
                task: {
                  type: "continue",
                  value: secondaryTabSubmitResponse,
                },
              };
            }
            break;
          case "tab/submit":
            responseBody = primaryTabSubmitResponse;
            break;
          case "tab/fetch":
          default:
            if (ctx.activity.value.tabContext.tabEntityId === "workday") {
              responseBody = primaryTabResponse;
            } else {
              responseBody = secondaryTabResponse;
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
    try {
      const response = await fetch("https://graph.microsoft.com/v1.0/me/", {
        headers: {
          Authorization: "Bearer " + authState.accessToken,
        },
      });

      const profile = await response.json();
      return profile.error == null ? profile : undefined;
    } catch (error) {
      log(
        "***************** Error fetching user profile from graph ***************"
      );
      log(error);
    }
  }
}
