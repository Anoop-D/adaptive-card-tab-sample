import { PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import {
  CardFactory,
  TurnContext,
  MemoryStorage,
  ConversationState,
  InvokeResponse,
  ActivityHandler,
} from "botbuilder";
import fetch from "node-fetch";
import WelcomeCard from "./dialogs/WelcomeDialog";
import VideoPlayerCard from "./dialogs/VideoPlayerCard";
import AdminCard from "./dialogs/AdminCard";
import QuickActionCard from "./dialogs/QuickActionsCard";
import ManagerDashboardCard from "./dialogs/ManagerDashboard";
import InterviewCandidatesCard from "./dialogs/interviewCandidates";
import SuccessCard from "./dialogs/SuccessCard";
import RecommendationCard from "./dialogs/errors/RecommendationsCard";
import ErrorAdminCard from "./dialogs/errors/ErrorAdminCard";
import ErrorInterviewCandidatesCard from "./dialogs/errors/ErrorInterviewCandidates";
import ErrorManagerDashboardCard from "./dialogs/errors/ErrorManagerDashboard";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/acPrototypeBot/acProtoBotTab.html")
export class AcPrototypeBot extends ActivityHandler {
  private readonly conversationState: ConversationState;
  private loggedInMemberOIDs: Map<string, object> = new Map();
  /**
   * The constructor
   * @param conversationState
   */
  public constructor(
    memoryStorage: MemoryStorage,
    conversationState: ConversationState
  ) {
    super();
    this.conversationState = conversationState;

    // Set up the Activity processing
    this.onInvokeActivity = async (
      context: TurnContext
    ): Promise<InvokeResponse> => {
      const ctx: any = context;
      // Verify state and retrieve stored accessToken.
      if (ctx.activity.value.state != null) {
        const authCode = await memoryStorage.read([ctx.activity.value.state]);
        this.loggedInMemberOIDs.set(
          ctx.activity.from.aadObjectId,
          authCode[ctx.activity.value.state]
        );
      }
      const profile = await this.getUserProfile(ctx.activity.from.aadObjectId);

      // Card for tab containing error tab scenarios
      const recommendationCard = CardFactory.adaptiveCard(RecommendationCard);
      const recommendationCardResponse:any = {
        tab: {
          type: "continue",
          value: {
            cards: [
              { card: recommendationCard.content }
            ],
          },
        }
      };

      if (ctx.activity.value.tabContext.tabEntityId === "workday" && !profile) {
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
      } else if (ctx.activity.value.tabContext.tabEntityId === "error_tab" && ctx.activity.name === "tab/fetch") {
        return { status: 200, body: recommendationCardResponse };
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
          if (ctx.activity.value.tabContext.tabEntityId === "error_tab") {
            return await this.handleErrorTabScenarios(ctx.activity.value.data, recommendationCardResponse);
          }

          if (ctx.activity.value.data.shouldLogout === true) {
            this.loggedInMemberOIDs.delete(ctx.activity.from.aadObjectId);
          }
          if (ctx.activity.value.tabContext.tabEntityId === "workday") {
            responseBody = primaryTabSubmitResponse;
          } else {
            responseBody = secondaryTabSubmitResponse;
          }
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
    };
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
      log("Error fetching user profile from graph: ", error);
    }
  }

  private async handleErrorTabScenarios(data: any, defaultResponse: any): Promise<InvokeResponse> {
    if (data == null) {
      return { status: 200, body: defaultResponse };
    }

    if (data.shouldSetHTTPError === true) {
      return { status: 430, body: null  };
    } 

    let responseBody: any;

    const errorAdminCard:any = CardFactory.adaptiveCard(ErrorAdminCard);
    const errorInterviewCandidatesCard:any = CardFactory.adaptiveCard(ErrorInterviewCandidatesCard);
    const errorManagerDashboardCard:any = CardFactory.adaptiveCard(ErrorManagerDashboardCard);

    const malformedTabResponse: any = {
      tab: {
        type: "continue",
        value: {
          cards: "",
        },
      },
    };

    const malformedCardResponse: any = {
      tab: {
        type: "continue",
        value: {
          cards: [
            { card: errorAdminCard.content },
            { card: errorInterviewCandidatesCard.content },
            { card: errorManagerDashboardCard.content }
          ],
        },
      },
    }

    if (data.shouldSetEmptyResponse === true) {
      responseBody = null;
    } else if (data.shouldSetMalformedTab === true) {
      responseBody = malformedTabResponse;
    } else if (data.shouldSetMalformedCard === true) {
      responseBody = malformedCardResponse;
    } else {
      responseBody = defaultResponse;
    }

    return { status: 200, body: responseBody };
  }

}
