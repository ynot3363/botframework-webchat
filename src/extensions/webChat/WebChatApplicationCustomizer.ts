import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { override } from "@microsoft/decorators";
import * as strings from "WebChatApplicationCustomizerStrings";
import * as WebChat from "botframework-webchat";

export interface IBotFrameworkApplicationCustomizerProperties {
  directLineToken: string;
}

export default class BotFrameworkApplicationCustomizer extends BaseApplicationCustomizer<IBotFrameworkApplicationCustomizerProperties> {
  private botContainerId: string = "botFrameworkContainer";

  @override
  public onInit(): Promise<void> {
    console.log(strings.Title);
    this.renderWebChat();
    return Promise.resolve();
  }

  private renderWebChat(): void {
    // Create a container for the BotFramework WebChat
    const botContainer = document.createElement("div");
    botContainer.id = this.botContainerId;
    botContainer.style.position = "fixed";
    botContainer.style.bottom = "20px";
    botContainer.style.right = "20px";
    botContainer.style.width = "400px";
    botContainer.style.height = "500px";
    botContainer.style.zIndex = "1000";
    botContainer.style.border = "1px solid #ccc";
    botContainer.style.borderRadius = "8px";
    botContainer.style.overflow = "hidden";
    botContainer.style.backgroundColor = "#ffffff";
    document.body.appendChild(botContainer);

    // Get current user's ID and username from SPFx context
    const userID = this.context.pageContext.legacyPageContext.userId.toString(); // SharePoint User ID
    const username = this.context.pageContext.user.displayName; // Display name

    // Initialize BotFramework WebChat
    const directLine = WebChat.createDirectLine({
      token: this.properties.directLineToken || "<YOUR DIRECT LINK TOKEN>",
    });

    WebChat.renderWebChat(
      {
        directLine: directLine,
        userID: userID, // Set the current user's ID
        username: username, // Set the current user's display name
        locale: "en-US",
      },
      botContainer
    );
  }

  @override
  public onDispose(): void {
    // Clean up the container when the customizer is disposed
    const botContainer = document.getElementById(this.botContainerId);
    if (botContainer) {
      document.body.removeChild(botContainer);
    }
  }
}
