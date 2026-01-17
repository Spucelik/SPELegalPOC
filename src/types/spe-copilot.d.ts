/**
 * Type declarations for the SharePoint Embedded Copilot React SDK.
 * 
 * These provide TypeScript support for the SDK.
 * Package: microsoft-sharepointembedded-copilotchat-react
 */
declare module "microsoft-sharepointembedded-copilotchat-react" {
  import { ComponentType } from "react";

  export interface IChatEmbeddedApiAuthProvider {
    hostname: string;
    getToken(): Promise<string>;
  }

  export interface ChatLaunchConfig {
    header?: string;
    zeroQueryPrompts?: {
      headerText: string;
      promptSuggestionList: Array<{
        suggestionText: string;
      }>;
    };
    suggestedPrompts?: string[];
    instruction?: string;
    locale?: string;
  }

  export interface ChatEmbeddedAPI {
    openChat(config?: ChatLaunchConfig): void;
    closeChat(): void;
  }

  export interface ChatEmbeddedProps {
    authProvider: IChatEmbeddedApiAuthProvider;
    containerId: string;
    onApiReady?: (api: ChatEmbeddedAPI) => void;
  }

  export const ChatEmbedded: ComponentType<ChatEmbeddedProps>;
}
