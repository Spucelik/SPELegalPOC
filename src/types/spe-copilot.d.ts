/**
 * Type declarations for the SharePoint Embedded Copilot React SDK.
 * 
 * These stubs allow TypeScript compilation before the SDK is installed.
 * After installing the SDK via:
 *   npm install "https://download.microsoft.com/download/970802a5-2a7e-44ed-b17d-ad7dc99be312/microsoft-sharepointembedded-copilotchat-react-1.0.9.tgz"
 * 
 * You can delete this file as the SDK provides its own types.
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
