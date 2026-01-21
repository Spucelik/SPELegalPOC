import { SHAREPOINT_CONFIG, SHAREPOINT_CONTAINER_SCOPES } from "@/config/sharepoint";

/**
 * Authentication provider for the SharePoint Embedded Copilot SDK.
 * Implements IChatEmbeddedApiAuthProvider interface from the SDK.
 * 
 * Reference: https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/tutorials/spe-da-vscode
 */
export interface IChatEmbeddedApiAuthProvider {
  hostname: string;
  getToken(): Promise<string>;
}

export class CopilotAuthProvider implements IChatEmbeddedApiAuthProvider {
  public readonly hostname: string;
  private getAccessToken: (scopes: string[]) => Promise<string | null>;
  private initialized: boolean = false;

  constructor(getAccessToken: (scopes: string[]) => Promise<string | null>) {
    this.hostname = SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME;
    this.getAccessToken = getAccessToken;
  }

  /**
   * Initialize the auth provider by testing token acquisition.
   * Call this before using the ChatEmbedded component.
   */
  async initialize(): Promise<void> {
    const token = await this.getToken();
    if (!token) {
      throw new Error("Failed to initialize auth provider - could not acquire token");
    }
    this.initialized = true;
    console.log("CopilotAuthProvider: Initialized successfully");
  }

  /**
   * Get access token with Container.Selected scope.
   * Required by the SDK: ${hostname}/Container.Selected
   */
  async getToken(): Promise<string> {
    const token = await this.getAccessToken(SHAREPOINT_CONTAINER_SCOPES);
    if (!token) {
      throw new Error("Failed to acquire SharePoint Container.Selected token");
    }
    console.log("CopilotAuthProvider: Acquired Container.Selected token");
    return token;
  }

  get isInitialized(): boolean {
    return this.initialized;
  }
}
