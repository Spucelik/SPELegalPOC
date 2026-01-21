// Copilot Chat Components
// Use CustomCopilotChat for Graph API fallback (works without SDK)
// Use SDKCopilotChat when official SDK is installed and configured

export { default as CustomCopilotChat } from "../CustomCopilotChat";
export { default as SDKCopilotChat } from "./SDKCopilotChat";
export { CopilotAuthProvider } from "./CopilotAuthProvider";
export type { IChatEmbeddedApiAuthProvider } from "./CopilotAuthProvider";
