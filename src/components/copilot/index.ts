// Copilot Chat Components
// Use CustomCopilotChat for Graph API fallback (works without SDK)
// Use SDKCopilotChat when official SDK is installed and configured

export { default as CustomCopilotChat } from "../CustomCopilotChat";
export { default as SDKCopilotChat } from "./SDKCopilotChat";
export { default as CopilotDesktopView } from "./CopilotDesktopView";
export { CopilotAuthProvider } from "./CopilotAuthProvider";
export { CopilotErrorBoundary } from "./CopilotErrorBoundary";
export type { IChatEmbeddedApiAuthProvider } from "./CopilotAuthProvider";
