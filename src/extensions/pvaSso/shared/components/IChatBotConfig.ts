export interface IChatBotConfig {
  botURL: string;
  clientID: string;
  authority: string;
  customScope: string;
  userEmail: string;
  userFriendlyName: string;
  botName: string;
  buttonLabel: string;
  greet?: boolean;
  position?: string;
} 