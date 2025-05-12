import { IChatBotConfig } from '../components/IChatBotConfig';

export const validateConfig = (config: IChatBotConfig): boolean => {
  return !!(config.botURL && 
           config.clientID && 
           config.authority && 
           config.customScope && 
           config.userEmail && 
           config.userFriendlyName);
};

export const getDefaultConfig = (): Partial<IChatBotConfig> => {
  return {
    greet: true,
    position: 'bottom-right',
    botName: 'FiTo',
    buttonLabel: 'Chat with FiTo'
  };
};

export const mergeConfig = (baseConfig: IChatBotConfig, overrideConfig: Partial<IChatBotConfig>): IChatBotConfig => {
  return {
    ...baseConfig,
    ...overrideConfig
  };
}; 