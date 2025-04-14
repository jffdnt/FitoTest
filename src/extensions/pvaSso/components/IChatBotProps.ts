export interface IChatbotProps {
     botURL: string;
     buttonLabel?: string;
     botName?: string;
     userEmail: string;
     userFriendlyName: string;
     botAvatarImage?: string;
     botAvatarInitials?: string;
     userAvatarImage?: string;
     greet?: boolean;
     customScope: string;
     clientID: string;
     authority: string;
     useFiToTemplate?: boolean;
   }