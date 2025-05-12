import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner } from '@fluentui/react/lib/Spinner';
import WebChat from 'botframework-webchat';
import MSALWrapper from '../../components/MSALWrapper';
import { IChatBotConfig } from './IChatBotConfig';
import { validateConfig, getDefaultConfig, mergeConfig } from '../utilities/chatBotUtils';
import { IModalStyles } from '@fluentui/react/lib/Modal';

export interface IChatBotProps {
  config: IChatBotConfig;
  onDismiss: () => void;
}

export const ChatBot: React.FunctionComponent<IChatBotProps> = (props) => {
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [token, setToken] = React.useState<string | null>(null);

  const { config, onDismiss } = props;
  const mergedConfig = mergeConfig(config, getDefaultConfig());

  React.useEffect(() => {
    const initializeChat = async () => {
      try {
        if (!validateConfig(mergedConfig)) {
          throw new Error('Invalid configuration');
        }

        const msalWrapper = new MSALWrapper(
          mergedConfig.clientID,
          mergedConfig.authority
        );

        const tokenResponse = await msalWrapper.acquireAccessToken(
          [mergedConfig.customScope],
          mergedConfig.userEmail
        );
        if (!tokenResponse) {
          console.error('Failed to acquire token');
          return;
        }
        const token = tokenResponse.accessToken;
        setToken(token);
        setIsLoading(false);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'An unknown error occurred');
        setIsLoading(false);
      }
    };

    initializeChat();
  }, [mergedConfig]);

  const dialogContentProps = {
    type: DialogType.normal,
    title: mergedConfig.botName,
    closeButtonAriaLabel: 'Close'
  };

  const modalStyles: IModalStyles = {
    main: {
      maxWidth: 450,
      minWidth: 450,
      maxHeight: '80vh',
      position: 'fixed',
      top: '50%',
      right: '0',
      transform: 'translateY(-50%)'
    },
    root: {},
    scrollableContent: {},
    layer: {},
    keyboardMoveIconContainer: {},
    keyboardMoveIcon: {}
  };

  const modalProps = {
    isBlocking: false,
    styles: modalStyles
  };

  const canvasStyleOptions = {
    rootHeight: '100%',
    rootWidth: '100%',
    sendBox: {
      position: 'absolute',
      bottom: '0',
      left: '0',
      right: '0',
      width: '100%'
    },
    transcript: {
      height: 'calc(100% - 100px)',
      padding: '16px'
    }
  } as const;

  if (error) {
    return (
      <Dialog
        hidden={false}
        onDismiss={onDismiss}
        dialogContentProps={dialogContentProps}
        modalProps={modalProps}
      >
        <div style={{ padding: '20px', textAlign: 'center' }}>
          <p>{error}</p>
        </div>
        <DialogFooter>
          <DefaultButton onClick={onDismiss} text="Close" />
        </DialogFooter>
      </Dialog>
    );
  }

  return (
    <Dialog
      hidden={false}
      onDismiss={onDismiss}
      dialogContentProps={dialogContentProps}
      modalProps={modalProps}
    >
      <div style={{ height: '100%', position: 'relative' }}>
        {isLoading ? (
          <div style={{ 
            position: 'absolute', 
            top: '50%', 
            left: '50%', 
            transform: 'translate(-50%, -50%)',
            zIndex: 10,
            backgroundColor: 'rgba(255, 255, 255, 0.8)',
            padding: '20px',
            borderRadius: '4px'
          }}>
            <Spinner label="Loading chat..." />
          </div>
        ) : (
          <WebChat
            directLine={{
              secret: mergedConfig.botURL,
              token: token
            }}
            userID={mergedConfig.userEmail}
            username={mergedConfig.userFriendlyName}
            styleOptions={canvasStyleOptions}
          />
        )}
      </div>
    </Dialog>
  );
};