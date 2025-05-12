import * as React from "react";
import { useBoolean, useId } from '@uifabric/react-hooks';
import * as ReactWebChat from 'botframework-webchat';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dispatch } from 'redux';
import { useRef, useEffect } from "react";
import { Providers, Person, PersonViewType } from '@microsoft/mgt-react';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';

import { IChatbotProps } from "./IChatBotProps";
import MSALWrapper from "./MSALWrapper";

export const ChatBot: React.FC<IChatbotProps> = (props): JSX.Element => {

    // Dialog properties and states
    const dialogContentProps = {
        type: DialogType.normal,
        title: props.botName,
        closeButtonAriaLabel: 'Close',
        styles: {
            header: {
                backgroundColor: '#009FDB',
                color: 'white',
            },
            title: {
                color: 'white'
            }
        }
    };

    const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
    const labelId: string = useId('dialogLabel');
    const subTextId: string = useId('subTextLabel');
    
    const modalProps = {
        isBlocking: false,
        containerClassName: 'chat-dialog-container'
    };

    // Initialize MGT Provider
    useEffect(() => {
        const provider = new Msal2Provider({
            clientId: props.clientID,
            authority: props.authority,
            scopes: [props.customScope]
        });
        Providers.globalProvider = provider;
    }, [props.clientID, props.authority, props.customScope]);

    // Your bot's token endpoint
    const botURL = props.botURL;
    const environmentEndPoint = botURL.slice(0,botURL.indexOf('/powervirtualagents'));
    const apiVersion = botURL.slice(botURL.indexOf('api-version')).split('=')[1];
    const regionalChannelSettingsURL = `${environmentEndPoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`;

    // Refs for WebChat area and spinner
    const webChatRef = useRef<HTMLDivElement>(null);
    const loadingSpinnerRef = useRef<HTMLDivElement>(null);

    // Utility to get OAuthCard resource URI from activity
    function getOAuthCardResourceUri(activity: any): string | undefined {
        const attachment = activity?.attachments?.[0];
        if (attachment?.contentType === 'application/vnd.microsoft.card.oauth' && attachment.content.tokenExchangeResource) {
            return attachment.content.tokenExchangeResource.uri;
        }
    }

    const handleLayerDidMount = async () => {
        console.log('Starting chat initialization...');
        const MSALWrapperInstance = new MSALWrapper(props.clientID, props.authority);
        console.log('MSALWrapper initialized with:', { clientID: props.clientID, authority: props.authority });

        let responseToken = await MSALWrapperInstance.handleLoggedInUser([props.customScope], props.userEmail);
        console.log('Logged in user token response:', responseToken ? 'Success' : 'Failed');

        if (!responseToken) {
            responseToken = await MSALWrapperInstance.acquireAccessToken([props.customScope], props.userEmail);
            console.log('Acquire token response:', responseToken ? 'Success' : 'Failed');
        }
        const token = responseToken?.accessToken || null;
        console.log('Final token status:', token ? 'Token obtained' : 'No token available');

        console.log('Fetching regional channel settings from:', regionalChannelSettingsURL);
        let regionalChannelURL;
        const regionalResponse = await fetch(regionalChannelSettingsURL);
        if(regionalResponse.ok){
            const data = await regionalResponse.json();
            regionalChannelURL = data.channelUrlsById.directline;
            console.log('Regional channel URL obtained:', regionalChannelURL);
        }
        else {
            console.error(`Failed to get regional channel settings. Status: ${regionalResponse.status}`);
        }

        console.log('Fetching bot token from:', botURL);
        let directline: any;
        const response = await fetch(botURL);
        if (response.ok) {
            const conversationInfo = await response.json();
            console.log('Bot token obtained successfully');
            directline = ReactWebChat.createDirectLine({
                token: conversationInfo.token,
                domain: regionalChannelURL + 'v3/directline',
            });
            console.log('DirectLine object created');
        } else {
            console.error(`Failed to get bot token. Status: ${response.status}`);
        }

        const store = ReactWebChat.createStore(
            {},
            ({ dispatch }: { dispatch: Dispatch }) => (next: any) => (action: any) => {
                if (props.greet) {
                    if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
                        console.log("Action:" + action.type); 
                        dispatch({
                            meta: { method: "keyboard" },
                            payload: {
                                activity: {
                                    channelData: { postBack: true },
                                    name: 'startConversation',
                                    type: "event"
                                },
                            },
                            type: "DIRECT_LINE/POST_ACTIVITY",
                        });
                        return next(action);
                    }
                }
                
                if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
                    const activity = action.payload.activity;
                    if (activity.from && activity.from.role === 'bot' && (getOAuthCardResourceUri(activity))) {
                        directline.postActivity({
                            type: 'invoke',
                            name: 'signin/tokenExchange',
                            value: {
                                id: activity.attachments[0].content.tokenExchangeResource.id,
                                connectionName: activity.attachments[0].content.connectionName,
                                token
                            },
                            "from": {
                                id: props.userEmail,
                                name: props.userFriendlyName,
                                role: "user"
                            }
                        }).subscribe(
                            (id: any) => {
                                if(id === "retry") {
                                    console.log("bot was not able to handle the invoke, so display the oauthCard");
                                    return next(action);
                                }
                            },
                            (error: any) => {
                                console.log("An error occurred so display the oauthCard");
                                return next(action);
                            }
                        );
                        return;
                    }
                }
                return next(action);
            }
        );

        // WebChat style options (unchanged for this layout)
        const canvasStyleOptions = {
            hideSendBox: false,
            hideUploadButton: true,
            sendBoxBackground: '#F3F4F6',
            sendBoxButtonColor: '#8B8B8B',
            sendBoxButtonColorOnDisabled: '#CCC',
            sendBoxButtonColorOnFocus: '#333',
            sendBoxButtonColorOnHover: '#333',
            sendBoxDisabledTextColor: 'White',
            sendBoxHeight: 100,
            sendBoxMaxHeight: 200,
            sendBoxTextColor: 'Black',
            sendBoxBorderBottom: 'solid 2px #8B8B8B',
            sendBoxBorderLeft: 'solid 2px #8B8B8B',
            sendBoxBorderRight: 'solid 2px #8B8B8B',
            sendBoxBorderTop: 'solid 2px #8B8B8B',
            sendBoxPlaceholderColor: '#8B8B8B',
            sendBoxTextWrap: false,
            sendBoxPosition: 'absolute',
            sendBoxBottom: '0',
            sendBoxPadding: '8px 8px 8px 8px',
            sendBoxMargin: '0',

            bubbleBackground: '#EBEBED',
            bubbleBorderColor: '#EBEBED',
            bubbleBorderRadius: 12,
            bubbleBorderStyle: 'solid',
            bubbleBorderWidth: 1,
            bubbleFromUserBackground: '#0057B8',
            bubbleFromUserBorderColor: '#0057B8',
            bubbleFromUserBorderRadius: 12,
            bubbleFromUserBorderStyle: 'solid',
            bubbleFromUserBorderWidth: 1,
            bubbleFromUserNubOffset: 'bottom',
            bubbleFromUserNubSize: 0,
            bubbleFromUserTextColor: 'White',
            bubbleImageHeight: 240,
            bubbleMaxWidth: 480,
            bubbleMinHeight: 40,
            bubbleMinWidth: 250,
            bubbleNubOffset: 'bottom',
            bubbleNubSize: 5,
            bubbleTextColor: 'Black',

            avatarSize: 50,
            botAvatarBackgroundColor: '#EBEBED',
            botAvatarImage: 'https://freepngimg.com/thumb/dog/163165-puppy-dog-face-free-transparent-image-hd.png',
            botAvatarInitials: '',
            userAvatarBackgroundColor: 'white',
            userAvatarImage: 'https://zetaphotoservice.azurewebsites.net/microsoft/photo/' + props.userEmail,
            userAvatarInitials: ''
        };

        // (Optional) Retrieve user avatar via MGT Person component ...
        const userAvatarUrl = await new Promise<string>((resolve) => {
            const person = document.createElement('mgt-person') as any;
            person.personQuery = 'me';
            person.view = PersonViewType.avatar;
            person.avatarSize = 'large';
            person.fetchImage = true;
            person.addEventListener('imageRendered', (e: any) => {
                resolve(e.detail);
            });
            document.body.appendChild(person);
            setTimeout(() => {
                document.body.removeChild(person);
                resolve('');
            }, 1000);
        });
        console.log("User Avatar URL resolved:", 'https://zetaphotoservice.azurewebsites.net/microsoft/photo/' + props.userEmail);
        canvasStyleOptions.userAvatarImage = 'https://zetaphotoservice.azurewebsites.net/microsoft/photo/' + props.userEmail;

        if (token && directline) {
            console.log('Attempting to render webchat...');
            if (webChatRef.current && loadingSpinnerRef.current) {
                webChatRef.current.style.minHeight = '50vh';
                loadingSpinnerRef.current.style.display = 'none';
                try {
                    ReactWebChat.renderWebChat(
                        {
                            directLine: directline,
                            store: store,
                            styleOptions: canvasStyleOptions,
                            userID: props.userEmail,
                            locale: 'en-US'
                        },
                        webChatRef.current
                    );
                    console.log('Webchat render completed');
                } catch (error) {
                    console.error('Error rendering webchat:', error);
                }
            } else {
                console.error("Webchat or loading spinner not found", {
                    hasWebChatRef: !!webChatRef.current,
                    hasLoadingSpinnerRef: !!loadingSpinnerRef.current
                });
            }
        } else {
            console.error('Cannot render webchat: missing required components', {
                hasToken: !!token,
                hasDirectline: !!directline
            });
        }
    };

    return (
        <div style={{ 
            display: "flex", 
            flexDirection: "column", 
            alignItems: "center", 
            position: "fixed",
            top: "50%",
            right: "0",
            transform: "translateY(-50%)",
            zIndex: 1000,
            marginRight: "15px"
        }}>
            <DefaultButton 
                secondaryText={props.buttonLabel} 
                text=""
                onClick={toggleHideDialog}
                iconProps={{ iconName: 'Message' }}
                styles={{
                    root: {
                        backgroundColor: '#0057B8',
                        color: 'white',
                        borderRadius: '10px',
                        border: 'none',
                        padding: '16px 16px',
                        display: 'flex',
                        alignItems: 'center',
                        minHeight: '80px'
                    },
                    rootHovered: {
                        backgroundColor: '#009FDB',
                        color: 'white',
                        borderRadius: '10px',
                        border: 'none',
                        padding: '16px 16px'
                    },
                    rootPressed: {
                        backgroundColor: '#009FDB',
                        color: 'white',
                        borderRadius: '10px',
                        border: 'none',
                        padding: '16px 16px'
                    },
                    label: { display: 'none' },
                    icon: {
                        marginRight: '8px',
                        color: 'white',
                        fontSize: '28px'
                    }
                }}
            >
                <Stack tokens={{ childrenGap: 1 }} verticalAlign="center" styles={{ root: { padding: '3px 3px' } }}>
                    <Text variant="mediumPlus" styles={{ root: { fontWeight: 'bold', color: 'white', lineHeight: '1.2' } }}>CHAT</Text>
                    <Text variant="mediumPlus" styles={{ root: { color: 'white', lineHeight: '1.2' } }}>with</Text>
                    <Text variant="mediumPlus" styles={{ root: { fontWeight: 'bold', color: 'white', lineHeight: '1.2' } }}>FiTo</Text>
                </Stack>
            </DefaultButton>
            <Dialog 
                styles={{
                    main: { 
                        selectors: { 
                            ['@media (min-width: 480px)']: { 
                                width: 450, 
                                minWidth: 450, 
                                maxWidth: '1000px' 
                            } 
                        } 
                    }
                }} 
                hidden={hideDialog} 
                onDismiss={toggleHideDialog} 
                onLayerDidMount={handleLayerDidMount} 
                dialogContentProps={dialogContentProps} 
                modalProps={modalProps}
                aria-labelledby={labelId}
                aria-describedby={subTextId}
            >
                <div id="chatContainer" style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
                    <div ref={webChatRef} role="main" style={{ width: "100%", height: "0rem" }}></div>
                    <div ref={loadingSpinnerRef}><Spinner label="Loading..." style={{ paddingTop: "1rem", paddingBottom: "1rem" }} /></div>
                </div>
            </Dialog>

        </div>
    );
};

export default ChatBot;
