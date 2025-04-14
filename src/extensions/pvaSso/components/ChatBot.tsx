import * as React from "react";
import { useBoolean, useId } from '@uifabric/react-hooks';
import * as ReactWebChat from 'botframework-webchat';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Dispatch } from 'redux'
import { useRef, useState} from "react";

import { IChatbotProps } from "./IChatBotProps";
import MSALWrapper from "./MSALWrapper";
import { fitoStyleOptions } from "./FiToStyles";
import { GraphAvatarService } from "./GraphAvatarService";

export const PVAChatbotDialog: React.FunctionComponent<IChatbotProps> = (props) => {
    
    // Dialog properties and states
    const dialogContentProps = {
        type: DialogType.normal,
        title: props.useFiToTemplate ? '' : props.botName, // Remove title when using FiTo template as it has its own banner
        closeButtonAriaLabel: 'Close'
    };
    
    const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
    const labelId: string = useId('dialogLabel');
    const subTextId: string = useId('subTextLabel');
    
    // State for user avatar
    const [userAvatar, setUserAvatar] = useState<string | null>(null);
    
    const modalProps = React.useMemo(
        () => ({
            isBlocking: false,
        }),
        [labelId, subTextId],
    );

    // Your bot's token endpoint
    const botURL = props.botURL;

    // constructing URL using regional settings
    const environmentEndPoint = botURL.slice(0,botURL.indexOf('/powervirtualagents'));
    const apiVersion = botURL.slice(botURL.indexOf('api-version')).split('=')[1];
    const regionalChannelSettingsURL = `${environmentEndPoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`;

    // Using refs instead of IDs to get the webchat and loading spinner elements
    const webChatRef = useRef<HTMLDivElement>(null);
    const loadingSpinnerRef = useRef<HTMLDivElement>(null);

    // A utility function that extracts the OAuthCard resource URI from the incoming activity or return undefined
    function getOAuthCardResourceUri(activity: any): string | undefined {
        const attachment = activity?.attachments?.[0];
        if (attachment?.contentType === 'application/vnd.microsoft.card.oauth' && attachment.content.tokenExchangeResource) {
            return attachment.content.tokenExchangeResource.uri;
        }
    }

    const handleLayerDidMount = async () => {
        
        const MSALWrapperInstance = new MSALWrapper(props.clientID, props.authority);

        // Trying to get token if user is already signed-in
        let responseToken = await MSALWrapperInstance.handleLoggedInUser([props.customScope], props.userEmail);

        if (!responseToken) {
            // Trying to get token if user is not signed-in
            responseToken = await MSALWrapperInstance.acquireAccessToken([props.customScope], props.userEmail);
        }

        const token = responseToken?.accessToken || null;

        // Fetch user avatar if using FiTo template
        if (props.useFiToTemplate) {
            try {
                const graphAvatarService = new GraphAvatarService(props.clientID, props.authority);
                const avatarUrl = await graphAvatarService.getUserAvatar(props.userEmail);
                if (avatarUrl) {
                    setUserAvatar(avatarUrl);
                }
            } catch (error) {
                console.error('Error fetching user avatar:', error);
                // Continue without avatar - will use default
            }
        }

        // Get the regional channel URL
        let regionalChannelURL;

        const regionalResponse = await fetch(regionalChannelSettingsURL);
        if(regionalResponse.ok){
            const data = await regionalResponse.json();
            regionalChannelURL = data.channelUrlsById.directline;
        }
        else {
            console.error(`HTTP error! Status: ${regionalResponse.status}`);
        }


        // Create DirectLine object
        let directline: any;

        const response = await fetch(botURL);
        
        if (response.ok) {
            const conversationInfo = await response.json();
            directline = ReactWebChat.createDirectLine({
            token: conversationInfo.token,
            domain: regionalChannelURL + 'v3/directline',
        });
        } else {
        console.error(`HTTP error! Status: ${response.status}`);
        }

        const store = ReactWebChat.createStore(
            {},
               ({ dispatch }: { dispatch: Dispatch }) => (next: any) => (action: any) => {
                   
                // Checking whether we should greet the user
                if (props.greet)
                {
                    if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
                        console.log("Action:" + action.type); 
                            dispatch({
                                meta: {
                                    method: "keyboard",
                                  },
                                    payload: {
                                      activity: {
                                              channelData: {
                                                  postBack: true,
                                              },
                                              //Web Chat will show the 'Greeting' System Topic message which has a trigger-phrase 'hello'
                                              name: 'startConversation',
                                              type: "event"
                                          },
                                  },
                                  type: "DIRECT_LINE/POST_ACTIVITY",
                              });
                              return next(action);
                          }
                    }
                    
                    // Checking whether the bot is asking for authentication
                    if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
                        const activity = action.payload.activity;
                        if (activity.from && activity.from.role === 'bot' &&
                        (getOAuthCardResourceUri(activity))){
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
                                      if(id === "retry"){
                                        // bot was not able to handle the invoke, so display the oauthCard (manual authentication)
                                        console.log("bot was not able to handle the invoke, so display the oauthCard")
                                            return next(action);
                                      }
                                    },
                                    (error: any) => {
                                      // an error occurred to display the oauthCard (manual authentication)
                                      console.log("An error occurred so display the oauthCard");
                                          return next(action);
                                    }
                                  )
                                // token exchange was successful, do not show OAuthCard
                                return;
                        }
                      } else {
                        return next(action);
                      }
                    
                    return next(action);
                }
            );

            // Define style options based on whether FiTo template is used
            const styleOptions = props.useFiToTemplate 
                ? {
                    ...fitoStyleOptions,
                    botAvatarImage: props.botAvatarImage || 'images/fito.png',
                    userAvatarImage: userAvatar || 'images/fito_user.png'
                  }
                : {
                    hideUploadButton: true
                  };
        
            // Render webchat
            if (token && directline) {
                if (webChatRef.current && loadingSpinnerRef.current) {
                    // Set appropriate height based on template
                    if (props.useFiToTemplate) {
                        webChatRef.current.style.height = 'calc(100% - 50px)';
                    } else {
                        webChatRef.current.style.minHeight = '50vh';
                    }
                    
                    loadingSpinnerRef.current.style.display = 'none';
                    ReactWebChat.renderWebChat(
                        {
                            directLine: directline,
                            store: store,
                            styleOptions: styleOptions,
                            userID: props.userEmail,
                        },
                    webChatRef.current
                    );
                } else {
                    console.error("Webchat or loading spinner not found");
                }
        }
    };

    return (
        <>
            <DefaultButton secondaryText={props.buttonLabel} text={props.buttonLabel} onClick={toggleHideDialog}/>
            <Dialog 
                styles={{
                    main: { 
                        selectors: { 
                            ['@media (min-width: 480px)']: { 
                                width: props.useFiToTemplate ? 600 : 450, 
                                minWidth: props.useFiToTemplate ? 600 : 450, 
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
            >
                {props.useFiToTemplate ? (
                    <div style={{ height: '60vh', display: 'flex', flexDirection: 'column' }}>
                        <div id="banner" style={{ 
                            alignItems: 'center', 
                            backgroundColor: '#009edb', 
                            display: 'flex', 
                            height: '50px' 
                        }}>
                            <h1 style={{ 
                                color: 'whitesmoke', 
                                fontFamily: 'Segoe UI', 
                                fontSize: '16px', 
                                lineHeight: '20px', 
                                margin: 0, 
                                padding: '0 20px' 
                            }}>
                                {props.botName || 'FiTo AI (Powered by Ask AT&T)'}
                            </h1>
                        </div>
                        <div id="webchat" role="main" ref={webChatRef} style={{ 
                            height: 'calc(100% - 50px)', 
                            overflow: 'hidden', 
                            position: 'relative', 
                            width: '100%' 
                        }}></div>
                        <div ref={loadingSpinnerRef}><Spinner label="Loading..." style={{ paddingTop: "1rem", paddingBottom: "1rem" }} /></div>
                    </div>
                ) : (
                    // Original dialog content
                    <div id="chatContainer" style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
                        <div ref={webChatRef} role="main" style={{ width: "100%", height: "0rem" }}></div>
                        <div ref={loadingSpinnerRef}><Spinner label="Loading..." style={{ paddingTop: "1rem", paddingBottom: "1rem" }} /></div>
                    </div>
                )}
            </Dialog>
        </>
    );
};

export default class Chatbot extends React.Component<IChatbotProps> {
    constructor(props: IChatbotProps) {
        super(props);
    }
    public render(): JSX.Element {
        return (
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", paddingBottom: "1rem" }}>
                <PVAChatbotDialog
                {...this.props}/>
            </div>
        );
    }
}