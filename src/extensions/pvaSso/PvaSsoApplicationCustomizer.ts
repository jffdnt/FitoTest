import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
  ApplicationCustomizerContext
} from '@microsoft/sp-application-base';
//import { Dialog } from '@microsoft/sp-dialog';
import * as ReactDOM from "react-dom";
import * as React from "react";
import Chatbot from './components/ChatBot';
import * as strings from 'PvaSsoApplicationCustomizerStrings';
import { IChatbotProps } from './components/IChatBotProps';

const LOG_SOURCE: string = 'PvaSsoApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
/**
 * Properties for the PvaSsoApplicationCustomizer.
 */
export interface IPvaSsoApplicationCustomizerProperties {
  /**
   * The URL of the bot.
   */
  botURL: string;
  /**
   * The name of the bot.
   */
  botName?: string;
  /**
   * The label for the button.
   */
  buttonLabel?: string;
  /**
   * The email of the user.
   */
  userEmail: string;
  /**
   * The URL of the bot's avatar image.
   */
  botAvatarImage?: string;
  /**
   * The initials of the bot's avatar.
   */
  botAvatarInitials?: string;
  /**
   * Whether or not to greet the user.
   */
  greet?: boolean;
  /**
   * The custom scope defined in the Azure AD app registration for the bot.
   */
  customScope: string;
  /**
   * The client ID from the Azure AD app registration for the bot.
   */
  clientID: string;
  /**
   * Azure AD tenant login URL
   */
  authority: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PvaSsoApplicationCustomizer
  extends BaseApplicationCustomizer<IPvaSsoApplicationCustomizerProperties> {

  protected readonly context!: ApplicationCustomizerContext;

  private _bottomPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    const context = this.context;
    const properties = this.properties;
    
    if (!context || !properties) {
      return Promise.reject(new Error('Context or properties not initialized'));
    }
    
    Log.info(LOG_SOURCE, `Bot URL ${properties.botURL}`);
    
    if (!properties.buttonLabel || properties.buttonLabel === "") {
      properties.buttonLabel = strings.DefaultButtonLabel;
    }
    
    if (!properties.botName || properties.botName === "") {
      properties.botName = strings.DefaultBotName;
    }
    
    if (properties.greet !== true) {
      properties.greet = false;
    }
    
    context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    
    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    const context = this.context;
    const properties = this.properties;

    if (!context || !properties) {
      return;
    }

    // Handling the bottom placeholder
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }
      const user = context.pageContext.user;
      const elem: React.ReactElement = React.createElement<IChatbotProps>(Chatbot, { ...properties, userEmail: user.email, userFriendlyName: user.displayName });  
      ReactDOM.render(elem, this._bottomPlaceholder.domElement);
    }
  }

  private _onDispose(): void {
  }

}
