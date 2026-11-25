import * as React from 'react';
import * as ReactDom from 'react-dom';
import { ISigningDocument } from '../../models/ISigningRequest';

import { SigningForm } from '../../components/SigningForm/SigningForm';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { Logger, ILoggerContext } from '../../common/Logger';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'StartSigningCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IStartSigningCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'StartSigningCommandSet';

export default class StartSigningCommandSet extends BaseListViewCommandSet<IStartSigningCommandSetProperties> {

  private _logger: Logger;

  @override
  public onInit(): Promise<void> {

    Log.info(LOG_SOURCE, 'Initialized StartSigningCommandSet');

    this._logger = new Logger(
      this.context.spHttpClient,
      'http://sharepointdev.com/enh/site01',
      'E-signering - Fellogg',
      'StartSigningCommandSet'
    );

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        const selectedRow = event.selectedRows[0];

        const itemId: string =
          selectedRow.getValueByName('ID');
        
        const fileName: string =
          selectedRow.getValueByName('FileLeafRef');
        
        const fileUrl: string =
          selectedRow.getValueByName('FileRef');
        
        const uniqueId: string =
          selectedRow.getValueByName('UniqueId');
        
        const webUrl: string =
          this.context.pageContext.web.absoluteUrl;
        
        const listId: string =
          this.context.pageContext.list.id.toString();
        
        const correlationId: string =
          this._generateCorrelationId();

        const logContext: ILoggerContext = {
          correlationId: correlationId,
          fileName: fileName,
          siteUrl: webUrl,
          libraryName: this.context.pageContext.list.title,
          listId: listId,
          itemId: itemId,
          fileUniqueId: uniqueId,
          fileRef: fileUrl
        };

        const signingDocument: ISigningDocument = {
          webUrl: webUrl,
          listId: listId,
          itemId: itemId,
          uniqueId: uniqueId,
          fileName: fileName,
          fileRef: fileUrl
        };

        this._logger.info(
          'onExecute',
          'User selected Start signing',
          logContext
        );

        this._showSigningForm(
          fileName,
          signingDocument,
          correlationId
        );


        break;

      case 'COMMAND_2':
        window.alert('Command Two');
        break;
  
      default:
        throw new Error('Unknown command');
    }
  }

  private _generateCorrelationId(): string {

    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'
      .replace(/[xy]/g, function(c: string): string {
  
        const r: number =
          Math.random() * 16 | 0;
  
        const v: number =
          c === 'x'
            ? r
            : (r & 0x3 | 0x8);
  
        return v.toString(16);
      });
  }

  private _closeSigningForm(): void {

    const formContainer: HTMLElement =
      document.getElementById('esign-signing-form-container');
  
    const overlay: HTMLElement =
      document.getElementById('esign-signing-form-host');
  
    if (formContainer) {
      ReactDom.unmountComponentAtNode(formContainer);
    }
  
    if (overlay && overlay.parentElement) {
      overlay.parentElement.removeChild(overlay);
    }
  }

  private _showSigningForm(
    fileName: string,
    signingDocument: ISigningDocument,
    correlationId: string
  ): void {

    // Ta bort eventuell tidigare instans
    const existingHost: HTMLElement =
      document.getElementById('esign-signing-form-host');
  
    if (existingHost) {
      ReactDom.unmountComponentAtNode(existingHost);
  
      if (existingHost.parentElement) {
        existingHost.parentElement.removeChild(existingHost);
      }
    }
  
    // Bakgrund
    const overlay: HTMLDivElement = document.createElement('div');
  
    overlay.id = 'esign-signing-form-host';
  
    overlay.style.position = 'fixed';
    overlay.style.left = '0';
    overlay.style.top = '0';
    overlay.style.right = '0';
    overlay.style.bottom = '0';
    overlay.style.backgroundColor = 'rgba(0, 0, 0, 0.4)';
    overlay.style.zIndex = '100000';
  
    // Formuläryta
    const formContainer: HTMLDivElement = document.createElement('div');
  
    formContainer.id = 'esign-signing-form-container';

    formContainer.style.position = 'absolute';
    formContainer.style.top = '0';
    formContainer.style.right = '0';
    formContainer.style.bottom = '0';
    formContainer.style.width = '440px';
    formContainer.style.maxWidth = '90vw';
    formContainer.style.backgroundColor = '#ffffff';
    formContainer.style.padding = '24px';
    formContainer.style.boxSizing = 'border-box';
    formContainer.style.boxShadow =
      '-4px 0 16px rgba(0, 0, 0, 0.25)';
    formContainer.style.overflowY = 'auto';
  
    overlay.appendChild(formContainer);
    document.body.appendChild(overlay);
  
    ReactDom.render(
      React.createElement(
        SigningForm,
        {
          fileName: fileName,
          context: this.context,
          correlationId: correlationId,
          document: signingDocument,
          onCancel: (): void => {
            this._closeSigningForm();
          }
        }
      ),
      formContainer
    );
  }
}
