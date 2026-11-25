import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ILoggerContext {
  correlationId: string;

  fileName?: string;
  siteUrl?: string;
  libraryName?: string;
  listId?: string;
  itemId?: string;
  fileUniqueId?: string;
  fileRef?: string;
}

export class Logger {

  private _spHttpClient: SPHttpClient;
  private _logWebUrl: string;
  private _logListName: string;
  private _component: string;

  constructor(
    spHttpClient: SPHttpClient,
    logWebUrl: string,
    logListName: string,
    component: string
  ) {
    this._spHttpClient = spHttpClient;
    this._logWebUrl = logWebUrl;
    this._logListName = logListName;
    this._component = component;
  }

  public info(
    functionName: string,
    message: string,
    context?: ILoggerContext
  ): void {

    console.log(
      '[ESIGN][INFO]',
      'Function:', functionName,
      'Message:', message,
      'Context:', context
    );
  }

  public warn(
    functionName: string,
    message: string,
    context?: ILoggerContext,
    error?: any
  ): Promise<void> {

    console.warn(
      '[ESIGN][WARN]',
      'Function:', functionName,
      'Message:', message,
      'Context:', context,
      'Error:', error
    );

    return this._writeToSharePoint(
      'WARN',
      functionName,
      message,
      context,
      error
    );
  }

  public error(
    functionName: string,
    message: string,
    context?: ILoggerContext,
    error?: any
  ): Promise<void> {

    console.error(
      '[ESIGN][ERROR]',
      'Function:', functionName,
      'Message:', message,
      'Context:', context,
      'Error:', error
    );

    return this._writeToSharePoint(
      'ERROR',
      functionName,
      message,
      context,
      error
    );
  }

  private _writeToSharePoint(
    level: string,
    functionName: string,
    message: string,
    context?: ILoggerContext,
    error?: any
  ): Promise<void> {

    try {

      const url: string =
        this._logWebUrl +
        "/_api/web/lists/getbytitle('" +
        encodeURIComponent(this._logListName) +
        "')/items";

      const body: any = {
        Title:
            level +
            ': ' +
            (context && context.fileName
                ? context.fileName + ' - '
                : '') +
            functionName,

        ESignLevel: level,
        ESignComponent: this._component,
        ESignFunction: functionName,
        ESignMessage: message,

        ESignCorrelationId:
          context ? context.correlationId : '',

        ESignFileName:
          context ? context.fileName || '' : '',

        ESignSiteUrl:
          context ? context.siteUrl || '' : '',

        ESignLibraryName:
          context ? context.libraryName || '' : '',

        ESignListId:
          context ? context.listId || '' : '',

        ESignItemId:
          context ? context.itemId || '' : '',

        ESignFileUniqueId:
          context ? context.fileUniqueId || '' : '',

        ESignFileRef:
          context ? context.fileRef || '' : '',

        ESignErrorDetails:
          this._errorToString(error)
      };

      return this._spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata'
          },
          body: JSON.stringify(body)
        }
      )
      .then((response: SPHttpClientResponse): void => {

        if (!response.ok) {

          console.error(
            '[ESIGN][LOGGER ERROR]',
            'Could not write to SharePoint log list.',
            'HTTP:', response.status,
            response.statusText
          );
        }
      })
      .catch((loggerError: any): void => {

        console.error(
          '[ESIGN][LOGGER ERROR]',
          'Logger failed while writing to SharePoint.',
          loggerError
        );
      });

    }
    catch (loggerError) {

      /*
       * Logging must never break the signing process.
       */
      console.error(
        '[ESIGN][LOGGER ERROR]',
        'Unexpected error in Logger._writeToSharePoint.',
        loggerError
      );

      return Promise.resolve();
    }
  }

  private _errorToString(error: any): string {

    if (!error) {
      return '';
    }
  
    try {
  
      if (error instanceof Error) {
  
        return (
          'Name = ' + error.name + '\n' +
          'Message = ' + error.message
        );
      }
  
      if (typeof error === 'string') {
        return error;
      }
  
      return JSON.stringify(error, null, 2);
    }
    catch (conversionError) {
  
      return 'Could not serialize error object.';
    }
  }
}
