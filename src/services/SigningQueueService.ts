import { IHttpClient } from '../http/IHttpClient';
import { ISigningRequest } from '../models/ISigningRequest';

export interface IQueueCreateResult {
  id: number;
  raw: any;
}

export class SigningQueueService {

  private _httpClient: IHttpClient;
  private _queueWebUrl: string;
  private _queueListTitle: string;

  constructor(
    httpClient: IHttpClient,
    queueWebUrl: string,
    queueListTitle: string
  ) {
    this._httpClient = httpClient;
    this._queueWebUrl = this._trimTrailingSlash(queueWebUrl);
    this._queueListTitle = queueListTitle;
  }

  public create(
    request: ISigningRequest
  ): Promise<IQueueCreateResult> {

    return Promise.all([
      this._getListItemEntityTypeFullName(),
      this._ensureSigners(request)
    ])
    .then((values: any[]): Promise<any> => {

      const entityType: string = values[0];
      const signerIds: number[] = values[1];

      const url: string = this._getItemsUrl();
      const body: any = this._mapRequestToQueueItem(
        request,
        entityType,
        signerIds
      );

      console.log('[ESIGN][QUEUE SERVICE] Entity type:', entityType);
      console.log('[ESIGN][QUEUE SERVICE] Signer IDs:', signerIds);
      console.log('[ESIGN][QUEUE SERVICE] POST URL:', url);
      console.log('[ESIGN][QUEUE SERVICE] POST BODY:', body);

      return this._httpClient.postJson(url, body);
    })
    .then((response: any): IQueueCreateResult => {

      const result: any = response && response.d
        ? response.d
        : response;

      const id: number = result && result.Id
        ? result.Id
        : 0;

      if (!id) {
        throw new Error(
          'Queue item was created but no SharePoint item ID was returned.'
        );
      }

      return {
        id: id,
        raw: result
      };
    });
  }

  /**
   * Ensures every signer in the queue web/site collection and returns the
   * local SharePoint user IDs for the queue site's User Information List.
   *
   * We intentionally use accountName/login claim, not the user ID from the
   * source site collection. SharePoint user IDs are site-collection specific.
   */
  private _ensureSigners(
    request: ISigningRequest
  ): Promise<number[]> {

    const promises: Array<Promise<number>> =
      request.signers.map(
        (signer: any): Promise<number> => {

          if (!signer.accountName) {
            return Promise.reject(
              new Error(
                'Signer has no accountName/claim: ' +
                (signer.displayName || 'unknown signer')
              )
            );
          }

          return this._ensureUser(
            signer.accountName
          );
        }
      );

    return Promise.all(promises);
  }

  /**
   * Resolves/adds a SharePoint user in the queue web and returns the ID that
   * must be used when writing a Person/Group field in the queue list.
   */
  private _ensureUser(
    accountName: string
  ): Promise<number> {

    const url: string =
      this._queueWebUrl +
      '/_api/web/ensureuser';

    const body: any = {
      logonName: accountName
    };

    console.log(
      '[ESIGN][QUEUE SERVICE] Ensure user:',
      accountName
    );

    return this._httpClient
      .postJson(url, body)
      .then((response: any): number => {

        const user: any = response && response.d
          ? response.d
          : response;

        const userId: number = user && user.Id
          ? user.Id
          : 0;

        if (!userId) {
          throw new Error(
            'Could not ensure queue user: ' +
            accountName
          );
        }

        console.log(
          '[ESIGN][QUEUE SERVICE] Ensured user:',
          accountName,
          'Queue user ID:',
          userId
        );

        return userId;
      });
  }

  private _getListItemEntityTypeFullName(): Promise<string> {

    const url: string =
      this._getListUrl() +
      '?$select=ListItemEntityTypeFullName';

    return this._httpClient
      .getJson(url)
      .then((response: any): string => {

        const result: any = response && response.d
          ? response.d
          : response;

        const entityType: string =
          result && result.ListItemEntityTypeFullName
            ? result.ListItemEntityTypeFullName
            : '';

        if (!entityType) {
          throw new Error(
            'Could not get ListItemEntityTypeFullName for queue list.'
          );
        }

        return entityType;
      });
  }

  private _getListUrl(): string {
    return this._queueWebUrl +
      "/_api/web/lists/getbytitle('" +
      this._escapeODataString(this._queueListTitle) +
      "')";
  }

  private _getItemsUrl(): string {
    return this._getListUrl() + '/items';
  }

  private _mapRequestToQueueItem(
    request: ISigningRequest,
    entityType: string,
    signerIds: number[]
  ): any {

    return {
      __metadata: {
        type: entityType
      },

      Title: request.document.fileName,

      ESignCorrelationId: request.correlationId,
      ESignQueueStatus: 'New',

      ESignSiteUrl: request.document.webUrl,
      ESignListId: request.document.listId,
      ESignItemId: this._toNumber(request.document.itemId),
      ESignFileUniqueId: request.document.uniqueId,
      ESignFileName: request.document.fileName,
      ESignFileRef: request.document.fileRef,

      // Multi-value Person/Group field.
      // IDs are always resolved against the queue site's User Information List.
      ESignSignersId: {
        results: signerIds
      },

      ESignMessage: request.message || '',
      ESignDeadline: request.deadline || null,

      // Temporary initiator text fields. These can be replaced by a single
      // Person/Group field when we make the same simplification for initiator.
      ESignInitiatorName: request.initiator.displayName || '',
      ESignInitiatorEmail: request.initiator.email || '',
      ESignInitiatorAccount: request.initiator.accountName || '',
      ESignInitiatorUserName: request.initiator.userName || '',
      ESignInitiatorUPN: request.initiator.userPrincipalName || '',
      ESignInitiatorSource: request.initiator.source || '',
      ESignInitiatorIdentifier: request.initiator.identifier || ''
    };
  }

  private _toNumber(value: string): number {
    const parsed: number = parseInt(value, 10);
    return isNaN(parsed) ? 0 : parsed;
  }

  private _escapeODataString(value: string): string {
    return String(value || '').replace(/'/g, "''");
  }

  private _trimTrailingSlash(value: string): string {
    if (!value) {
      return '';
    }

    return value.charAt(value.length - 1) === '/'
      ? value.substring(0, value.length - 1)
      : value;
  }
}
