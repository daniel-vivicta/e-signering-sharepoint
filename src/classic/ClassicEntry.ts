import { IHttpClient } from '../http/IHttpClient';
import { UserProfileService, IUserIdentityFallback } from '../services/UserProfileService';
import { ISigner } from '../models/ISigner';
import { ISigningDocument, ISigningInitiator, ISigningRequest } from '../models/ISigningRequest';

export class ClassicEntry {

  private _userProfileService: UserProfileService;
  private _httpClient: IHttpClient;
  private _webUrl: string;

  constructor(
    httpClient: IHttpClient,
    webUrl: string
  ) {
    this._httpClient = httpClient;
    this._webUrl = this._trimTrailingSlash(webUrl);

    this._userProfileService =
      new UserProfileService(
        this._httpClient,
        this._webUrl
      );
  }

  /**
   * Builds the initial document context from the Classic Custom Action.
   * UniqueId is enriched asynchronously by getDocument().
   */
  public open(
    itemId: string,
    itemUrl: string,
    listId: string,
    siteUrl: string
  ): ISigningDocument {

    const document: ISigningDocument = {
      webUrl: this._trimTrailingSlash(siteUrl),
      listId: this._stripBraces(listId),
      itemId: String(itemId || ''),
      uniqueId: '',
      fileName: this._getFileName(itemUrl),
      fileRef: itemUrl || ''
    };

    if (window.console && window.console.log) {
      window.console.log(
        '[ESIGN][CLASSIC ENTRY] Initial document:',
        document
      );
    }

    return document;
  }

  /**
   * Enriches an ISigningDocument from SharePoint using FileRef.
   */
  public getDocument(
    document: ISigningDocument
  ): Promise<ISigningDocument> {

    const fileRef: string = document.fileRef || '';
    const encodedFileRef: string = encodeURIComponent(fileRef);

    const url: string =
      this._trimTrailingSlash(document.webUrl) +
      "/_api/web/GetFileByServerRelativeUrl(@v)" +
      "?@v='" + encodedFileRef + "'" +
      '&$select=UniqueId,Name,ServerRelativeUrl';

    return this._httpClient
      .getJson(url)
      .then((response: any): ISigningDocument => {

        const file: any = response && response.d
          ? response.d
          : response;

        const completeDocument: ISigningDocument = {
          webUrl: this._trimTrailingSlash(document.webUrl),
          listId: this._stripBraces(document.listId),
          itemId: String(document.itemId || ''),
          uniqueId: file.UniqueId || '',
          fileName: file.Name || document.fileName,
          fileRef: file.ServerRelativeUrl || document.fileRef
        };

        if (window.console && window.console.log) {
          window.console.log(
            '[ESIGN][CLASSIC ENTRY] Complete document:',
            completeDocument
          );
        }

        return completeDocument;
      });
  }

  public getSignerFromPickerPerson(
    person: any
  ): Promise<ISigner> {

    const fallback: IUserIdentityFallback = {
      id:
        person && person.EntityData && person.EntityData.SPUserID
          ? parseInt(person.EntityData.SPUserID, 10)
          : undefined,
      displayName:
        person && person.DisplayText
          ? person.DisplayText
          : '',
      email:
        person && person.EntityData && person.EntityData.Email
          ? person.EntityData.Email
          : '',
      accountName:
        person && person.Key
          ? person.Key
          : '',
      userName: ''
    };

    return this._userProfileService.getSigner(fallback);
  }

  public getSignersFromPickerPeople(
    people: any[]
  ): Promise<ISigner[]> {

    const signerPromises: Array<Promise<ISigner>> =
      (people || []).map(
        (person: any): Promise<ISigner> => {
          return this.getSignerFromPickerPerson(person);
        }
      );

    return Promise.all(signerPromises);
  }

  public getInitiator(
    fallback: IUserIdentityFallback
  ): Promise<ISigningInitiator> {

    return this._userProfileService
      .getSigner(fallback)
      .then((signer: ISigner): ISigningInitiator => {
        return {
          id: signer.id,
          displayName: signer.displayName,
          email: signer.email,
          accountName: signer.accountName,
          userName: signer.userName,
          userPrincipalName: signer.userPrincipalName,
          source: signer.source,
          identifier: signer.identifier
        };
      });
  }

  public createSigningRequest(
    correlationId: string,
    document: ISigningDocument,
    signers: ISigner[],
    message: string,
    deadline: string,
    initiator: ISigningInitiator
  ): ISigningRequest {

    return {
      correlationId: correlationId,
      document: document,
      signers: signers,
      message: message || '',
      deadline: deadline || '',
      initiator: initiator
    };
  }

  private _getFileName(fileRef: string): string {
    if (!fileRef) {
      return '';
    }

    const parts: string[] = fileRef.split('/');
    return decodeURIComponent(parts[parts.length - 1]);
  }

  private _stripBraces(value: string): string {
    return String(value || '')
      .replace('{', '')
      .replace('}', '');
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
