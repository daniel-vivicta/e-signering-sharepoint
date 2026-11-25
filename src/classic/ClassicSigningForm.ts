import { ClassicEntry } from './ClassicEntry';
import { ISigner } from '../models/ISigner';
import {
  ISigningDocument,
  ISigningInitiator,
  ISigningRequest
} from '../models/ISigningRequest';

import {
    SigningQueueService,
    IQueueCreateResult
  } from '../services/SigningQueueService';
  import { ClassicHttpClientAdapter } from '../http/ClassicHttpClientAdapter';
  

export interface IClassicDocumentContext {
  itemId: string;
  itemUrl: string;
  listId: string;
  siteUrl: string;
}

export class ClassicSigningForm {

  private _entry: ClassicEntry;
  private _documentContext: IClassicDocumentContext;
  private _document: ISigningDocument;
  private _queueService: SigningQueueService;

  private _overlayId: string = 'esign-classic-panel-overlay';
  private _pickerId: string = 'esignClassicPeoplePicker';

  constructor(
    entry: ClassicEntry,
    documentContext: IClassicDocumentContext,
    document: ISigningDocument
  ) {
    this._entry = entry;
    this._documentContext = documentContext;
    this._document = document;
    this._queueService =
        new SigningQueueService(
            new ClassicHttpClientAdapter(),
            'http://sharepointdev.com/enh/site01',
            'E-signering - Kö'
        );
  }

  public open(): void {
    this.close();
    this._createPanel();
    this._loadPeoplePickerScripts(
      (): void => this._initializePeoplePicker()
    );
  }

  public close(): void {
    const existing: HTMLElement =
      document.getElementById(this._overlayId);

    if (existing && existing.parentElement) {
      existing.parentElement.removeChild(existing);
    }
  }

  private _createPanel(): void {
    const overlay: HTMLDivElement = document.createElement('div');
    overlay.id = this._overlayId;
    overlay.style.position = 'fixed';
    overlay.style.top = '0';
    overlay.style.left = '0';
    overlay.style.right = '0';
    overlay.style.bottom = '0';
    overlay.style.backgroundColor = 'rgba(0,0,0,0.35)';
    overlay.style.zIndex = '100000';

    const panel: HTMLDivElement = document.createElement('div');
    panel.style.position = 'absolute';
    panel.style.top = '0';
    panel.style.right = '0';
    panel.style.bottom = '0';
    panel.style.width = '440px';
    panel.style.maxWidth = '90vw';
    panel.style.backgroundColor = '#ffffff';
    panel.style.padding = '24px 32px';
    panel.style.boxSizing = 'border-box';
    panel.style.boxShadow = '-4px 0 16px rgba(0,0,0,0.25)';
    panel.style.overflowY = 'auto';
    panel.style.fontFamily = 'Segoe UI, Arial, sans-serif';
    panel.style.color = '#323130';

    const iconUrl: string =
      this._trimTrailingSlash(this._document.webUrl) +
      '/SiteAssets/signing.png';

    panel.innerHTML =
      '<div style="display:flex;align-items:center;margin-bottom:24px;">' +
        '<img src="' + this._escapeHtml(iconUrl) + '" alt="" ' +
          'style="width:40px;height:40px;margin-right:12px;object-fit:contain;" />' +
        '<div>' +
          '<div style="font-size:21px;font-weight:600;color:#323130;">' +
            'Starta signering' +
          '</div>' +
          '<div style="font-size:13px;color:#605e5c;margin-top:2px;overflow-wrap:anywhere;">' +
            this._escapeHtml(this._document.fileName) +
          '</div>' +
        '</div>' +
      '</div>' +

      '<div style="font-size:14px;font-weight:600;margin-bottom:6px;">' +
        'Undertecknare' +
      '</div>' +
      '<div id="' + this._pickerId + '"></div>' +
      '<div id="esign-classic-validation" ' +
        'style="display:none;color:#a4262c;font-size:13px;margin-top:6px;">' +
        'Välj minst en undertecknare.' +
      '</div>' +
      '<div id="esign-classic-error" ' +
        'style="display:none;color:#a4262c;font-size:13px;margin-top:6px;">' +
      '</div>' +

      '<div style="margin-top:24px;">' +
        '<label for="esign-classic-message" ' +
          'style="display:block;font-size:14px;font-weight:600;color:#323130;margin-bottom:6px;">' +
          'Meddelande' +
        '</label>' +
        '<textarea id="esign-classic-message" rows="5" ' +
          'style="width:100%;box-sizing:border-box;border:1px solid #8a8886;' +
          'padding:8px;font-family:Segoe UI;font-size:14px;resize:vertical;">' +
        '</textarea>' +
      '</div>' +

      '<div style="margin-top:20px;">' +
        '<label for="esign-classic-deadline" ' +
          'style="display:block;font-size:14px;font-weight:600;color:#323130;margin-bottom:6px;">' +
          'Deadline' +
        '</label>' +
        '<input id="esign-classic-deadline" type="date" ' +
          'style="height:32px;border:1px solid #8a8886;padding:0 8px;' +
          'font-family:Segoe UI;font-size:14px;box-sizing:border-box;" />' +
      '</div>' +

      '<div style="display:flex;justify-content:flex-end;margin-top:32px;gap:8px;">' +
        '<button id="esign-classic-cancel" type="button" ' +
          'style="min-width:80px;height:32px;border:1px solid #8a8886;background:#fff;' +
          'color:#323130;font-family:Segoe UI;font-size:14px;cursor:pointer;">' +
          'Avbryt' +
        '</button>' +
        '<button id="esign-classic-continue" type="button" ' +
          'style="min-width:80px;height:32px;border:1px solid #0078d4;background:#0078d4;' +
          'color:#fff;font-family:Segoe UI;font-size:14px;font-weight:600;cursor:pointer;">' +
          'Fortsätt' +
        '</button>' +
      '</div>';

    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    const cancelButton: HTMLElement =
      document.getElementById('esign-classic-cancel');

    const continueButton: HTMLElement =
      document.getElementById('esign-classic-continue');

    if (cancelButton) {
      cancelButton.onclick = (): void => this.close();
    }

    if (continueButton) {
      continueButton.onclick = (): void => {
        this._submit();
      };
    }
  }

  private _submit(): void {
    const people: any[] = this._getSelectedPeople();
  
    const validation: HTMLElement =
      document.getElementById('esign-classic-validation');
  
    if (!people || people.length === 0) {
      if (validation) {
        validation.style.display = 'block';
      }
      return;
    }
  
    if (validation) {
      validation.style.display = 'none';
    }
  
    const messageElement: HTMLTextAreaElement =
      document.getElementById('esign-classic-message') as HTMLTextAreaElement;
  
    const deadlineElement: HTMLInputElement =
      document.getElementById('esign-classic-deadline') as HTMLInputElement;
  
    const message: string = messageElement ? messageElement.value : '';
    const deadline: string = deadlineElement ? deadlineElement.value : '';
  
    this._setBusy(true);
  
    Promise.all([
      this._entry.getSignersFromPickerPeople(people),
      this._getInitiator()
    ])
    .then((values: any[]): Promise<IQueueCreateResult> => {
      const signers: ISigner[] = values[0];
      const initiator: ISigningInitiator = values[1];
  
      const request: ISigningRequest =
        this._entry.createSigningRequest(
          this._generateCorrelationId(),
          this._document,
          signers,
          message,
          deadline,
          initiator
        );
  
      if (window.console && window.console.log) {
        window.console.log(
          '[ESIGN][CLASSIC SIGNING REQUEST]',
          request
        );
        window.console.log(
          '[ESIGN][CLASSIC QUEUE] Calling SigningQueueService.create()'
        );
      }
  
      return this._queueService.create(request);
    })
    .then((result: IQueueCreateResult): void => {

        this._setBusy(false);
      
        if (window.console && window.console.log) {
          window.console.log(
            '[ESIGN][CLASSIC QUEUE ITEM CREATED]',
            result
          );
        }
      
        this.close();
      })

    .catch((error: any): void => {
      this._setBusy(false);
  
      this._showError(
        'Det gick inte att skapa signeringsbegaran i kon.',
        error
      );
    });
  }
  

  private _getInitiator(): Promise<ISigningInitiator> {
    return this._getJson(
      this._document.webUrl + '/_api/web/currentuser'
    )
    .then((response: any): Promise<ISigningInitiator> => {
      const currentUser: any = response && response.d
        ? response.d
        : response;

      return this._entry.getInitiator({
        id: currentUser.Id,
        displayName: currentUser.Title || '',
        email: currentUser.Email || '',
        accountName: currentUser.LoginName || '',
        userName: ''
      });
    })
    .catch((error: any): ISigningInitiator => {
      if (window.console && window.console.warn) {
        window.console.warn(
          '[ESIGN][CLASSIC INITIATOR] Profile lookup failed.',
          error
        );
      }

      return {
        displayName: '',
        email: '',
        accountName: '',
        userName: '',
        userPrincipalName: '',
        source: '',
        identifier: ''
      };
    });
  }

  private _getJson(url: string): Promise<any> {
    return new Promise<any>(
      (
        resolve: (value?: any) => void,
        reject: (reason?: any) => void
      ): void => {
        const request: XMLHttpRequest = new XMLHttpRequest();

        request.open('GET', url, true);
        request.setRequestHeader(
          'Accept',
          'application/json;odata=verbose'
        );

        request.onreadystatechange = (): void => {
          if (request.readyState !== 4) {
            return;
          }

          if (request.status >= 200 && request.status < 300) {
            try {
              resolve(JSON.parse(request.responseText));
            }
            catch (parseError) {
              reject(parseError);
            }
            return;
          }

          reject(
            new Error(
              'HTTP GET failed. Status: ' +
              request.status +
              ' ' + request.statusText +
              '. URL: ' + url
            )
          );
        };

        request.send();
      }
    );
  }

  private _getSelectedPeople(): any[] {
    const pickerKey: string = this._pickerId + '_TopSpan';
    const picker: any =
      SPClientPeoplePicker.SPClientPeoplePickerDict[pickerKey];

    return picker
      ? picker.GetAllUserInfo()
      : [];
  }

  private _loadPeoplePickerScripts(callback: () => void): void {
    const layouts: string =
      window.location.protocol +
      '//' +
      window.location.host +
      '/_layouts/15/';

    SP.SOD.registerSod(
      'clienttemplates.js',
      layouts + 'clienttemplates.js'
    );
    SP.SOD.registerSod(
      'clientforms.js',
      layouts + 'clientforms.js'
    );
    SP.SOD.registerSod(
      'clientpeoplepicker.js',
      layouts + 'clientpeoplepicker.js'
    );
    SP.SOD.registerSod(
      'autofill.js',
      layouts + 'autofill.js'
    );

    SP.SOD.registerSodDep(
      'clientforms.js',
      'clienttemplates.js'
    );
    SP.SOD.registerSodDep(
      'clientpeoplepicker.js',
      'clientforms.js'
    );
    SP.SOD.registerSodDep(
      'clientpeoplepicker.js',
      'autofill.js'
    );

    SP.SOD.executeFunc(
      'clientpeoplepicker.js',
      'SPClientPeoplePicker',
      callback
    );
  }

  private _initializePeoplePicker(): void {
    const schema: any = {};

    schema.PrincipalAccountType = 'User';
    schema.SearchPrincipalSource = 15;
    schema.ResolvePrincipalSource = 15;
    schema.AllowMultipleValues = true;
    schema.MaximumEntitySuggestions = 10;
    schema.Width = '100%';

    SPClientPeoplePicker_InitStandaloneControlWrapper(
      this._pickerId,
      null,
      schema
    );
  }

  private _setBusy(isBusy: boolean): void {
    const button: HTMLButtonElement =
      document.getElementById(
        'esign-classic-continue'
      ) as HTMLButtonElement;

    if (!button) {
      return;
    }

    button.disabled = isBusy;
    button.innerHTML = isBusy
      ? 'Bearbetar...'
      : 'Fortsätt';
    button.style.opacity = isBusy ? '0.65' : '1';
  }

  private _showError(message: string, error: any): void {
    const errorElement: HTMLElement =
      document.getElementById('esign-classic-error');

    if (errorElement) {
      errorElement.innerHTML = this._escapeHtml(message);
      errorElement.style.display = 'block';
    }

    if (window.console && window.console.error) {
      window.console.error(
        '[ESIGN][CLASSIC ERROR]',
        message,
        error
      );
    }
  }

  private _generateCorrelationId(): string {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'
      .replace(/[xy]/g, (c: string): string => {
        const r: number = Math.random() * 16 | 0;
        const v: number = c === 'x'
          ? r
          : (r & 0x3 | 0x8);

        return v.toString(16);
      });
  }

  private _escapeHtml(value: string): string {
    const div: HTMLDivElement = document.createElement('div');
    div.appendChild(
      document.createTextNode(value || '')
    );
    return div.innerHTML;
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
