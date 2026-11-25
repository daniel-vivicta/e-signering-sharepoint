import * as React from 'react';

import { PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { SPHttpClient } from '@microsoft/sp-http';

import { ISigner } from '../../models/ISigner';
import {
  ISigningRequest,
  ISigningDocument,
  ISigningInitiator
} from '../../models/ISigningRequest';

import { UserProfileService } from '../../services/UserProfileService';
import {
  SigningQueueService,
  IQueueCreateResult
} from '../../services/SigningQueueService';
import { SpfxHttpClientAdapter } from '../../http/SpfxHttpClientAdapter';

export interface ISigningFormProps {
  fileName: string;
  context: any;
  onCancel: () => void;
  correlationId: string;
  document: ISigningDocument;
}

export interface ISigningFormState {
  signers: ISigner[];
  message: string;
  deadline: string;
  showValidationError: boolean;
  isSubmitting: boolean;
  submitError: string;
}

export class SigningForm
  extends React.Component<ISigningFormProps, ISigningFormState> {

  private _userProfileService: UserProfileService;
  private _queueService: SigningQueueService;

  constructor(props: ISigningFormProps) {
    super(props);

    const httpClient: SpfxHttpClientAdapter =
      new SpfxHttpClientAdapter(
        this.props.context.spHttpClient
      );

    this._userProfileService =
      new UserProfileService(
        httpClient,
        this.props.context.pageContext.web.absoluteUrl
      );

    this._queueService =
      new SigningQueueService(
        httpClient,
        'http://sharepointdev.com/enh/site01',
        'E-signering - Kö'
      );

    this.state = {
      signers: [],
      message: '',
      deadline: '',
      showValidationError: false,
      isSubmitting: false,
      submitError: ''
    };

    this._onPeopleChanged = this._onPeopleChanged.bind(this);
    this._onMessageChanged = this._onMessageChanged.bind(this);
    this._onDeadlineChanged = this._onDeadlineChanged.bind(this);
    this._onSubmit = this._onSubmit.bind(this);
  }

  public render(): JSX.Element {
    return (
      <div>

        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            marginBottom: '24px'
          }}
        >
          <img
            src="https://localhost:4321/temp/signing.png"
            alt=""
            style={{
              width: '40px',
              height: '40px',
              marginRight: '12px'
            }}
          />

          <div>
            <div
              style={{
                fontSize: '21px',
                fontWeight: 600,
                color: '#323130'
              }}
            >
              Starta signering
            </div>

            <div
              style={{
                fontSize: '13px',
                color: '#605e5c',
                marginTop: '2px'
              }}
            >
              {this.props.fileName}
            </div>
          </div>
        </div>

        <div>
          <PeoplePicker
            context={this.props.context}
            titleText="Undertecknare"
            personSelectionLimit={10}
            showtooltip={true}
            isRequired={false}
            selectedItems={this._onPeopleChanged}
          />
        </div>

        {this.state.showValidationError && this.state.signers.length === 0 &&
          <div
            style={{
              marginTop: '6px',
              color: '#a4262c',
              fontSize: '13px'
            }}
          >
            Välj minst en undertecknare.
          </div>
        }

        {this.state.submitError &&
          <div
            style={{
              marginTop: '12px',
              color: '#a4262c',
              fontSize: '13px'
            }}
          >
            {this.state.submitError}
          </div>
        }

        <div style={{ marginTop: '24px' }}>
          <label
            style={{
              display: 'block',
              fontSize: '14px',
              fontWeight: 600,
              color: '#323130',
              marginBottom: '6px'
            }}
          >
            Meddelande
          </label>

          <textarea
            value={this.state.message}
            onChange={this._onMessageChanged}
            rows={5}
            style={{
              width: '100%',
              boxSizing: 'border-box',
              border: '1px solid #8a8886',
              padding: '8px',
              fontFamily: 'Segoe UI',
              fontSize: '14px',
              resize: 'vertical'
            }}
          />
        </div>

        <div style={{ marginTop: '20px' }}>
          <label
            style={{
              display: 'block',
              fontSize: '14px',
              fontWeight: 600,
              color: '#323130',
              marginBottom: '6px'
            }}
          >
            Deadline
          </label>

          <input
            type="date"
            value={this.state.deadline}
            onChange={this._onDeadlineChanged}
            style={{
              height: '32px',
              border: '1px solid #8a8886',
              padding: '0 8px',
              fontFamily: 'Segoe UI',
              fontSize: '14px',
              boxSizing: 'border-box'
            }}
          />
        </div>

        <div
          style={{
            display: 'flex',
            justifyContent: 'flex-end',
            marginTop: '32px'
          }}
        >
          <button
            onClick={this.props.onCancel}
            disabled={this.state.isSubmitting}
            style={{
              minWidth: '80px',
              height: '32px',
              marginRight: '8px',
              border: '1px solid #8a8886',
              backgroundColor: '#ffffff',
              color: '#323130',
              fontFamily: 'Segoe UI',
              fontSize: '14px',
              cursor: this.state.isSubmitting ? 'default' : 'pointer'
            }}
          >
            Avbryt
          </button>

          <button
            onClick={this._onSubmit}
            disabled={this.state.isSubmitting}
            style={{
              minWidth: '80px',
              height: '32px',
              border: '1px solid #0078d4',
              backgroundColor: '#0078d4',
              color: '#ffffff',
              fontFamily: 'Segoe UI',
              fontSize: '14px',
              fontWeight: 600,
              cursor: this.state.isSubmitting ? 'default' : 'pointer',
              opacity: this.state.isSubmitting ? 0.65 : 1
            }}
          >
            {this.state.isSubmitting ? 'Skapar...' : 'Fortsätt'}
          </button>
        </div>

      </div>
    );
  }

  private _onPeopleChanged(items: any[]): void {
    const signerPromises: Array<Promise<ISigner>> =
      items.map(
        (item: any): Promise<ISigner> => {
          return this._mapPeoplePickerItem(item);
        }
      );

    Promise.all(signerPromises)
      .then((signers: ISigner[]): void => {
        this.setState({
          signers: signers,
          showValidationError: signers.length === 0,
          submitError: ''
        });
      })
      .catch((error: any): void => {
        console.error('[ESIGN][SIGNER MAPPING ERROR]', error);
        this.setState({
          signers: [],
          submitError: 'Det gick inte att hämta uppgifter om undertecknaren.'
        });
      });
  }

  private _mapPeoplePickerItem(item: any): Promise<ISigner> {
    const userId: number = item.id;

    return this._getSharePointUser(userId)
      .then((spUser: any): Promise<any> => {
        return this._userProfileService.getProfile(
          spUser.LoginName
        );
      })
      .then((profile: any): ISigner => {
        const signer: ISigner =
          this._userProfileService.mapProfileToSigner(
            profile,
            {
              id: userId,
              accountName: profile.AccountName || '',
              displayName: profile.DisplayName || '',
              email: profile.Email || ''
            }
          );

        console.log('[ESIGN][SIGNER]', signer);
        return signer;
      });
  }

  private _getSharePointUser(userId: number): Promise<any> {
    const webUrl: string =
      this.props.context.pageContext.web.absoluteUrl;

    const url: string =
      webUrl + '/_api/web/getuserbyid(' + userId + ')';

    return this.props.context.spHttpClient
      .get(
        url,
        SPHttpClient.configurations.v1
      )
      .then((response: any): Promise<any> => {
        if (!response.ok) {
          throw new Error(
            'Could not get SharePoint user. HTTP ' +
            response.status
          );
        }

        return response.json();
      });
  }

  private _getInitiator(): Promise<ISigningInitiator> {
    const accountName: string =
      this.props.context.pageContext.user.loginName;

    return this._userProfileService
      .getProfile(accountName)
      .then((profile: any): ISigningInitiator => {
        const userName: string =
          this._userProfileService.getProfileProperty(
            profile,
            'UserName'
          );

        const userPrincipalName: string =
          this._userProfileService.getProfileProperty(
            profile,
            'SPS-UserPrincipalName'
          );

        return {
          displayName: profile.DisplayName || '',
          email: profile.Email || '',
          accountName: profile.AccountName || accountName,
          userName: userName,
          userPrincipalName: userPrincipalName,
          source: this._userProfileService.determineSource(
            userPrincipalName,
            profile.AccountName || accountName
          ),
          identifier: userName ? 'cn=' + userName : ''
        };
      });
  }

  private _onMessageChanged(
    event: React.ChangeEvent<HTMLTextAreaElement>
  ): void {
    this.setState({
      message: event.target.value
    });
  }

  private _onDeadlineChanged(
    event: React.ChangeEvent<HTMLInputElement>
  ): void {
    this.setState({
      deadline: event.target.value
    });
  }

  private _onSubmit(): void {
    if (!this.state.signers || this.state.signers.length === 0) {
      this.setState({
        showValidationError: true
      });
      return;
    }

    this.setState({
      showValidationError: false,
      isSubmitting: true,
      submitError: ''
    });

    this._getInitiator()
      .then((initiator: ISigningInitiator): Promise<IQueueCreateResult> => {
        const request: ISigningRequest = {
          correlationId: this.props.correlationId,
          document: this.props.document,
          signers: this.state.signers,
          message: this.state.message,
          deadline: this.state.deadline,
          initiator: initiator
        };

        console.log(
          '[ESIGN][SIGNING REQUEST]',
          request
        );
        
        console.log(
          '[ESIGN][QUEUE] Calling SigningQueueService.create()'
        );
        
        return this._queueService.create(request);
      })
      .then((result: IQueueCreateResult): void => {

        console.log(
          '[ESIGN][QUEUE ITEM CREATED]',
          result
        );
      
        this.setState({
          isSubmitting: false
        });
      
        this.props.onCancel();
      })
      
      .catch((error: any): void => {
        console.error('[ESIGN][QUEUE ERROR]', error);

        this.setState({
          isSubmitting: false,
          submitError: 'Det gick inte att skapa signeringsbegäran i kön.'
        });
      });
  }
}
