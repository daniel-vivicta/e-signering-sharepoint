import { ISigner } from '../models/ISigner';
import { IHttpClient } from '../http/IHttpClient';

export interface IUserIdentityFallback {
  id?: number;
  displayName?: string;
  email?: string;
  accountName: string;
  userName?: string;
}

/**
 * Shared profile and identity service.
 *
 * This class has no dependency on SPFx, React, Classic SharePoint,
 * SPHttpClient or XMLHttpRequest. HTTP is supplied through IHttpClient.
 */
export class UserProfileService {

  private _httpClient: IHttpClient;
  private _webUrl: string;

  constructor(
    httpClient: IHttpClient,
    webUrl: string
  ) {
    this._httpClient = httpClient;
    this._webUrl = this._trimTrailingSlash(webUrl);
  }

  public getProfile(accountName: string): Promise<any> {

    const encodedAccount: string =
      encodeURIComponent(accountName);

    const url: string =
      this._webUrl +
      '/_api/SP.UserProfiles.PeopleManager/' +
      'GetPropertiesFor(accountName=@v)' +
      "?@v='" + encodedAccount + "'";

    return this._httpClient.getJson(url)
      .then((response: any): any => {
        return response && response.d
          ? response.d
          : response;
      });
  }

  public getProfileProperty(
    profile: any,
    propertyName: string
  ): string {

    if (!profile || !profile.UserProfileProperties) {
      return '';
    }

    let properties: any[] =
      profile.UserProfileProperties;

    if (profile.UserProfileProperties.results) {
      properties =
        profile.UserProfileProperties.results;
    }

    const property: any = properties.filter(
      (item: any): boolean => {
        return item.Key === propertyName;
      }
    )[0];

    return property && property.Value
      ? property.Value
      : '';
  }

  public mapProfileToSigner(
    profile: any,
    fallback: IUserIdentityFallback
  ): ISigner {

    let userName: string =
      this.getProfileProperty(
        profile,
        'UserName'
      );

    const userPrincipalName: string =
      this.getProfileProperty(
        profile,
        'SPS-UserPrincipalName'
      );

    const accountName: string =
      profile.AccountName ||
      fallback.accountName ||
      '';

    if (!userName) {
      userName = fallback.userName || '';
    }

    return {
      id: fallback.id,
      displayName:
        profile.DisplayName ||
        fallback.displayName ||
        '',
      email:
        profile.Email ||
        fallback.email ||
        '',
      accountName: accountName,
      userName: userName,
      userPrincipalName: userPrincipalName,
      source: this.determineSource(
        userPrincipalName,
        accountName
      ),
      identifier:
        userName
          ? 'cn=' + userName
          : ''
    };
  }

  public getSigner(
    fallback: IUserIdentityFallback
  ): Promise<ISigner> {

    return this.getProfile(fallback.accountName)
      .then((profile: any): ISigner => {
        return this.mapProfileToSigner(
          profile,
          fallback
        );
      });
  }

  public determineSource(
    userPrincipalName: string,
    accountName: string
  ): string {

    const identity: string =
      (
        (userPrincipalName || '') +
        ' ' +
        (accountName || '')
      ).toLowerCase();

    if (
      identity.indexOf('@eduad.stockholm.se') >= 0 ||
      identity.indexOf('eduad\\') >= 0
    ) {
      return 'EDUAD';
    }

    if (
      identity.indexOf('@ad.stockholm.se') >= 0 ||
      identity.indexOf('fobad\\') >= 0
    ) {
      return 'FOBAD';
    }

    // Development fallback only.
    // Verify the final rule with real claims in TEST.
    return 'FOBAD';
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
