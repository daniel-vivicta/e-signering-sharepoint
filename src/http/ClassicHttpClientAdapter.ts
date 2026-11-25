import { IHttpClient } from './IHttpClient';

/**
 * Classic SharePoint implementation of IHttpClient.
 * GET uses same-origin XMLHttpRequest.
 * POST first obtains a request digest from the target SharePoint web.
 */
export class ClassicHttpClientAdapter implements IHttpClient {

  public getJson(url: string): Promise<any> {
    return this._request('GET', url);
  }

  public postJson(url: string, body: any): Promise<any> {
    return this._getRequestDigest(url)
      .then((digest: string): Promise<any> => {
        return this._request(
          'POST',
          url,
          body,
          digest
        );
      });
  }

  private _getRequestDigest(targetUrl: string): Promise<string> {
    const apiIndex: number = targetUrl.toLowerCase().indexOf('/_api/');

    if (apiIndex < 0) {
      return Promise.reject(
        new Error(
          'Could not determine SharePoint web URL from: ' + targetUrl
        )
      );
    }

    const webUrl: string = targetUrl.substring(0, apiIndex);
    const contextInfoUrl: string = webUrl + '/_api/contextinfo';

    return new Promise<string>(
      (
        resolve: (value?: string) => void,
        reject: (reason?: any) => void
      ): void => {
        const request: XMLHttpRequest = new XMLHttpRequest();

        request.open('POST', contextInfoUrl, true);
        request.setRequestHeader(
          'Accept',
          'application/json;odata=verbose'
        );
        request.setRequestHeader(
          'Content-Type',
          'application/json;odata=verbose'
        );
        request.setRequestHeader('odata-version', '');

        request.onreadystatechange = (): void => {
          if (request.readyState !== 4) {
            return;
          }

          if (request.status >= 200 && request.status < 300) {
            try {
              const response: any = JSON.parse(request.responseText);
              const digest: string =
                response.d &&
                response.d.GetContextWebInformation &&
                response.d.GetContextWebInformation.FormDigestValue
                  ? response.d.GetContextWebInformation.FormDigestValue
                  : '';

              if (!digest) {
                reject(
                  new Error(
                    'SharePoint contextinfo returned no FormDigestValue.'
                  )
                );
                return;
              }

              resolve(digest);
            }
            catch (parseError) {
              reject(parseError);
            }

            return;
          }

          reject(
            new Error(
              'Could not get SharePoint request digest. Status: ' +
              request.status + ' ' + request.statusText +
              '. URL: ' + contextInfoUrl +
              '. Response: ' + request.responseText
            )
          );
        };

        request.onerror = (): void => {
          reject(
            new Error(
              'Network error while getting request digest: ' +
              contextInfoUrl
            )
          );
        };

        request.send('');
      }
    );
  }

  private _request(
    method: string,
    url: string,
    body?: any,
    requestDigest?: string
  ): Promise<any> {
    return new Promise<any>(
      (
        resolve: (value?: any) => void,
        reject: (reason?: any) => void
      ): void => {
        const request: XMLHttpRequest = new XMLHttpRequest();
        request.open(method, url, true);

        request.setRequestHeader(
          'Accept',
          'application/json;odata=verbose'
        );

        if (method === 'POST') {
          request.setRequestHeader(
            'Content-Type',
            'application/json;odata=verbose'
          );
          request.setRequestHeader('odata-version', '');

          if (requestDigest) {
            request.setRequestHeader(
              'X-RequestDigest',
              requestDigest
            );
          }
        }

        request.onreadystatechange = (): void => {
          if (request.readyState !== 4) {
            return;
          }

          if (request.status >= 200 && request.status < 300) {
            if (!request.responseText) {
              resolve({});
              return;
            }

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
              'HTTP ' + method + ' failed. Status: ' +
              request.status + ' ' + request.statusText +
              '. URL: ' + url +
              '. Response: ' + request.responseText
            )
          );
        };

        request.onerror = (): void => {
          reject(
            new Error(
              'Network error while requesting: ' + url
            )
          );
        };

        request.send(
          method === 'POST'
            ? JSON.stringify(body)
            : undefined
        );
      }
    );
  }
}
