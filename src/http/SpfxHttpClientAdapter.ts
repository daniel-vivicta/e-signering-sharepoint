import {
    SPHttpClient,
    SPHttpClientResponse
  } from '@microsoft/sp-http';
  
  import { IHttpClient } from './IHttpClient';
  
  export class SpfxHttpClientAdapter implements IHttpClient {
  
    private _spHttpClient: SPHttpClient;
  
    constructor(spHttpClient: SPHttpClient) {
      this._spHttpClient = spHttpClient;
    }
  
    public getJson(url: string): Promise<any> {
      return this._spHttpClient
        .get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse): Promise<any> => {
          if (!response.ok) {
            throw new Error(
              'HTTP GET failed. Status: ' + response.status +
              ' ' + response.statusText + '. URL: ' + url
            );
          }
          return response.json();
        });
    }
  
    public postJson(url: string, body: any): Promise<any> {
      return this._spHttpClient
        .post(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'odata-version': ''
              },
            body: JSON.stringify(body)
          }
        )
        .then((response: SPHttpClientResponse): Promise<any> => {
            if (!response.ok) {

                return response.text()
                  .then((errorText: string): any => {
              
                    throw new Error(
                      'HTTP POST failed. Status: ' +
                      response.status +
                      ' ' +
                      response.statusText +
                      '. URL: ' +
                      url +
                      '. Response: ' +
                      errorText
                    );
                  });
              }
              
              return response.json();
        });
    }
  }
  