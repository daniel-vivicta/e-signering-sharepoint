/**
 * Small HTTP abstraction shared by Modern SPFx and Classic SharePoint.
 */
export interface IHttpClient {
    getJson(url: string): Promise<any>;
    postJson(url: string, body: any): Promise<any>;
  }
  