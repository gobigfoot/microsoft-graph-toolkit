/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider, ProviderState, createFromProvider } from '@microsoft/mgt-element';

/**
 * AadTokenProvider
 *
 * @interface AadTokenProvider
 */
declare interface AadTokenProvider {
  /**
   * get token with x
   *
   * @param {string} x
   * @memberof AadTokenProvider
   */
  getToken(x: string);
}

/**
 * contains the contextual services available to a web part
 *
 * @export
 * @interface WebPartContext
 */
declare interface WebPartContext {
  // tslint:disable-next-line: completed-docs
  aadTokenProviderFactory: any;
  // tslint:disable-next-line: completed-docs
  msGraphClientFactory: any;
}

/**
 * SharePoint Provider handler
 *
 * @export
 * @class SharePointProvider
 * @extends {IProvider}
 */
export class SharePointProvider extends IProvider {
  /**
   * returns _provider
   *
   * @readonly
   * @memberof SharePointProvider
   */
  get provider() {
    return this._provider;
  }

  /**
   * returns _idToken
   *
   * @readonly
   * @type {boolean}
   * @memberof SharePointProvider
   */
  get isLoggedIn(): boolean {
    return !!this._idToken;
  }

  /**
   * privilege level for authenication
   *
   * @type {string[]}
   * @memberof SharePointProvider
   */
  public scopes: string[];

  /**
   * authority
   *
   * @type {string}
   * @memberof SharePointProvider
   */
  public authority: string;
  private _idToken: string;

  /**
   * returns _baseUrl
   *
   * @readonly
   * @type {string}
   * @memberof SharePointProvider
   */
  get baseUrl(): string {
    return this._baseUrl;
  }
  private _baseUrl: string;

  private _provider: AadTokenProvider;

  constructor(context: WebPartContext, baseUrl?: string) {
    super();
    this.setBaseUrl(context, baseUrl)
      .then(() => {
        context.aadTokenProviderFactory.getTokenProvider().then(
          (tokenProvider: AadTokenProvider): void => {
            this._provider = tokenProvider;
            this.graph = createFromProvider(this);
            this.internalLogin();
          }
        );

      })
  }

  private async setBaseUrl(context: WebPartContext, baseUrl?: string) {
    if(!baseUrl) {
      baseUrl = await context.msGraphClientFactory.getClient().then(client => client.constructor._graphBaseUrl).catch(e => 'https://graph.microsoft.com');
    }
    this._baseUrl = baseUrl;
  }

  /**
   * uses provider to receive access token via SharePoint Provider
   *
   * @returns {Promise<string>}
   * @memberof SharePointProvider
   */
  public async getAccessToken(): Promise<string> {
    let accessToken: string;
    try {
      accessToken = await this.provider.getToken(this.baseUrl);
    } catch (e) {
      throw e;
    }
    return accessToken;
  }
  /**
   * update scopes
   *
   * @param {string[]} scopes
   * @memberof SharePointProvider
   */
  public updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }

  private async internalLogin(): Promise<void> {
    this._idToken = await this.getAccessToken();
    this.setState(this._idToken ? ProviderState.SignedIn : ProviderState.SignedOut);
  }
}
