import * as microsoftTeams from '@microsoft/teams-js';
import AuthenticationContext from 'adal-angular';

import { IAuthService, Resource } from '.';

/*
 * Use ADAL.js and Teams.js library to authenticate against AAD v1
 */
class TeamsAuthService implements IAuthService {
  private _authParams: URLSearchParams;
  private _authContext: AuthenticationContext;

  private loginPromise: Promise<AuthenticationContext.UserInfo>;

  constructor() {
    microsoftTeams.initialize();
    microsoftTeams.getContext(function (context) {});
    const url = new URL(window.location.href);
    this._authParams = new URLSearchParams(url.search);
    this._authContext = new AuthenticationContext(this.config);
  }

  login() {
    if (!this.loginPromise) {
      this.loginPromise = new Promise<AuthenticationContext.UserInfo>((resolve, reject) => {
        this.ensureLoginHint().then(() => {
          // Start the login flow
          microsoftTeams.authentication.authenticate({
            url: `${window.location.origin}/tab/silent-start`,
            width: 600,
            height: 535,
            successCallback: () => {
              resolve(this.getUser());
            },
            failureCallback: (reason) => {
              reject(reason);
            },
          });
        });
      });
    }
    return this.loginPromise;
  }

  logout() {
    this._authContext.logOut();
  }

  isCallback() {
    return this._authContext.isCallback(window.location.hash);
  }

  getUser() {
    return new Promise<AuthenticationContext.UserInfo>((resolve, reject) => {
      this._authContext.getUser((error, user) => {
        if (!error) {
          resolve(user);
        } else {
          reject(error);
        }
      });
    });
  }

  getToken(resource: Resource) {
    return new Promise<string>((resolve, reject) => {
      this.ensureLoginHint().then(() => {
        this._authContext.acquireToken(resource, (reason, token, error) => {
          if (!error) {
            resolve(token);
          } else {
            reject({ error, reason });
          }
        });
      });
    });
  }

  ensureLoginHint() {
    return new Promise((resolve, reject) => {
      microsoftTeams.getContext((context) => {
        const scopes = encodeURIComponent(
          'email profile User.ReadBasic.All, User.Read.All, Group.Read.All, Directory.Read.All'
        );

        // Setup extra query parameters for ADAL
        // - openid and profile scope adds profile information to the id_token
        // - login_hint provides the expected user name
        if (context.loginHint) {
          this._authContext.config.extraQueryParameter = `prompt=consent&scope=${scopes}&login_hint=${encodeURIComponent(
            context.loginHint
          )}`;
        } else {
          this._authContext.config.extraQueryParameter = `prompt=consent&scope=${scopes}`;
        }
        resolve();
      });
    });
  }

  get config() {
    return {
      cacheLocation: 'localStorage' as 'localStorage' | 'sessionStorage',
      clientId: process.env.CLIENT_ID,
      endpoints: { ...Resource },
      extraQueryParameter: '',
      instance: 'https://login.microsoftonline.com/',
      navigateToLoginRequestUrl: false,
      postLogoutRedirectUri: `${window.location.origin}/${process.env.ADAL_REDIRECT_PATH}`,
      redirectUri: `${window.location.origin}/${process.env.ADAL_REDIRECT_PATH}`,
      tenant: this._authParams.get('tenantId') || 'common',
    };
  }
}

export default TeamsAuthService;
