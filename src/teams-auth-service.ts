import * as microsoftTeams from '@microsoft/teams-js';
import AuthenticationContext from 'adal-angular';
import { IAuthService, Resource } from './types';

/*
 * Use ADAL.js and Teams.js library to authenticate against AAD v1
 */
class TeamsAuthService implements IAuthService {
  private authParams: URLSearchParams;

  private authContext: AuthenticationContext;

  private loginPromise: Promise<AuthenticationContext.UserInfo>;

  constructor() {
    microsoftTeams.initialize();
    microsoftTeams.getContext(function getContext() {});
    const url = new URL(window.location.href);
    this.authParams = new URLSearchParams(url.search);
    this.authContext = new AuthenticationContext(this.config);
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
    this.authContext.logOut();
  }

  isCallback() {
    return this.authContext.isCallback(window.location.hash);
  }

  getUser() {
    return new Promise<AuthenticationContext.UserInfo>((resolve, reject) => {
      this.authContext.getUser((error, user) => {
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
        this.authContext.acquireToken(resource, (reason, token, error) => {
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
    return new Promise((resolve) => {
      microsoftTeams.getContext((context) => {
        const scopes = encodeURIComponent(
          'email profile User.ReadBasic.All, User.Read.All, Group.Read.All, Directory.Read.All'
        );

        // Setup extra query parameters for ADAL
        // - openid and profile scope adds profile information to the id_token
        // - login_hint provides the expected user name
        if (context.loginHint) {
          this.authContext.config.extraQueryParameter = `prompt=consent&scope=${scopes}&login_hint=${encodeURIComponent(
            context.loginHint
          )}`;
        } else {
          this.authContext.config.extraQueryParameter = `prompt=consent&scope=${scopes}`;
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
      tenant: this.authParams.get('tenantId') || 'common',
    };
  }
}

export default TeamsAuthService;
