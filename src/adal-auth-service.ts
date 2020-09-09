import AuthenticationContext from 'adal-angular';

import { IAuthService, Resource } from '.';

/*
 * Use ADAL.js to authenticate against AAD v1
 */
class AdalAuthService implements IAuthService {
  private _authParams: URLSearchParams;
  private _authContext: AuthenticationContext;

  private loginPromise: Promise<AuthenticationContext.UserInfo>;
  private loginPromiseResolve: (value?: AuthenticationContext.UserInfo) => void;
  private loginPromiseReject: (err: Error) => void;

  constructor() {
    const url = new URL(window.location.href);
    this._authParams = new URLSearchParams(url.search);
    this._authContext = new AuthenticationContext(this.config);
  }

  login() {
    if (!this.loginPromise) {
      this.loginPromise = new Promise((resolve, reject) => {
        this.loginPromiseResolve = resolve;
        this.loginPromiseReject = reject;
        // Start the login flow
        this._authContext.login();
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

  loginCallback = (reason: any, token: any, error: any) => {
    if (this.loginPromise) {
      if (!error) {
        this.getUser()
          .then((user) => this.loginPromiseResolve(user))
          .catch((error) => {
            this.loginPromiseReject(error);
            this.loginPromise = undefined;
          });
      } else {
        this.loginPromiseReject(new Error(`${error}: ${reason}`));
        this.loginPromise = undefined;
      }
    }
  };

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

  getToken = () => {
    return new Promise<string>((resolve, reject) => {
      this._authContext.acquireToken(this.config.endpoints.graph, (reason, token, error) => {
        if (!error) {
          resolve(token);
        } else {
          if (error === 'login required') {
            this.login()
              .then(this.getToken)
              .then((token) => resolve(token))
              .catch(({ error, reason }) => reject({ error, reason }));
          } else {
            reject({ error, reason });
          }
        }
      });
    });
  };

  get config() {
    const scopes = encodeURIComponent(
      'email profile User.ReadBasic.All, User.Read.All, Group.Read.All, Directory.Read.All'
    );

    return {
      cacheLocation: 'localStorage' as 'localStorage' | 'sessionStorage',
      callback: this.loginCallback,
      clientId: process.env.CLIENT_ID,
      endpoints: { ...Resource },
      extraQueryParameter: `prompt=consent&scope=${scopes}`,
      instance: 'https://login.microsoftonline.com/',
      navigateToLoginRequestUrl: false,
      popUp: !(window.navigator as any).standalone,
      postLogoutRedirectUri: `${window.location.origin}/${process.env.ADAL_REDIRECT_PATH}`,
      redirectUri: `${window.location.origin}/${process.env.ADAL_REDIRECT_PATH}`,
      tenant: this._authParams.get('tenantId') || 'common',
    };
  }
}

export default AdalAuthService;
