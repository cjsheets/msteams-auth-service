import AuthenticationContext from 'adal-angular';
import { IAuthService, Resource } from './types';

/*
 * Use ADAL.js to authenticate against AAD v1
 */
class AdalAuthService implements IAuthService {
  private authParams: URLSearchParams;

  private authContext: AuthenticationContext;

  private loginPromise: Promise<AuthenticationContext.UserInfo>;

  private loginPromiseResolve: (value?: AuthenticationContext.UserInfo) => void;

  private loginPromiseReject: (err: Error) => void;

  constructor() {
    const url = new URL(this.window.location.href);
    this.authParams = new URLSearchParams(url.search);
    this.authContext = new AuthenticationContext(this.config);
  }

  login() {
    if (!this.loginPromise) {
      this.loginPromise = new Promise((resolve, reject) => {
        this.loginPromiseResolve = resolve;
        this.loginPromiseReject = reject;
        // Start the login flow
        this.authContext.login();
      });
    }
    return this.loginPromise;
  }

  logout() {
    this.authContext.logOut();
  }

  isCallback() {
    return this.authContext.isCallback(this.window.location.hash);
  }

  loginCallback = (reason: any, token: any, error: any) => {
    if (this.loginPromise) {
      if (!error) {
        this.getUser()
          .then((user) => this.loginPromiseResolve(user))
          .catch((err) => {
            this.loginPromiseReject(err);
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
      this.authContext.getUser((error, user) => {
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
      this.authContext.acquireToken(this.config.endpoints.graph, (reason, token, error) => {
        if (!error) {
          resolve(token);
        } else if (error === 'login required') {
          this.login()
            .then(this.getToken)
            .then((_token) => resolve(_token))
            .catch(({ _error, _reason }) => reject({ error: _error, reason: _reason }));
        } else {
          reject({ error, reason });
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
      popUp: !(this.window.navigator as any).standalone,
      postLogoutRedirectUri: `${this.window.location.origin}/${process.env.ADAL_REDIRECT_PATH}`,
      redirectUri: `${this.window.location.origin}/${process.env.ADAL_REDIRECT_PATH}`,
      tenant: this.authParams.get('tenantId') || 'common',
    };
  }

  private get window() {
    return window || global;
  }
}

export default AdalAuthService;
