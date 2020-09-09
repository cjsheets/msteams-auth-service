import AuthenticationContext from 'adal-angular';

import { insideIframe } from './utility';
import AdalAuthService from './adal-auth-service';
import MockAuthService from './mock-auth-service';
import MsalAuthService from './msal-auth-service';
import TeamsAuthService from './teams-auth-service';
import TeamsSsoAuthService from './teams-sso-auth-service';

export interface IAuthService {
  config: AuthenticationContext.Options | MsalAuthService['config'];
  getToken(resource: Resource): Promise<string>;
  login(): Promise<AuthenticationContext.UserInfo>;
  logout(): void;
  getUser(): Promise<AuthenticationContext.UserInfo>;
  isCallback(location: string): boolean;
}

export enum Resource {
  graph = 'https://graph.microsoft.com/',
}

class AuthService implements IAuthService {
  private _authService: IAuthService;

  constructor() {
    this.initAuthService();
  }

  get config() {
    return this._authService.config;
  }

  isCallback() {
    return this._authService.isCallback(window.location.hash);
  }

  login() {
    return this._authService.login();
  }

  logout() {
    this._authService.logout();
  }

  getToken(resource: Resource) {
    return this._authService.getToken(resource);
  }

  getUser() {
    return this._authService.getUser();
  }

  private initAuthService() {
    const url = new URL(window.location.href);
    const params = new URLSearchParams(url.search);

    if (params.get('mockData')) {
      this._authService = new MockAuthService();
    } else if (params.get('isTeamsFrame') || insideIframe()) {
      // Teams doesn't allow query parameters for Team scope URIs
      this._authService = new TeamsAuthService();
    } else if (params.get('isTeamsFrameSSO')) {
      this._authService = new TeamsSsoAuthService();
    } else if (
      params.get('useV2') ||
      url.pathname.indexOf(`/${process.env.ADAL_REDIRECT_PATH}`) !== -1
    ) {
      this._authService = new MsalAuthService();
    } else {
      this._authService = new AdalAuthService();
    }
  }
}

export default new AuthService();
