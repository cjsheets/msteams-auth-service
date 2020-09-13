import { insideIframe } from './utility';
import AdalAuthService from './adal-auth-service';
import MockAuthService from './mock-auth-service';
import MsalAuthService from './msal-auth-service';
import TeamsAuthService from './teams-auth-service';
import TeamsSsoAuthService from './teams-sso-auth-service';
import { IAuthService, Resource } from './types';

export default class AuthService implements IAuthService {
  private authService: IAuthService;

  constructor() {
    this.initAuthService();
  }

  get config() {
    return this.authService.config;
  }

  isCallback() {
    return this.authService.isCallback(this.window.location.hash);
  }

  login() {
    return this.authService.login();
  }

  logout() {
    this.authService.logout();
  }

  getToken(resource: Resource) {
    return this.authService.getToken(resource);
  }

  getUser() {
    return this.authService.getUser();
  }

  private initAuthService() {
    const url = new URL(this.window.location.href);
    const params = new URLSearchParams(url.search);

    if (params.get('mockData')) {
      this.authService = new MockAuthService();
    } else if (params.get('isTeamsFrame') || insideIframe()) {
      // Teams doesn't allow query parameters for Team scope URIs
      this.authService = new TeamsAuthService();
    } else if (params.get('isTeamsFrameSSO')) {
      this.authService = new TeamsSsoAuthService();
    } else if (
      params.get('useV2') ||
      url.pathname.indexOf(`/${process.env.ADAL_REDIRECT_PATH}`) !== -1
    ) {
      this.authService = new MsalAuthService();
    } else {
      this.authService = new AdalAuthService();
    }
  }

  private get window() {
    return window || global;
  }
}
