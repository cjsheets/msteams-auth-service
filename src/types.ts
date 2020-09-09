import AuthenticationContext from 'adal-angular';

export interface IAuthService {
  config: AuthenticationContext.Options | MsalConfig;
  getToken(resource: Resource): Promise<string>;
  login(): Promise<AuthenticationContext.UserInfo>;
  logout(): void;
  getUser(): Promise<AuthenticationContext.UserInfo>;
  isCallback(location: string): boolean;
}

export enum Resource {
  graph = 'https://graph.microsoft.com/',
}

interface MsalConfig {
  auth: {
    clientId: string;
    redirectUri: string;
  };
}
