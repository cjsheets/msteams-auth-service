import AuthenticationContext from 'adal-angular';

export interface IAuthService {
  config: AuthenticationContext.Options | MsalOptions;
  getToken(resource: Resource): Promise<string>;
  login(): Promise<AuthenticationContext.UserInfo>;
  logout(): void;
  getUser(): Promise<AuthenticationContext.UserInfo>;
  isCallback(location: string): boolean;
}

export enum Resource {
  graph = 'https://graph.microsoft.com/',
}

export interface MsalOptions {
  auth: {
    clientId: string;
    redirectUri: string;
  };
}
