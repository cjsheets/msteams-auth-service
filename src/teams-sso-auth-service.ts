import * as microsoftTeams from '@microsoft/teams-js';
import AuthenticationContext from 'adal-angular';
import { IAuthService, Resource } from './types';
import TeamsAuthService from './teams-auth-service';

/*
 * Use Teams.js library to request tokens for logged in user
 */
class TeamsSsoAuthService implements IAuthService {
  private token: string;

  private authService: TeamsAuthService;

  constructor() {
    microsoftTeams.initialize();

    this.token = null;
  }

  login() {
    if (!this.authService) {
      this.authService = new TeamsAuthService();
    }
    return this.authService.login();
  }

  logout() {
    this.authService.logout();
  }

  isCallback() {
    if (!this.authService) {
      this.authService = new TeamsAuthService();
    }
    return this.authService.isCallback();
  }

  static parseTokenToUser(token: string): AuthenticationContext.UserInfo {
    // parse JWT token to object
    const base64Url = token.split('.')[1];
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const parsedToken = JSON.parse(window.atob(base64));
    const nameParts = parsedToken.name.split(' ');
    return {
      userName: parsedToken.name,
      profile: {
        family_name: nameParts.length > 1 ? nameParts[1] : 'n/a',
        given_name: nameParts.length > 0 ? nameParts[0] : 'n/a',
        upn: parsedToken.preferred_username,
        name: parsedToken.name,
      },
    };
  }

  getUser() {
    return new Promise<AuthenticationContext.UserInfo>((resolve, reject) => {
      if (this.token) {
        resolve(TeamsSsoAuthService.parseTokenToUser(this.token));
      } else {
        this.getToken(Resource.graph)
          .then((token) => {
            resolve(TeamsSsoAuthService.parseTokenToUser(token));
          })
          .catch((reason) => {
            reject(reason);
          });
      }
    });
  }

  getToken(resource: Resource) {
    return new Promise<string>((resolve, reject) => {
      if (this.token) {
        resolve(this.token);
      } else {
        microsoftTeams.authentication.getAuthToken({
          resources: [resource],
          successCallback: (result) => {
            this.token = result;
            resolve(result);
          },
          failureCallback: (reason) => {
            reject(reason);
          },
        });
      }
    });
  }

  get config() {
    return this.authService.config;
  }
}

export default TeamsSsoAuthService;
