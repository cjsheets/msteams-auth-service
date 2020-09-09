import * as microsoftTeams from '@microsoft/teams-js';
import AuthenticationContext from 'adal-angular';

import { IAuthService, Resource } from '.';
import TeamsAuthService from './teams-auth-service';

/*
 * Use Teams.js library to request tokens for logged in user
 */
class TeamsSsoAuthService implements IAuthService {
  private _token: string;
  private _authService: TeamsAuthService;

  constructor() {
    microsoftTeams.initialize();

    this._token = null;
  }

  login() {
    if (!this._authService) {
      this._authService = new TeamsAuthService();
    }
    return this._authService.login();
  }

  logout() {
    this._authService.logout();
  }

  isCallback() {
    if (!this._authService) {
      this._authService = new TeamsAuthService();
    }
    return this._authService.isCallback();
  }

  parseTokenToUser(token: string): AuthenticationContext.UserInfo {
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
      if (this._token) {
        resolve(this.parseTokenToUser(this._token));
      } else {
        this.getToken(Resource.graph)
          .then((token) => {
            resolve(this.parseTokenToUser(token));
          })
          .catch((reason) => {
            reject(reason);
          });
      }
    });
  }

  getToken(resource: Resource) {
    return new Promise<string>((resolve, reject) => {
      if (this._token) {
        resolve(this._token);
      } else {
        microsoftTeams.authentication.getAuthToken({
          resources: [resource],
          successCallback: (result) => {
            this._token = result;
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
    return this._authService.config;
  }
}

export default TeamsSsoAuthService;
