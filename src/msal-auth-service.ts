import * as MSAL from '@azure/msal';

import { IAuthService, Resource } from '.';

/*
 * Use MSAL.js to authenticate AAD or MSA accounts against AAD v2
 */
class MsalAuthService implements IAuthService {
  app: MSAL.UserAgentApplication;

  constructor() {
    this.app = new MSAL.UserAgentApplication(this.config);
  }

  login() {
    const scopes = [
      `api://${this.config.auth.clientId}/access_as_user`,
      'https://graph.microsoft.com/User.Read',
    ];

    return ((window.navigator as any).standalone
      ? Promise.resolve(this.app.loginRedirect({ scopes }) as any)
      : this.app.loginPopup({ scopes })
    ).then(() => {
      return this.getUser();
    });
  }

  logout() {
    this.app.logout();
  }

  isCallback() {
    return this.app.isCallback(window.location.hash);
  }

  getUser() {
    return Promise.resolve((this.app as any).getUser());
  }

  getToken(resource: Resource) {
    const scopes = [resource];
    return this.app
      .acquireTokenSilent({ scopes })
      .then((res) => res.accessToken)
      .catch(() => {
        return this.app
          .acquireTokenPopup({ scopes })
          .then((res) => res.accessToken)
          .catch((error) => {
            console.error(error);
            throw error;
          });
      });
  }

  get config() {
    return {
      auth: {
        clientId: process.env.CLIENT_ID,
        redirectUri: `${window.location.origin}/${process.env.ADAL_REDIRECT_PATH}`,
      },
    };
  }
}

export default MsalAuthService;
