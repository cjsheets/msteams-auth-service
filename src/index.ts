import AuthenticationContext from 'adal-angular';
import * as microsoftTeamsSdk from '@microsoft/teams-js';
import AuthService from './auth-service';

export const MsTeamsAuthService = AuthService;

export const AdalAuthenticationContext = AuthenticationContext;

export const TeamsJs = microsoftTeamsSdk;
