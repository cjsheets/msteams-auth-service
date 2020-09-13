import AuthenticationContext from 'adal-angular';
import AuthService from './auth-service';
import { MicrosoftTeams } from './utility';

export const MsTeamsAuthService = AuthService;

export const AdalAuthenticationContext = AuthenticationContext;

export const TeamsJs = MicrosoftTeams;
export type TeamsJS = typeof MicrosoftTeams;
