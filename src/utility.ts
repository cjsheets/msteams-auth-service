import * as TeamsJs from '@microsoft/teams-js';

export function insideIframe() {
  try {
    return window && window.self !== window.top;
  } catch (e) {
    return true;
  }
}

let teams = {} as typeof TeamsJs;
if (typeof window !== 'undefined') {
  teams = TeamsJs;
}

export const MicrosoftTeams = teams;
