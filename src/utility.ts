import type * as TeamsJs from '@microsoft/teams-js';

export function insideIframe() {
  try {
    return window && window.self !== window.top;
  } catch (e) {
    return true;
  }
}

let teams = {};
if (typeof window !== 'undefined') {
  // eslint-disable-next-line global-require
  teams = require('@microsoft/teams-js');
}

export const MicrosoftTeams = teams as typeof TeamsJs;
