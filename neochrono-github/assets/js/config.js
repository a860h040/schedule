export const APP_NAME = 'NeoChrono';
export const APP_VERSION = '2026.09.17-GITHUB-1';

export const DEFAULT_REPOSITORY = Object.freeze({
  owner: '',
  repo: 'neochrono-data',
  branch: 'main',
  dataDir: 'data'
});

export const DATA_FILES = Object.freeze({
  meta: 'meta.json',
  admins: 'admins.json',
  users: 'users.json',
  skills: 'skills.json',
  shifts: 'shifts.json',
  staffing: 'staffing.json',
  requests: 'requests.json',
  weeklyAvailability: 'weekly-availability.json',
  preceptorCalendar: 'preceptor-calendar.json',
  schedule: 'schedule.json',
  settings: 'settings.json',
  audit: 'audit.json',
  swaps: 'swaps.json'
});

export const SESSION_KEYS = Object.freeze({
  githubToken: 'neochronoGithubToken',
  appSession: 'neochronoAppSession'
});

export const LOCAL_KEYS = Object.freeze({
  repoConfig: 'neochronoRepoConfig'
});
