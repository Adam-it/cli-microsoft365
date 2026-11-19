import { unavailable } from '../shims/unavailable.js';

// replaces the proper-lockfile package, used for the token cache file
export const lock = unavailable('proper-lockfile.lock');
export const unlock = unavailable('proper-lockfile.unlock');

export default { lock, unlock };
