import { unavailable } from './unavailable.js';

export const spawn = unavailable('child_process.spawn');
export const spawnSync = unavailable('child_process.spawnSync');
export const exec = unavailable('child_process.exec');
export const execSync = unavailable('child_process.execSync');
export const execFile = unavailable('child_process.execFile');
export const execFileSync = unavailable('child_process.execFileSync');
export const fork = unavailable('child_process.fork');

export default { spawn, spawnSync, exec, execSync, execFile, execFileSync, fork };
