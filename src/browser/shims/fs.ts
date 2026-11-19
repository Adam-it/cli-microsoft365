import { unavailable } from './unavailable.js';

export const existsSync = (): boolean => false;
export const readFileSync = unavailable('fs.readFileSync');
export const writeFileSync = unavailable('fs.writeFileSync');
export const appendFileSync = unavailable('fs.appendFileSync');
export const mkdirSync = unavailable('fs.mkdirSync');
export const readdirSync = unavailable('fs.readdirSync');
export const statSync = unavailable('fs.statSync');
export const lstatSync = unavailable('fs.lstatSync');
export const unlinkSync = unavailable('fs.unlinkSync');
export const rmSync = unavailable('fs.rmSync');
export const copyFileSync = unavailable('fs.copyFileSync');
export const renameSync = unavailable('fs.renameSync');
export const openSync = unavailable('fs.openSync');
export const closeSync = unavailable('fs.closeSync');
export const readSync = unavailable('fs.readSync');
export const createReadStream = unavailable('fs.createReadStream');
export const createWriteStream = unavailable('fs.createWriteStream');
export const readFile = unavailable('fs.readFile');
export const writeFile = unavailable('fs.writeFile');
export const mkdir = unavailable('fs.mkdir');
export const stat = unavailable('fs.stat');
export const access = unavailable('fs.access');
export const constants = {};
export const promises = {
  readFile: unavailable('fs.promises.readFile'),
  writeFile: unavailable('fs.promises.writeFile'),
  mkdir: unavailable('fs.promises.mkdir'),
  stat: unavailable('fs.promises.stat'),
  access: unavailable('fs.promises.access'),
  unlink: unavailable('fs.promises.unlink'),
  readdir: unavailable('fs.promises.readdir')
};

export default {
  existsSync,
  readFileSync,
  writeFileSync,
  appendFileSync,
  mkdirSync,
  readdirSync,
  statSync,
  lstatSync,
  unlinkSync,
  rmSync,
  copyFileSync,
  renameSync,
  openSync,
  closeSync,
  readSync,
  createReadStream,
  createWriteStream,
  readFile,
  writeFile,
  mkdir,
  stat,
  access,
  constants,
  promises
};
