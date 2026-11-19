import { unavailable } from './unavailable.js';

export const randomUUID = (): string => globalThis.crypto.randomUUID();
export const randomBytes = (size: number): Uint8Array => globalThis.crypto.getRandomValues(new Uint8Array(size));
export const createHash = unavailable('crypto.createHash');
export const createHmac = unavailable('crypto.createHmac');

export default { randomUUID, randomBytes, createHash, createHmac };
