export const EOL = '\n';
export const homedir = (): string => '/';
export const tmpdir = (): string => '/tmp';
export const platform = (): string => 'browser';
export const hostname = (): string => 'browser';
export const release = (): string => '';
export const type = (): string => 'Browser';
export const arch = (): string => 'javascript';
export const cpus = (): any[] => [];
export const userInfo = (): { username: string; homedir: string } => ({ username: '', homedir: '/' });

export default { EOL, homedir, tmpdir, platform, hostname, release, type, arch, cpus, userInfo };
