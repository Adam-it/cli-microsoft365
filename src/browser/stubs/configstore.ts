// replaces the configstore package, which stores settings in a file. The
// browser host replaces cli.getConfig, so this only has to exist.
export default class Configstore {
  private settings = new Map<string, unknown>();

  public get all(): { [key: string]: unknown } {
    return Object.fromEntries(this.settings);
  }

  public get(key: string): unknown {
    return this.settings.get(key);
  }

  public set(key: string, value: unknown): void {
    this.settings.set(key, value);
  }

  public has(key: string): boolean {
    return this.settings.has(key);
  }

  public delete(key: string): void {
    this.settings.delete(key);
  }

  public clear(): void {
    this.settings.clear();
  }
}
