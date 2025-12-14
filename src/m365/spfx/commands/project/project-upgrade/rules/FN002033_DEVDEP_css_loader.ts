import { DependencyRule } from "./DependencyRule";

export class FN002033_DEVDEP_css_loader extends DependencyRule {
  constructor(packageVersion: string) {
    super('css-loader', packageVersion, true);
  }

  get id(): string {
    return 'FN002033';
  }
}