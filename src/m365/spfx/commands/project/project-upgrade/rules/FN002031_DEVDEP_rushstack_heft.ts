import { DependencyRule } from "./DependencyRule";

export class FN002031_DEVDEP_rushstack_heft extends DependencyRule {
  constructor(packageVersion: string) {
    super('@rushstack/heft', packageVersion, true);
  }

  get id(): string {
    return 'FN002031';
  }
}