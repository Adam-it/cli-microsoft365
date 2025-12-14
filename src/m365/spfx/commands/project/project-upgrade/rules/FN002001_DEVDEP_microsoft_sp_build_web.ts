import { DependencyRule } from "./DependencyRule";

export class FN002001_DEVDEP_microsoft_sp_build_web extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@microsoft/sp-build-web', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002001';
  }
}