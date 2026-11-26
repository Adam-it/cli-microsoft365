import { DependencyRule } from './DependencyRule';

export class FN001036_DEP_microsoft_sp_copilot_component extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super('@microsoft/sp-copilot-component', options.packageVersion, true);
  }

  get id(): string {
    return 'FN001036';
  }
}