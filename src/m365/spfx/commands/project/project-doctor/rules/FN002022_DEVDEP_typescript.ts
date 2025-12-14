import { DependencyRule } from './DependencyRule';

export class FN002022_DEVDEP_typescript extends DependencyRule {
  constructor(version: string) {
    super('typescript', version, true);
  }

  get id(): string {
    return 'FN002022';
  }
}