import { JsonRule } from "../../JsonRule";
import { Project } from "../../project-model/index";
import { Finding } from "../../report-model/index";

export class FN021007_PKG_scripts_start extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021007';
  }

  get title(): string {
    return 'package.json scripts.start';
  }

  get description(): string {
    return 'Update package.json scripts.start property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "start": "${this.script}"
  }
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './package.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageJson) {
      return;
    }

    if (!project.packageJson.scripts ||
      typeof project.packageJson.scripts !== 'object' ||
      !project.packageJson.scripts.start ||
      project.packageJson.scripts.start !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.start');
      this.addFindingWithPosition(findings, node);
    }
  }
}