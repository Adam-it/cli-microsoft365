import appInsights from "./appInsights";
import { Cli } from "./cli/Cli";
import { settingsNames } from "./settingsNames";

class Telemetry {
  public trackEvent(commandName: string, properties: any): void {
    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.disableTelemetry, true)) {
      return;
    }

    appInsights.trackEvent({
      name: commandName,
      properties
    });
    appInsights.flush();
  }

  public trackException(exception: any): void {
    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.disableTelemetry, true)) {
      return;
    }

    appInsights.trackException({
      exception
    });
    appInsights.flush();
  }
}

export const telemetry = new Telemetry();