// replaces src/telemetry.ts in the browser build. Nothing is tracked there.
export const telemetry = {
  trackEvent: async (): Promise<void> => {
    // no telemetry in the browser build
  }
};
