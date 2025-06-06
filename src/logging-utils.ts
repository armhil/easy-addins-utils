export const LoggingUtils = {
  Trace: function (key: string) {
    // Don't log if we're in localhost.
    if (window.location.hostname.indexOf('localhost') >= 0) return;
    // Check if app insights is available, if yes, track event.
    if (window.appInsights && window.appInsights.trackEvent) {
      window.appInsights.trackEvent({ name: key });
    }
  },
};
