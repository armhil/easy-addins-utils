export {};

export type AppInsights = {
  trackEvent: (event: { name: string }) => void;
};

export type GoogleScriptRun = {
  script: {
    run: {
      withSuccessHandler: (
        callback: (result: object) => void
      ) => GoogleScriptRun['script']['run'];
      withFailureHandler: (
        errorCallback: (error: object) => void
      ) => GoogleScriptRun['script']['run'];
      insertImageFromBase64String: (image: string) => void;
      insertPlainText: (text: string) => void;
      getSelectedText: () => void;
    };
  };
};

declare global {
  interface Window {
    // potential extending object for G-Suite.
    google: GoogleScriptRun;
    // extending object for app insights.
    appInsights: AppInsights;
  }
}
