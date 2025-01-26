export {};

declare global {
  interface Window {
    // potential extending object for Office.js
    Office: any;
    // potential extending object for G-Suite.
    google: any;
    // extending object for app insights.
    appInsights: any;
  }
}