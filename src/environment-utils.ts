export const EnvironmentUtils = {
  // If we're in an office add-in
  IsOffice: function(): boolean {
    return !!(window.Office && window.Office.context && window.Office.context.document);
  },
  // If we're in g-suite
  IsGsuite: function(): boolean {
    return !!(window.google && window.google.script && window.google.script.run);
  },
  // Is localhost?
  IsLocalhost: function(): boolean {
    return (window.location.hostname.indexOf('localhost') >= 0);
  }
}