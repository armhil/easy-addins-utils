import { EnvironmentUtils } from './environment-utils';

export const AddinUtils = {
  /**
 * Initializes the addin.
   * Required for Office addins, tbd for G-Suite.
 * @param {function} successCallback - Success callback
 */
  Initialize: function(successCallback: Function) {
    // We're in localhost 
    if (EnvironmentUtils.IsLocalhost()) {
      console.log('AddinUtils.Initialize -> We are in localhost');
    }
    // Microsoft Office
    else if (EnvironmentUtils.IsOffice()) {
      window.Office.initialize = function() {
        successCallback();
      };
    }
  },

  /**
   * Inserts image - we expect the underlying APIs for PPT and Word to be
   * relatively similar for this.
   * @param {string} image Base64 image.
   * @param {function} callback Callback function
   */
  InsertImage: function(image: string, callback: Function) {
    if (EnvironmentUtils.IsOffice()) {
      window.Office.context.document.setSelectedDataAsync(
        image,
        {coercionType: window.Office.CoercionType.Image},
        function (asyncResult: any) {
            callback(asyncResult);
        }
      );
    }
  },

  /**
   * Gets the selected text
   * @param {function} callback - Callback function
   */
  GetText: function(callback: Function) {
    if (EnvironmentUtils.IsOffice()) {
      window.Office.context.document.getSelectedDataAsync(
          window.Office.CoercionType.Text,
          function (asyncResult: any) {
              callback(asyncResult);
          }
      );
    }
  },

  /**
   * Saves a setting
   * @param {string} key - Setting key
   * @param {string} value - Setting value
  */
  SaveSetting: function(key: string, value: string) {
    if (EnvironmentUtils.IsOffice()) {
      window.Office.context.document.settings.set(key, value);
      window.Office.context.document.settings.saveAsync();
    }
    else if (EnvironmentUtils.IsLocalhost()){
            if (!this.localDictionary) 
                this.localDictionary={};
      this.localDictionary[key] = value;
    }
  },

  /**
   * Gets a setting
   * @param {string} key - Setting key
   */
  GetSetting: function(key: string) {
    if (EnvironmentUtils.IsOffice()) {
      return window.Office.context.document.settings.get(key);
    }
    else if (EnvironmentUtils.IsLocalhost()) {
      return this.localDictionary ? this.localDictionary[key] : undefined;
    }
  }
}