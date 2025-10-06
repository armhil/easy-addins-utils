import { EnvironmentUtils } from './environment-utils';

export const AddinUtils = {
  /**
   * Initializes the addin.
   * Required for Office addins, tbd for G-Suite.
   */
  Initialize: function (): Promise<any> {
    return new Promise((resolve, reject) => {
      if (EnvironmentUtils.IsGsuite()) {
        console.log('AddinUtils.Initialize -> We are in Google Docs');
        resolve(true);
      } else if (EnvironmentUtils.IsLocalhost()) {
        console.log('AddinUtils.Initialize -> We are in localhost');
        resolve(true);
      }
      // Microsoft Office
      // We need to keep this last, and ideally not rely on the IsOffice
      else {
        window.Office.onReady((info: any) => {
          resolve(info);
        });
      }
    });
  },
  /**
   * Inserts text
   * @param {string} text - Text to insert
   * @param {function} callback - Callback function
   */
  InsertText: function (
    text: string,
    insertionType: 'Text' | 'Html' = 'Text'
  ): Promise<any> {
    return new Promise((resolve, reject) => {
      if (EnvironmentUtils.IsLocalhost()) {
        // Do nothing
        return resolve(`No action in localhost, invoked with text: ${text}`);
      } else if (EnvironmentUtils.IsGsuite()) {
        window.google.script.run
          .withSuccessHandler((result: any) => resolve(result))
          .withFailureHandler((err: any) => reject(err))
          .insertPlainText(text);
      } else {
        window.Office.context.document.setSelectedDataAsync(
          text,
          { coercionType: window.Office.CoercionType[insertionType] },
          (successParams: Office.AsyncResult<void>) => {
            if (
              successParams.status === window.Office.AsyncResultStatus.Failed
            ) {
              return reject(successParams.error.message);
            }
            return resolve(successParams);
          }
        );
      }
    });
  },
  /**
   * Inserts image - we expect the underlying APIs for PPT and Word to be
   * relatively similar for this.
   * @param {string} image Base64 image.
   * @param {function} callback Callback function
   */
  InsertImage: function (image: string): Promise<any> {
    return new Promise((resolve, reject) => {
      if (EnvironmentUtils.IsOffice()) {
        window.Office.context.document.setSelectedDataAsync(
          image,
          { coercionType: window.Office.CoercionType.Image },
          (successParams: Office.AsyncResult<void>) => {
            if (
              successParams.status === window.Office.AsyncResultStatus.Failed
            ) {
              return reject(successParams.error.message);
            }
            return resolve(successParams);
          }
        );
      } else if (EnvironmentUtils.IsGsuite()) {
        window.google.script.run
          .withSuccessHandler((result: any) => resolve(result))
          .withFailureHandler((err: any) => reject(err))
          .insertImageFromBase64String(image);
      }
    });
  },

  /**
   * Gets the selected text
   * @param {function} callback - Callback function
   */
  GetText: function (): Promise<string> {
    if (EnvironmentUtils.IsOffice()) {
      return new Promise((resolve, reject) => {
        window.Office.context.document.getSelectedDataAsync(
          window.Office.CoercionType.Text,
          function (asyncResult: any) {
            if (asyncResult.status == window.Office.AsyncResultStatus.Failed) {
              return reject(asyncResult.error.message);
            } else {
              return resolve(asyncResult.value);
            }
          }
        );
      });
    } else if (EnvironmentUtils.IsGsuite()) {
      return new Promise((resolve, reject) => {
        window.google.script.run
          .withSuccessHandler((result: any) => resolve(result))
          .withFailureHandler((err: any) => reject(err))
          .getSelectedText();
      });
    }
  },

  /**
   * Saves a setting
   * @param {string} key - Setting key
   * @param {string} value - Setting value
   */
  SaveSetting: function (key: string, value: string) {
    if (EnvironmentUtils.IsOffice()) {
      window.Office.context.document.settings.set(key, value);
      window.Office.context.document.settings.saveAsync();
    } else if (EnvironmentUtils.IsLocalhost()) {
      if (!this.localDictionary) this.localDictionary = {};
      this.localDictionary[key] = value;
    }
  },

  /**
   * Gets a setting
   * @param {string} key - Setting key
   */
  GetSetting: function (key: string) {
    if (EnvironmentUtils.IsOffice()) {
      return window.Office.context.document.settings.get(key);
    } else if (EnvironmentUtils.IsLocalhost()) {
      return this.localDictionary ? this.localDictionary[key] : undefined;
    }
  },
};
