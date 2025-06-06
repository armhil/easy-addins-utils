import { AddinUtils } from './../addin-utils';
import '../types';

describe('AddinUtils tests', () => {
  beforeEach(() => {
    delete window.location;
    delete window.Office;
    delete window.google;
  });

  it('Initialize:Office should call the callback fn if there is one', () => {
    // This test is slighly tricky, because in the actual environment,
    // we rely on Office.js to invoke the callback fn.
    const successMockFn = jest.fn();

    (window as any).Office = {
      onReady: () => successMockFn(),
      context: { document: 'SomeValue' },
    };
    (window as any).location = { hostname: 'https://testing:1234' } as Location;

    AddinUtils.Initialize(successMockFn);
    // now our callback should have been called
    expect(successMockFn).toHaveBeenCalledTimes(1);
  });

  it('Initialize:Office should work without callback', () => {
    (window as any).Office = {
      onReady: () => {},
      context: { document: 'SomeValue' },
    };
    (window as any).location = { hostname: 'https://testing:1234' } as Location;

    AddinUtils.Initialize();
  });

  it('GetSetting:localhost returns undefined for no settings', () => {
    (window as any).location = {
      hostname: 'https://localhost:1234',
    } as Location;
    expect(AddinUtils.GetSetting('randomSetting')).toBe(undefined);
  });

  it('GetSetting:localhost returns expected value for setting', () => {
    (window as any).location = {
      hostname: 'https://localhost:1234',
    } as Location;
    const settingName = 'randomSetting';
    const settingValue = 'settingValue';
    AddinUtils.SaveSetting(settingName, settingValue);

    expect(AddinUtils.GetSetting(settingName)).toBe(settingValue);
  });

  it('GetSetting:Office calls settings.get', () => {
    const mockFn = jest.fn();
    const settingName = 'randomSetting';
    (window as any).Office = {
      context: { document: { settings: { get: mockFn } } },
    };

    AddinUtils.GetSetting(settingName);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(mockFn).toHaveBeenCalledWith(settingName);
  });

  it('SaveSetting:Office calls settings.saveAsync', () => {
    const setMockFn = jest.fn();
    const saveAsyncMockFn = jest.fn();
    const settingName = 'randomSetting';
    const settingValue = 'settingValue';
    (window as any).Office = {
      context: {
        document: { settings: { set: setMockFn, saveAsync: saveAsyncMockFn } },
      },
    };

    AddinUtils.SaveSetting(settingName, settingValue);
    expect(setMockFn).toHaveBeenCalledTimes(1);
    expect(setMockFn).toHaveBeenCalledWith(settingName, settingValue);
    expect(saveAsyncMockFn).toHaveBeenCalledTimes(1);
  });

  it('GetText:Office calls getSelectedDataAsync with correct params', () => {
    const mockFn = jest.fn();
    const successFn = jest.fn();
    (window as any).Office = {
      CoercionType: { Text: 'Text' },
      context: { document: { getSelectedDataAsync: mockFn } },
    };

    AddinUtils.GetText(successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(successFn).toHaveBeenCalledTimes(0);
  });

  it('GetText:google calls getSelectedText with correct params', () => {
    const mockFn = jest.fn();
    const successFn = jest.fn();
    (window as any).google = {
      script: {
        run: {
          getSelectedText: mockFn,
          withSuccessHandler: () => window.google.script.run,
          withFailureHandler: () => window.google.script.run,
        },
      },
    };

    AddinUtils.GetText(successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(successFn).toHaveBeenCalledTimes(0);
  });

  it('InsertText:Office calls setSelectedDataAsync with correct params', () => {
    const mockFn = jest.fn();
    const successFn = jest.fn();
    const textContent = 'sample text';
    (window as any).location = { hostname: 'https://testing:1234' } as Location;
    (window as any).Office = {
      CoercionType: { Text: 'Text' },
      context: { document: { setSelectedDataAsync: mockFn } },
    };

    AddinUtils.InsertText(textContent, 'Text', successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(mockFn.mock.lastCall[0]).toBe(textContent);
    //expect(mockFn.mock.lastCall[0]).toBe(textContent);
    expect(successFn).toHaveBeenCalledTimes(0);
    // simulate Office callback invocation
    mockFn.mock.lastCall[2]();
    expect(successFn).toHaveBeenCalledTimes(1);
  });

  it('InsertText:google calls insertText with correct params', () => {
    const mockFn = jest.fn();
    const successFn = jest.fn();
    const textContent = 'sample text';
    (window as any).location = { hostname: 'https://testing:1234' } as Location;
    (window as any).google = {
      script: {
        run: {
          insertPlainText: mockFn,
        },
      },
    };

    AddinUtils.InsertText(textContent, 'Text', successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(mockFn.mock.lastCall[0]).toBe(textContent);
  });

  it('InsertImage:Office calls setSelectedDataAsync with correct params', () => {
    const mockFn = jest.fn();
    const successFn = jest.fn();
    const imageText = 'imageText';
    (window as any).Office = {
      CoercionType: { Image: 'Image' },
      context: { document: { setSelectedDataAsync: mockFn } },
    };

    AddinUtils.InsertImage(imageText, successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(mockFn.mock.lastCall[0]).toBe(imageText);
    expect(mockFn.mock.lastCall[1]).toEqual({
      coercionType: Office.CoercionType.Image,
    });
    expect(successFn).toHaveBeenCalledTimes(0);
    // manually call the callback of InsertImage - which should call our callback
    mockFn.mock.lastCall[2]();
    expect(successFn).toHaveBeenCalledTimes(1);
  });

  it('InsertImage:google calls insertImageFromBase64String with correct params', () => {
    const mockFn = jest.fn();
    // there is no real way to test the successFn callback in this case
    // since google handles these on their end.
    // https://developers.google.com/apps-script/guides/html/reference/run
    const successFn = jest.fn();
    const imageText = 'imageText';
    window.google = {
      script: {
        run: {
          withSuccessHandler: () => window.google.script.run,
          withFailureHandler: () => window.google.script.run,
          insertImageFromBase64String: mockFn,
          insertPlainText: () => undefined,
          getSelectedText: () => undefined,
        },
      },
    };

    AddinUtils.InsertImage(imageText, successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(mockFn.mock.lastCall[0]).toBe(imageText);
  });
});
