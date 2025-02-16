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
    window.Office = { initialize : undefined, context : { document: "SomeValue"}};
    window.location = { hostname: 'https://testing:1234' } as Location;

    AddinUtils.Initialize(successMockFn);
    // which means the function assignment should have been done
    expect(window.Office.initialize).not.toBe(undefined);
    // explicitly invoke the function, mocking Office.js behaviour
    window.Office.initialize();
    // now our callback should have been called
    expect(successMockFn).toHaveBeenCalledTimes(1);
  });

  it('Initialize:Office should work without callback', () => {
    window.Office = { initialize : undefined, context : { document: "SomeValue"}};
    window.location = { hostname: 'https://testing:1234' } as Location;

    AddinUtils.Initialize();
  });

  it('GetSetting:localhost returns undefined for no settings', () => {
    window.location = { hostname: 'https://localhost:1234' } as Location;
    expect(AddinUtils.GetSetting('randomSetting')).toBe(undefined);
  });

  it('GetSetting:localhost returns expected value for setting', () => {
    window.location = { hostname: 'https://localhost:1234' } as Location;
    const settingName = 'randomSetting';
    const settingValue = 'settingValue';
    AddinUtils.SaveSetting(settingName, settingValue);

    expect(AddinUtils.GetSetting(settingName)).toBe(settingValue);
  });

  it('GetSetting:Office calls settings.get', () => {
    const mockFn = jest.fn();
    const settingName = 'randomSetting';
    window.Office =  { context : { document: { settings: { get: mockFn }}}};

    AddinUtils.GetSetting(settingName);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(mockFn).toHaveBeenCalledWith(settingName);
  });

  it('SaveSetting:Office calls settings.saveAsync', () => {
    const setMockFn = jest.fn();
    const saveAsyncMockFn = jest.fn();
    const settingName = 'randomSetting';
    const settingValue = 'settingValue';
    window.Office =  { context : { document: { settings: { set: setMockFn, saveAsync: saveAsyncMockFn }}}};

    AddinUtils.SaveSetting(settingName, settingValue);
    expect(setMockFn).toHaveBeenCalledTimes(1);
    expect(setMockFn).toHaveBeenCalledWith(settingName, settingValue);
    expect(saveAsyncMockFn).toHaveBeenCalledTimes(1);
  });

  it('GetText:Office calls getSelectedDataAsync with correct params', () => {
    const mockFn = jest.fn();
    const successFn = jest.fn();
    window.Office =  { CoercionType: { Text: 'Text' }, context : { document: { getSelectedDataAsync: mockFn }}};

    AddinUtils.GetText(successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(successFn).toHaveBeenCalledTimes(0);
  });

  it('GetText:google calls getSelectedText with correct params', () => {
    const mockFn = jest.fn();
    const successFn = jest.fn();
    window.google =  {
      script: {
        run: {
          withSuccessHandler: () => window.google.script.run,
          withFailureHandler: () => window.google.script.run,
          getSelectedText: mockFn
        }
      }
    };

    AddinUtils.GetText(successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(successFn).toHaveBeenCalledTimes(0);
  });

  it('InsertImage:Office calls setSelectedDataAsync with correct params', () => {
    const mockFn = jest.fn();
    const successFn = jest.fn();
    const imageText = 'imageText';
    window.Office =  { CoercionType: { Image: 'Image' }, context : { document: { setSelectedDataAsync: mockFn }}};

    AddinUtils.InsertImage(imageText, successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(mockFn.mock.lastCall[0]).toBe(imageText);
    expect(mockFn.mock.lastCall[1]).toEqual({coercionType: window.Office.CoercionType.Image});
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
    window.google =  {
      script: {
        run: {
          withSuccessHandler: () => window.google.script.run,
          withFailureHandler: () => window.google.script.run,
          insertImageFromBase64String: mockFn
        }
      }
    };

    AddinUtils.InsertImage(imageText, successFn);
    expect(mockFn).toHaveBeenCalledTimes(1);
    expect(mockFn.mock.lastCall[0]).toBe(imageText);
  });
});