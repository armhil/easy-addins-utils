import { AddinUtils } from './../addin-utils';

describe('AddinUtils tests', () => {
  beforeEach(() => {
    delete window.location;
    delete window.Office;
  });

  it('Initialize:Office should call the callback fn', () => {
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
    // first parameter
    expect(mockFn.mock.lastCall[0]).toBe('Text');
    // manually call the callback of GetText - which should call our call-back
    mockFn.mock.lastCall[1]();
    expect(successFn).toHaveBeenCalledTimes(1);
  });
});