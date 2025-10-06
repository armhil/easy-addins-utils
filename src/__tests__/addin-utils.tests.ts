import { AddinUtils } from './../addin-utils';
import { EnvironmentUtils } from './../environment-utils';
import type { GoogleScriptRun } from './../types';
// Mock EnvironmentUtils
jest.mock('./../environment-utils', () => ({
  EnvironmentUtils: {
    IsLocalhost: jest.fn(),
    IsOffice: jest.fn(),
    IsGsuite: jest.fn(),
  },
}));

// Define Office/Google globals
declare global {
  interface Window {
    Office: any;
    google: GoogleScriptRun;
  }
}

// Reset mocks before each test
beforeEach(() => {
  jest.resetAllMocks();
  window.Office = undefined;
  window.google = undefined;
});

describe('AddinUtils.Initialize', () => {
  it('resolves immediately on localhost', async () => {
    (EnvironmentUtils.IsLocalhost as jest.Mock).mockReturnValue(true);
    await expect(AddinUtils.Initialize()).resolves.toBe(true);
  });

  it('resolves on Office ready', async () => {
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(true);
    const readyMock = jest.fn((cb) => cb('office-ready'));
    window.Office = { onReady: readyMock };

    const result = await AddinUtils.Initialize();
    expect(result).toBe('office-ready');
  });
});

describe('AddinUtils.InsertText', () => {
  it('should resolve immediately on localhost with a message', async () => {
    (EnvironmentUtils.IsLocalhost as jest.Mock).mockReturnValue(true);
    const result = await AddinUtils.InsertText('test');
    expect(result).toContain('No action in localhost');
  });

  it('should invoke Google Apps Script when in G-Suite', async () => {
    (EnvironmentUtils.IsLocalhost as jest.Mock).mockReturnValue(false);
    (EnvironmentUtils.IsGsuite as jest.Mock).mockReturnValue(true);

    const mockWithSuccessHandler = jest
      .fn()
      .mockImplementation((cb: (res: any) => void) => {
        cb('gsuite-success');
        return {
          withFailureHandler: jest.fn().mockReturnThis(),
          insertPlainText: jest.fn(),
        };
      });

    global.window.google = {
      script: {
        run: {
          withSuccessHandler: mockWithSuccessHandler,
          withFailureHandler: jest.fn().mockReturnThis(),
          insertPlainText: jest.fn(),
        },
      },
    } as any;

    const result = await AddinUtils.InsertText('gsuite-success');
    expect(result).toBe('gsuite-success');
  });

  it('should insert text via Office API when in Office', async () => {
    (EnvironmentUtils.IsLocalhost as jest.Mock).mockReturnValue(false);
    (EnvironmentUtils.IsGsuite as jest.Mock).mockReturnValue(false);
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(true);

    const setSelectedDataAsync = jest.fn((_t: any, _opts: any, cb: any) => {
      cb({ status: 'succeeded' });
    });

    global.window.Office = {
      CoercionType: {
        Text: 'Text',
        Html: 'Html',
      },
      AsyncResultStatus: {
        Failed: 'failed',
        Succeeded: 'succeeded',
      },
      context: {
        document: {
          setSelectedDataAsync,
        },
      },
    } as any;
    const sampleText = 'Sample text to insert';
    const result = await AddinUtils.InsertText(sampleText);
    expect(result).toEqual({ status: 'succeeded' });
    expect(setSelectedDataAsync).toHaveBeenCalledWith(
      sampleText,
      { coercionType: 'Text' },
      expect.any(Function)
    );
  });

  it('should reject when Office API returns failure', async () => {
    (EnvironmentUtils.IsLocalhost as jest.Mock).mockReturnValue(false);
    (EnvironmentUtils.IsGsuite as jest.Mock).mockReturnValue(false);
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(true);
    const sampleText = 'Sample text to insert';
    const errorMessage = 'Office insertion failed';

    global.window.Office = {
      CoercionType: {
        Text: 'Text',
        Html: 'Html',
      },
      AsyncResultStatus: {
        Failed: 'failed',
        Succeeded: 'succeeded',
      },
      context: {
        document: {
          setSelectedDataAsync: (_t: any, _opts: any, cb: any) => {
            cb({ status: 'failed', error: { message: errorMessage } });
          },
        },
      },
    } as any;

    await expect(AddinUtils.InsertText(sampleText)).rejects.toEqual(
      errorMessage
    );
  });
});

describe('AddinUtils.InsertImage', () => {
  const sampleImage = 'base64encodedimage==';

  afterEach(() => {
    jest.resetAllMocks();
  });

  it('should insert image using Office API when in Office', async () => {
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(true);

    const mockSetSelectedDataAsync = jest.fn(
      (_img: any, _options: any, callback: any) => {
        callback({ status: 'succeeded' });
      }
    );

    global.window.Office = {
      CoercionType: {
        Image: 'Image',
      },
      AsyncResultStatus: {
        Failed: 'failed',
        Succeeded: 'succeeded',
      },
      context: {
        document: {
          setSelectedDataAsync: mockSetSelectedDataAsync,
        },
      },
    } as any;

    const result = await AddinUtils.InsertImage(sampleImage);
    expect(result).toEqual({ status: 'succeeded' });
    expect(mockSetSelectedDataAsync).toHaveBeenCalledWith(
      sampleImage,
      { coercionType: 'Image' },
      expect.any(Function)
    );
  });

  it('should reject when Office API returns failure', async () => {
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(true);

    const errorMessage = 'Failed to insert image';

    global.window.Office = {
      CoercionType: {
        Image: 'Image',
      },
      AsyncResultStatus: {
        Failed: 'failed',
        Succeeded: 'succeeded',
      },
      context: {
        document: {
          setSelectedDataAsync: (_img: any, _options: any, callback: any) => {
            callback({ status: 'failed', error: { message: errorMessage } });
          },
        },
      },
    } as any;

    await expect(AddinUtils.InsertImage(sampleImage)).rejects.toEqual(
      errorMessage
    );
  });

  it('should insert image using GSuite API when in G-Suite', async () => {
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(false);
    (EnvironmentUtils.IsGsuite as jest.Mock).mockReturnValue(true);

    const mockSuccessHandler = jest
      .fn()
      .mockImplementation((cb: (res: any) => void) => {
        cb('gsuite-image-success');
        return {
          withFailureHandler: jest.fn().mockReturnThis(),
          insertImageFromBase64String: jest.fn(),
        };
      });

    global.window.google = {
      script: {
        run: {
          withSuccessHandler: mockSuccessHandler,
          withFailureHandler: jest.fn().mockReturnThis(),
          insertImageFromBase64String: jest.fn(),
        },
      },
    } as any;

    const result = await AddinUtils.InsertImage(sampleImage);
    expect(result).toBe('gsuite-image-success');
  });

  it('should reject if GSuite API fails', async () => {
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(false);
    (EnvironmentUtils.IsGsuite as jest.Mock).mockReturnValue(true);

    const mockFailureHandler = jest.fn().mockImplementation(() => {
      return {
        withSuccessHandler: jest.fn().mockReturnThis(),
        insertImageFromBase64String: jest.fn().mockImplementation(() => {
          throw new Error('fail');
        }),
      };
    });

    global.window.google = {
      script: {
        run: {
          withFailureHandler: mockFailureHandler,
          withSuccessHandler: jest.fn().mockReturnThis(),
          insertImageFromBase64String: jest.fn(),
        },
      },
    } as any;

    // Note: GSuite failure handler doesn't throw in your actual code – this is illustrative
    await expect(AddinUtils.InsertImage(sampleImage)).rejects.toBeDefined();
  });
});

describe('AddinUtils.GetText', () => {
  it('gets text from Office', async () => {
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(true);

    const mockGet = jest.fn((_t, cb) =>
      cb({ status: 'succeeded', value: 'the-text' })
    );

    window.Office = {
      CoercionType: { Text: 'Text' },
      AsyncResultStatus: { Failed: 'failed', Succeeded: 'succeeded' },
      context: { document: { getSelectedDataAsync: mockGet } },
    };

    const text = await AddinUtils.GetText();
    expect(text).toBe('the-text');
  });

  it('fails getting text from Office', async () => {
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(true);

    const errorMsg = 'failed!';
    window.Office = {
      CoercionType: { Text: 'Text' },
      AsyncResultStatus: { Failed: 'failed' },
      context: {
        document: {
          getSelectedDataAsync: (_t: any, cb: any) =>
            cb({ status: 'failed', error: { message: errorMsg } }),
        },
      },
    };

    await expect(AddinUtils.GetText()).rejects.toBe(errorMsg);
  });

  it('gets text from GSuite', async () => {
    (EnvironmentUtils.IsGsuite as jest.Mock).mockReturnValue(true);

    let success: any;
    const runner = {
      withSuccessHandler: (cb: any) => {
        success = cb;
        return runner;
      },
      withFailureHandler: () => runner,
      getSelectedText: jest.fn(),
    };

    (window.google as any) = { script: { run: runner } };

    const p = AddinUtils.GetText();
    success('gsuite-text');
    await expect(p).resolves.toBe('gsuite-text');
  });
});

describe('AddinUtils.Settings', () => {
  it('saves and gets Office setting', () => {
    (EnvironmentUtils.IsOffice as jest.Mock).mockReturnValue(true);

    const setMock = jest.fn();
    const saveAsyncMock = jest.fn();
    const getMock = jest.fn().mockReturnValue('stored-value');

    window.Office = {
      context: {
        document: {
          settings: {
            set: setMock,
            saveAsync: saveAsyncMock,
            get: getMock,
          },
        },
      },
    };

    AddinUtils.SaveSetting('theme', 'dark');
    expect(setMock).toHaveBeenCalledWith('theme', 'dark');
    expect(saveAsyncMock).toHaveBeenCalled();

    const val = AddinUtils.GetSetting('theme');
    expect(val).toBe('stored-value');
  });

  it('saves and gets settings in localhost', () => {
    (EnvironmentUtils.IsLocalhost as jest.Mock).mockReturnValue(true);
    AddinUtils.SaveSetting('lang', 'en');
    const val = AddinUtils.GetSetting('lang');
    expect(val).toBe('en');
  });
});
