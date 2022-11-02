import { LoggingUtils } from './../logging-utils';

describe('LoggingUtils tests', () => {
  afterEach(() => {
    delete window.appInsights;
    delete window.location;
  });

  it.each([
    ['https://localhost:1234', 0],
    ['https://testing', 1]
  ])('should call appropriate number of times', (hostname: string, times: number) => {
    const mockTrackFn = jest.fn();
    window.appInsights = { trackEvent: mockTrackFn };
    window.location = { hostname: hostname } as Location;

    LoggingUtils.Trace('testKey');
    expect(window.appInsights.trackEvent).toHaveBeenCalledTimes(times);
  });
});