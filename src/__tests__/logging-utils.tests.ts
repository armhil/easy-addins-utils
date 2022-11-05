import { LoggingUtils } from '../logging-utils';

describe('LoggingUtils tests', () => {
  afterEach(() => {
    delete window.appInsights;
    delete window.location;
  });

  it('should not call external library for the localhost setup', () => {
    const mockTrackFn = jest.fn();
    const loggingKey = 'testKey';
    window.appInsights = { trackEvent: mockTrackFn };
    window.location = { hostname: "https://localhost:1234" } as Location;

    LoggingUtils.Trace(loggingKey);
    expect(mockTrackFn).toHaveBeenCalledTimes(0);
  });

  it('should call external library for non-localhost setup', () => {
    const mockTrackFn = jest.fn();
    const loggingKey = 'testKey';
    window.appInsights = { trackEvent: mockTrackFn };
    window.location = { hostname: "https://testing:1234" } as Location;

    LoggingUtils.Trace(loggingKey);
    expect(mockTrackFn).toHaveBeenCalledTimes(1);
    expect(mockTrackFn).toHaveBeenCalledWith({name: loggingKey});
  });
});