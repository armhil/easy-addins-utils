import { EnvironmentUtils } from '../environment-utils';

describe('EnvironmentUtils tests', () => {
    // cleanup
    afterEach(() => {
        delete window.location;
        delete window.Office;
        delete window.google;
    });

    // localhost tests
    it.each([
        ["localhost:1234", true],
        ["easyaddins.net", false]
    ])
    ('detects localhost correctly', (url: string, expectedResult: boolean) => {
        window.location = { hostname: url } as Location;
        expect(EnvironmentUtils.IsLocalhost()).toBe(expectedResult);
    });

    // Office tests
    it.each([
        [undefined, false],
        [{context: undefined}, false],
        [{context: {document: undefined}}, false],
        [{context: {document: {context: "TestValue"}}}, true]
    ])
    ('detects Office correctly', (officeObject, expectedResult) => {
        window.Office = officeObject;
        expect(EnvironmentUtils.IsOffice()).toBe(expectedResult);
    });

    // G-suite tests
    it.each([
        [undefined, false],
        [{script: undefined}, false],
        [{script: {run: "TestValue"}}, true]
    ])
    ('detects G-Suite correctly', (gObject, expectedResult) => {
        window.google = gObject;
        expect(EnvironmentUtils.IsGsuite()).toBe(expectedResult);
    });
});