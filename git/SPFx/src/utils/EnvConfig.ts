/* eslint-disable */
import * as config from './../config.json';

interface IEnvConfig {
    BackendHost: string;
    BakendAppClientId: string;
    SPTenantHostUrl: string;
    ApprodotSiteUrl: string;
    AppDiagnosticsprefix: string;
    AppInsightsConnectionString: string;
    FrontendCacheMinutes: string;
}

const rootKey = "default" as keyof typeof config;
const configObject = config[rootKey]

const hostSpecificConfigKey = window.location.host as keyof typeof configObject;
const defaultConfigKey = "defaults" as keyof typeof configObject;

let envConfig = { ...(configObject[defaultConfigKey] as any) };
if (window.location.host && configObject[hostSpecificConfigKey]) {
    envConfig = { ...envConfig, ...(configObject[hostSpecificConfigKey] as any) };
}
const EnvConfig: IEnvConfig = envConfig as IEnvConfig;

export { EnvConfig };