/* eslint-disable */
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ServiceKey } from '@microsoft/sp-core-library';
import { EnvConfig } from './EnvConfig';
import { PageContext } from '@microsoft/sp-page-context';
import { getBodyIdFormUrl } from '../service/RoleService';

type Logprops = { [k: string]: string };
const CONSOLE_testFIX: string = EnvConfig.AppDiagnosticsprefix;
const SERVICE_KEY: string = `${CONSOLE_testFIX}:IDiagnosticsService`;
declare global {
    interface Window { AppDiagnistics: DiagnosticsService; O365ShellContext: any; }
}

export enum DiagnosticsLevel {
    Error = 'Error',
    Warning = 'Warning',
    Info = 'Info',
    Debug = 'Debug',
    Trace = 'Trace'
}

export interface IDiagnosticsService {
    error(message: string, exception?: Error, props?: PageContext | any): void;
    log(message: string, category?: DiagnosticsLevel, exception?: Error, props?: PageContext | any): void;
    info(message: string, exception?: Error, props?: PageContext | any): void
    warning(message: string, exception?: Error, props?: PageContext | any): void
    debug(message: string, exception?: Error, props?: PageContext | any): void 
    trackEvent(eventName: string, props?: PageContext | any): void;
}

class DiagnosticsService implements IDiagnosticsService {
    public static readonly serviceKey: ServiceKey<IDiagnosticsService> = ServiceKey.create<IDiagnosticsService>(
        SERVICE_KEY,
        DiagnosticsService,
    );
    private static appInsights: ApplicationInsights;
    private static instance: DiagnosticsService;

    static getInstance(): IDiagnosticsService {
        if (!DiagnosticsService.instance) {
            DiagnosticsService.instance = new DiagnosticsService();

            // only save requests to backend domain
            const domain = EnvConfig.BackendHost.replace('https://', '');
            const requestExcludePattern = [new RegExp('^((?!' + domain + ').)*$')];
            const role = 'spfx:' + window.location.host;

            DiagnosticsService.appInsights = new ApplicationInsights({
                config: {
                    connectionString: EnvConfig.AppInsightsConnectionString,
                    enableCorsCorrelation: true,
                    disableExceptionTracking: true,
                    excludeRequestFromAutoTrackingPatterns: requestExcludePattern,
                    correlationHeaderExcludePatterns: requestExcludePattern,
                    enableAutoRouteTracking: true, // TODO: check if works with modern navigation.
                },
            });
            DiagnosticsService.appInsights.addTelemetryInitializer(envelope => {
                envelope.data = envelope.data || {};
                // add tags
                envelope.tags = envelope.tags || [];
                envelope.tags['ai.cloud.role'] = role;
                envelope.tags['ai.cloud.roleInstance'] = role;

                // try add custom properties from shell
                envelope.data['properties'] = envelope.data['properties'] || {};
                envelope.data['properties']['user_email'] = window?.O365ShellContext?.BootHeaderState?.userEmail;
                envelope.data['properties']['user_displayName'] = window?.O365ShellContext?.BootHeaderState?.meDisplayName;
            });
            DiagnosticsService.appInsights.loadAppInsights();
            DiagnosticsService.appInsights.trackPageView();
        }
        return DiagnosticsService.instance;
    }

    error(message: string, exception?: Error, props?: PageContext | any): void {
        this.log(message, DiagnosticsLevel.Error, exception, props);
    }

    debug(message: string, exception?: Error, props?: PageContext | any): void {
        this.log(message, DiagnosticsLevel.Debug, exception, props);
    }

    info(message: string, exception?: Error, props?: PageContext | any): void {
        this.log(message, DiagnosticsLevel.Info, exception, props);
    }

    warning(message: string, exception?: Error, props?: PageContext | any): void {
        this.log(message, DiagnosticsLevel.Warning, exception, props);
    }

    log(message: string, category?: DiagnosticsLevel, exception?: Error, props?: PageContext | any): void {
        const logprops: Logprops = this.handleprops(props);

        const traceText = `${CONSOLE_testFIX}${category ? ' (' + category + ')' : ''}: ${message}${exception?.message ? ' ::> ' + exception.message : ''}`;
        console.log(traceText);

        switch (category) {
            case DiagnosticsLevel.Error:
                DiagnosticsService.appInsights?.trackException({
                    id: message,
                    exception: {
                        name: message,
                        message: traceText,
                        stack: exception?.stack,
                    },
                    properties: { ...logprops },
                    severityLevel: SeverityLevel.Error,
                });
                break;
            case DiagnosticsLevel.Warning:
                DiagnosticsService.appInsights?.trackTrace({
                    message: message,
                    properties: { ...logprops },
                    severityLevel: SeverityLevel.Warning,
                });
                break;
            case DiagnosticsLevel.Info:
                DiagnosticsService.appInsights?.trackTrace({
                    message: message,
                    properties: { ...logprops },
                    severityLevel: SeverityLevel.Information,
                });
                break;
            default:
                DiagnosticsService.appInsights?.trackTrace({
                    message: message,
                    properties: { ...logprops },
                    severityLevel: SeverityLevel.Verbose,
                });
                break;
        }
    }

    trackEvent(eventName: string, props?: PageContext | any): void {
        DiagnosticsService.appInsights?.trackEvent({ name: eventName }, { ...this.handleprops(props) });
    }

    handleprops(props?: PageContext | any): Logprops {
        const logprops: Logprops = {};

        try {
            if (props) {
                if (props instanceof PageContext) {
                    logprops.CS_UPN = props.user?.loginName;
                    logprops.CS_BodyId = getBodyIdFormUrl(props.site?.absoluteUrl);
                    logprops.siteAbsUrl = props.site?.absoluteUrl;
                    logprops.pageAbsUrl = props.legacyPageContext?.serverRequestPath;
                } else {
                    const symbolKeys = Object.getOwnpropertyNames(props);
                    if (symbolKeys.length > 0) {
                        symbolKeys.forEach(key => {
                            logprops[key] = JSON.stringify(props[key]);
                        });
                    }
                }
            }
            logprops.CategoryName = "CNT.SPFx";
        } catch (error) {
            console.log('WARNING: Error while handling props in log (props will be ignored)', error);
        }

        return logprops;
    }
}

const Logger = DiagnosticsService.getInstance() as IDiagnosticsService;

export { Logger };
