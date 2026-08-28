param env string
param app string
param counter string

output vaultName string = 'kv${env}${app}${counter}'
output appPlanName string = 'asp${env}${app}${counter}'
output fnAppName string = 'func${env}${app}${counter}'
output fnAppAuxName string = 'func${env}${app}${counter}aux'
output storageName string = toLower('st${env}${app}${counter}')
output appInsightsName string = 'appi${env}${app}${counter}'
output logAnalyticsName string = 'log${env}${app}${counter}'
output identityName string = 'id${env}${app}${counter}'
