Import-module .\_functions.ps1 -Force

write-host "hello world"

Get-ALMAwareExchangeOnlineConnection
#Get-DistributionGroup -Identity "maetd-cs-ae_members" 

$group = New-DistributionGroup -Name "pruebaCristianDev" -Type Security

$group | select *

ls 

gci env:* | sort-object name