function Write-Log {
    param(
        [Parameter(Mandatory = $false)][string]$Method,
        [Parameter(Mandatory = $false)][string]$Message,
        [Parameter(Mandatory = $false)][string]$Detail,
        [Parameter(Mandatory = $false)][object]$MessageType
    )
	
    $log = [PSCustomObject]@{
        Method      = $Method
        Message     = $Message
        Detail      = $Detail
        MessageType = if ($null -ne $MessageType) { $MessageType }else { [CustomLogType]::Info }
        Date        = Get-Date    
        Guid        = $global:LogGuid
    }
    $log | Export-Csv -Path $global:LogPath -NoTypeInformation -Append -Force
}

function Write-LogWarning {
    param(
        [Parameter(Mandatory = $false)][string]$Method,
        [Parameter(Mandatory = $false)][string]$Message,
        [Parameter(Mandatory = $false)][string]$Detail
    )
	
    Write-host -Object $Method -BackgroundColor Black -ForegroundColor Yellow
    Write-host -Object $Message -BackgroundColor Black -ForegroundColor Yellow
    if (![string]::IsNullOrECNTy($Detail)) {
        Write-host -Object $Detail -BackgroundColor Black -ForegroundColor Yellow
    }
    Write-Log -Method $Method -Message $Message -Detail $Detail -MessageType [CustomLogType]::Warning
}

function Write-LogError {
    param(
        [Parameter(Mandatory = $false)][string]$Method,
        [Parameter(Mandatory = $false)][string]$Message,
        [Parameter(Mandatory = $false)][string]$Detail
    )
	
    Write-host -Object $Method -BackgroundColor Black -ForegroundColor Red
    Write-host -Object $Message -BackgroundColor Black -ForegroundColor Red
    if (![string]::IsNullOrECNTy($Detail)) {
        Write-host -Object $Detail -BackgroundColor Black -ForegroundColor Red
    }
    Write-Log -Method $Method -Message $Message -Detail $Detail -MessageType [CustomLogType]::Error
}

function Write-LogInfo {
    param(
        [Parameter(Mandatory = $false)][string]$Method,
        [Parameter(Mandatory = $false)][string]$Message,
        [Parameter(Mandatory = $false)][string]$Detail
    )
	
    Write-Log -Method $Method -Message $Message -Detail $Detail -MessageType [CustomLogType]::Info
}

function Write-LogException {
    param(
        [Parameter(Mandatory = $false)][string]$Method,
        [Parameter(Mandatory = $true)][object]$Exception,
        [Parameter(Mandatory = $true)][string]$Message
    )
    Write-Log $Method $Message $Exception.Exception.Message [CustomLogType]::Exception
    Write-Log $Method $Message "Exception: $($Exception)"  [CustomLogType]::Exception
    Write-Log $Method $Message "Exception line: $($Exception.InvocationInfo.ScriptLineNumber)" [CustomLogType]::Exception
}

function Add-LogPath {
    param(
        [Parameter(Mandatory = $true)][string] $scriptPath
    )
    
}

function Start-Log {

    Add-Type -TypeDefinition @"
   public enum CustomLogType
   {
      Info,
      Warning,
      Error,
      Exception
   }
"@

    
    $dateNow = [System.DateTime]::Now
    $logFile = [string]::Format("Log_{0}{1}{2}-{3}{4}-pnplog.txt", $dateNow.Year.ToString(), $dateNow.Month.ToString("D2"), $dateNow.Day.ToString("D2"), $dateNow.Hour.ToString("D2"), $dateNow.Minute.ToString("D2"))
    $logFilePath = Join-Path $global:LogGeneralPath $logFile
    Set-PnPTraceLog -On -LogFile $logFilePath -Level Debug -AutoFlush:$true
}

function Stop-Log {
    param()
    Set-PnPTraceLog -Off
}