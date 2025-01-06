
#param ([int]$dayornight)  

function Ensure-AdminPrivileges {
    param (
        [Parameter(ValueFromRemainingArguments=$true)]
        [string[]]$RemainingArgs
    )

    # ����Ƿ��Թ���Ա�������
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        # ���������ű����������ԱȨ��
        $params = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        #$params = "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        if ($RemainingArgs) {$params += " -RemainingArgs `"$($RemainingArgs -join ' ')`"" }
        Start-Process powershell.exe -Verb RunAs -ArgumentList $params
        exit
    }
}




    
function Get-LocationFromWindowsAPI {
    try {
        # ����Windows.Location API
        Add-Type -AssemblyName System.Device
        # ����λ���ṩ�߶���
        $locationProvider = New-Object System.Device.Location.GeoCoordinateWatcher
        # ��ʼ��ȡλ����Ϣ
        $locationProvider.Start()
        # �ȴ�λ����Ϣ����
        while ($locationProvider.Status -ne 'Ready' -and $locationProvider.Status -ne 'Initialized') {
            Start-Sleep -Milliseconds 100
        }
        # ��ȡ����
        $coordinate = $locationProvider.Position.Location
        # ����Ƿ�ɹ���ȡλ����Ϣ
        if ($coordinate.Latitude -ne 0 -and $coordinate.Longitude -ne 0) {
            # ���ؾ�γ��
            return @{
                Latitude = $coordinate.Latitude
                Longitude = $coordinate.Longitude
            }
        } else {
            Write-Host "Failed to retrieve location data."
            return $null
        }

        # ֹͣλ���ṩ��
        $locationProvider.Stop()
    } catch {
        Write-Host "An error occurred: $_"
        return $null
    }
}#end function




function Get-LocationFromIP {
    try {
        # ����HTTP�����ȡIP��Ϣ
        $response = Invoke-RestMethod -Uri "http://ipinfo.io/json" -Method Get

        # ��������Ƿ�ɹ�
        if ($response) {
            # ��ȡ��γ����Ϣ
            $latitude = $response.loc.Split(',')[0]
            $longitude = $response.loc.Split(',')[1]
            # ���ؾ�γ��
            return @{
                Latitude = $latitude
                Longitude = $longitude
            }
        } else {
            Write-Host "Failed to retrieve location data."
            return $null
        }
    } catch {
        Write-Host "An error occurred: $_"
        return $null
    }
}#end function

function Get-SunriseSunset {
    param (
        [double]$Latitude,
        [double]$Longitude,
        [string]$Date = (Get-Date).ToString("yyyy-MM-dd")
    )

    try {
        # ����API����URL
        $url = "https://api.sunrise-sunset.org/json?lat=$Latitude&lng=$Longitude&date=$Date"

        # ����HTTP����
        $response = Invoke-RestMethod -Uri $url -Method Get

        # ��������Ƿ�ɹ�
        if ($response.status -eq 'OK') {
            # ��ȡ�ճ������䡢���������ʱ��
            $sunrise = $response.results.sunrise
            $sunset = $response.results.sunset
            $twilightBegin = $response.results.civil_twilight_begin
            $twilightEnd = $response.results.civil_twilight_end

            # ���ʱ���Ƿ�Ϊ��
            if (-not $sunrise -or -not $sunset -or -not $twilightBegin -or -not $twilightEnd) {
                Write-Host "One or more of the times is null or empty."
                return $null
            }

            # ���Խ�ʱ��ת��Ϊ24Сʱ�Ʋ�ת��Ϊ����ʱ��
            try {
                # ����ʹ�� h:mm:ss tt ��ʽ
                $sunriseUtc = [datetime]::ParseExact($sunrise, 'h:mm:ss tt', [System.Globalization.CultureInfo]::InvariantCulture)
                $sunsetUtc = [datetime]::ParseExact($sunset, 'h:mm:ss tt', [System.Globalization.CultureInfo]::InvariantCulture)
                $twilightBeginUtc = [datetime]::ParseExact($twilightBegin, 'h:mm:ss tt', [System.Globalization.CultureInfo]::InvariantCulture)
                $twilightEndUtc = [datetime]::ParseExact($twilightEnd, 'h:mm:ss tt', [System.Globalization.CultureInfo]::InvariantCulture)
            } catch {
                try {
                    # ���������ʽ��ƥ�䣬����ʹ�� h:mm tt ��ʽ
                    $sunriseUtc = [datetime]::ParseExact($sunrise, 'h:mm tt', [System.Globalization.CultureInfo]::InvariantCulture)
                    $sunsetUtc = [datetime]::ParseExact($sunset, 'h:mm tt', [System.Globalization.CultureInfo]::InvariantCulture)
                    $twilightBeginUtc = [datetime]::ParseExact($twilightBegin, 'h:mm tt', [System.Globalization.CultureInfo]::InvariantCulture)
                    $twilightEndUtc = [datetime]::ParseExact($twilightEnd, 'h:mm tt', [System.Globalization.CultureInfo]::InvariantCulture)
                } catch {
                    Write-Host "Failed to parse time: $_"
                    return $null
                }
            }

            # ��UTCʱ��ת��Ϊ����ʱ��
            $sunriseLocal = $sunriseUtc.ToLocalTime().ToString('HH:mm')
            $sunsetLocal = $sunsetUtc.ToLocalTime().ToString('HH:mm')
            $twilightBeginLocal = $twilightBeginUtc.ToLocalTime().ToString('HH:mm')
            $twilightEndLocal = $twilightEndUtc.ToLocalTime().ToString('HH:mm')

            # ����ճ������䡢���������ʱ��
            return @{
                Sunrise = $sunriseLocal
                Sunset = $sunsetLocal
                Dawn = $twilightBeginLocal
                Dusk= $twilightEndLocal
            }
        } else {
            Write-Host "API request failed with status: $($response.status)"
            return $null
        }
    } catch {
        Write-Host "An error occurred: $_"
        return $null
    }
}
function Get-DayNightStatus {
    param (
        [datetime]$Sunrise,
        [datetime]$Sunset
    )

    # ��ȡ��ǰ����ʱ��
    $currentTime = Get-Date

    # �жϵ�ǰ�ǰ��컹�Ǻ�ҹ
    if ($currentTime -ge $Sunrise -and $currentTime -lt $Sunset) {
        return 1 #"day"
    } else {
        return 0 #"night"
    }
}
function Get-IsNightNow {
    param (
        [datetime]$Sunrise,
        [datetime]$Sunset
    )

    # ��ȡ��ǰ����ʱ��
    $currentTime = Get-Date

    # �жϵ�ǰ�ǰ��컹�Ǻ�ҹ
    if ($currentTime -ge $Sunrise -and $currentTime -lt $Sunset) {
        return $false
    } else {
        return $true
    }
}


function Set-WindowsTheme {
    param ([int]$Mode)

    switch ($Mode) {
        1 {
            # ����ǳɫģʽ
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 1
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 1
            Write-Host "Switched to Light Mode."
            return $true
        }
        0 {
            # ������ɫģʽ
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0
            Write-Host "Switched to Dark Mode."
            return $true
        }
        default {
            Write-Host "Unknown mode: $Mode"
            return $false
        }
    }
}


# ���崴�����޸�����ƻ��ĺ���
function Set-ScheduledTaskForScript {
    param (
        [int]$ThemeMode , 
        [datetime]$StartTime
    )
   
    # �������Ϳ�ʼʱ���Ƿ��ṩ
    if ($null -eq $ThemeMode -or -not $StartTime) {
        Write-Host "Mode and start time must be provided."
        return $false
    }

    # ��ȡ��ǰ�ű�������·��
    $scriptPath = $MyInvocation.PSCommandPath

    # �����ű������ַ���
    $paramsString = $ThemeMode
    $TaskName = if( $ThemeMode -eq 1)  {"LightTheme"} else {"DarkTheme"}
    <#
    foreach ($key in $ScriptParameters.Keys) {
        $paramsString += " -$key $($ScriptParameters[$key])"
    }
    #>
    $userName = "$env:USERDOMAIN\$env:USERNAME"

    # �������񴥷�����ÿ����ָ��ʱ������
    $trigger = New-ScheduledTaskTrigger -Daily -At $StartTime

    # ����������������� PowerShell �ű�
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-windowstyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`" $paramsString"

    # ��������ԭ��ȷ���������û�δ��¼ʱҲ�����У������洢����
    $principal  = New-ScheduledTaskPrincipal -UserId $userName -LogonType S4U -RunLevel Highest
  
    # ��������Ƿ��Ѵ���
    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

    if ($existingTask) {
        # �����Ѵ��ڣ��޸Ĵ�����
        try {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Register-ScheduledTask -TaskName $TaskName -Trigger $trigger -Action $action -Principal $principal -Force 
            Write-Host "Task '$TaskName' modified successfully."
        } catch {
            Write-Host "Failed to modify task: $_"
        }
    } else {
        # ���񲻴��ڣ�����������
        try {
            Register-ScheduledTask -TaskName $TaskName -Trigger $trigger -Action $action -Principal $principal -Force 
            Write-Host "Task '$TaskName' created successfully."
        } catch {
            Write-Host "Failed to create task: $_"
        }
    }
}

# ���崴������ƻ��ĺ���
function SlefScheduledTask {
    param ([string]$TaskName)

    # ������������Ƿ��ṩ
    if (-not $TaskName) {
        Write-Host "Task name must be provided."
        return $false
    }

    # ��ȡ��ǰ�ű�������·��
    $scriptPath = $MyInvocation.PSCommandPath
    # �������񴥷������û���¼ʱ
    $trigger = New-ScheduledTaskTrigger -AtLogon
    # ����������������� PowerShell �ű�
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-windowstyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

    # ע������ƻ�
    try {
        Register-ScheduledTask -TaskName $TaskName -Trigger $trigger -Action $action -User $env:USERNAME -RunLevel Highest -Force
        Write-Host "Task '$TaskName' created successfully."
    } catch {
        Write-Host "Failed to create task: $_"
    }
}

function Test-NetworkConnection {
    param (
        [string]$Target = "http://www.example.com", # Ĭ��Ŀ��Ϊhttp://www.example.com
        [int]$Port = 0,                            # �����Ҫ�����ض��˿ڣ����ṩ�˿ں�
        [switch]$UseICMP,                          # ʹ��ICMP��Ping������
        [switch]$UseTCP,                           # ʹ��TCP����
        [switch]$UseHTTP,                          # ʹ��HTTP(S)���ԣ�Ĭ�Ͽ���
        [switch]$CheckAdapters                     # �������������״̬
    )

    # ���û��ָ���κβ��Է�������Ĭ��ʹ��HTTP����
    if (-not ($UseICMP -or $UseTCP -or $CheckAdapters)) {
        $UseHTTP = $true
    }

    $isConnected = $false

    if ($UseICMP) {
        Write-Host "����ʹ��ICMP���Ե� $($Target.Replace('http://','').Replace('https://','')) ������..."
        $targetHost = [System.Uri]::new($Target).Host
        $isConnected = Test-Connection -ComputerName $targetHost -Count 1 -Quiet
    } elseif ($UseTCP -and $Port -gt 0) {
        Write-Host "����ʹ��TCP���Ե� $Target,$Port ������..."
        try {
            $result = Test-NetConnection -ComputerName $Target -Port $Port -InformationLevel Quiet -ErrorAction Stop
            $isConnected = $result.TcpTestSucceeded
        } catch {
            Write-Host "TCP����ʧ��: $_"
            $isConnected = $false
        }
    } elseif ($UseHTTP) {
        Write-Host "����ʹ��HTTP���Ե� $Target ������..."
        try {
            $response = Invoke-WebRequest -Uri $Target -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
            $isConnected = $true
        } catch {
            Write-Host "HTTP����ʧ��: $_"
            $isConnected = $false
        }
    } elseif ($CheckAdapters) {
        Write-Host "���ڼ������������״̬..."
        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
        $isConnected = $adapters.Count -gt 0
    }

    return $isConnected
}


 


#__________main function____________

Ensure-AdminPrivileges @args

if ($args.Count -eq 1) {
     $intMode=$args[0]  
     if ([int]::TryParse($args[0],[ref]$intMode)) {
        $result = Set-WindowsTheme -Mode $intMode        
     }
    exit
}

if (Test-NetworkConnection) {
    $Location=$Location=Get-LocationFromIP
    if (-not $Location) {$Location=Get-LocationFromWindowsAPI}
    if (-not $Location) {
        ��Location = @{
        # ������γ��
        Latitude = 39.9042  
        Longitude = 116.4074  
        }
    }
    $sunriseSunset = Get-SunriseSunset -Latitude $Location.Latitude -Longitude $Location.Longitude
}
$today = Get-Date
if ($sunriseSunSet -eq $null) {
    $sunriseSunSet = @{
        Sunrise = [datetime]::ParseExact($today.ToString("yyyy-MM-dd") + " 07:00", "yyyy-MM-dd HH:mm", $null)
        Sunset  = [datetime]::ParseExact($today.ToString("yyyy-MM-dd") + " 17:00", "yyyy-MM-dd HH:mm", $null)
        Dawn    = [datetime]::ParseExact($today.ToString("yyyy-MM-dd") + " 07:30", "yyyy-MM-dd HH:mm", $null)
        Dusk    = [datetime]::ParseExact($today.ToString("yyyy-MM-dd") + " 17:30", "yyyy-MM-dd HH:mm", $null)
    }
}


Write-Host $sunriseSunset.Sunrise,$sunriseSunset.Sunset

$dayNightStatus = Get-DayNightStatus -Sunrise $sunriseSunset.Sunrise -Sunset $sunriseSunset.Sunset
$result = Set-WindowsTheme -Mode $dayNightStatus
if ($dayNightStatus -eq 1) {
    Set-ScheduledTaskForScript -ThemeMode 0 -StartTime $($sunriseSunset.Sunset)
}else{
    Set-ScheduledTaskForScript -ThemeMode 1 -StartTime $($sunriseSunset.Sunrise)
}

$SlefTaskName="Theme" 
$existingTask = Get-ScheduledTask -TaskName $SlefTaskName -ErrorAction SilentlyContinue
if (-not $existingTask) {SlefScheduledTask -TaskName $SlefTaskName}


