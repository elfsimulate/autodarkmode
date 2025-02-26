
#param ([int]$dayornight)  

function Ensure-AdminPrivileges {
    param (
        [Parameter(ValueFromRemainingArguments=$true)]
        [string[]]$RemainingArgs
    )

    # 检查是否以管理员身份运行
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        # 重新启动脚本并请求管理员权限
        $params = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        #$params = "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        if ($RemainingArgs) {$params += " -RemainingArgs `"$($RemainingArgs -join ' ')`"" }
        Start-Process powershell.exe -Verb RunAs -ArgumentList $params
        exit
    }
}




    
function Get-LocationFromWindowsAPI {
    try {
        # 加载Windows.Location API
        Add-Type -AssemblyName System.Device
        # 创建位置提供者对象
        $locationProvider = New-Object System.Device.Location.GeoCoordinateWatcher
        # 开始获取位置信息
        $locationProvider.Start()
        # 等待位置信息可用
        while ($locationProvider.Status -ne 'Ready' -and $locationProvider.Status -ne 'Initialized') {
            Start-Sleep -Milliseconds 100
        }
        # 获取坐标
        $coordinate = $locationProvider.Position.Location
        # 检查是否成功获取位置信息
        if ($coordinate.Latitude -ne 0 -and $coordinate.Longitude -ne 0) {
            # 返回经纬度
            return @{
                Latitude = $coordinate.Latitude
                Longitude = $coordinate.Longitude
            }
        } else {
            Write-Host "Failed to retrieve location data."
            return $null
        }

        # 停止位置提供者
        $locationProvider.Stop()
    } catch {
        Write-Host "An error occurred: $_"
        return $null
    }
}#end function




function Get-LocationFromIP {
    try {
        # 发送HTTP请求获取IP信息
        $response = Invoke-RestMethod -Uri "http://ipinfo.io/json" -Method Get

        # 检查请求是否成功
        if ($response) {
            # 提取经纬度信息
            $latitude = $response.loc.Split(',')[0]
            $longitude = $response.loc.Split(',')[1]
            # 返回经纬度
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
        # 构造API请求URL
        $url = "https://api.sunrise-sunset.org/json?lat=$Latitude&lng=$Longitude&date=$Date"

        # 发送HTTP请求
        $response = Invoke-RestMethod -Uri $url -Method Get

        # 检查请求是否成功
        if ($response.status -eq 'OK') {
            # 提取日出、日落、天亮和天黑时间
            $sunrise = $response.results.sunrise
            $sunset = $response.results.sunset
            $twilightBegin = $response.results.civil_twilight_begin
            $twilightEnd = $response.results.civil_twilight_end

            # 检查时间是否为空
            if (-not $sunrise -or -not $sunset -or -not $twilightBegin -or -not $twilightEnd) {
                Write-Host "One or more of the times is null or empty."
                return $null
            }

            # 尝试将时间转换为24小时制并转换为本地时间
            try {
                # 尝试使用 h:mm:ss tt 格式
                $sunriseUtc = [datetime]::ParseExact($sunrise, 'h:mm:ss tt', [System.Globalization.CultureInfo]::InvariantCulture)
                $sunsetUtc = [datetime]::ParseExact($sunset, 'h:mm:ss tt', [System.Globalization.CultureInfo]::InvariantCulture)
                $twilightBeginUtc = [datetime]::ParseExact($twilightBegin, 'h:mm:ss tt', [System.Globalization.CultureInfo]::InvariantCulture)
                $twilightEndUtc = [datetime]::ParseExact($twilightEnd, 'h:mm:ss tt', [System.Globalization.CultureInfo]::InvariantCulture)
            } catch {
                try {
                    # 如果上述格式不匹配，尝试使用 h:mm tt 格式
                    $sunriseUtc = [datetime]::ParseExact($sunrise, 'h:mm tt', [System.Globalization.CultureInfo]::InvariantCulture)
                    $sunsetUtc = [datetime]::ParseExact($sunset, 'h:mm tt', [System.Globalization.CultureInfo]::InvariantCulture)
                    $twilightBeginUtc = [datetime]::ParseExact($twilightBegin, 'h:mm tt', [System.Globalization.CultureInfo]::InvariantCulture)
                    $twilightEndUtc = [datetime]::ParseExact($twilightEnd, 'h:mm tt', [System.Globalization.CultureInfo]::InvariantCulture)
                } catch {
                    Write-Host "Failed to parse time: $_"
                    return $null
                }
            }

            # 将UTC时间转换为本地时间
            $sunriseLocal = $sunriseUtc.ToLocalTime().ToString('HH:mm')
            $sunsetLocal = $sunsetUtc.ToLocalTime().ToString('HH:mm')
            $twilightBeginLocal = $twilightBeginUtc.ToLocalTime().ToString('HH:mm')
            $twilightEndLocal = $twilightEndUtc.ToLocalTime().ToString('HH:mm')

            # 输出日出、日落、天亮和天黑时间
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

    # 获取当前本地时间
    $currentTime = Get-Date

    # 判断当前是白天还是黑夜
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

    # 获取当前本地时间
    $currentTime = Get-Date

    # 判断当前是白天还是黑夜
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
            # 设置浅色模式
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 1
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 1
            Write-Host "Switched to Light Mode."
            return $true
        }
        0 {
            # 设置深色模式
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


# 定义创建或修改任务计划的函数
function Set-ScheduledTaskForScript {
    param (
        [int]$ThemeMode , 
        [datetime]$StartTime
    )
   
    # 检查参数和开始时间是否提供
    if ($null -eq $ThemeMode -or -not $StartTime) {
        Write-Host "Mode and start time must be provided."
        return $false
    }

    # 获取当前脚本的完整路径
    $scriptPath = $MyInvocation.PSCommandPath

    # 构建脚本参数字符串
    $paramsString = $ThemeMode
    $TaskName = if( $ThemeMode -eq 1)  {"LightTheme"} else {"DarkTheme"}
    <#
    foreach ($key in $ScriptParameters.Keys) {
        $paramsString += " -$key $($ScriptParameters[$key])"
    }
    #>
    $userName = "$env:USERDOMAIN\$env:USERNAME"

    # 定义任务触发器：每天在指定时间运行
    $trigger = New-ScheduledTaskTrigger -Daily -At $StartTime

    # 定义任务操作：运行 PowerShell 脚本
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-windowstyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`" $paramsString"

    # 设置任务原则，确保任务在用户未登录时也能运行，但不存储密码
    $principal  = New-ScheduledTaskPrincipal -UserId $userName -LogonType S4U -RunLevel Highest
  
    # 检查任务是否已存在
    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

    if ($existingTask) {
        # 任务已存在，修改触发器
        try {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Register-ScheduledTask -TaskName $TaskName -Trigger $trigger -Action $action -Principal $principal -Force 
            Write-Host "Task '$TaskName' modified successfully."
        } catch {
            Write-Host "Failed to modify task: $_"
        }
    } else {
        # 任务不存在，创建新任务
        try {
            Register-ScheduledTask -TaskName $TaskName -Trigger $trigger -Action $action -Principal $principal -Force 
            Write-Host "Task '$TaskName' created successfully."
        } catch {
            Write-Host "Failed to create task: $_"
        }
    }
}

# 定义创建任务计划的函数
function SlefScheduledTask {
    param ([string]$TaskName)

    # 检查任务名称是否提供
    if (-not $TaskName) {
        Write-Host "Task name must be provided."
        return $false
    }

    # 获取当前脚本的完整路径
    $scriptPath = $MyInvocation.PSCommandPath
    # 定义任务触发器：用户登录时
    $trigger = New-ScheduledTaskTrigger -AtLogon
    $trigger.Delay ='PT3M'
    # 定义任务操作：运行 PowerShell 脚本
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-windowstyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
    $userName = "$env:USERDOMAIN\$env:USERNAME"
    # 设置任务原则，确保任务在用户未登录时也能运行，但不存储密码
    $principal  = New-ScheduledTaskPrincipal -UserId $userName -LogonType S4U -RunLevel Highest
    # 注册任务计划
    try {
        $ot=Register-ScheduledTask -TaskName $TaskName -Trigger $trigger -Action $action -Principal $principal -Force
        #Register-ScheduledTask -TaskName $TaskName -Trigger $trigger -Action $action -User $env:USERNAME -RunLevel Highest -Force
        Write-Host "Task '$TaskName' created successfully."
    } catch {
        Write-Host "Failed to create task: $_"
    }
}

function Test-NetworkConnection {
    param (
        [string]$Target = "http://www.example.com", # 默认目标为http://www.example.com
        [int]$Port = 0,                            # 如果需要测试特定端口，请提供端口号
        [switch]$UseICMP,                          # 使用ICMP（Ping）测试
        [switch]$UseTCP,                           # 使用TCP测试
        [switch]$UseHTTP,                          # 使用HTTP(S)测试，默认开启
        [switch]$CheckAdapters                     # 检查网络适配器状态
    )

    # 如果没有指定任何测试方法，则默认使用HTTP测试
    if (-not ($UseICMP -or $UseTCP -or $CheckAdapters)) {
        $UseHTTP = $true
    }

    $isConnected = $false

    if ($UseICMP) {
        Write-Host "正在使用ICMP测试到 $($Target.Replace('http://','').Replace('https://','')) 的连接..."
        $targetHost = [System.Uri]::new($Target).Host
        $isConnected = Test-Connection -ComputerName $targetHost -Count 1 -Quiet
    } elseif ($UseTCP -and $Port -gt 0) {
        Write-Host "正在使用TCP测试到 $Target,$Port 的连接..."
        try {
            $result = Test-NetConnection -ComputerName $Target -Port $Port -InformationLevel Quiet -ErrorAction Stop
            $isConnected = $result.TcpTestSucceeded
        } catch {
            Write-Host "TCP测试失败: $_"
            $isConnected = $false
        }
    } elseif ($UseHTTP) {
        Write-Host "正在使用HTTP测试到 $Target 的连接..."
        try {
            $response = Invoke-WebRequest -Uri $Target -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
            $isConnected = $true
        } catch {
            Write-Host "HTTP请求失败: $_"
            $isConnected = $false
        }
    } elseif ($CheckAdapters) {
        Write-Host "正在检查网络适配器状态..."
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
        ＄Location = @{
        # 北京的纬度
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


