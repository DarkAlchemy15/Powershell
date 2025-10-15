$all_hosts = Get-vmhost

$AdvSettingsStigs = @(
    [PSCustomObject]@{GroupID = 'V-265977'; Setting = 'Net.BMCNetworkEnable'; STIGValue = '0' }
    #[PSCustomObject]@{GroupID = 'V-258797'; Setting = 'Syslog.global.logDir'; STIGValue = 'Needs persistent storage location' }
    [PSCustomObject]@{GroupID = 'V-258787'; Setting = 'Syslog.global.auditRecord.storageEnable'; STIGValue = 'true' }
    [PSCustomObject]@{GroupID = 'V-258779'; Setting = 'Syslog.global.certificate.checkSSLCerts'; STIGValue = 'true' }
    [PSCustomObject]@{GroupID = 'V-258757'; Setting = 'UserVars.DcuiTimeOut'; STIGValue = '600' }
    [PSCustomObject]@{GroupID = 'V-258756'; Setting = 'UserVars.ESXiShellTimeOut'; STIGValue = '600' }
    [PSCustomObject]@{GroupID = 'V-258744'; Setting = 'Syslog.global.logHost'; STIGValue = 'Valid Syslog Server' }
    [PSCustomObject]@{GroupID = 'V-258742'; Setting = 'Security.AccountUnlockTime'; STIGValue = '900' }
    [PSCustomObject]@{GroupID = 'V-258739'; Setting = 'UserVars.ESXiShellInteractiveTimeOut'; STIGValue = '900' }
    [PSCustomObject]@{GroupID = 'V-258735'; Setting = 'Security.PasswordHistory'; STIGValue = '5' }
    #[PSCustomObject]@{GroupID = 'V-258734'; Setting = 'Security.PasswordQualityControl'; STIGValue = 'similar=deny retry=3 min=disabled,disabled,disabled,disabled,15' }
    [PSCustomObject]@{GroupID = 'V-258798'; Setting = 'VMkernel.Boot.execInstalledOnly'; STIGValue = 'true' }
    [PSCustomObject]@{GroupID = 'V-258796'; Setting = 'Config.HostAgent.plugins.hostsvc.esxAdminsGroup'; STIGValue = 'Not Set to ESX Admins' }
    [PSCustomObject]@{GroupID = 'V-258790'; Setting = 'Syslog.global.logLevel'; STIGValue = 'info' }
    [PSCustomObject]@{GroupID = 'V-258789'; Setting = 'Syslog.global.certificate.strictX509Compliance'; STIGValue = 'true' }
    [PSCustomObject]@{GroupID = 'V-258788'; Setting = 'Syslog.global.auditRecord.remoteEnable'; STIGValue = 'true' }
    [PSCustomObject]@{GroupID = 'V-258782'; Setting = 'Security.PasswordMaxDays'; STIGValue = '90' }
    [PSCustomObject]@{GroupID = 'V-258781'; Setting = 'Config.HostAgent.vmacore.soap.sessionTimeout'; STIGValue = '30' }
    [PSCustomObject]@{GroupID = 'V-258780'; Setting = 'Mem.MemEagerZero'; STIGValue = '1' }
    [PSCustomObject]@{GroupID = 'V-258778'; Setting = 'UserVars.SuppressHyperthreadWarning'; STIGValue = '0' }
    [PSCustomObject]@{GroupID = 'V-258777'; Setting = 'UserVars.SuppressShellWarning'; STIGValue = '0' }
    [PSCustomObject]@{GroupID = 'V-258774'; Setting = 'Net.DVFilterBindIpAddress'; STIGValue = '' }
    [PSCustomObject]@{GroupID = 'V-258770'; Setting = 'Net.BlockGuestBPDU'; STIGValue = '1' }
    [PSCustomObject]@{GroupID = 'V-258768'; Setting = 'Mem.ShareForceSalting'; STIGValue = '2' }
    #[PSCustomObject]@{GroupID = 'V-258752'; Setting = 'Config.Etc.issue'; STIGValue = "I've read & consent to terms in IS user agreem't." }
    [PSCustomObject]@{GroupID = 'V-258751'; Setting = 'DCUI.Access'; STIGValue = 'root' }
    [PSCustomObject]@{GroupID = 'V-258743'; Setting = 'Syslog.global.auditRecord.storageCapacity'; STIGValue = '100' }
    [PSCustomObject]@{GroupID = 'V-258736'; Setting = 'Config.HostAgent.plugins.solo.enableMob'; STIGValue = 'false' }
    [PSCustomObject]@{GroupID = 'V-258733'; Setting = 'Config.HostAgent.log.level'; STIGValue = 'info' }
    [PSCustomObject]@{GroupID = 'V-258731'; Setting = 'UserVars.HostClientSessionTimeout'; STIGValue = '900' }
    #[PSCustomObject]@{GroupID = 'V-258729'; Setting = 'Annotations.WelcomeMessage'; STIGValue = "I've read & consent to terms in IS user agreem't." }
    [PSCustomObject]@{GroupID = 'V-258728'; Setting = 'Security.AccountLockFailures'; STIGValue = '3' }
)

### The $msg.running feild is a boolean. You can't compare a boolean to an object, so make sure the PSCustomObject is a boolean for the running field
$VMHostServices = @(
    [PSCustomObject]@{GroupID = 'V-258745'; Service = 'NTP Daemon'; STIGValue = 'on'; SRunning = $true }
    [PSCustomObject]@{GroupID = 'V-258745'; Service = 'PTP Daemon'; STIGValue = 'off'; SRunning = $false }
    [PSCustomObject]@{GroupID = 'V-258786'; Service = 'slpd'; STIGValue = 'off'; SRunning = $false }
    [PSCustomObject]@{GroupID = 'V-258783'; Service = 'CIM Server'; STIGValue = 'off'; SRunning = $false }
    [PSCustomObject]@{GroupID = 'V-258755'; Service = 'ESXi Shell'; STIGValue = 'off'; SRunning = $false }
    [PSCustomObject]@{GroupID = 'V-258754'; Service = 'SSH'; STIGValue = 'off'; SRunning = $false }
)


$Results_AdvSettings = @()
$Results_Services = @()

Foreach ($ESXi in $all_hosts) {
    $GetEsxi = get-vmhost -name $ESXi.Name
    
    foreach ($Stig in $AdvSettingsStigs) {
        $ESXiData = [PSCustomObject]@{}
        $msg = $GetEsxi | Get-AdvancedSetting -name $stig.Setting #| Select-Object Entity, Name, Value, @{n = "STIG-Value"; e = { $stig.STIGValue } }
        
        $ESXiData | Add-Member -MemberType NoteProperty -Name 'ESXi Host' -Value $msg.Entity
        $ESXiData | Add-Member -MemberType NoteProperty -Name 'Group ID' -Value $Stig.GroupID
        $ESXiData | Add-Member -MemberType NoteProperty -Name 'Advanced Setting' -Value $msg.Name
        $ESXiData | Add-Member -MemberType NoteProperty -Name 'Value' -Value $msg.Value
        $ESXiData | Add-Member -MemberType NoteProperty -Name 'STIG Value' -Value $stig.STIGValue
        if ($msg.Value -ne $stig.STIGValue) {
            $ESXiData | Add-Member -MemberType NoteProperty -Name 'Status' -Value "This is a Finding"
        }
        elseif ($msg.Value -eq $stig.STIGValue) {
            $ESXiData | Add-Member -MemberType NoteProperty -Name 'Status' -Value "NOT a Finding"
        }
        else {
            $ESXiData | Add-Member -MemberType NoteProperty -Name 'Status' -Value " "
        }
        $Results_AdvSettings += $ESXiData
    }
    
    ForEach ($Service in $VMHostServices) {
        $serviceInfo = [PSCustomObject]@{}
        $msg = Get-VMHostService -VMHost $ESXi.name | Where-Object "label" -eq $Service.Service
        $serviceInfo | Add-Member -MemberType NoteProperty -Name 'ESXi Host' -Value $ESXi.name
        $serviceInfo | Add-Member -MemberType NoteProperty -Name 'Group ID' -Value $Service.GroupID
        $serviceInfo | Add-Member -MemberType NoteProperty -Name 'Service' -Value $msg.Label
        $serviceInfo | Add-Member -MemberType NoteProperty -Name 'Value' -Value $msg.Policy
        $serviceInfo | Add-Member -MemberType NoteProperty -Name 'STIG Value' -Value $Service.STIGValue
        $serviceInfo | Add-Member -MemberType NoteProperty -Name 'Running' -Value $msg.Running
        $serviceInfo | Add-Member -MemberType NoteProperty -Name 'STigRunning' -Value $Service.SRunning
        ### The $msg.running feild is a boolean. You can't compare a boolean to an object, so make sure the PSCustomObject is a boolean for the running field
        if ($msg.Policy -eq $Service.STIGValue -and $msg.Running -eq $Service.SRunning) {
            $serviceInfo | Add-Member -MemberType NoteProperty -Name 'Status' -Value "NOT a Finding"
        }
        elseif ($msg.Policy -ne $Service.STIGValue -or $msg.Running -ne $Service.SRunning) {
            $serviceInfo | Add-Member -MemberType NoteProperty -Name 'Status' -Value "This is a Finding"
        }
        else {
            $serviceInfo | Add-Member -MemberType NoteProperty -Name 'Status' -Value " "
        }
        
        $Results_Services += $serviceInfo
    }
}

$Results_AdvSettings | Format-Table `
    'ESXi Host', `
    'Group ID', `
    'Advanced Setting', `
    'Value', 'STIG Value', `
@{label = 'Status'; e = { $color = if ($_.status -eq 'This is a Finding') { $PSStyle.Background.Black + $PSStyle.Foreground.Red }"$color$($_.Status)$($PSStyle.Reset)" } }


$Results_Services | Format-Table `
    'ESXi Host', `
    'Group ID', `
    'Service', `
    'Value', `
    'STIG Value', `
    'Running', `
    'STigRunning', `
@{label = 'Status'; e = { $color = if ($_.status -eq 'This is a Finding') { $PSStyle.Background.Black + $PSStyle.Foreground.Red }"$color$($_.Status)$($PSStyle.Reset)" } }
