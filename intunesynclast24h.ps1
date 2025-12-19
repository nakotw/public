Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All"

$Threshold = (Get-Date).AddHours(-24)

$Devices = Get-MgDeviceManagementManagedDevice -All `
    -Filter "operatingSystem eq 'Windows'"

$DevicesToSync = $Devices | Where-Object {
    $_.LastSyncDateTime -lt $Threshold
}

foreach ($Device in $DevicesToSync) {
    Write-Host "Syncing $($Device.DeviceName) (Last sync: $($Device.LastSyncDateTime))" `
        -ForegroundColor Cyan

    $Uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($Device.Id)/syncDevice"
    Invoke-MgGraphRequest -Method POST -Uri $Uri

    Start-Sleep -Seconds 2   # avoid throttling
}
