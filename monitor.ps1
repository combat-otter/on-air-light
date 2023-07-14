
function Confirm-MicrophoneActivity
{
    $confirmed = $false

    $root_key = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"
    $sub_keys = Get-ChildItem -Path $root_key -Recurse

    foreach ($key in $sub_keys)
    {
        $started = $key.GetValue("LastUsedTimeStart")
        $stopped = $key.GetValue("LastUsedTimeStop")
        if ($started -ne $null -and $stopped -ne $null)
        {
            if ($stopped -le 0)
            {
                if ($key.PSParentPath.EndsWith('NonPackaged'))
                {
                    $app_path = $key.PSChildName -replace '#', '\'
                    $app_name = Split-Path $app_path -Leaf
                }
                else 
                {
                    $app_name = $key.PSChildName -replace '_[0-9a-z]+$', ''
                }
                
                Write-Host "'$app_name' is currently using the microphone"
                $confirmed = $true
            }
        }
        
    }

    return $confirmed
}

while ($true)
{
    $activity = Confirm-MicrophoneActivity
    Start-Sleep -Seconds 5
}