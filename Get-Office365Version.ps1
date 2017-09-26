function Get-Office365Version {
    [cmdletbinding()]
    param (
        [ValidateSet("x86", "x64")]
        [string] 
        $Architecture = "x86",

        [ValidateSet("Semi-Annual", "Semi-Annual (Targeted)", "Monthly")]
        [string] 
        $Channel = "Semi-Annual",

        [ValidateSet('O365ProPlusRetail', 'VisioProRetail', 'ProjectProRetail')]
        [string] 
        $ProductId = 'O365ProPlusRetail',

        [ValidateSet('en-us', 'ja-jp')]
        [string] 
        $Language = 'en-us'
    )

    # Determines the channel CDN
    $FFN = switch ($Channel){
        "Semi-Annual"  {'7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'}
        "Semi-Annual (Targeted)" {'b8f9b850-328d-4355-9145-c59439a0c4cf'}
        "Monthly" {'492350f6-3a01-4f97-b9c0-c7c6ddf67d60'}
    }

    # Product IDs
    #$prids = $ProductIds -join('%7c')
    $prids = "$($ProductId).16_$($Language)._x-none"

    # OS info
    $os = Get-CimInstance Win32_OperatingSystem | select ProductType, Version
    $osver = Switch ($os.ProductType) {
        1 {"Client%7c$($os.Version.toString())"}
        default {"Server%7c$($os.Version.toString())"}
    }

    # ???
    $tid = ''

    # ???
    $omid = ''

    # WSUS ID
    $susid = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate' -Name SusClientId

    # ???
    $werid = ''

    $url = "https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData?audienceFFN=$($FFN)&prids=$($prids)&osver=$($osver)&bit=$($Architecture)&tid=$tid&omid=$($omid)&susid=$($susid)&werid=$($werid)"

    
    $result = Invoke-RestMethod -Method Get -Uri $url

    return $result
}

