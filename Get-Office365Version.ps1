<#
.SYNOPSIS
Call a Microsoft webservice to get the latest O365 build numbers per a particular channel.

.DESCRIPTION
Calls a Microsoft webservice that was gleaned from the O365 2016 Deployment Tool. Returns an object with version numbers and other data.

.PARAMETER Architecture
The bitness of O365 software you want to install

.PARAMETER Channel
Chanel describes how often you update O365 in your environement.

.PARAMETER ProductId
Describes which product you are searching for (Pro Plus, Visio Pro etc.) Matches the product id used by the deployment tool.
https://support.microsoft.com/en-us/help/2842297/product-ids-that-are-supported-by-the-office-deployment-tool-for-click

.PARAMETER Language
Lanugague of the product.

.PARAMETER Version
Only seems to support 2016 (16) currently. 

.EXAMPLE
Query for Visio Pro in the semi-annual for Spanish language

Get-Office365Version -Architecture x64 -Channel Semi-Annual -Language es-es -ProductId VisioProRetail 

#>

function Get-Office365Version {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$false)]
        [ValidateSet("x86", "x64")]
        [string] 
        $Architecture = "x86",

        [Parameter(Mandatory=$false)]
        [ValidateSet("Semi-Annual", "Semi-Annual (Targeted)", "Monthly")]
        [string] 
        $Channel = "Semi-Annual",

        [Parameter(Mandatory=$false)]
        [ValidateSet('O365ProPlusRetail', 'O365BusinessRetail', 'VisioProRetail', 'ProjectProRetail', 'SPDRetail', 'AccessRuntimeRetail', 'LanguagePack')]
        [string] 
        $ProductId = 'O365ProPlusRetail',

        [Parameter(Mandatory=$false)]
        [ValidateSet('en-us', 'ar-sa', 'bg-bg', 'zh-cn', 'zh-tw', 'hr-hr', 'cs-cz', 'da-dk', 'nl-nl', 'et-ee', 'fi-fi', 'fr-fr', 'de-de', 'el-gr', 'he-il', 'hi-in', 'hu-hu', 'id-id', 'it-it', 'ja-jp', 'kk-kz', 'ko-kr', 'lv-lv', 'lt-lt', 'ms-my', 'nb-no', 'pl-pl', 'pt-br', 'pt-pt', 'ro-ro', 'ru-ru', 'sr-latn-rs', 'sk-sk', 'sl-si', 'es-es', 'sv-se', 'th-th', 'tr-tr', 'uk-ua', 'vi-vn' )]
        [string] 
        $Language = 'en-us',

        [Parameter(Mandatory=$false)]
        [ValidateSet('16')]
        [string] 
        $Version = '16'
    )

    # Determines the channel CDN
    $FFN = switch ($Channel){
        "Semi-Annual"  {'7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'}
        "Semi-Annual (Targeted)" {'b8f9b850-328d-4355-9145-c59439a0c4cf'}
        "Monthly" {'492350f6-3a01-4f97-b9c0-c7c6ddf67d60'}
    }

    # Product IDs
    $prids = "$($ProductId).$($Version)_$($Language)._x-none"

    # OS info
    $os = Get-CimInstance Win32_OperatingSystem | select ProductType, Version
    $osver = Switch ($os.ProductType) {
        1 {"Client%7c$($os.Version.toString())"}
        default {"Server%7c$($os.Version.toString())"}
    }

    # Unknown, seems to accept blank.
    $tid = ''

    # Unknown, seems to accept blank.
    $omid = ''

    # WSUS ID, seems to accept blank.
    $susid = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate' -Name SusClientId

    # Unknown, seems to accept blank.
    $werid = ''

    $url = "https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData?audienceFFN=$($FFN)&prids=$($prids)&osver=$($osver)&bit=$($Architecture)&tid=$tid&omid=$($omid)&susid=$($susid)&werid=$($werid)"
    Write-Verbose "Contacting web service: $url"
    
    Try {
        $result = Invoke-RestMethod -Method Get -Uri $url -ErrorAction Stop
    } Catch {
        Write-Error "Unable to contact web service.`n$($_.exception)"
    }

    return $result
}

Get-Office365Version