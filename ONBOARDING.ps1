################################################
# 
# AUTHOR:  Eddie
# EMAIL:   eddie@directbox.de
# BLOG:    https://exchangeblogonline.de
# COMMENT: Migrate OnPremise Mailbox to Exchange Online
#
################################################

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, HelpMessage = "Bitte Mailbox UPN eingeben")]
    [ValidateNotNullorEmpty()] [string] $SourceMailbox,
    
    [Parameter(Mandatory = $true, HelpMessage = "Bitte die externe Exchange URL eintragen")]
    [ValidateNotNullorEmpty()] [string] $ExchangeFQDN    

)

$Host.ui.RawUI.WindowTitle = "EXCHANGE ONLINE ONBOARDING"
$ErrorActionPreference = "SilentlyContinue"

Clear-Host
function header {
    $datum = Get-Date -Format ("HH:mm  dd/MM/yyyy")
    Write-Host "
 --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
 |e| |x| |c| |h| |a| |n| |g| |e| |b| |l| |o| |g| |o| |n| |l| |i| |n| |e| |.| |d| |e|
 --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
 Powered by Eddie  |  https://exchangeblogonline.de

" -F Green
    Write-Host "$datum                       `n" -b Blue
			
}
header

#query active sessions
if ((Get-PSSession).ComputerName -notmatch "outlook.office365"){    

    Get-PSSession | Remove-PSSession  -ea 0

    #########################################################
    #cloud credentials
    Write-Host "Bitte die Cloud-Credentials eingeben:" -f Yellow
    $credential = Get-Credential "user@domain.de"

    #import office 365 session
    $proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential `
        -Authentication Basic -AllowRedirection
    ######################################################################
    #o365 session
    Import-PSSession $Session -wa 0 -AllowClobber
    Start-Sleep 5

    if (!(Get-PSSession $Session)) {
        Write-Host "connect via ie proxy settings..." -ForegroundColor Yellow
 
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -SessionOption $proxysettings
        Start-Sleep 3
        Import-PSSession $Session -AllowClobber -ea 0 SilentlyContinue -wa 0
    }  
    if (!(Get-PSSession $Session)) {
        Write-Host "Connection to Office 365 has failed!" -ForegroundColor Red 
        break   
    }

}


cls
header

#exchange OnPremise login
Write-Host "Bitte die OnPremise-Credentials eingeben:" -f Yellow
$credential = Get-Credential "DOMAIN\logonname"

#remove last move request information
Remove-MoveRequest $SourceMailbox -Confirm:$false -ea 0 -wa 0

#targetDeliveryDomain
$targetdomain = (Get-AcceptedDomain | ? { $_.DomainName -match "mail" }).name

Write-Host "Is this target domain correct? $targetdomain [y / n]" -ForegroundColor Yellow

$input = Read-Host
if($input -match "n"){
    Write-Host "please enter the target mail domain [Get-AcceptedDomain]"
    $global:targetdomain = Read-Host "Bitte die target mail domain eintragen [TENANTNAME.mail.onmicrosoft.com]"
}

cls
header

Write-Host "Migration wird vorbereitet... `n" -f Yellow

function moverequest($var){
        new-moverequest -identity $SourceMailbox -remote -remotehostname $var `
            -targetdeliverydomain $targetdomain  `
            -baditemlimit unlimited  `
            -LargeItemLimit unlimited `
            -AcceptLargeDataLoss `
            -remotecredential $credential #`            
            #-CompleteAfter (Get-Date).AddDays(+5)
        # kann auskommentiert werden, um die Migration zu einem bestimmten Zeitpunkt abzuschliessen
    }

moverequest($ExchangeFQDN)

Start-Sleep 3

###############################################################################################
$status = get-moverequest -identity $sourcemailbox -ea 0 -wa 0

if (!($status)) {
    cls
    header

    Write-Host "
    Der Zugriff auf den Exchange MigrationsEndPunkt ist fehlgeschlagen.`n" -f Red
    do { 
        Write-Host "Bitte OnPremise Exchange Domain eintragen - z.B. [exchangeblogonline.de]:" -f Yellow
        $opURI = Read-Host "Externe Exchange URL"
        moverequest($opURI)
    }until($opURI -ne [string]::empty)

}ifelse (!($status)) {
    Write-host "Die Migration war leider nicht erfolgreich!" -ForegroundColor Red; 
    Write-host "SMTP Domain der Mailbox bereits erfolgreich registriert?" -ForegroundColor Red; 
    break
}
else {
    Write-Host "Die Migration wurde angestossen. `n" -ForegroundColor Yellow
}

#migration status feedback
do {
	cls
    header
    $progress = (Get-MoveRequest $sourcemailbox).Status	
	
    if ($progress -match "Failed") {
        Suspend-MoveRequest $sourcemailbox -Confirm:$false
        Set-MoveRequest $sourcemailbox -BadItemLimit unlimited -LargeItemLimit unlimited -AcceptLargeDataLoss 
        Start-Sleep 5
        Resume-MoveRequest $sourcemailbox
    }
    if (!($progress)) {
        Write-Host "Der MoveRequest ist nicht [mehr] vorhanden!" -ForegroundColor Red
        break        
    }	
	    
    Write-Host "Migrations-Status..." -Fore Yellow	
    Get-MoveRequestStatistics $sourcemailbox | ft DisplayName, StatusDetail, TotalMailboxSize, TotalArchiveSize, PercentComplete
    start-sleep -s 5
    cls           
           
}until ($progress -match "Complete") 

if ($progress -match "Complete") { 
    Write-Host "Die Migration wurde erfolgreich abgeschlossen!" -ForegroundColor Green
    pause
} else {
     Write-Host "Skript wurde beendet!"
}

#END
