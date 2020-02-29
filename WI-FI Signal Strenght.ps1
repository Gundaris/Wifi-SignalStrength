<#

    .SYNOPSIS
    This is a simple, but yet a super Powershell script to measure signal strength and document the signal and the Access Points the signal is received from.

    .DESCRIPTION
    The script monitor your Wifi signal strenght and if desired write the result to a file for documentation.

    .PARAMETER <$Delay>
    The delay variable sets the amount of seconds the script will wait after each run of the script, before it runs again. 
    The complete wait time will be the time to run the script plus the specified delay time. The default delay time is 3 seconds.

    .PARAMETER <$PopUpWarning>
    Should you for some strange reason wan to have a popup warning everytime the measured signal strenght is below the configured limit (see paramter $Limit), 
    then you have to specify this parameter with the value Yes.
    Please note that this will pause your script from running until you click on OK. 
    The default value is No.

    .PARAMETER <$Limit>
    The wonderfull $Limit parameter is not just here for fun. No, it is here so you can specify a procentage limit. If the signal strenght is below the 
    specified limit, then you can choose to get a warning in a popup prompt (see parameter $PopUpWarning) or you can choose to get a verbal warning (see $Talk_Droid_All_Mute).
    
    .Parameter <$Talk_Droid_All_Mute>
    This parameter value is by default Yes, so it is muted. If you give the parameter the value No, then it will be able to talk to you, speaking about the signal 
    strenght. Please note that the script can´t hear you, so you cannot have a conversation with the script.

    .Parameter <$Talk_Droid_Warning_Only>
    If the parameter $Talk_Droid_All_Mute configured to be muted, then you can configure the $Talk_Droid_Warning_Only with the value Yes, which limits the talk Droid only to 
    speak when the signal strength is below the specified signal strength value in the parameter $Limit, so you are warned about bad signal strength.

    .EXAMPLE
    & '.\WI-FI Signal Strenght.ps1' -Delay 8 -PopUpWarning No -CreateLog Yes

    .NOTES
    Script created by Gunnar "Gundaris" Hermansen as a private Hoppy project. 
    Twitter: @Gundaris
    This is version 42.01.4 (not that i have made 42 versions, i just like the number)

    
    .LINK
    Oh no .... There are no related links :-(
    If this help is not good enough, then You are on your own. You could try with voodoo or magic. 
    Please let me know if Voodoo or Magic works and include a step by step guide for me. Thanks.

    Ups... forgot this:     Find new releases on https://github.com/Gundaris/Wifi-SignalStrength.git
    
#>


Param (
    [Int]$ProcentageLimit = 50,
    [Int]$Delay = 3,
    [String]$Highest,
    [String]$lowest,
    [DateTime]$LowTime,
    [String]$BSSID,
    [String]$BSSID2,
    [String]$PopUpWarning = 'No',
    [String]$CreateLog = 'Yes',
    [String]$ShowJob = 'No',
    [String]$Talk_Droid_All_Mute = 'Yes',
    [String]$Talk_Droid_Warning_Only = 'Yes',
    [String]$Date = (Get-Date -Format ddMMyyyy)
)

Add-Type -AssemblyName System.speech
$Speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

function Talk_droid ($Text){
    $Speak.Speak($Text) 
}

Function WriteData ($Data){
    $Filename = $Date +".txt"
    $Path = [environment]::getfolderpath("mydocuments")+"\"+"Wifi Logs\"
    If(!(Test-path $path)){
        New-Item -ItemType Directory -Force -Path $path
    }
    $Logfile = $Path+$Filename
    if(!(Test-path $Logfile)) {
        New-item -ItemType File -Path $Logfile
    }
    $Data | Add-Content -Path $Logfile
}

$BSSID2 = ((netsh wlan show interfaces) -Match 'BSSID' -Replace '^\s+BSSID\s+:\s+','')

While ($Date -eq (Get-Date -Format ddMMyyyy)){
    $Procentage = ((netsh wlan show interfaces) -Match '^\s+Signal' -Replace '^\s+Signal\s+:\s+','')
    $BSSID = ((netsh wlan show interfaces) -Match 'BSSID' -Replace '^\s+BSSID\s+:\s+','')
    $Data = 'Signal Strength '+$Procentage+$("("+(Get-Date)) +"). Accesspoint BSSID: "+$BSSID+". "

    If([String]::IsNullOrEmpty($BSSID)){
        If([String]::IsNullOrEmpty(($Procentage -replace "%",''))){
            $Data = "I could not find WIFI. "+$("("+(Get-Date)) +"). Are you sure you are connected?"    
            If ($CreateLog = 'Yes'){
                WriteData($Data)
            }
            $Data
        }
    }
    Else{
        If($BSSID -eq $BSSID2){
            If(($Procentage -replace "%","") -lt $lowest){
                $lowest = $Procentage
                $LowTime = Get-Date
            }
            If(($Procentage -replace "%","") -gt $Highest){
                $Highest = $Procentage
            }
            If([String]::IsNullOrEmpty(($Lowest))){
                    $lowest = $Procentage
                    $LowTime = Get-Date
            }
        $APSwitched=""
        }
        Else{
            $Highest = $Procentage
            $Lowest  = $Highest
            $LowTime = Get-date
            $APSwitched = "- AP Switched"
        }
        $BSSID2 = $BSSID
        $Data = $Data +'Best: '+$Highest+'and Worst:'+$lowest+"measured at "+$LowTime+" "+$APSwitched
        If (($Procentage -replace "%","") -le $ProcentageLimit){
            If ($CreateLog = 'Yes'){
                $Data = $Data +'- Under limit.'
                WriteData($Data)
                If ($PopUpWarning -eq 'Yes'){
                    $JobWarning = {
                        param($Data)
                        Add-Type -AssemblyName PresentationFramework
                        ([System.Windows.MessageBox]::Show($Data)).TopMost
                    }
                    $JobID = (Start-Job -ScriptBlock $JobWarning -ArgumentList $Data).Id
                }
            }
            If($ShowJob -eq 'Yes'){
                Get-Job    
            }
            If ($PopUpWarning -eq 'Yes'){
                Receive-Job -wait -AutoRemoveJob -ID $JobID
            }
            $Data
            If($Talk_Droid_All_Mute -eq 'No'){
                If($Talk_Droid_Warning_Only -eq 'Yes'){
                    Talk_Droid($Data)
                }
            }
        }
        Else{
            If ($CreateLog = 'Yes'){
                WriteData($Data)
            }
            $Data
        }
        If($Talk_Droid_All_Mute -eq 'No'){
            If($Talk_Droid_Warning_Only -eq 'No'){
                Talk_Droid($Data)
            }
        }
    }
    Start-Sleep -Seconds $Delay     
    #Cleaning up jobs
    get-job | remove-job
}
