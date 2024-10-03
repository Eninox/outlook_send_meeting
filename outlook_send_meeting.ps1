Import-Module ActiveDirectory

# Fonction d'envoi d'invitation polyvalente
Function Send-Meeting() {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][String]$Subject,
        [Parameter(Mandatory=$true)][DateTime]$MeetingStart,
        [Parameter()][String]$Recipients,
        [Parameter()][String]$Categories,
        [Parameter()][String]$Location,
        [Parameter()][String]$Body=$Subject,
        [Parameter()][int]$ReminderMinutesBeforeStart=15,
        [Parameter()][int]$DurationInMinutes=30
    )

    $ol = New-Object -ComObject Outlook.Application
    $meeting = $ol.CreateItem('olAppointmentItem')

    if ($ReminderMinutesBeforeStart -gt 0) {
        $meeting.ReminderSet = $true
        $meeting.ReminderMinutesBeforeStart = $ReminderMinutesBeforeStart
    }

    if ($PSBoundParameters.ContainsKey('Recipients')) {
        foreach ($email in $Recipients -split ";") {
            if ($email -ne '') {
                $meeting.Recipients.Add($email)
            }
        }
        $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
    } else {
        $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olNonMeeting
    }

    if ($PSBoundParameters.ContainsKey('Location')) {
        $meeting.Location = $Location
    }

    if ($PSBoundParameters.ContainsKey('Categories')) {
        $meeting.Categories = $Categories
    }

    $meeting.Subject = $Subject
    $meeting.Body = $Body
    $meeting.Start = $MeetingStart
    $meeting.Duration = $DurationInMinutes
    $meeting.Save()
    $meeting.Send()
}

# Fichiers avec données variables et logs invitation
$sendFile = "chemin\du\fichier\envoi_invitations_astreintes.txt"
$logFile = "chemin\du\fichier\suivi_invitations_envoyees.txt"

$logDate = Get-Date -Format "dddd dd/MM/yyyy HH:mm"
$logSeparation = "*******************************************************************************************"
$logStart = "`r$logSeparation`rLancement du script pour invitations astreinte en masse le $logDate`r$logSeparation"
Add-Content -Path $logFile -Value $logStart -PassThru

# Durées de meeting en minutes
$durationWeekend = 2280
$durationNuit = 840
$durationFerie = 1439
$durationFerieWeekend = 4320

$members = Import-Csv -Delimiter ";" -Path $sendFile -Encoding UTF8

# Boucle composition des données variables (identité, type astreinte, debut d'invitation...) et envoi unitaire d'invitation avec conditions et logs
foreach ($member in $members) {
    $dateMeeting = $member.date
    $dayNumber = [Int] (Get-Date $dateMeeting).DayOfWeek
    $meetingStart1 = (Get-Date $dateMeeting -Format "MM/dd/yyyy 18:00:00").ToString()
    $meetingStart2 = (Get-Date $dateMeeting -Format "MM/dd/yyyy 08:00:00").ToString()
    $type = $member.type
    $firstName = $member.prenom
    $recipient = $member.mail
    $userAd = Get-ADUser -Filter {EmailAddress -eq $recipient} -Properties mail
    $userAdMail = $userAD.mail
    
    try {

        if ($userAdMail -eq $recipient) {

            if ($type -eq "WEEKEND" -and $dayNumber -eq 5) {
                $hideOutput = Send-Meeting -Subject "Astreinte de $type $firstName" -MeetingStart $meetingStart1 -DurationInMinutes $durationWeekend -Recipients $recipient
            } elseif ($type -eq "NUIT") {
                $hideOutput = Send-Meeting -Subject "Astreinte de $type $firstName" -MeetingStart $meetingStart1 -DurationInMinutes $durationNuit -Recipients $recipient
            } elseif ($type -eq "FERIE" -and $dayNumber -eq 5) {
                $hideOutput = Send-Meeting -Subject "Astreinte de $type + WEEKEND $firstName" -MeetingStart $meetingStart2 -DurationInMinutes $durationFerieWeekend -Recipients $recipient 
            } elseif ($type -eq "FERIE") {
                $hideOutput = Send-Meeting -Subject "Astreinte de $type $firstName" -MeetingStart $meetingStart2 -DurationInMinutes $durationFerie -Recipients $recipient
            }

            $logSuccess = "Invitation astreinte de $type transmise a $firstname pour le $dateMeeting"
            Add-Content -Path $logFile -Value $logSuccess -PassThru

        } else {
            $logError = "Erreur survenue invitation astreinte $type de $firstname le $dateMeeting, verifier calendrier"
            Add-Content -Path $logFile -Value $logError -PassThru
        }

    } catch {
        Add-Content -Path $logFile -Value $logError -PassThru
    }

    Start-Sleep -Seconds 2
}

# Pause