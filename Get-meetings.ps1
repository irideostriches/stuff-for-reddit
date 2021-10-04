
#========================================================================
# Created on  : 04/09/2019
# Created by  : 
# Filename    : Get-Meetings.ps1
# Version     : 1.1
# Parameters  : 
# Usage       : This script was created to get all the skype meetings in
#               a user's outlook calendar. It then generates a report
#               for all those meetings, and a separate report with allget-
#               the meetings that have more than 250 recipients. This 
#               script uses Microsoft Graph to pull this information. 
# 
# Other Notes : 
#========================================================================
# Change Log  :
#
# 1.1 
# - Added functionality to get the meeting end date/time
# - Added new function to calculate meeting length
# - Added new export to export the number of meetings that a user has, as well as the 
#   sum of the length of all of their meetings
#
# 1.0
# - Initial script creation
##========================================================================

$ImportPath = import-csv "C:\Temp\List_of_users.csv"
Import-Module "\\Fileserver\Modules\Logging"

#These are used in the URI to get the next 6 months worth of meetings from the user calendars starting with the current date$
$DatePlus180 = (get-date).AddDays(90)
$StartDate = get-date -Format "o" #Format them to ISO format for Graph
$EndDate = get-date -Format "o" -date $DatePlus180
#$EndDate = get-date -Format "o" -year 2021
$Date = CurrentDate

$Date = CurrentDate #This is used for logging

Start-Transcript -Path "\\fileserver\Logs\Get-Meetings-$Date.log" #Log is stored in this directory

Function Get-AuthToken($AppSecret) {
    $TenantName = "tenant.mail.onmicrosoft.com"
    Import-Module Azure
    $clientId = "00000000-0000-0000-0000-000000000000" #This is the graph API app that is in Azure AD
    $resourceAppIdURI = "https://graph.microsoft.com"
    $authority = "https://login.microsoftonline.com/$TenantName"
    $credentials = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList $clientId, $appSecret
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext"-ArgumentList $authority
    #$Credential = Get-Credential #This is used if you want to authenticate using credentials instead of the appSecret
    #$AADCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential" -ArgumentList $credential.UserName,$credential.Password
    $authResult = $authContext.AcquireToken($resourceAppIdURI, $credentials)
    return $authResult
}

Function Get-MeetingsOver250($Meetings) {
    $MeetingsOver250 = @()
    ForEach ($Meeting in $Meetings) {
        if ($Meeting.AttendeeCount -ge 250) {
            if ($MeetingsOver250 -ne $null) {
                if ($Meeting.Type -like "occurrence") {
                    foreach ($Instance in $MeetingsOver250) {
                        if ($Instance.Subject -eq $Meeting.Subject -and $Instance.Organizer -eq $Meeting.Organizer -and $Instance.AttendeeCount -eq $Meeting.AttendeeCount) {
                            Write-Host "INFO: Recurring meeting with subject of $($Instance.Subject) with more than 250 attendees has already been added to the list - ignoring duplicate instance"
                        } else {
                            Write-Host "INFO: Adding recurring meeting with subject of $($Instance.Subject) to the list - it doesn't exist in the list yet"
                            $MeetingsOver250 += $Meeting
                        }
                    } else {
                        $MeetingsOver250 += $Meeting
                    }
                } else {
                    $MeetingsOver250 += $Meeting
                }
            }
        }
        return $MeetingsOver250
    }
}
    Function Get-MeetingInfo($Meeting) {
        #Figure out some licensing minutes nonsense
        $MeetingEndDate = $Meeting.end.dateTime | get-date
        $MeetingStartDate = $Meeting.start.dateTime | get-date
        $MeetingDuration = (New-TimeSpan -Start $MeetingStartDate -End $MeetingEndDate).TotalMinutes
        $LicensingMinutes = ($MeetingDuration * $meeting.attendees.count)

        #Transfer meeting info from the variable to a custom PSObject that can be easily exported to a csv
        $MeetingInfo = New-Object -TypeName PSObject -Property @{
        OnCalendarOfUser= $meeting.OnCalendarOfUser
        Organizer= $meeting.organizer.emailAddress.address
        Subject= $meeting.subject
        AttendeeCount = $meeting.attendees.count
        Type = $meeting.type
        SkypeURL = $meeting.onlineMeetingUrl
        MeetingStart = $MeetingStartDate
        MeetingEnd = $MeetingEndDate
        MeetingDuration= $MeetingDuration
        LicensingMinutes= $LicensingMinutes
        }

        return $MeetingInfo
    }

    Function Get-Meetings($User) {
        Write-Host "INFO: Getting meetings for $($User.PrimarySMTPAddress)"
        #Call the Graph API and store the meetings from the calendar of the mailbox into a variable
        try {
            $BatchSize = 250
            $Uri = "https://graph.microsoft.com/v1.0/users/$($User.PrimarySMTPAddress)/calendar/calendarView?startDateTime=$StartDate&endDateTime=$EndDate&top=$BatchSize"
            $ObjCapture= @()
            $Meetings = @()
            $MeetingInfo = @()
            #This will keep going and dynamically change the Graph URI that's used to pull everything in the date range specified until it pulls it all
            #It pulls in batches of 100 by default, or whatever you change the batchsize variable to that's above
            do {
                $Objects = Invoke-RestMethod -Uri $uri –Headers $authHeader
                $FoundObjects = ($Objects.value).count
                write-host "URI" $uri " | Found:" $FoundObjects
                #------Batched URI Construction------#
                $Uri = $Objects.'@odata.nextlink'
                $ObjCapture = $ObjCapture + $users.value

                #This saves that current URI's objects to the Meetings array to be sorted through afterward
                foreach ($Object in $Objects.value) {
                    $Meetings += $Object
                }
        
            } until ($Uri -eq $null)

        } catch {
            Write-Host "ERROR: Unable to pull meetings for $($User.PrimarySMTPAddress) - please double check the email address" -ForegroundColor Red
            return $null
        }

        #This sorts through each meeting gathered from the graph URIs and drops any event that isn't a skype meeting
        foreach ($Meeting in $Meetings) {
            if ($Meeting.onlineMeetingUrl.length -eq 0) {
                Write-Host "INFO: Meeting with subject of '$($Meeting.subject)' isn't a skype meeting - dropping from list"
            } else {
                $Meeting | Add-Member -NotePropertyName 'OnCalendarOfUser' -NotePropertyValue "$($user.PrimarySMTPAddress)" -Force
                $MeetingInfo += Get-MeetingInfo -Meeting $Meeting        
            }
        }
        #Return object to an array list
        return $MeetingInfo
        #return $meetings #This is used to get all of the meeting object's content before it's sorted using the foreach loop above. Use this if you want to see what other data is pulled
                        #From Graph before it's sorted
    }

    Function Get-TotalMeetings ($Meetings, $Users) {
        $ListOfTotals = @()

        ForEach ($user in $users.PrimarySMTPAddress) {
            Write-Host "INFO: Getting meeting totals for $user"
            $amount = ($Meetings.organizer | where-object {$_ -imatch $user}).count

            $totalduration = 0
            $totallicensing = 0
            ForEach ($Meeting in $Meetings) {
                if ($user -eq $meeting.organizer) {
                    $totalduration += $meeting.MeetingDuration
                    $totallicensing += $meeting.LicensingMinutes
                }
            }
            $container = new-object -typename psobject -property @{
                User= $user
                MeetingCount= $amount
                TotalDurationOfMeetings= $totalduration
                TotalLicensingMinutesNeeded= $totallicensing
            }
            $ListOfTotals += $container
        }
        return $ListOfTotals
    }

    #Get a token to authenticate with
    $appSecret = Read-Host -AsSecureString "Enter App Secret" #This is the secret for the App
    $token = Get-AuthToken -AppSecret $appSecret

    #Build an authentication header using JSON to authenticate with
    $authHeader = @{
        'Content-Type'='application\json'
        'Authorization'=$token.CreateAuthorizationHeader()
    }

    #Import the csv with users that you want to run this script against. Uses the PrimarySMTPAddress field of a csv
    $Users = $ImportPath

    #Create array to store sorted meetings in
    $MasterMeetingObject = @()
    #This is a timer to renew the authentication token every 25 URI pulls to make sure it doesn't expire before the script completes
    $TokenTimer = 0

    #Run the script against all the user's in the csv
    ForEach ($User in $users) {
        $MasterMeetingObject += Get-Meetings -User $User
        $TokenTimer++

        if ($TokenTimer -eq 25) {
            $token = Get-AuthToken -AppSecret $appSecret
            $TokenTimer = 0
        }
    }

    #Get all the meetings that have more than 250 attendees from the master array of meetings
    $MeetingsOver250 = Get-MeetingsOver250 -Meetings $MasterMeetingObject
    $TotalMeetingsPerUser = Get-TotalMeetings -Meetings $MasterMeetingObject -Users $Users

    #Export arrays to csvs
    $MeetingsOver250 | Export-Csv "\\fileserver\Reports\MeetingsOver250-$Date.csv" -NoTypeInformation
    $MasterMeetingObject | Export-Csv "\\fileserver\Reports\AllMeetings-$Date.csv" -NoTypeInformation
    $TotalMeetingsPerUser |  Export-Csv "\\fileserver\Reports\TotalMeetingsPerUser-$Date.csv" -NoTypeInformation
    . "\\fileserver\Count-Meetings.ps1" -MasterMeetingObject $MasterMeetingObject -Date $Date
    Stop-Transcript
