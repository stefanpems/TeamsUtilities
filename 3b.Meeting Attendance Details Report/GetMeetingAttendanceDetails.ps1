<#
  .VERSION AND AUTHOR
    Script version: v-2020.05.14-2 
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
    This script downloads the details of the users attendance to all the Teams meetings planned in the 
    organizational's Office 365 groups/teams. 
    In particular, this script correlates all the user attendance logs to all the meetings 
    in the specified period with the single teams in whose calendars these meetings were planned. 
    The script separately returns also le list of the user attendance logs in meetings for which it 
    is not possible to identify a reference team (no calendar appointment found).
    This script may be useful where it is necessary to correlate user attendance logs with the teams
    where the meetings were planned (e.g., in schools, where it is necessary to make meeting attendance 
    statistics by Classes/Subject)
    
    IMPORTANT: The main source of data for this report is the Teams Call Analytics / Call Quality Dashboard service. 
    This data source can the base for very interesting shapes on the work of the different teams - possibly 
    to identify which teachers or students may need help in the adoption - but it can’t be guaranteed as 
    certainly complete. There may be situations, though rare, in which an end-user's Teams client app cannot 
    report its telemetry data to Teams Call Analytics. Because of this evidence, the reports based on this 
    data source should never be used as a certification or proof for anything official like student absence 
    or teachers’ activity tracking, funds or other relevant benefit assignments, etc….

  .RESULTS OF THE EXECUTION
    1) [MAIN RESULT] --> File: "AttendanceToPlannedMeetings_Results_<execution-date-time>.csv"
        CSV file containing the details of the users attendance to the meetings found in the 
        groups/teams calendars. Header row:
        "UserLoginName","StartDateTime","EndDateTime","TeamEmail","TeamDisplayName","TeamVisibility"
    2) File: "AttendanceToUnknownMeetings_Results_<execution-date-time>.csv"
        CSV file containing the details of the users attendance to the meetings NOT found in the groups/teams 
        calendars. Header row:
        "UserLoginName","StartDateTime","EndDateTime","MeetingId"
    3) File: "QueryGroupCalendars_Meetings_<execution-date-time>.csv"
        CSV file containing the list of the meetings found in the Office 365 groups/teams calendars.
        Header row:
        "calendarMail","calendarDisplayName","calendarVisibility","Day","StartHour","EndHour","OrganizerEmail","MeetingId"
    4) File: "QueryGroupCalendars_Unaccessed_<execution-date-time>.csv"
        CSV file containing the list of the Office 365 groups/teams whose calendar were NOT accessible to the current 
        user. Header row:
        "calendarMail","calendarDisplayName","calendarVisibility","Exception"
    5) File: "QueryCQD_Attendance_<execution-date-time>.csv"
        CSV file containing the raw list of the attendance to the conferences donwloaded from Teams CQD.
        Header row:
        "First UPN","Second UPN","Meeting Id","Start Time","End Time","Total Stream Count"
    6) File: "GetMeetingAttendanceDetails_<execution-date-time>.log"
        Log text file with the details of the execution (apart from the progress details of phase 3, all the output 
        messages shown during the execution are also recorded in this log file).

  .EXECUTION DETAILS
    The script's execution has 3 main actions (phases); each phase can be skipped in needed.
    * [PHASE 1] Retrieve the list of meetings scheduled in the Office 365 groups/teams
    * [PHASE 2] Retrieve the list of users' attendance to the conferences held by Teams
    * [PHASE 3] Correlate the two lists above and produce the final 2 output CSV files (1 & 2 in the list above)

  .CONSTRAINTS
    * In order to read the meetings in a group/teams calendar (PHASE 1) the PnP.PowerShell application must have 
      permissions to access groups' calendars in Office 365. This permission can be grant if the script is run 
      by a tenant admin (a permission grant pop-up appears).
    * The match is possible only with the meetings saved in the group/teams calendars. 
      The users' attendance (start/end time) to the Teams meeting created as "Immediate meetings" within a Teams 
      channel cannot be matched with the Team/channel because there is no correlation information available.
      These unmatched attendance details are part of the CSV file #2 in the list above.
      
  .RECOMMENDATIONS
    If needed, each of the 3 phases can be skipped by setting appropiately the documented script variables. 
    For those who have Microsoft Access or other RDBMS programs available, it is HIGHLY advised to skip PHASE 3 and to 
    execute the correlation of meetings and attendance (CSV #3 and #5 in the list of results) within such a program. 
    The PHASE 3 in PowerShell is EXTREMELY slow: it can take different hours for parsing the meetings of the last 30 days.
    In a RBMS, it's a matter of a few seconds by using a simple SQL JOIN statement.
    A sample matching file is provided in MS Access together with this script file (Alternative-Phase3.accdb).

  .PREREQUISITES
    1) PowerShell module SharePointPnPPowerShellOnline
        (https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps)
    2) PowerShell module CQDPowerShell 
        (https://www.powershellgallery.com/packages/CQDPowerShell/2.0.0)
    3) A first administrative access must have been done to CQD for its provisining: https://cqd.teams.microsoft.com/
    4) The Internet Explorer first launch configuration wizard must have been completed (otherwise the execution of 
        CQDPowerShell returns the error: "Invoke-WebRequest : The response content cannot be parsed").
    5) At the very first launch, you need to approve the pop-up window with the request to authorize the app PnP.PowerShell
        to read calendars, users' base profile and data for which the current user is authorized.
    6) The credentials specified in the first authentication request (used for the Graph API queries to all the calendars
       existing in Exchange Online) must have at least read access to the calendars of the groups' (Teams) that need to 
       be accessed.
    7) The credentials specified in the secon authentication request (used for running CQDPowerShell) must have at least 
       read access to the data managed by Teams Call Analytics / Call Quality Dashboard.
     
         
        To install the PowerShell modules 1) and 2), open PowerShell as administrator and type:
                Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
                Install-Module -Name SharePointPnPPowerShellOnline
                Install-Module -Name CQDPowerShell

  .VARIABLES TO BE SET
    $daysBack : number of days before today for executing the queries on the groups calendar
    $outDirPath: path of the local folder where the script generates the output files
    $matchingDomainName: DNS name of the school. Only upn with that extension are considered in the CQD report. 
    $addHoursFromUTC: number of hours to be added (or subtracted) to the UTC time (from the CQD report)
    $existingMeetingsCsvFileName: name of the possible existing CSV file with the list of the group meetings 
                already generated from a previous execution of the script and to be found in $outDirPath.
                Leave empty if the script must query the Office 365 groups/teams calendars.
                If this parameter is populated, the script will skip this query.
    $existingAttendanceCsvFileName: name of the possible existing CSV file with the details of the users'  
                conf. attendance already generated from a previous execution of the script and to be 
                found in $outDirPath. Leave empty if the script must query the Office 365 groups/teams calendars.
                If this parameter is populated, the script will skip this query.
    $skipPhase3: set it to $true if you can correlate team meetings and attendance outside of PowerShell (for
                example, with the MS Access example delivered together with this script).
    $inPhase3RemoveDuplicates: set it to $false if you can remove duplicates outside of PowerShell (for
                example, in MS Access, by using a SELECT DISTINCT statement).
                NOTE: it takes almost 2 hours on Intel I7 to remove duplicates on a 70K CSV file (about 200K lines) 

  .DISCLAIMER 
    The script is given "as-is", without any explicit or implicit warranties, obligations or conditions.
    This is NOT an official content delivered by Microsoft: it is just a script produced and shared voluntarly 
    by the author and never officially reviewed by none.
    Please double check the correctness of the results: I have a poor knowledge of the Teams' CQD concepts so, 
    while parsing the results of the CQDPowerShell query, I may have done wrong assumptions/operations.

  .TEST PLAN
    Tests that still need to be executed:
        1) For a specific meeting planned in a channel (then found in the calendar and reported by this 
        script) are the users possibly added to the meeting reported in the main CSV result of this 
        script? This should be checked in both cases if they or they do  not belong to the team.

  .NOTES AND UPDATES
    Preable: I'm not a code developer; please do not expect elegant code... ;-)    
    2020.04.19 - I'm correlating the Teams membership (downloaded from Azure AD) with the report of this script
        in order to see which team members (students) missed a meeting (lesson). I'll give updates on this.
    2020.04.19 - For the "istant meetings" created in a channel (not found in the calendar), is it possible to identify  
        the classes to which they are related by comparing the users attendance (got from CQD) with the 
        classes team membership (got from AAD)? 
    2020.04.21 - Phase 3, if done by this script (instead of outside) may cause an inappropriate parsing of the dates 
        and time of the occurred meetings as retrieved by CQDPowerShell. The issue is most probably due to the missing 
        use of locale information. I'll fix it ASAP. Please check carfully the correctness of the dates reported
        by phase 3
#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################

$daysBack = 30 
$outDirPath = "C:\Temp\OUT"
$matchingDomainName = "schoolname.edu.it"
$addHoursFromUTC = 2
$existingMeetingsCsvFileName = "" #Leave empty if the script must generate it; fill it to skip Phase 1
$existingAttendanceCsvFileName = "" #Leave empty if the script must generate it; fill it to skip Phase 2
$skipPhase3 = $false #Recommended: set to $true and use MS Access or SQL or other RDBMS to join meetings and attendance.
$inPhase3RemoveDuplicates = $true  #Recommended: set to $false and use MS Access or SQL or other RDBMS to SELECT DISTINCT

#########################################################################################################################
# FUNCTIONS:
#########################################################################################################################

function QueryGraph{

    param(
         [DateTime]
         $startDate,
         [DateTime]
         $endDate,
         [string]
         $accessToken
    )

    $GraphURL = "https://graph.microsoft.com/beta" 

    try
    {        
        $sDate = $startDate.ToString("yyyy-MM-dd")
        $eDate = $endDate.ToString("yyyy-MM-dd")

        $graphGroupReqUrl = $GraphURL+"/groups?$filter=groupTypes/any(c:c+eq+'Unified')&select=id,mail,displayName,visibility"

        $groupTotCount = 0;
        $groupQueryExecutionCount = 0;
        $accessedGroupTotCount = 0;
        $unaccessedGroupTotCount = 0;

        do{
            $groupQueryExecutionCount++;
            $hds = @{
                Authorization = "Bearer $accesstoken"
            }

            $groupResponse = Invoke-RestMethod -Uri $graphGroupReqUrl -Headers $hds -Method Get 

            $groupResultInQueryCount = 0;
            if($groupResponse){
                $graphGroupReqUrl = $groupResponse.'@odata.nextLink'
                $groupResponse.value | ForEach-Object{  
                    $groupTotCount++;
                    $groupResultInQueryCount++;
                        
                    $groupId = $_.id
                    $groupMail = $_.mail
                    $groupDN = $_.displayName
                    $groupVisibility = $_.visibility
                        
                    $info = "["+$groupTotCount+" ("+$accessedGroupTotCount+", "+$unaccessedGroupTotCount+") - "+$groupQueryExecutionCount+" - "+$groupResultInQueryCount+"] Found group "+$groupMail+" ("+$groupDN+") - "+$groupVisibility
                    Write-Host $info -ForegroundColor Cyan; $info | Out-File $outLogFilePath -Append
                        
                    $foundMeetingsCount = 0
                    $graphCalendarReqUrl = $GraphURL+"/groups/$groupId/calendarview?$top=50&startdatetime="+$sDate+"T00:00:00Z&enddatetime="+$eDate+"T23:59:00Z"            
                    $foundMeetingsCount = DownloadCalendarEventsToCsv -currentTotMeetingCount $meetingsTotCount -calendarMail $groupMail -calendarDisplayName $groupDN -calendarVisibility $groupVisibility -graphCalendarReqUrl $graphCalendarReqUrl -accessToken $accessToken                        

                    if($foundMeetingsCount -ge 0){
                        $meetingsTotCount += $foundMeetingsCount;
                        $accessedGroupTotCount++;
                    }
                    else{
                        $unaccessedGroupTotCount++;
                    }
                }
            }  
            else{
                $graphGroupReqUrl = ""
                $info = "["+$groupTotCount+" ("+$accessedGroupTotCount+", "+$unaccessedGroupTotCount+") - "+$groupQueryExecutionCount+"] Empty response for query "+$graphCalendarReqUrl
                Write-Host $info -ForegroundColor Red; $info | Out-File $outLogFilePath -Append
            }
        }
        while( -not([string]::IsNullOrEmpty($graphGroupReqUrl)) -and ($groupQueryExecutionCount -le 200) )

                 
    }
    catch
    {
        $info = "Error while executing query - "+$_.Exception.Message 
        Write-Host $info -ForegroundColor DarkRed -BackgroundColor Yellow; $info | Out-File $outLogFilePath -Append
        throw $_
    }

    $res = @{
        groupTotCount = $groupTotCount
        accessedGroupTotCount = $accessedGroupTotCount
        unaccessedGroupTotCount = $unaccessedGroupTotCount
        meetingsTotCount = $meetingsTotCount
    }
    $res
}

function DownloadCalendarEventsToCsv{

    param(
         [int]
         $currentTotMeetingCount,
         [string]
         $calendarMail,
         [string]
         $calendarDisplayName,
         [string]
         $calendarVisibility,
         [string]
         $graphCalendarReqUrl,
         [string]
         $accessToken
    )

    $GraphURL = "https://graph.microsoft.com/beta" 

    try
    {
        $meetingInFunctionCount=0;              
        $calendarQueryExecutionCount = 0;
        
        do{
            $calendarQueryExecutionCount++;
            $hds = @{
                Authorization = "Bearer $accesstoken"
                Prefer = 'outlook.timezone="Europe/Berlin"'
            }

            $calendarResponse = Invoke-RestMethod -Uri $graphCalendarReqUrl -Headers $hds -Method Get 
            
            $calendarResultInQueryCount = 0;
            if($calendarResponse){
                $graphCalendarReqUrl = $calendarResponse.'@odata.nextLink'
                $calendarResponse.value | ForEach-Object{   
                    $currentTotMeetingCount++;   
                    $meetingInFunctionCount++;              
                    $calendarResultInQueryCount++;  
                    $omProvider = $_.onlineMeetingProvider

                    if($omProvider = "teamsForBusiness"){                                                                               
                        [datetime]$meetingSdt = $_.start.dateTime
                        [datetime]$meetingEdt = $_.end.dateTime
                        $teamId = $_.organizer.emailAddress.address 
                        $joinUrl = $_.onlineMeeting.joinUrl

                        if($joinUrl){
                            $meetId = ""; $confId = ""
                            $aSplUrl = ($joinUrl -split "%22")
                            if($aSplUrl){
                                $meetId = ($aSplUrl)[0].replace("https://teams.microsoft.com/l/meetup-join/","").replace("%3a",":").replace("%40","@")
                                $meetId = ($meetId.split('/'))[0]
                            }                    
                        }
                        $info = "   [" + $currentTotMeetingCount + " - " + $calendarQueryExecutionCount + " - " + $calendarResultInQueryCount + "] "+$meetingSdt.Date.ToShortDateString()+" - "+$meetingSdt.AddMinutes(30).Hour+" - "+$meetingEdt.AddMinutes(30).Hour+" - "+$teamId+" - "+$meetId
                        Write-Host $info -ForegroundColor Green; $info | Out-File $outLogFilePath -Append
                        '"'+$calendarMail+'","'+$calendarDisplayName+'","'+$calendarVisibility+'","'+$meetingSdt.Date.ToShortDateString()+'","'+$meetingSdt.AddMinutes(30).Hour+'","'+$meetingEdt.AddMinutes(30).Hour+'","'+$teamId+'","'+$meetId+'"' | Out-File $outMeetingsCsvFilePath -Append

                        if($meetingSdt.Date -ne $meetingEdt.Date){
                            $info = "   [" + $currentTotMeetingCount + " - " + $calendarQueryExecutionCount + " - " + $calendarResultInQueryCount + "] NOTE: this is a multi-day event " + $meetingSdt.Date.ToShortDateString() + " - " + $meetingSdt.AddMinutes(30).Hour + " - " + $meetingEdt.Date.ToShortDateString() + " - " + $meetingEdt.AddMinutes(30).Hour + " - " + $teamId 
                            Write-Host $info -ForegroundColor Gray; $info | Out-File $outLogFilePath -Append
                        }
                     
                    }
                    else{
                        $info = "   [" + $currentTotMeetingCount + " - " + $calendarQueryExecutionCount + " - " + $calendarResultInQueryCount + "] Skipped meeting with provider "+$omProvider
                        Write-Host $info -ForegroundColor Magenta; $info | Out-File $outLogFilePath -Append
                    }
                }
            }  
            else{
                $graphCalendarReqUrl = ""
                $info = "   [" + $currentTotMeetingCount + " - " + $calendarQueryExecutionCount + "] Empty response for query "+$graphCalendarReqUrl
                Write-Host $info -ForegroundColor Red; $info | Out-File $outLogFilePath -Append
            }
        }
        while( -not([string]::IsNullOrEmpty($graphCalendarReqUrl)) -and ($calendarQueryExecutionCount -le 200) )
            
    }
    catch
    {
        $meetingInFunctionCount = -1;
        
        if($_.Exception.Message.Contains("(403) Forbidden.")){
            $info = "   Skipping the group - Reason: access denied..."  
            Write-Host $info -ForegroundColor Gray; $info | Out-File $outLogFilePath -Append            
        }
        else{ #Totally unexpected condition
            $info = "   Skipping the group - Reason: error while executing query "+$graphCalendarReqUrl+" - Exception: "+$_.Exception.Message 
            Write-Host $info -ForegroundColor DarkRed -BackgroundColor Yellow; $info | Out-File $outLogFilePath -Append
        }
        '"'+$calendarMail+'","'+$calendarDisplayName+'","'+$calendarVisibility+'","'+$_.Exception.Message.replace('"','#')+'"' | Out-File $outUnaccessedCsvFilePath -Append
        #DO NOT throw $_
    }

    $meetingInFunctionCount;
}

function JoinMeetingsAndAttendanceDetails{

    param(
         [string]
         $meetingsCsvFileFullPath,
         [string]
         $attendanceCsvFileFullPath,
         [string]
         $resultsMatchedCsvFileFullPath,
         [string]
         $resultsUnmatchedCsvFileFullPath
    )

    '"UserLoginName","StartDateTime","EndDateTime","TeamEmail","TeamDisplayName","TeamVisibility"' | Out-File $resultsMatchedCsvFileFullPath
    '"UserLoginName","StartDateTime","EndDateTime","MeetingId"' | Out-File $resultsUnmatchedCsvFileFullPath

    $outerCount = 0;
    $skippedMeetings = 0;
    $matchedMeetings = 0;
    $unmatchedMeetings = 0;
    $fullAttendanceCsvFileFullPath = Import-Csv $attendanceCsvFileFullPath | where 'Meeting Id' -ne '' | where 'Start Time' -ne '' | where 'End Time' -ne ''
    $outerTotCount = $fullAttendanceCsvFileFullPath.Count
    $fullAttendanceCsvFileFullPath | ForEach-Object{
        $outerCount++; 
        $fUpn = $_.'First UPN'
        $sUpn = $_.'Second UPN'
        $MId = $_.'Meeting Id'
        $sDt = $_.'Start Time'
        $eDt = $_.'End Time'
        $tSc = $_.'Total Stream Count'

        $skipRow = $false
        $upn = ""

        if([string]::IsNullOrEmpty($sDt)){
            $skipRow = $true
        }

        if([string]::IsNullOrEmpty($eDt)){
            $skipRow = $true
        }

        if(-not($skipRow)){
            if($sUpn.EndsWith($matchingDomainName)){
                $upn = $sUpn #Priority is given to the second Upn        
            }
            else{
                if($fUpn.EndsWith($matchingDomainName)){
                    $upn = $fUpn
                }
                else{
                    $skipRow = $true
                }
            }
        }

        if(-not($skipRow)){
            $parsedStartDateTime = [datetime]::Parse($sDt).AddHours($addHoursFromUTC)
            $parsedEndDateTime = [datetime]::Parse($eDt).AddHours($addHoursFromUTC)
            $meetingFound = Import-Csv -Path $meetingsCsvFileFullPath | where 'MeetingId' -eq $MId 

            if($meetingFound){
                $matchedMeetings++;
                Write-Host "[" $outerCount "/" $outerTotCount "] Matched meeting " -ForegroundColor Green
                $meetingFound | ForEach-Object{
                    $tem = $_.calendarMail
                    $tdn = $_.calendarDisplayName
                    $tvs = $_.calendarVisibility
                    '"'+$upn+'","'+$parsedStartDateTime.ToString("dd/MM/yyy hh:mm:ss")+'","'+$parsedEndDateTime.ToString("dd/MM/yyy hh:mm:ss")+'","'+$tem+'","'+$tdn+'","'+$tvs+'"' | Out-File $resultsMatchedCsvFileFullPath -Append
                }
            }
            else{
                $unmatchedMeetings++;
                Write-Host "[" $outerCount "/" $outerTotCount "] Unmatched meeting " -ForegroundColor Yellow
                '"'+$upn+'","'+$sDt+'","'+$eDt+'","'+$MId+'"' | Out-File $resultsUnmatchedCsvFileFullPath -Append
            }
        }
        else{
            $skippedMeetings++;
            Write-Host "[" $outerCount "/" $outerTotCount "] Skipped" -ForegroundColor Gray
        }
    }
    
    #--> Remove duplicates 
    if($inPhase3RemoveDuplicates){
        $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
        $info = "Removing duplicates in CSV file. It may take hours: please wait..."; Write-Host $info -ForegroundColor Cyan; $info | Out-File $outLogFilePath -Append
        $tmp = Import-Csv $resultsMatchedCsvFileFullPath | Select-Object * -Unique 
        $tmp | Export-Csv $resultsMatchedCsvFileFullPath -Force -NoTypeInformation
        $info="Done" ; Write-Host $info -ForegroundColor Cyan; $info | Out-File $outLogFilePath -Append
        $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    }

    $res = @{
        skippedMeetings = $skippedMeetings
        unmatchedMeetings = $unmatchedMeetings
        matchedMeetings = $matchedMeetings
    }
    $res

}

#########################################################################################################################
# MAIN:
#########################################################################################################################

cls

#########: PREREQUISITES AND AUTHENTICATIONS
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outLogFilePath = "$outDirPath\GetMeetingAttendanceDetails_$LogStartTime.log"
If (Test-Path $outLogFilePath){
	Remove-Item $outLogFilePath
}
$info="Execution started - "+(Get-Date -Format "yyyy-MM-dd_hh.mm.ss") ; Write-Host $info; $info | Out-File $outLogFilePath -Append

$skipPhase1 = $false
if($existingMeetingsCsvFileName -ne ""){
    $outMeetingsCsvFilePath = "$outDirPath\$existingMeetingsCsvFileName"
    If (-not(Test-Path $outMeetingsCsvFilePath)){
	    throw [System.IO.FileNotFoundException] "$outMeetingsCsvFilePath not found."
    }
    $skipPhase1 = $true
}
else
{
    $info="Prerequisites - SharePointPnPPowerShellOnline - Module import..." ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue 

    $info="Prerequisites - SharePointPnPPowerShellOnline - Authentication..." ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $arrayOfScopes = @("Group.Read.All","Calendars.Read", "Directory.Read.All")
    Connect-PnPOnline -Scopes $arrayOfScopes
    $accessToken = Get-PnPAccessToken

    $outMeetingsCsvFilePath = "$outDirPath\QueryGroupCalendars_Meetings_$LogStartTime.csv"
    If (Test-Path $outMeetingsCsvFilePath){
	    Remove-Item $outMeetingsCsvFilePath
    }
}

$outUnaccessedCsvFilePath = "$outDirPath\QueryGroupCalendars_Unaccessed_$LogStartTime.csv"
If (Test-Path $outUnaccessedCsvFilePath){
	Remove-Item $outUnaccessedCsvFilePath
}

$skipPhase2 = $false
if($existingAttendanceCsvFileName -ne ""){
    $outAttendanceCsvFilePath = "$outDirPath\$existingAttendanceCsvFileName"
    If (-not(Test-Path $outAttendanceCsvFilePath)){
	    throw [System.IO.FileNotFoundException] "$outAttendanceCsvFilePath not found."
    }
    $skipPhase2 = $true
}
else{
    #Immediate fake execution just to obtain immediately the auth token
    $info="Prerequisites - CQDPowerShell - Module import..." ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    Import-Module CQDPowerShell 

    $info="Prerequisites - CQDPowerShell - Authentication..." ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $fakeDate = Get-Date;
    Get-CQDData -CQDVer "V3" -StartDate $fakeDate -EndDate $fakeDate -OutPutType DataTable -Dimensions "AllStreams.Start Time" -Measures "Measures.PSTN Total Attempts Count" | Out-Null

    $outAttendanceCsvFilePath = "$outDirPath\QueryCQD_Attendance_$LogStartTime.csv"
    If (Test-Path $outAttendanceCsvFilePath){
	    Remove-Item $outAttendanceCsvFilePath
    }
}

if(-not($skipPhase3)){
    $resultsMatchedFilePath = "$outDirPath\AttendanceToPlannedMeetings_Results_$LogStartTime.csv"
    If (Test-Path $resultsMatchedFilePath){
	    Remove-Item $resultsMatchedFilePath
    }

    $resultsUnmatchedFilePath = "$outDirPath\AttendanceToUnknownMeetings_Results_$LogStartTime.csv"
    If (Test-Path $resultsUnmatchedFilePath){
	    Remove-Item $resultsUnmatchedFilePath
    }
}

$info="Execution started - "+(Get-Date -Format "yyyy-MM-dd_hh.mm.ss") ; Write-Host $info; $info | Out-File $outLogFilePath -Append

$ed = Get-Date
$sd = $ed.AddDays(-$daysBack)

#########: PHASE 1
if(-not($skipPhase1)){
    $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info="PHASE 1: download meetings from group/teams calendar - Execution started - "+(Get-Date -Format "yyyy-MM-dd_hh.mm.ss")+" - Please wait..."  
    Write-Host $info; $info | Out-File $outLogFilePath
    $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append

    $info="#################################" ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    Write-Host "Format of the trace line prefix:" 
    $info = "[Tot. group count (Accessed group count, Unaccessed group count) - Group query execution count - Result in group query count]"
    Write-Host $info -ForegroundColor Cyan
    $info = "   [Tot. meetings found count - Calendar query execution count - Result in group query count] "
    Write-Host $info -ForegroundColor Green
    $info="#################################" ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append

    '"calendarMail","calendarDisplayName","calendarVisibility","Day","StartHour","EndHour","OrganizerEmail","MeetingId"' | Out-File $outMeetingsCsvFilePath
    '"calendarMail","calendarDisplayName","calendarVisibility","Exception"' | Out-File $outUnaccessedCsvFilePath

    $groupTotCount = 0;
    $meetingsTotCount = 0;
    
    $res = QueryGraph -startDate $sd -endDate $ed -accessToken $accessToken

    $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info="PHASE 1 - Execution ended - "+(Get-Date -Format "yyyy-MM-dd_hh.mm.ss")
    Write-Host $info; $info | Out-File $outLogFilePath -Append
    if($res){
        $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
        $info="   Found groups: "+$res.groupTotCount ; Write-Host $info; $info | Out-File $outLogFilePath -Append 
        $info="      Successfully accessed groups: "+$res.accessedGroupTotCount ; Write-Host $info; $info | Out-File $outLogFilePath -Append 
        $info="      Unaccessed groups: "+$res.unaccessedGroupTotCount ; Write-Host $info; $info | Out-File $outLogFilePath -Append 
        $info="   Found Meetings: "+$res.meetingsTotCount ; Write-Host $info; $info | Out-File $outLogFilePath -Append 
        $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    }
}
else{
    $info = "PHASE 1 Skipped" ; Write-Host $info; $info | Out-File $outLogFilePath -Append
}

#########: PHASE 2
if(-not($skipPhase2)){
    $info="PHASE 2: download conference attendance details - Execution started - "+(Get-Date -Format "yyyy-MM-dd_hh.mm.ss")+" - Please wait..."; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info="   Running: " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info='   Get-CQDData -CQDVer "V3" -StartDate "'+$sd+'" -EndDate "'+$ed+'" -IsServerPair "Client : Server" -Dimensions "AllStreams.First UPN","AllStreams.Second UPN","AllStreams.Meeting Id","AllStreams.Start Time","AllStreams.End Time" -Measures "Measures.Total Stream Count" -OutPutType "CSV" -OutPutFilePath "'+$outAttendanceCsvFilePath+'" -OverWriteOutput -LargeQuery'
    Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append

    Get-CQDData -CQDVer "V3" -StartDate $sd -EndDate $ed -IsServerPair 'Client : Server' -Dimensions "AllStreams.First UPN","AllStreams.Second UPN","AllStreams.Meeting Id","AllStreams.Start Time","AllStreams.End Time" -Measures "Measures.Total Stream Count" -OutPutType "CSV" -OutPutFilePath $outAttendanceCsvFilePath -OverWriteOutput -LargeQuery
    $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info="PHASE 2 - Execution ended - "+(Get-Date -Format "yyyy-MM-dd_hh.mm.ss"); Write-Host $info; $info | Out-File $outLogFilePath -Append
}
else{
    $info = "PHASE 2 Skipped" ; Write-Host $info; $info | Out-File $outLogFilePath -Append
}

#########: PHASE 3
if(-not($skipPhase3)){
    $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info="PHASE 3: join meeting and attendance details - Execution started - "+(Get-Date -Format "yyyy-MM-dd_hh.mm.ss")+" - Please wait..."; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append

    $resMatch = JoinMeetingsAndAttendanceDetails -meetingsCsvFileFullPath $outMeetingsCsvFilePath -attendanceCsvFileFullPath $outAttendanceCsvFilePath -resultsMatchedCsvFileFullPath $resultsMatchedFilePath -resultsUnmatchedCsvFileFullPath  $resultsUnmatchedFilePath
    $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
    $info="PHASE 3 - Execution ended - "+(Get-Date -Format "yyyy-MM-dd_hh.mm.ss")
    Write-Host $info; $info | Out-File $outLogFilePath -Append

    if($resMatch){
        $info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
        $info="   Matched Meetings: "+$resMatch.matchedMeetings ; Write-Host $info; $info | Out-File $outLogFilePath -Append 
        $info="   Unmatched Meetings: "+$resMatch.unmatchedMeetings ; Write-Host $info; $info | Out-File $outLogFilePath -Append 
        $info="   Skipped Meetings: "+$resMatch.skippedMeetings ; Write-Host $info; $info | Out-File $outLogFilePath -Append 
    }
}
else{
    $info = "PHASE 3 Skipped" ; Write-Host $info; $info | Out-File $outLogFilePath -Append
}

#########: FINALIZATIONS
$info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
$info="Execution completed - "+(Get-Date -Format "yyyy-MM-dd_hh.mm.ss") ; Write-Host $info; $info | Out-File $outLogFilePath -Append
$info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
$info="Output files generated :" ; Write-Host $info; $info | Out-File $outLogFilePath -Append
$info="   " ; Write-Host $info; $info | Out-File $outLogFilePath -Append
$info="   Verbose Log File:"+$outLogFilePath ; Write-Host $info; $info | Out-File $outLogFilePath -Append  

if(-not($skipPhase1)){
    $info="   Identified meetings CSV:"+$outMeetingsCsvFilePath ; Write-Host $info; $info | Out-File $outLogFilePath -Append  
    $info="   Unaccessed groups CSV:"+$outUnaccessedCsvFilePath ; Write-Host $info; $info | Out-File $outLogFilePath -Append   
}
if(-not($skipPhase2)){
    $info="   Attendance to conference CSV:"+$outAttendanceCsvFilePath ; Write-Host $info; $info | Out-File $outLogFilePath -Append  
}
if(-not($skipPhase3)){
    $info="   Matched Meetings/attendance CSV:"+$resultsMatchedFilePath ; Write-Host $info; $info | Out-File $outLogFilePath -Append  
    $info="   Unmatched Meetings/attendance CSV:"+$resultsUnmatchedFilePath ; Write-Host $info; $info | Out-File $outLogFilePath -Append  
}

#########################################################################################################################


