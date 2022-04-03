function New-BreakerTask {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER TaskPath
    Parameter description
    
    .PARAMETER TaskName
    Parameter description
    
    .PARAMETER DaysOfWeek
    Parameter description
    
    .PARAMETER TimeOfDay
    Parameter description
    
    .PARAMETER ExpireAfter
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    TODO:
    * Documentation.
    #>
    param(
        [CmdletBinding(SupportsShouldProcess = $true)]

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('TaskFolder', 
            'Path')]
        [string]
        $TaskPath,

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('Name')]
        [string]
        $TaskName,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Monday', 
            'Tuesday', 
            'Wednesday', 
            'Thursday', 
            'Friday', 
            'Saturday', 
            'Sunday')]
        [string[]]
        $DaysOfWeek,

        [Parameter(Mandatory = $false)]
        $TimeOfDay,

        [Parameter(Mandatory = $false)]
        [int]
        $ExpireAfter = 6
    )

    begin {
        Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): begin"
        if ($VerbosePreference) {
            foreach ($Parameter in $MyInvocation.BoundParameters.GetEnumerator()) {
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): param: $($Parameter.Key): $($Parameter.Value)"
            }
        }

        # Verify that the path looks correct and capture the path without leading and trailing backslashes.
        if ($TaskPath -match '^\\?(?<Path>([^\\/:*?"<>|]+)(\\[^\\/:*?"<>|]+)*)\\?$') {
            # Re-create the path with exactly one leading and trailing backslash.
            $TaskPath = '\' + $Matches.Path + '\'
            Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Real TaskPath: '$($TaskPath)'..."
        }

        # Ensure that the name does not contain invalid characters.
        $TaskName = $TaskName -replace '[\\/:*?"<>|]+', '_'
        Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Real TaskName: '$($TaskName)'..."

        $TaskDescription = 'Breaker - Take a Break!'
        $Executable = 'rundll32.exe'
        $Arguments  = @('user32.dll,LockWorkStation')
        $DeleteExpiredTaskAfter = New-TimeSpan -Seconds 0
        $ExecutionTimeLimit = New-TimeSpan -Minutes 1
        $EndBoundary = (Get-Date).AddMonths($ExpireAfter).Date.ToString('u').Replace(' ', 'T')
    }

    process {
        try {
            Write-Verbose -Message "Trying to get information for task '$($TaskPath + $TaskName)'."
            $TaskInfo = Get-ScheduledTaskInfo -TaskName $($TaskPath + $TaskName) -ErrorAction SilentlyContinue
        }
        catch {
            throw "Could not get information for task '$($TaskPath + $TaskName)'. Error: $($PSItem)"
        }
        if ($TaskInfo) {
            throw "The task '$($TaskInfo.TaskName)' already exists."
        }
        $Action   = New-ScheduledTaskAction -Execute $Executable -Argument $($Arguments -join ' ')
        $Trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DaysOfWeek -At $TimeOfDay
        $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DeleteExpiredTaskAfter $DeleteExpiredTaskAfter -ExecutionTimeLimit $ExecutionTimeLimit -MultipleInstances IgnoreNew
        $Task     = New-ScheduledTask -Description $TaskDescription -Action $Action -Trigger $Trigger -Settings $Settings
        # Add some additional settings to the trigger(s).
        $Task.Triggers | ForEach-Object {
            $PSItem.EndBoundary = $EndBoundary
            $PSItem.ExecutionTimeLimit = 'PT30S' # 30 seconds
        }
        $Task | Register-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName
    }

    end {
    }
}

function Get-BreakerTask {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER TaskPath
    Parameter description
    
    .PARAMETER TaskName
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    TODO:
    * Documentation
    #>
    param(
        [CmdletBinding(SupportsShouldProcess = $true)]

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('TaskFolder', 
            'Path')]
        [string]
        $TaskPath,

        [Parameter(Mandatory = $false, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('Name')]
        [string]
        $TaskName
    )

    begin {
        Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): begin"
        if ($VerbosePreference) {
            foreach ($Parameter in $MyInvocation.BoundParameters.GetEnumerator()) {
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): param: $($Parameter.Key): $($Parameter.Value)"
            }
        }
    }

    process {
        if ($TaskPath -match '^\\?(?<Path>([^\\/:*?"<>|]+)(\\[^\\/:*?"<>|]+)*)\\?$') {
            $TaskPath = '\' + $Matches.Path + '\'
            Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Real TaskPath: '$($TaskPath)'..."
        }

        if ($TaskName) {
            try {
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Trying to get task '$($TaskName)' from path '$($TaskPath)'."
                $Result = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Task: '$($Result.TaskName)'."
            }
            catch {
                throw "Error trying to get task '$($TaskName)' from path '$($TaskPath)'. Error: $($PSItem)"
            }
        }
        else {
            try {
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Trying to get all tasks from path '$($TaskPath)'."
                $Result = Get-ScheduledTask -TaskPath $TaskPath -ErrorAction Stop
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Task: '$($Result.TaskName)'."
            }
            catch {
                throw "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Error trying to get tasks from path '$($TaskPath)'. Error: $($PSItem)"
            }
        }

        return $Result
    }

    end {
    }
}

function Remove-BreakerTask {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER TaskPath
    Parameter description
    
    .PARAMETER TaskName
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    TODO:
    * Documentation
    #>
    param(
        [CmdletBinding(SupportsShouldProcess = $true, 
            ConfirmImpact = 'High')]

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('TaskFolder', 
            'Path')]
        [string]
        $TaskPath,

        [Parameter(Mandatory = $false, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('Name')]
        [string]
        $TaskName
    )

    begin {
        Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): begin"
        if ($VerbosePreference) {
            foreach ($Parameter in $MyInvocation.BoundParameters.GetEnumerator()) {
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): param: $($Parameter.Key): $($Parameter.Value)"
            }
        }

        $ScheduleService = New-Object -ComObject Schedule.Service
        if (!$ScheduleService.Connected) {
            $ScheduleService.Connect()
        }
    }

    process {
        Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): TaskPath: $($TaskPath)"
        # Verify that the path looks correct and capture the path without leading and trailing backslashes.
        if ($TaskPath -match '^\\?(?<Path>([^\\/:*?"<>|]+)(\\[^\\/:*?"<>|]+)*)\\?$') {
            $TaskPath = '\' + $Matches.Path
            Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Real TaskPath: $($TaskPath)"
        }

        # Ensure that the name does not contain invalid characters.
        $TaskName = $TaskName -replace '[\\/:*?"<>|]+', '_'
        Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Real TaskName: $($TaskName)"

        if ($TaskName) {
            Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Deleting task: $($TaskPath)$($TaskName)"
            $ScheduleService.GetFolder($TaskPath).DeleteTask($TaskName, $null)
        }
        else {
            $ScheduleService.GetFolder($TaskPath).GetTasks($null) | ForEach-Object {
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Deleting task: $($TaskPath)$($PSItem.Name)"
                $ScheduleService.GetFolder($TaskPath).DeleteTask($PSItem.Name, $null)
            }
        }

        if (!$ScheduleService.GetFolder($TaskPath).GetTasks($null).Count) {
            Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): Deleting Folder: $($TaskPath)"
            $ScheduleService.GetFolder('\').DeleteFolder($TaskPath, $null)
        }
    }

    end {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ScheduleService) | Out-Null
    }
}

function New-BreakerAppointment {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER Subject
    Parameter description
    
    .PARAMETER Body
    Parameter description
    
    .PARAMETER Start
    Parameter description
    
    .PARAMETER Duration
    Parameter description
    
    .PARAMETER Location
    Parameter description
    
    .PARAMETER EnableReminder
    Parameter description
    
    .PARAMETER Reminder
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    TODO:
    * Add handling of the remaining parameters.
    * Documentation
    #>
    param (
        [CmdletBinding(SupportsShouldProcess = $true)]

        [Parameter(Mandatory = $true)]
        [string]
        $Subject,

        [Parameter(Mandatory = $false)]
        [string]
        $Body,

        [Parameter(Mandatory = $false)]
        [datetime]
        $Start =(Get-Date),

        [Parameter(Mandatory = $false)]
        [int]
        $Duration = 15,

        [Parameter(Mandatory = $false)]
        [string]
        $Location,

        [Parameter(Mandatory = $false)]
        [bool]
        $EnableReminder = $true,

        [Parameter(Mandatory = $false)]
        [int]
        $Reminder = 1
    )

    begin {
        Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): begin"
        if ($VerbosePreference) {
            foreach ($Parameter in $MyInvocation.BoundParameters.GetEnumerator()) {
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): param: $($Parameter.Key): $($Parameter.Value)"
            }
        }

        $OutlookApplication = New-Object -ComObject 'Outlook.Application'
        $AppointmentItem    = $OutlookApplication.CreateItem('olAppointmentItem')
        $PatternEndDate = (Get-Date).AddMonths(6).Date

    }

    process {
        $AppointmentItem.Subject = $Subject
        $AppointmentItem.Body = $Body + "`r`n`r`n`r`n`r`n# DO NOT REMOVE # Created using Breaker # DO NOT REMOVE #"
        $AppointmentItem.Location  = $Location
        $AppointmentItem.ReminderSet = $EnableReminder
	    $AppointmentItem.ReminderMinutesBeforeStart = $Reminder
	    $AppointmentItem.Start = $Start
	    $AppointmentItem.Duration = $Duration

        $RecurrencePattern = $AppointmentItem.GetRecurrencePattern()
        $RecurrencePattern.PatternEndDate = $PatternEndDate
    }

    end {
        $AppointmentItem.Save()
        #$AppointmentItem.Display($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($OutlookApplication) | Out-Null
    }
}

function Get-BreakerAppointment {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER EntryID
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    TODO:
    * Add handing of parameter(s)
    * Documentation
    #>
    param(
        [CmdletBinding()]

        [Parameter(Mandatory = $false, 
            ValueFromPipelineByPropertyName = $true)]
        [string[]]
        $EntryID
    )

    begin {
        Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): begin"
        if ($VerbosePreference) {
            foreach ($Parameter in $MyInvocation.BoundParameters.GetEnumerator()) {
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): param: $($Parameter.Key): $($Parameter.Value)"
            }
        }

        $OutlookApplication = New-Object -ComObject Outlook.Application
        $OutlookNamespace   = $OutlookApplication.GetNamespace('MAPI')
        $OutlookCalendar    = $OutlookNamespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
    }

    process {
        return $OutlookCalendar.Items | Where-Object { $PSItem.Body -match 'Created using Breaker.' }
        #$OutlookItem | Select-Object -Property Subject,Start,End,Duration,BusyStatus,EntryID,GlobalAppointmentID,ConversationID,ConversationIndex,ConversationTopic,CreationTime,LastModificationTime,Location,Organizer,StartUTC,EndUTC
    }

    end {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($OutlookApplication) | Out-Null
    }
}

function Remove-BreakerAppointment {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER EntryID
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    TODO:
    * Documentation
    #>
    param(
        [CmdletBinding(SupportsShouldProcess = $true, 
            ConfirmImpact = 'High')]

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [string[]]
        $EntryID
    )

    begin {
        Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): begin"
        if ($VerbosePreference) {
            foreach ($Parameter in $MyInvocation.BoundParameters.GetEnumerator()) {
                Write-Verbose -Message "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name): param: $($Parameter.Key): $($Parameter.Value)"
            }
        }

        $OutlookApplication = New-Object -ComObject Outlook.Application
        $OutlookNamespace   = $OutlookApplication.GetNamespace('MAPI')
        $OutlookCalendar    = $OutlookNamespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
    }

    process {
        $OutlookItem = $OutlookNamespace.GetItemFromID($EntryID, $OutlookCalendar.StoreID)
        $OutlookItem.Delete()
    }

    end {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($OutlookApplication) | Out-Null
    }
}

# Variables for easy testing
$BreakerPath = 'Breaker - Take a Break'
$BreakerMessages = @(
    'Breaker One-Five.',
    'Breakdance!',
    'Timeout! Breaker disruption!',
    'Time for a break!',
    "Knock, knock! Who's there? Sorry can't answer, taking a break.",
    'Breakerfast time.'
)
$BreakerName = $BreakerMessages | Get-Random
$WeekDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')
$DaysOfWeek = $WeekDays | Get-Random -Count (Get-Random -Minimum 1 -Maximum 7)
$QuarterHours = @(0, 15, 30, 45)
$Hour = Get-Random -Minimum 0 -Maximum 23
$Minute = $QuarterHours | Get-Random
$TimeOfDay = "$($Hour):$($Minute)"

#New-BreakerAppointment -Verbose -Subject $BreakerName
#Get-BreakerAppointment -Verbose | Remove-BreakerAppointment -Verbose
#New-BreakerTask -Verbose -DaysOfWeek $DaysOfWeek -TimeOfDay $TimeOfDay -TaskPath $BreakerPath -Name $BreakerName
#Get-BreakerTask -Verbose -TaskPath $BreakerPath -TaskName $BreakerName
#Get-BreakerTask -Verbose -TaskPath $BreakerPath
#Get-BreakerTask -Verbose -TaskPath $BreakerPath -TaskName $BreakerName | Remove-BreakerTask -Verbose
#Get-BreakerTask -Verbose -TaskPath $BreakerPath | Remove-BreakerTask -Verbose
#Remove-BreakerTask -Verbose -TaskPath $BreakerPath -TaskName $BreakerName
#Remove-BreakerTask -Verbose -TaskPath $BreakerPath
