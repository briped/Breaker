function New-BreakerTask {
    param(
        [CmdletBinding(SupportsShouldProcess = $true)]

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('TaskFolder', 
            'TaskPath')]
        [string]
        $Path,

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('TaskName')]
        [string]
        $Name,

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
        $TimeOfDay
    )

    begin {
        $Executable = 'rundll32.exe'
        $Arguments  = @('user32.dll,LockWorkStation')
    }

    process {
        $Action  = New-ScheduledTaskAction -Execute $Executable -Argument $($Arguments -join ' ')
        $Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DaysOfWeek -At $TimeOfDay
        Register-ScheduledTask -Action $Action -Trigger $Trigger -TaskPath $Path -TaskName $Name
    }

    end {}
}

function Get-BreakerTask {
    param(
        [CmdletBinding(SupportsShouldProcess = $true)]

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('TaskFolder', 
            'TaskPath')]
        [string[]]
        $Path,

        [Parameter(Mandatory = $false, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('TaskName')]
        [string[]]
        $Name
    )

    begin {
        $ScheduleService = New-Object -ComObject Schedule.Service
        if (!$ScheduleService.Connected) {
            $ScheduleService.Connect()
        }
    }

    process {
        if ($Name) {
            $ScheduleService.GetFolder("\$($Path)").GetTask($Name)
        }
        else {
            $ScheduleService.GetFolder("\$($Path)").GetTasks($null)
        }
    }

    end {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ScheduleService) | Out-Null
    }
}


function Remove-BreakerTask {
    param(
        [CmdletBinding(SupportsShouldProcess = $true, 
            ConfirmImpact = 'High')]

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('TaskFolder', 
            'TaskPath')]
        [string]
        $Path,

        [Parameter(Mandatory = $false, 
            ValueFromPipelineByPropertyName = $true)]
        [Alias('TaskName')]
        [string]
        $Name
    )

    begin {
        Write-Verbose -Message "Path: '$($Path)'. Name: '$($Name)'."
        $ScheduleService = New-Object -ComObject Schedule.Service
        if (!$ScheduleService.Connected) {
            $ScheduleService.Connect()
        }
    }

    process {
        Write-Verbose -Message "Path: '$($Path)'. Name: '$($Name)'."
        $Path = "\$($Path)"
        if ($Path -like "*\$($Name)") {
            Write-Verbose -Message "Fixing piped path: '$($Path)'."
            $Path = Split-Path -Path $Path -Parent
        }
        $Path = $Path -replace '\+', '\'

        if ($Name) {
            Write-Verbose -Message "Deleting task: $($Path)\$($Name)"
            $ScheduleService.GetFolder($Path).DeleteTask($Name, $null)
        }
        else {
            $ScheduleService.GetFolder($Path).GetTasks($null) | ForEach-Object {
                Write-Verbose -Message "Deleting task: $($Path)\$($PSItem.Name)"
                $ScheduleService.GetFolder($Path).DeleteTask($PSItem.Name, $null)
            }
        }

        if (!$ScheduleService.GetFolder($Path).GetTasks($null).Count) {
            Write-Verbose -Message "Deleting Folder: $($Path)"
            $ScheduleService.GetFolder('\').DeleteFolder($Path, $null)
        }
    }

    end {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ScheduleService) | Out-Null
    }
}

function New-BreakerAppointment {
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
        $MeetingStart =(Get-Date),

        [Parameter(Mandatory = $false)]
        [int]
        $MeetingDuration = 15,

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
        $OutlookApplication = New-Object -ComObject 'Outlook.Application'
        $AppointmentItem    = $OutlookApplication.CreateItem('olAppointmentItem')
    }

    process {
        $AppointmentItem.Subject = $Subject
        $AppointmentItem.Body = $Body + "`r`n`r`n`r`n`r`n# DO NOT REMOVE # Created using Breaker # DO NOT REMOVE #"
        $AppointmentItem.Location  = $Location
        $AppointmentItem.ReminderSet = $EnableReminder
	    $AppointmentItem.ReminderMinutesBeforeStart = $Reminder
	    $AppointmentItem.Start = $MeetingStart
	    $AppointmentItem.Duration = $MeetingDuration
    }

    end {
        $AppointmentItem.Save()
        $AppointmentItem.Display($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($OutlookApplication) | Out-Null
    }
}

function Get-BreakerAppointment {
    param(
        [CmdletBinding()]

        [Parameter(Mandatory = $false, 
            ValueFromPipelineByPropertyName = $true)]
        [string[]]
        $EntryID
    )

    begin {
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
    param(
        [CmdletBinding(SupportsShouldProcess = $true, 
            ConfirmImpact = 'High')]

        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)]
        [string[]]
        $EntryID
    )

    begin {
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

$BreakerMessages = @(
    'Breaker One-Five.',
    'Breakdance!',
    'Timeout! Breaker disruption!',
    'Time for a break!',
    "Knock, knock! Who's there? Breaker! Breaker who? Breaker Taker!",
    'Breakerfast time.'
)

$BreakerPath = 'Breaker - Take a Break'
$BreakerName = 'No Time!'

#New-BreakerAppointment -Subject $BreakerName
#Get-BreakerAppointment | Remove-BreakerAppointment
#New-BreakerTask -DaysOfWeek Monday -TimeOfDay '13:00' -Path $BreakerPath -Name $BreakerName
#Get-BreakerTask -Path $BreakerPath | Remove-BreakerTask -Verbose
#Remove-BreakerTask -Path $BreakerPath -Verbose