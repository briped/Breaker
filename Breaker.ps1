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
        $Action   = New-ScheduledTaskAction -Execute $Executable -Argument $($Arguments -join ' ')
        $Trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DaysOfWeek -At $TimeOfDay
        $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DeleteExpiredTaskAfter $(New-TimeSpan -Seconds 0) -ExecutionTimeLimit $(New-TimeSpan -Minutes 5) -MultipleInstances IgnoreNew
        $Task     = New-ScheduledTask -Description '' -Action $Action -Trigger $Trigger -Settings $Settings
        Register-ScheduledTask -TaskPath $Path -TaskName $Name -InputObject $Task
    }

    end {
    }
}

function Get-BreakerTask {
    param(
        [CmdletBinding(SupportsShouldProcess = $true)]

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
    }

    process {
        Write-Verbose -Message "Path: '$($Path)'... Name: '$($Name)'..."
        if ($Path -match '^\\?(?<Path>([^\\/:*?"<>|]*)(\\[^\\/:*?"<>|]*)*)\\?$') {
            $Path = '\' + $Matches.Path + '\'
            Write-Verbose -Message "Sanitized Path: '$($Path)'..."
        }

        if ($Name) {
            try {
                Write-Verbose -Message "Trying to get task '$($Name)' from path '$($Path)'."
                $Result = Get-ScheduledTask -TaskPath $Path -TaskName $Name
                Write-Verbose -Message "Task: $($Result.TaskName)"
            }
            catch {
                throw "Error trying to get task '$($Name)' from path '$($Path)'. Error: $($PSItem)"
            }
        }
        else {
            try {
                Write-Verbose -Message "Trying to get all tasks from path '$($Path)'."
                $Result = Get-ScheduledTask -TaskPath $Path -ErrorAction Stop
                Write-Verbose -Message "Tasks: $($Result.TaskName)"
            }
            catch {
                throw "Error trying to get tasks from path '$($Path)'. Error: $($PSItem)"
            }
        }

        return $Result
    }

    end {
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
    "Knock, knock! Who's there? Soory can't answer, taking a break.",
    'Breakerfast time.'
)

$BreakerPath = 'Breaker - Take a Break'
#$BreakerMessages | Get-Random

#New-BreakerAppointment -Subject $BreakerName
#Get-BreakerAppointment
#Remove-BreakerAppointment

#New-BreakerTask -DaysOfWeek Monday -TimeOfDay '13:00' -Path $BreakerPath -Name $($BreakerMessages | Get-Random)
[xml]$TaskXml = (Get-BreakerTask -Path $BreakerPath -Name 'Timeout! Breaker disruption!').Xml
Get-BreakerTask -Path $BreakerPath -Name 'Timeout! Breaker disruption!'
<#
Name               : Timeout! Breaker disruption!
Path               : \Breaker - Take a Break\Timeout! Breaker disruption!

State              : 3
Enabled            : True
LastRunTime        : 30/11/1999 00.00.00
LastTaskResult     : 267011
NumberOfMissedRuns : 0
NextRunTime        : 04/04/2022 13.00.00
Definition         : System.__ComObject

Xml                : <?xml version="1.0" encoding="UTF-16"?>
                     <Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
                       <RegistrationInfo>
                         <Author>DAREDEVIL\briped</Author>
                         <URI>\Breaker - Take a Break\Timeout! Breaker disruption!</URI>
                       </RegistrationInfo>
                       <Principals>
                         <Principal id="Author">
                           <UserId>S-1-5-21-572002506-2244596241-3529022954-1001</UserId>
                           <LogonType>InteractiveToken</LogonType>
                         </Principal>
                       </Principals>
                       <Settings>

                         <DeleteExpiredTaskAfter>PT0S</DeleteExpiredTaskAfter>
                         <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
                         <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
                         <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
                         <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
                         <IdleSettings>
                           <StopOnIdleEnd>true</StopOnIdleEnd>
                           <RestartOnIdle>false</RestartOnIdle>
                         </IdleSettings>
                       </Settings>
                       <Triggers>
                         <CalendarTrigger>

                           <StartBoundary>2022-04-02T13:00:00+02:00</StartBoundary>

                           <EndBoundary>2022-05-01T17:00:00</EndBoundary>

                           <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>

                           <ScheduleByWeek>
                             <WeeksInterval>1</WeeksInterval>
                             <DaysOfWeek>

                               <Monday />
                               <Tuesday />
                               <Wednesday />
                               <Thursday />
                               <Friday />

                             </DaysOfWeek>
                           </ScheduleByWeek>
                         </CalendarTrigger>
                       </Triggers>
                       <Actions Context="Author">
                         <Exec>
                           <Command>rundll32.exe</Command>
                           <Arguments>user32.dll,LockWorkStation</Arguments>
                         </Exec>
                       </Actions>
                     </Task>
#>

#Remove-BreakerTask -Path $BreakerPath -Verbose