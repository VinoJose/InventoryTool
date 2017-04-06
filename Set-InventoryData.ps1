#region Init
#This is a function by Warren Frame with few tweaks.
function Get-RunspaceData {
    [cmdletbinding()]
    param( 
        [switch]$Wait,
        $SleepTimer,
        $RunspaceTimeout,
        $Logfile,
        $Quiet
     )
    #loop through runspaces
    #if $wait is specified, keep looping until all complete
    Do {

        #set more to false for tracking completion
        $More = $false

        #Progress bar if we have inputobject count (bound parameter)
        if (-not $Quiet) {
			Write-Progress  -Activity "Running Query" -Status "Starting threads"`
				-CurrentOperation "$StartedCount threads defined - $TotalCount input objects - $Script:CompletedCount input objects processed"`
				-PercentComplete $( Try { $Script:CompletedCount / $TotalCount * 100 } Catch {0} )
		}

        #run through each runspace.           
        Foreach($Runspace in $Runspaces) {
                    
            #get the duration - inaccurate
            $Currentdate = Get-Date
            $RunTime = $Currentdate - $Runspace.startTime
            $RunMin = [math]::Round( $RunTime.totalminutes ,2 )

            #set up log object
            $Log = "" | select Date, Action, Runtime, Status, Details
            $Log.Action = "Remoting:'$($Runspace.object.ServerName)'"
            $Log.Date = $Currentdate
            $Log.Runtime = "$RunMin minutes"

            #If runspace completed, end invoke, dispose, recycle, counter++
            If ($Runspace.Runspace.isCompleted) {
                            
                $Script:CompletedCount++
                        
                #check if there were errors
                if($Runspace.PowerShell.Streams.Error.Count -gt 0) {
                                
                    #set the logging info and move the file to completed
                    $Log.status = "CompletedWithErrors"
                    $Log.Details = $Runspace.PowerShell.Streams.Error[0]
                    Write-Verbose ($Log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                    foreach($ErrorRecord in $Runspace.PowerShell.Streams.Error) {
                        Write-Error -ErrorRecord $ErrorRecord
                    }
                }
                else {
                                
                    #add logging details and cleanup
                    $Log.status = "Completed"
                    Write-Verbose ($Log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                }

                #everything is logged, clean up the runspace
                $Runspace.PowerShell.EndInvoke($Runspace.Runspace)
                $Runspace.PowerShell.dispose()
                $Runspace.Runspace = $Null
                $Runspace.PowerShell = $Null

            }

            #If runtime exceeds max, dispose the runspace
            ElseIf ( $RunspaceTimeout -ne 0 -and $RunTime.totalseconds -gt $RunspaceTimeout) {
                            
                $Script:CompletedCount++
                $TimedOutTasks = $true
                            
				#add logging details and cleanup
                $Log.status = "TimedOut"
                Write-Verbose ($Log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                Write-Error "Runspace timed out at $($RunTime.totalseconds) seconds for the object:`n$($Runspace.object | out-string)"

                #Depending on how it hangs, we could still get stuck here as dispose calls a synchronous method on the PowerShell instance
                $Runspace.PowerShell.dispose()
                $Runspace.Runspace = $Null
                $Runspace.PowerShell = $Null
                $CompletedCount++

            }
                   
            #If runspace isn't null set more to true  
            ElseIf ($Runspace.Runspace -ne $Null ) {
                $Log = $Null
                $More = $true
            }

            #log the results if a log file was indicated
            if($LogFile -and $Log){
                $Log | Export-Csv $LogFile -append -NoTypeInformation
            }
        }

        #Clean out unused runspace jobs
        $TempHash = $Runspaces.clone()
        $TempHash | Where { $_.runspace -eq $Null } | ForEach {
            $Runspaces.remove($_)
        }

        #sleep for a bit if we will loop again
        if($PSBoundParameters['Wait']){ Start-Sleep -milliseconds $SleepTimer }

    #Loop again only if -wait parameter and there are more runspaces to process
    } while ($More -and $PSBoundParameters['Wait'])
                
#End of runspace function
}

$Sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$Runspacepool = [Runspacefactory]::CreateRunspacePool(1,10,$Sessionstate,$Host)
$Runspacepool.Open()

Write-Verbose "Creating an empty collection to hold runspace jobs"
$Script:Runspaces = New-Object System.Collections.ArrayList


#endregion
$Input = Import-Csv -Path "F:\MyLab\InventoryTool\Input.csv"
$Headers = $Input |Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
#$Credential = Get-Credential -Message "Enter Credentials for remoting"

#counts for progress
$TotalCount = $allObjects.count
$Script:CompletedCount = 0
$StartedCount = 0

foreach ($Item in $Input){
    $PowerShell = [PowerShell]::Create()
    $PowerShell.RunspacePool = $Runspacepool
    $Parameters = @{
        Data = $Item
        Headers = $Headers
        Credential = $Credential
    }
    $PowerShell.AddScript({
        Param (
            [Object]$Data,
            [Array]$Headers,
            $Credential
        )
        $ScriptBlock = {
            Param (
                [Array]$Headers,
                $Item
            )
            Function Reg-Backup {
                Param (
                    $Path
                )
                $FileName = "$($env:computername)_$(get-date -Format ddmmyyyy-hhMMss)"
                $RegBackup = "C:\RegBackup"
                If(Test-Path $RegBackup)
                {
                    If((Get-ChildItem $RegBackup).count -gt 10)
                    {
                            Get-ChildItem $RegBackup | Sort-Object -Property CreationTime | select -First ((Get-ChildItem $RegBackup).count - 10) | Remove-Item -Force
                    }
                    $Backup = &Reg export $Path "$RegBackup\$FileName.reg"
                }
                Else
                {
                    New-Item -Path C:\ -Name RegBackup -Force -ItemType directory
                    sleep 5
                    $Backup = &Reg export $Path "$RegBackup\$FileName.reg"
                }
            }
            Function Set-Values {
                Param (
                    $Path ,$Headers,$Item
                )
                $FileName = "WoN_$(get-date -Format ddmmyyyy-hhMMss)"
                $RegBackup = "C:\RegBackup"
                ForEach($Header in $Headers)
                {
                    If(($Header -ne "ServerName") -and ($Item.($Header) -ne $NUll) -and ($Item.($Header) -ne ""))
                    {
                        If($Header -eq "Build")
                        {
                            If(Test-Path -Path HKLM:\SOFTWARE\WoN\Build) {
                                $Backup = &Reg export HKLM:\SOFTWARE\WoN\Build "$RegBackup\$FileName.reg"
                                Set-ItemProperty -Path HKLM:\SOFTWARE\WoN\Build -Name $Header -Value $Item.($Header)
                            }
                            ElseIf(Test-Path -Path HKLM:\SOFTWARE\WoN) {
                                $Null = new-item -Path HKLM:\SOFTWARE -Name "Build"
                                Sleep 5
                                Set-ItemProperty -Path HKLM:\SOFTWARE\WoN\Build -Name $Header -Value $Item.($Header)
                            }
                            ElseIf(Test-Path -Path HKLM:\SOFTWARE) {
                                $Null = new-item -Path HKLM:\SOFTWARE -Name "WoN"
                                Sleep 5
                                $Null = new-item -Path HKLM:\SOFTWARE\WoN -Name "Build"
                                Sleep 5
                                Set-ItemProperty -Path HKLM:\SOFTWARE\WoN\Build -Name $Header -Value $Item.($Header)
                            }

                        }
                        Else
                        {
                            Set-ItemProperty -Path $Path -Name $Header -Value $Item.($Header)
                        }

                    }
                } 
            }  
            If(Test-Path -Path HKLM:\SOFTWARE\Wipro\Computer) {
                Reg-Backup -Path HKLM\SOFTWARE\Wipro\Computer
                Set-Values -Path HKLM:\SOFTWARE\Wipro\Computer -Headers $Headers -Item $Item
            }
            ElseIf(Test-Path -Path HKLM:\SOFTWARE\Wipro) {
                $Null = new-item -Path HKLM:\SOFTWARE\Wipro -Name "Computer"
                Sleep 5
                Set-Values -Path HKLM:\SOFTWARE\Wipro\Computer -Headers $Headers -Item $Item
            }
            ElseIf(Test-Path -Path HKLM:\SOFTWARE) {
                $Null = new-item -Path HKLM:\SOFTWARE -Name "Wipro"
                Sleep 5
                $Null = new-item -Path HKLM:\SOFTWARE\Wipro -Name "Computer"
                Sleep 5
                Set-Values -Path HKLM:\SOFTWARE\Wipro\Computer -Headers $Headers -Item $Item
            }
        }
        Invoke-Command -ComputerName $Data.ServerName -ScriptBlock $ScriptBlock -ArgumentList $Headers,$Data -Credential $Credential
    }) | Out-Null    
    [void]$PowerShell.AddParameters($Parameters)

    $Temp = "" | Select PowerShell, Runspace, StartTime, Object
    $Temp.PowerShell = $PowerShell
    $Temp.Object = $Item
    $Temp.StartTime = Get-Date
    $Temp.Runspace = $PowerShell.BeginInvoke()
    $StartedCount++
    [void]$Runspaces.add($Temp)
}

#region Post activities
Try {
    Get-RunspaceData -wait -SleepTimer 10 -RunspaceTimeout 60 -Logfile C:\Temp\InventoryLog.csv | Export-Csv -NoTypeInformation -Path C:\Temp\InventoryResult.csv -ErrorAction Stop
}
Catch {
    "Caught an error $_"
}
Write-Verbose "Closing the runspace pool"
$Runspacepool.close()

#collect garbage
[gc]::Collect()
#endregion
