<#
.SYNOPSIS
    MSMQ Monitoring

.DESCRIPTION
    CheckMQ MSMQ monitoring. 
    To be able to check if MSMQ queue messages are processed.

.PARAMETER  <Parameter-Name>
    If bound parameters, no need to put them overhere

.EXAMPLE 
    .\thescript.ps1 -param1 

.INPUTS
    file

.OUTPUTS
    file

.NOTES
    -Date 2021-05-06 - Author Admin - Johan Bijnens

#>
#Requires -Modules MSMQ

#$DebugPreference = [System.Management.Automation.ActionPreference]::Continue ;
#$DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue ;

Trap {
    # Handle the error
    $err = $_.Exception
    #Want to save tons of time debugging a #Powershell script? Put this in your catch blocks: 
    $ErrorLineNumber = $_.InvocationInfo.ScriptLineNumber
    write-warning $('Trapped error at line [{0}] : [{1}]' -f  $ErrorLineNumber,  $err.Message );

    write-Error $err.Message
    while( $err.InnerException ) {
	    $err = $err.InnerException
	    write-error $err.Message
	    };
    # End the script.
    break
    }


Import-Module MSMQ -ErrorAction Stop ;
<#
get-command -Module MSMQ 
#>



if ( $DebugPreference -eq [System.Management.Automation.ActionPreference]::Continue ){
    # om geen last te hebben van de plaats die wordt ingenomen door Write-Progress
    1..10 | % { ('*' * $_ ) }    
    }

#$env:COMPUTERNAME


$TargetFile = 'c:\CheckMQ\CheckMQReference.csv' ;

if ( !( Test-Path -Path (Split-Path -path $TargetFile -Parent ) -PathType Container ) ) {

    MD (Split-Path -path $TargetFile -Parent ) 

    }

$CheckMQ_Data = @() ;
if ( Test-Path -Path $TargetFile -PathType Leaf ) {

    $CheckMQ_Data += import-csv -Path $TargetFile -Delimiter ';' ;

    }

Write-Debug $('{0} - [{1}] records in file' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $CheckMQ_Data.Count )  ;

# Get-MsmqQueueManager 
$ALLQueues =  Get-MsmqQueue | sort QueueName ; 

Write-Debug $('{0} - [{1}] queues found' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $ALLQueues.Count )  ;

$xCtr = 0 ;
foreach ( $Queue in $ALLQueues ) {
    $xCtr ++ ;
    $pct = $xCtr * 100 / $ALLQueues.count 
    #Write-Progress -Activity $( 'Progressing [{0}]' -f $Queue.QueueName ) -Status $('{0} / {1}' -f $xCtr, $ALLQueues.Count ) -PercentComplete $pct   -SourceId 1 ;

    #$Queue =  Get-MsmqQueue -Name 'PublicQalzdba.DL' ;
    #$Queue | fl ;

    Write-Debug $('{0} - n Messages in queue [{1}]' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Queue.MessageCount )  ;
    #write-host $('{0} - n Messages in queue [{1}]' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Queue.MessageCount )  -BackgroundColor Gray -ForegroundColor Black ;


    if ( $Queue.MessageCount -gt 0 ) {
        Write-Debug $('{0} - Process Queue procedure' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff') ) ;
        #write-host $('{0} - Process Queue procedure' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff') ) -BackgroundColor Yellow -ForegroundColor Black ;
        # identifier van eerste msg lezen en opslaan ter vergelijking
        ## Get Message

        # peek is non destructive
        $PeekMsg = $Queue | Receive-MsmqQueue -Peek -Count 1 ;

        <#
        1 message uit de queue halen
        Get-MsmqQueue -Name $Queue.QueueName | receive-msmqqueue -Count 1 ;
        #>

        if ( $PeekMsg ) {
            Write-Debug $('{0} - PeekMsg in Queue [{1}]' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $PeekMsg.Id ) ;
            #write-host $('{0} - PeekMsg in Queue [{1}]' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $PeekMsg.Id )  -BackgroundColor green -ForegroundColor Black ;
            $SavedData = $CheckMQ_Data | where QueueName -eq $Queue.QueueName | where ServerName -eq $env:COMPUTERNAME ;
            if ( $SavedData ) {
                if ( $SavedData.id -eq $PeekMsg.Id ) {
                    $SavedData.LastModifyTime = $Queue.LastModifyTime ;
                    $SavedData.MessageCount = $Queue.MessageCount ;
                    $SavedData.TsChecked = get-date ;
                    $SavedData.Changed = $false ;
                    }
                else {
                    $SavedData.LastModifyTime = $Queue.LastModifyTime ;
                    $SavedData.MessageCount = $Queue.MessageCount ;
                    $SavedData.Id = $PeekMsg.Id ;
                    $SavedData.SentTime = $PeekMsg.SentTime ;
                    $SavedData.TsChecked = get-date ;
                    $SavedData.Changed = $true ;
                    }
                }
            else {
                $CheckMQ_Data += $Queue | Select @{n='ServerName';e={$env:COMPUTERNAME}}, QueueName, LastModifyTime, MessageCount, @{n='Id';e={[string]$PeekMsg.Id}}, @{n='SentTime';e={$PeekMsg.SentTime}}, @{n='TsChecked';e={get-date}}, @{n='Changed';e={$true}}  ;  
                }
            }
        }
    else {
         Write-Debug $('{0} - Queue is empty procedure' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff') )
         # write-host $('{0} - Queue is empty procedure' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff') )
         $SavedData = $CheckMQ_Data | where QueueName -eq $Queue.QueueName
         if ( $SavedData ) {
            if ( $SavedData.MessageCount -eq $Queue.MessageCount ) {
                $SavedData.LastModifyTime = $Queue.LastModifyTime ;
                $SavedData.Id = $null ;
                $SavedData.SentTime = $null ;
                $SavedData.TsChecked = get-date ;
                $SavedData.Changed = $false ;
            
                }
            else {
                $SavedData.LastModifyTime = $Queue.LastModifyTime ;
                $SavedData.MessageCount = $Queue.MessageCount ;
                $SavedData.Id = $null ;
                $SavedData.SentTime = $null ;
                $SavedData.TsChecked = get-date ;
                $SavedData.Changed = $true ;
                }
            }
         else {
            $CheckMQ_Data += $Queue | Select @{n='ServerName';e={$env:COMPUTERNAME}}, QueueName, LastModifyTime, MessageCount, @{n='Id';e={[string]$null}}, @{n='SentTime';e={$null}}, @{n='TsChecked';e={get-date}}, @{n='Changed';e={$false}} ;
            }
        }
    }


if ( $DebugPreference -eq [System.Management.Automation.ActionPreference]::Continue ){
    $CheckMQ_Data | sort QueueName | ft -AutoSize ;
    }

if ( Test-Path -Path $TargetFile -PathType Leaf ) {
    Remove-Item -Path $TargetFile -Force ;
    }

$CheckMQ_Data | sort QueueName | export-csv -Path $TargetFile -NoClobber -NoTypeInformation  -Delimiter ';' ;

Write-output $('{0} - [{1}] nRecords saved to [{2}]' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $CheckMQ_Data.count, $TargetFile ) #-BackgroundColor Blue -ForegroundColor White ;

write-output $('{0} - The end' -f (get-date -Format 'yyyy-MM-dd HH:mm:ss.fff')) #-BackgroundColor Yellow -ForegroundColor Black ;




