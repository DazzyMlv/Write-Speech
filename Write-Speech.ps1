function Write-Speech{
<#
.SYNOPSIS
Speaks a sting object.

.EXAMPLE
Write-Speech "Hello World."
.EXAMPLE
Write-Speech "Hi, i'm $Env:USERNAME" -SpeechVoiceSpeakFlags SVSFlagsAsync
.EXAMPLE
"Hello World." | Write-Speech

.INPUTS
System.String

.LINKS
https://github.com/DazzyMlv

.COMPONENT
SAPI.SpVoice

.FUNCTIONALITY
Speak out stream of text past to its 'InputObject' parameter.

.FORWARDHELPCATEGORY Function
#>
    [CmdletBinding(PositionalBinding=$true, 
                   DefaultParameterSetName="Speak",
                   HelpUri='https://github.com/DazzyMlv', 
                   RemotingCapability='None')]
    param(
        [Parameter(Mandatory=$true,
                   ParameterSetName="Speak",
                   Position=0,
                   ValueFromPipeline=$True,
                   HelpMessage="Specifies the objects to speak then send down the pipeline. Enter a variable that contains the objects, or type a command or expression that gets the objects.")]
        [AllowEmptyString()]
        [AllowEmptyCollection()]
        [AllowNull()]
        [Alias('Msg', 'Message')]
        [string[]]$InputObject,

        [Parameter(HelpMessage="[Optionally] Specifies the Flags. Default value is SVSFDefault.")]
        [ValidateSet("SVSFDefault", "SVSFIsFilename", "SVSFIsNotXML", "SVSFIsXML", 
                     "SVSFlagsAsync", "SVSFNLPMask",
                     "SVSFNLPSpeakPunc", "SVSFParseAutodetect",
                     "SVSFParseMask", "SVSFParseSapi", 
                     "SVSFParseSsml", "SVSFPersistXML", 
                     "SVSFPurgeBeforeSpeak", "SVSFUnusedFlags", 
                     "SVSFVoiceMask")]
        [parameter(ParameterSetName="Speak")]
        [parameter(ParameterSetName="ReadRecentHistory")]
        [Alias('SVSF')]
        [string[]]$SpeechVoiceSpeakFlag = "SVSFDefault",

        [ValidateRange(0,100)]
        [Alias('Volume')]
        [parameter(ParameterSetName="ReadRecentHistory")]
        [parameter(ParameterSetName="Speak")]
        [Int]$SpVoiceVolume=50,

        [Parameter(HelpMessage="Specifies voice. Default value is 1.", 
                   ParameterSetName="Speak")]
        [parameter(ParameterSetName="ReadRecentHistory")]
        [ValidateSet(0,1)]
        [int]$SpVoice = [int]0,

        [parameter(ParameterSetName="ReadRecentHistory")]
        [parameter(ParameterSetName="Speak")]
        [Alias('Priority')]
        [Int]$SpVoicePriority=0,

        [parameter(ParameterSetName="ReadRecentHistory")]
        [parameter(ParameterSetName="Speak")]
        [Alias('Rate')]
        [Int]$SpVoiceRate = 0,

        [parameter(ParameterSetName="ReadRecentHistory")]
        [parameter(ParameterSetName="Speak")]
        [Alias('SynchronousSpeakTimeout')]
        [Int]$SpVoiceSynchronousSpeakTimeout = 10000,

        [parameter(ParameterSetName="Speak")]
        [switch]$SuppressConsoleOutput,

        [parameter(ParameterSetName="Speak")]
        [switch]$SuppressReturnObject,

        
        [parameter(ParameterSetName="ReadRecentHistory")]
        [switch]$ReadRecentHistory,

        [parameter(ParameterSetName="GetRecentRecords")]
        [parameter(ParameterSetName="ReadRecentHistory")]
        [parameter(ParameterSetName="SelectOject_First_Last_Skip_Unique")]
        [parameter(ParameterSetName="SelectOject_Unique")]
        [switch]$Unique,

        [switch]$Pompous,

        [parameter(ParameterSetName="ResetRecentHistory")]
        [switch]$ResetRecentHistory,

        [parameter(ParameterSetName="GetRecentRecords")]
        [switch]$GetRecentRecords,

        [parameter(ParameterSetName="GetRecentRecords")]
        [parameter(ParameterSetName="ReadRecentHistory")]
        [parameter(ParameterSetName="SelectOject_First_Last_Skip_Unique")]
        [int]$Last,

        [parameter(ParameterSetName="GetRecentRecords")]
        [parameter(ParameterSetName="ReadRecentHistory")]
        [parameter(ParameterSetName="SelectOject_First_Last_Skip_Unique")]
        [int]$First,

        [parameter(ParameterSetName="GetRecentRecords")]
        [parameter(ParameterSetName="ReadRecentHistory")]
        [parameter(ParameterSetName="SelectOject_First_Last_Skip_Unique")]
        [int]$Skip,

        [parameter(ParameterSetName="GetRecentRecords")]
        [parameter(ParameterSetName="ReadRecentHistory")]
        [parameter(ParameterSetName="SelectOject_Unique")]
        [int32]$Index,

        [parameter(ParameterSetName="Speak")]
        [Alias('NoRecords')]
        [switch]$NoHistory
        
    ) # END: param() Block
    BEGIN {
        Try{
            Write-Verbose "[+] Initialising $($MyInvocation.MyCommand.CommandType.ToString().ToLower()): '$($MyInvocation.InvocationName)'..."
            #$RegistryPath="HKLM:\SOFTWARE\DAZZY-ThinkCentre, Inc.\$($MyInvocation.InvocationName)"
            $AppData  = "$Env:APPDATA\DAZZY-ThinkCentre, Inc\$($MyInvocation.InvocationName)";
            [string]$LogFile = "$AppData\$($MyInvocation.InvocationName).log";
            [string]$LogString = "";
            [string]$ReturnableString = "";
            [string]$RecentRecords = "$AppData\$($MyInvocation.InvocationName).rec";

            Write-Verbose "[+] Enumerating 'Speech Voice Speak Flags'..."
            Write-Debug "[+] Enumerating 'Speech Voice Speak Flags'..."
            
            Enum SpeechVoiceSpeakFlags{
                #SpVoice flags
                SVSFDefault = 2;
                SVSFlagsAsync = 1;
                SVSFPurgeBeforeSpeak = 2;
                SVSFIsFilename = 4;
                SVSFIsXML = 8;
                SVSFIsNotXML = 16;
                SVSFPersistXML = 32;

                #Normalizer flags
                SVSFNLPSpeakPunc = 64;

                #Masks
                SVSFNLPMask = 64;
                SVSFVoiceMask = 127;
                SVSFUnusedFlags = -128;
            }
            Write-Verbose "[✓] 'Speech Voice Speak Flags' Enumerated."
            Write-Debug "[✓] 'Speech Voice Speak Flags' Enumerated."

            if($SpeechVoiceSpeakFlag -eq 'SVSFIsFilename'){
                Write-Warning "'$SpeechVoiceSpeakFlag' is not surported at the moment. Defaulting to 'SVSFDefault'."
                $SpeechVoiceSpeakFlag = "SVSFDefault"
            }

            Try{
                Write-Verbose "[+] Initialising the speech engine ..."
                Write-Debug "[+] Initialising the speech engine ..."
                $Flag = $SVSF = ([SpeechVoiceSpeakFlags]::$SpeechVoiceSpeakFlag)         
                $Sapi = New-Object -ComObject SAPI.SpVoice
                $Sapi.Volume = $SpVoiceVolume;
                $Sapi.Rate = $SpVoiceRate;
                $Sapi.Priority = $SpVoicePriority;
                $Sapi.Voice = @($Sapi.GetVoices())[$SpVoice];
                Write-Verbose "[✓] Speech engine initialized."
                Write-Debug "[✓] Speech engine initialized."
            }Catch{
                Write-Verbose "[-] $($_.Exception)."
                Write-Warning "[-] $($_.Exception.Message)."
                if( $Pompous ){
                    PompousSapi;
                }
                Write-Debug "[-] $($Error[0]) (`n`twhile $($MyInvocation.MyCommand.CommandType.ToString().ToLower()): '$($MyInvocation.InvocationName)' was setting up 'Sapi').";
            }

            
           
            Try {
                if( -NOT (Test-Path -Path $AppData -PathType Container) ){
                   New-Item -Path $AppData -ItemType Directory |Out-Null
     
                }
            } Catch{
                Write-Verbose "[-] $($_.Exception)."
                Write-Warning "[-] $($_.Exception.Message)."
                if( $Pompous ){
                    PompousSapi;
                }
                Write-Debug "$($Error[0]) (`n`twhile creating working directory for $($MyInvocation.MyCommand.CommandType.ToString().ToLower()): '$($MyInvocation.InvocationName)').";
            }

            #$SelectOjectCommonParams = @{}
            $SelectOjectCommonParams = [System.Collections.Hashtable]::new()
            
            if($First){
                $SelectOjectCommonParams.Add("First", $First)
            }

            if($Last){
                $SelectOjectCommonParams.Add('Last', $Last)
            }

            if($Skip){
                $SelectOjectCommonParams.Add('Skip', $Skip)
            }
            if($Index){
                $SelectOjectCommonParams.Add('Index', $Index)
            }
            
            function SpeakNewMsg{
                [CmdletBinding(PositionalBinding=$true)]
                param(
                    [string[]]$MsgToSpeak
                    )

                [bool]$SpeakNewMsgStatus = $false
                Try{
                    if(  -NOT (Test-Path -Path $RecentRecords -PathType Leaf) ){
                        New-Item -Path $RecentRecords -ItemType File -Force|Out-Null
                    }
                    foreach ($Msg in $MsgToSpeak) {
                        Write-Verbose "[+] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' processing '$Msg'";
                        if(!$NoHistory){
                            Try{
                                Write-Verbose "$($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' saving '$($Msg)' to '$($RecentRecords)'...";                                
                                Out-File -FilePath $RecentRecords -Append -Force -NoClobber -InputObject $Msg
                                Write-Verbose "$($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' saved '$($Msg)' to '$($RecentRecords)' sucessfully.";                                
                                $SpeakNewMsgStatus = $true
                            }Catch{
                                $SpeakNewMsgStatus = $false
                                Write-Verbose "[-] $($_.Exception)."
                                Write-Warning "[-] $($_.Exception.Message)."
                                if( $Pompous ){
                                    PompousSapi ($_.Exception.Message);
                                }
                                Write-Debug "$($Error[0]) (`n`twhile '$($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' was trying to save '$($Msg)' to '$($RecentRecords)'.";
                            } # End: Try{}Catch{}
                        }
                        Try {
                            $Sapi.Speak("$Msg", ($SVSF.value__) ) | Out-Null;
                            $SpeakNewMsgStatus = $true
                        } Catch {
                            $SpeakNewMsgStatus = $false;
                            $LogString += "[-] <$($MyInvocation.InvocationName)> Failed to Process '$Msg'"
                            Write-Verbose "[-] $($_.Exception)."
                            Write-Warning "[-] $($_.Exception.Message)."
                                if( $Pompous ){
                                    PompousSapi;
                                }
                            #Write-Log -InputObject "[-] $($MyInvocation.MyCommand.CommandType.ToString()):<$($MyInvocation.InvocationName)> $($Error[0])" -LogFileName $LogFile -LogType DEBUG
                            Write-Debug "[-] $($MyInvocation.MyCommand.CommandType.ToString()):<$($MyInvocation.InvocationName)> $($Error[0])"
                            #throw
                        } # End: Try-Catch
                        Write-Verbose "[!] '$($MyInvocation.InvocationName)' status: $($Sapi.Status)."
                    } # End: foreach ($WriteSpeechText in $WriteSpeechMsg)
                } Catch{
                    $SpeakNewMsgStatus = $false
                    Write-Verbose "[-] $($_.Exception)."
                    Write-Warning "[-] $($_.Exception.Message)."
                    if( $Pompous ){
                        PompousSapi;
                    }
                    Write-Debug "[-] $($Error[0]) (Running $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)').";
                }Finally{
                    Write-Verbose "[+] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' terminated." 
                }
                Return $SpeakNewMsgStatus;
            } # End: function SpeakNew{

            function Reset-RecentHistory{
                [bool]$ResetRecentHistoryStatus = $true;
                Try{
                    if( (Test-Path -Path $RecentRecords -PathType Leaf -ErrorAction Stop) ){
                        Out-File -FilePath  $RecentRecords -Force -NoNewline -InputObject ''
                    }else{
                        Write-Verbose "[!] No recent records found."
                    }
                    $ResetRecentHistoryStatus = $true;
                }Catch{
                    $ResetRecentHistoryStatus = $false;
                    Write-Verbose "[-] $($_.Exception)."
                    Write-Warning "[-] $($_.Exception.Message)."
                    PompousSapi ($_.Exception.Message);
                    Write-Debug "[-] $($Error[0]) (while $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' was trying to 'reset recent history').";
                }
                Return $ResetRecentHistoryStatus
            }
            function Get-RecentRecords{
                [CmdletBinding(PositionalBinding=$true)]
                param()
                [bool]$GetRecentRecordsStatus = $true
                $RecentRecordsCreationTime = [System.IO.File]::GetCreationTime($RecentRecords);
                $RecentRecordsLastAccessTime = [System.IO.File]::GetLastWriteTime($RecentRecords);

                if( (Test-Path -Path $RecentRecords -PathType Leaf -ErrorAction Stop) ){
                    Try{
                        $Speeches = ( (Get-Content -Path $RecentRecords -Force) )
                        if($Speeches.Length -le 0 ){
                            $GetRecentRecordsStatus = $false;
                            $TempWarningMsg = "No recent history. History was resently cleared by $Env:USERNAME";
                            Write-Warning "[!] $TempWarningMsg"
                            if( $Pompous ){
                                PompousSapi $TempWarningMsg |Out-Null;
                            }
                        } Else{
                            $GotRecentRecords = ($Speeches |Select @SelectOjectCommonParams );
                            $GetRecentRecordsStatus = $true
                        } # End: if($Speeches.Length -le 0 ){}Else{}
                    }Catch{
                        Write-Verbose "[-] $($_.Exception)."
                        Write-Warning "[-] $($_.Exception.Message)."
                        $GetRecentRecordsStatus = $false
                        if( $Pompous ){
                            PompousSapi -InputObject "$($_.Exception.Message)";
                        }
                        Write-Debug "[+] '$($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)': `t $($Error[0])";
                    }
                }Else{
                    $GetRecentRecordsStatus = $false
                    if( $RecentRecordsCreationTime -eq $RecentRecordsLastAccessTime  ){
                        $TempWarningMsg = "'Write-Speech' has never spoken.";
                        Write-Warning "[+] $TempWarningMsg"
                        if( $Pompous ){
                            PompousSapi $TempWarningMsg |Out-Null;
                        }
                    }Else{
                        $TempWarningMsg = "No recent history"
                        Write-Warning "[+] $TempWarningMsg.";
                        if( $Pompous ){
                            PompousSapi -InputObject "$TempWarningMsg.";
                        }
                    }
                    Write-Debug "[+] '$($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' reports '$TempWarningMsg'.";
                }
                Return $GotRecentRecords
            }

            function Read-RecentRecords{
                [CmdletBinding(PositionalBinding=$true)]
                param()
                [bool]$ReadRecentRecordsStatus = $true
                $RecentRecordsCreationTime = [System.IO.File]::GetCreationTime($RecentRecords);
                $RecentRecordsLastAccessTime = [System.IO.File]::GetLastWriteTime($RecentRecords);
                            
                if( (Test-Path -Path $RecentRecords -PathType Leaf -ErrorAction Stop) ){
                    Try{
                        $Speeches = ( (Get-Content -Path $RecentRecords -Force) )
                        if($Speeches.Length -eq 0 ){
                            $TempWarningMsg = "No recent history. History was resently cleared by $Env:USERNAME";
                            Write-Warning "[+] $TempWarningMsg"
                            if( $Pompous ){
                                PompousSapi $TempWarningMsg |Out-Null;
                            }
                        } Else{ # if($Speeches.Length -le 0 ){}
                            $Sapi.Speak( ($Speeches | Select @SelectOjectCommonParams), ($SVSF.value__) ) | Out-Null;
                                #$Sapi.CreateObjRef([ref])
                        } # End: if($Speeches.Length -le 0 ){}Else{}
                        $ReadRecentRecordsStatus = $true
                    }Catch{
                        Write-Verbose "[-] $($_.Exception)."
                        Write-Warning "[-] $($_.Exception.Message)."
                        $ReadRecentRecordsStatus = $false
                        if( $Pompous ){
                            PompousSapi -InputObject "$($_.Exception.Message)";
                        }
                        Write-Debug "[-] '$($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)': $($Error[0])";
                    }
                }Else{
                    $ReadRecentRecordsStatus = $false
                    if( $RecentRecordsCreationTime -eq $RecentRecordsLastAccessTime  ){
                        $TempWarningMsg = "'$($MyInvocation.InvocationName)' has never spoken.";
                        Write-Warning "[+] $TempWarningMsg"
                        if( $Pompous ){
                            PompousSapi $TempWarningMsg |Out-Null;
                        }
                    }Else{
                        $TempWarningMsg = "No recent history"
                        Write-Warning "[-] $TempWarningMsg.";
                        if( $Pompous ){
                            PompousSapi -InputObject "$TempWarningMsg.";
                        }
                    }
                    Write-Debug "[!] '$($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' reports '$TempWarningMsg'.";
                }
                Return $ReadRecentRecordsStatus
            } # End: function ReadRecentRecords{     
    
            function PompousSapi{
                [CmdletBinding()]
                param(
                    [string[]]$PompousSapiMsg = "WARNING: $($_.Exception.Message). ($Env:USERNAME, please consider 'Debugging')."
                )
                begin{
                    Try{
                        Write-Verbose "[+] Initialising $($MyInvocation.MyCommand.CommandType.ToString().ToLower()): '$($MyInvocation.InvocationName)'..."
                        $PompousSapi = New-Object -ComObject SAPI.SpVoice;
                        $PompousSapi.Volume = 100;
                        $PompousSapi.Rate = 0;
                        $PompousSapi.Priority = 0;
                        $PompousSapi.Voice = @($PompousSapi.GetVoices())[0];
                        $PompousSapiDefaultSpeech = "Internal Exception: {0}. ($Env:USERNAME, please consider 'Debugging').";
                        $PompousSapiWorkingDirectory = "$AppData\$($MyInvocation.InvocationName)";
                        $PompousSapiTempFile = "$PompousSapiWorkingDirectory\{BLANK}-$($MyInvocation.InvocationName).tmp";
                        $PompousSapiFlag = ($SVSF.value__)

                        if( !(Test-Path -Path $PompousSapiWorkingDirectory) ){
                            New-Item -Type Directory -Path $PompousSapiWorkingDirectory -Force | Out-Null;
                        }
                        if( !(Test-Path -Path $PompousSapiTempFile -PathType Leaf ) ){
                            New-Item -path "$PompousSapiWorkingDirectory\{BLANK}-$($MyInvocation.InvocationName).tmp" -Type file -Force | Out-Null
                        }
            
                        Write-Verbose "[+] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' initialisation completed." 
                    } Catch{
                        Write-Verbose "[-] $($_.Exception)."
                        Write-Warning "[-] $($_.Exception.Message)."
                        if( $Pompous ){
                            $PompousSapi.Speak( ($PompousSapiDefaultSpeech -f $_.Exception.Message), $PompousSapiFlag ) | Out-Null;
                        }
                        Write-Debug "[-] $($_.Exception.Message) (while initialising $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)').";
                    }Finally{
                        Write-Verbose "[+] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' initialisation process terminated." 
                    }
                } # End:  PompousSapi begin{}
                process{
                    try{
                        if($Pompous){
                            foreach($PompousSapitext in $PompousSapiMsg){
                                Try{
                                    $PompousSapi.Speak("$PompousSapitext", $PompousSapiFlag) | Out-Null;
                                }Catch{
                                    $PompousSapiTempWarning = "$($MyInvocation.InvocationName) is unable to process the msg: `"$PompousSapitext`" ";
                                    Write-Warning "[-] $PompousSapiTempWarning"
                                    if( $Pompous ){$PompousSapi.Speak( $PompousSapiTempWarning, $PompousSapiFlag ) | Out-Null;}
                                }Finally{
                                }
                            }
                        }
                    } Catch{
                        Write-Verbose "[-] $($_.Exception)."
                        Write-Warning "[-] $($_.Exception.Message)."
                        if( $Pompous ){$PompousSapi.Speak( ($PompousSapiDefaultSpeech -f $_.Exception.Message), $PompousSapiFlag ) | Out-Null;}
                        Write-Debug "[-] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' $($Error[0]) (while trying to speak $PompousSapiMsg)." 
                    }
                } # End: PompousSapi process{}
                end{
                    Try{
                        Write-Verbose "[+] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' doing some housekeeping...";
                    }Catch{
                        Write-Verbose "[-] $($_.Exception)."
                        Write-Warning "[-] $($_.Exception.Message)."
                        if( $Pompous ){$PompousSapi.Speak( ($PompousSapiDefaultSpeech -f $_.Exception.Message), $PompousSapiFlag ) | Out-Null;}
                        Write-Debug "[-] $($_.Exception.Message) (while  $($MyInvocation.MyCommand.CommandType.ToString().ToLower()): '$($MyInvocation.InvocationName)' was doing some housekeeping).";
                    } Finally{ # End: END{} Try{}Catch{}
                        Write-Verbose "[!] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' completed. ";
                    }
                } # End:  PompousSapi end{} 
            }# End: function PompousSapi{}

            Write-Verbose "[+] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' initialisation completed." 
        }Catch{
            Write-Verbose "[-] $($_.Exception)."
            Write-Warning "[-] $($_.Exception.Message)."
            if( $Pompous ){
                PompousSapi;
            }
            Write-Debug "[-] $($Error[0]) (while initialising $($MyInvocation.MyCommand.CommandType.ToString().ToLower()): '$($MyInvocation.InvocationName)')  $($MyInvocation.MyCommand.CommandType.ToString().ToLower()).";
        } # End: BEGIN{} Try{}Catch{}
    } # END:  BEGIN {}
    PROCESS {
        Try{
            Write-Verbose "[+] $($MyInvocation.MyCommand.CommandType.ToString()) '$($MyInvocation.InvocationName)' PROCESS started.";            
            if( $PSCmdlet.ParameterSetName -eq "Speak" ){
                SpeakNewMsg -MsgToSpeak $InputObject |Out-Null
            } Elseif( $PSCmdlet.ParameterSetName -eq "ReadRecentHistory" ){ # End if( $PSCmdlet.ParameterSetName -eq "Speak" ){}
                Read-RecentRecords |Out-Null
            } Elseif( $PSCmdlet.ParameterSetName -eq "ResetRecentHistory" ){
                Reset-RecentHistory |Out-Null
            } # End: if(...){}ElseIF...(...){}
        }Catch{
            Write-Verbose "[-] $($_.Exception)."
            Write-Warning "[-] $($_.Exception.Message)."
            if( $Pompous ){
                PompousSapi -PompousSapiMsg  
            }
            Write-Debug "[-] $($_.Exception.Message) (while '$($MyInvocation.InvocationName)' $($MyInvocation.MyCommand.CommandType.ToString().ToLower()) was on process).";
        } Finally{ # End: END{} Try{}Catch{}
            Write-Verbose "[+] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' PROCESS completed. ";
            Write-Debug "[+] $($MyInvocation.MyCommand.CommandType.ToString()):'$($MyInvocation.InvocationName)'  PROCESS complete.";
        } # End: PROCESS{} Try{}Catch{}
    } # End: PROCESS {}
    END {
        Try{
            if( $PSCmdlet.ParameterSetName -eq "GetRecentRecords" ){
                Return Get-RecentRecords
            } # End: if()
        }Catch{
            Write-Verbose "[-] $($_.Exception)."
            Write-Warning "[-] $($_.Exception.Message)."
            if( $Pompous ){
                PompousSapi   
            }
            Write-Debug "[-] $($Error[0]) (while '$($MyInvocation.InvocationName)' $($MyInvocation.MyCommand.CommandType.ToString().ToLower()) was doing some housekeeping).";
        } Finally{ # End: END{} Try{}Catch{}
            Write-Verbose "[+] $($MyInvocation.MyCommand.CommandType.ToString()): '$($MyInvocation.InvocationName)' completed. ";
            Write-Debug "[+] $($MyInvocation.MyCommand.CommandType.ToString()):'$($MyInvocation.InvocationName)'  complete.";
        }
    } # End: END {}
} # END: Function Write-Speech{}
