#requires -version 2
<#
.SYNOPSIS
  Script to manipulate xlsx data and create email templates

.DESCRIPTION
  Script will run through xlsx sheets/workbooks and pull out the required data and format and email using a hashtable

.PARAMETER <Parameter_Name>
   

.INPUTS
  XLSX,

.OUTPUTS
  Email templates

.NOTES
  Version:        1.0
  Author:         JR
  Creation Date:  09/02/2018
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#xls

#Logging Module
Import-Module PSLogging

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Log File Info
$sLogPath = "D:\Scripts\Logs\"
$sLogName = "xlxsmani.log"
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#-----------------------------------------------------------[Functions]------------------------------------------------------------



Function Start-Spunk{
  Param()
  
  Begin{
  Write-Host "Begin"
    #Start-Log -LogPath $sLogFile -LineValue "<description of what is going on>..."
  }
  
  Process{
    Try{
      Write-Host "Process"
    }
    
    Catch{
      #Write-LogError -LogPath $sLogFile -ErrorDesc $_.Exception -ExitGracefully $True
      Break
    }
  }
  
  End{
    If($?){
      #Log-Write -LogPath $sLogFile -LineValue "Completed Successfully."
      #Log-Write -LogPath $sLogFile -LineValue " "
     # Stop-Log -LogPath $sLogFile
     Write-Host "End"
    }
  }
}



#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion
Start-Spunk
#Log-Finish -LogPath $sLogFile