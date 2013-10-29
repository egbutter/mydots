Function Send-Mail ([String] $FromEmailID, [String] $ToEmailID,[String] $Subject,[String] $Body, [String] $Smtpserver=$SMTPServer)
{
        $Msg  = New-Object Net.Mail.MailMessage
        $SMTP = New-Object Net.Mail.SmtpClient($SMTPServer) 
        $Msg.From = $FromEmailID
        $Msg.To.Add($ToEmailID)
        $Msg.Subject = $Subject
        $Msg.Body = $Body
        $SMTP.Send($msg)
}



Function Execute-SQL
{

    Param
    (
        [Object]  $SQLConnection,        #OPEN SQL Connection object
        [String]  $SQL,                  #SQL Command
        [int]     $TimeOut=300
    )
    
    $ErrorActionPreference="SilentlyContinue"
[Object] $SQLHash=@{SQLErrMessage="";SQLResult="";SQLExitCode=0}
    trap [Exception] 
    {
        $SQLHash.SQLErrMessage=($_.exception.message).ToString()
        $SQLHash.SQLResult=$rowoutput
        $SQLHash.SQLExitCode=-1
        if ($_.exception.message -match 'Exception calling "ExecuteReader" with "0"')
        {
            Return $SQLHash            
            Continue
        }
    }

    [Boolean] $Help = $H

    #Define variables that would be used in this Script
    [int]        $Count            = 1
    [int]        $fieldmax         = 0
    [int]        $fieldcount       = 0
    [String]     $rowoutput        = ""
    [String]     $rowoutputtemp    = ""
    [int]        $readcolnameflag  = 0
    [Object]     $SqlCmd           = $NULL
    [Object]     $ObjReader        = $NULL
    [String]     $ERR1             = ""
    [boolean]    $IsRecordsetavail = $True


    if ($Help -or $ArgsCount -eq 0)
    {
        Show-Help -SourceScriptName $ScriptName
        End-Script
    }

    if ($SQLConnection -eq $NULL)
    {
        Write-Msg -ErrNumber 2 -Message "Parameters Missing."
        Show-Help -SourceScriptName $ScriptName
        End-Script
    }

    
    if ($SQLConnection.state -ne "Open") 
    {
        Write-Msg -ErrNumber 2 -Message "SQL Connection in $ScriptName is not Open. Cannot execute"
        End-Script  
    }

    $SQLCmd        = New-Object System.Data.SqlClient.SqlCommand

    if ($TimeOut -eq $False -or $TimeOut -eq "")
    {
        $SQLCmd.CommandTimeout = 300
    }
    else
    {
        $SQLCmd.CommandTimeout = $TimeOut
    }
    
    $SQLCmd.CommandTimeout = $TimeOut
    $SQLCmd.Connection  = $SqlConnection
    $SQLCmd.CommandText = $SQL

    #$SqlConnection.open() | Out-Null

    $objReader= $SqlCmd.ExecuteReader() 

    if ($?) 
    {
        while ($IsRecordsetavail -eq $True)
        {
            while ($objReader.Read()) 
            { 
               $FieldMax=$ObjReader.FieldCount
               $FieldCount=0
               if ($ReadColNameFlag -eq 0)
               {
                   while ($FieldCount -le $FieldMax-1)
                   {
                       $RowOutputTemp=$RowOutputTemp +","+(($objReader.GetName($FieldCount)).ToString()).Trim()
                       $FieldCount = $FieldCount +1
                   }

                   $ReadColNameFlag=1 
                   $FieldCount=0
               
                   if($RowOutputTemp.SubString(0,1) -eq ",")
                   {
                       $RowOutputTemp=$RowOutputTemp.SubString(1,($RowOutputTemp.Length -1))
                   }

                   $RowOutput=$RowOutput+$RowOutputTemp+"`n"    
                   $RowOutputTemp=""
                }

                while ($FieldCount -le $FieldMax-1)
                {
                    $RowOutputTemp=$RowOutputTemp +","+(($ObjReader.Item($FieldCount)).ToString()).Trim()
                    $FieldCount = $FieldCount +1
                }
                $RowOutput=$RowOutput+$RowOutputTemp.SubString(1,($RowOutputTemp.Length -1))
                $RowOutput=$RowOutput+"`n"    
                $RowOutputTemp=""
            }
            $IsRecordsetavail=$objReader.nextresult()
            $ReadColNameFlag =0
        }

        $SQLHash.SQLErrMessage="None." 
        $SQLHash.SQLResult=$rowoutput
        $SQLHash.SQLExitCode=0
        $ObjReader.Close()
    }
    else 
    {
        $ERR=$error[0].exception
        $SQLHash.SQLErrMessage=$ERR.message
        $SQLHash.SQLResult=$rowoutput
        $SQLHash.SQLExitCode=-1
    }
    Return $SQLHash
}

