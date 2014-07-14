#----------------------------------------------------------------------
# add-sysadminToAllRussellInstances.ps1
# Grant sysadmin access to a given login to all RI instances
#
# Created On Jan-29-2014 By Rodrigo Silva
#-----------------------------------------------------------------------


function SrvConnect($Server)
{
    $srv = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $Server 
    #Write-Host ===================== 	
    #Write-Host $srv.Name 				
    #Write-Host $srv.VersionString		
    #Write-Host $srv.Edition 			
    #Write-Host $srv.ProductLevel 		
    #Write-Host ===================== 	
    $Srv.ConnectionContext.StatementTimeout=0
    
    Return ,$srv
}

function to_exit()
{
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers() 
	exit
}

$dbaConfigServerName = 'DBPROD30B\DBSQL30B'
$dbaConfigDBName = 'DBA_Configuration'
$loginName = 'PROD\Russell_DBA_Support'

### Assemblies
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoExtended") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoEnum") | Out-Null

Write-Host Verifying Connection to $dbaConfigServerName ... 

$dbaConfigServ = SrvConnect($dbaConfigServerName)
$dbServerInfo = $dbaConfigServ.Databases[$dbaConfigDBName]
    if ($dbServerInfo)
    {
		Write-Host Connected to $dbaConfigServerName
        Write-Host
	}
	else
	{
		Write-Host Error connecting to $dbaConfigServerName -foregroundcolor red
		if ($Error[0])
		{
			$Error[0].Expetion.GetBaseException()
			to_exit
		}
	}
try
{
	
	$dsServerNameId = $dbServerInfo.ExecuteWithResults("SELECT [ServerInstanceName] FROM [DBA_Configuration].[dbo].[DB_SQL_SRV_LIST] WHERE [ServerInstanceName] is not NULL order by [ServerInstanceName];")
	$dtServerNameId = $dsServerNameId.Tables[0]
	foreach ($r in $dtServerNameId.Rows)
	{
		$strServerName = $r[0].ToString()
        try
        {
            $Serv = SrvConnect($strServerName)
	        if ($Serv)
	        {
		        Write-Host Connected to $strServerName
                $log =$Serv.Logins.Item($loginName)
                if (!($log))
                    {
						try
						{
							$login = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Login -ArgumentList $strServerName, $loginName   
							$login.LoginType = 'WindowsGroup'
							Write-Host Login $loginName created!
							
							try
							{
								$svrole = $serv.Roles["sysadmin"]
								$svrole.AddMember($loginName)
								Write-Host Granted sysadmin to $loginName
								Write-Host 
								Write-Host
							}
							catch
							{	
								Write-Host Error trying to grant sysadmin on $strServerName -foregroundcolor red
								Write-Host $error[0].ToString() -foregroundcolor red
							}
						}
						catch
						{
							Write-Host Error trying to create login on $strServerName -foregroundcolor red
							Write-Host $error[0].ToString() -foregroundcolor red
						}
                    }
                else
                    {
                        try
						{
							Write-HOst 2
							$svrole = $Serv.Roles["sysadmin"]
							Write-host 3
							$svrole.AddMember($loginName)
							Write-Host Granted sysadmin to $loginName
							Write-Host
							Write-Host
						}
						catch
						{
							Write-Host Error trying to grant sysadmin on $strServerName -foregroundcolor red
							Write-Host $error[0].ToString() -foregroundcolor red
						}
                    }
  	        }
            else
	        {
		        Write-Host $serverName does not exists -foregroundcolor red
	        }
	    }
		catch
		{
			Write-Host Error trying to connect to $strServerName -foregroundcolor red
			Write-Host $error[0].ToString() -foregroundcolor red
		}
	}
}
catch
{
	Write-Host Error trying to get server name and server ir from DBA Configuration -foregroundcolor red
	Write-Host $error[0].ToString() -foregroundcolor red
	to_exit
}