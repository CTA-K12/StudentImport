<#
	.SYNOPSIS
		A script that processes student data, and creates accounts in
		Active Directoy and Google, and other services.
	
	.DESCRIPTION
		Automatically Create Accounts for Students
	
	.PARAMETER District
		The Name of the District.
	
	.PARAMETER ConfigXML
		Path to the XML Configuration  file.
	
	.NOTES
		===========================================================================
		Created on:   	3/21/2017 1:06 PM
		Created by:   	Eden Nelson
		Organization: 	Cascade Technology Alliance
		Filename:     	StudentImport.ps1
		Version: 		1.0.50.0
		===========================================================================
#>
[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true)]
	[Alias('DistrictName')]
	[string]$District,
	[Parameter(Mandatory = $true)]
	[Alias('ConfigurationFile')]
	[string]$ConfigXML
)
#Requires -Version 3.0
#Requires -Modules ActiveDirectory
##Requires -PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Import-Module ActiveDirectory
#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

if (!([bool]((Get-Content $ConfigXML) -as [System.Xml.XmlDocument]))) { Write-Error "$ConfigXML does not exist or is not valid XML!"; break }
if (!(Select-Xml -Xml ([System.Xml.XmlDocument](Get-Content $ConfigXML)) -XPath "/Districts/District[@Name=`'$District`']")) { Write-Error "$ConfigXML does not contain district $District"; break }

$ObjRandom = new-object SYSTEM.Random
$output = @()
$todaysDate = Get-Date -Format 'yyyy/MM/dd'

if (!(Test-Path -Path 'Script:')) { New-PSDrive -Name Script -PSProvider FileSystem -Root $PSScriptRoot | Out-Null }
if ((Get-Location).Path -ne 'Script:\') { Push-Location -Path Script:\ }

#region Functions
function Clear-File {
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[Alias('Name')]
		$File
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [system.diagnostics.stopwatch]::StartNew()
	}
	process {
		foreach ($filename in $File) { if (Test-Path -Path ($PSScriptRoot, '\Files\', $filename -join '')) { Remove-Item ($PSScriptRoot, '\Files\', $filename -join '') -Force -Confirm:$false } }
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Get-AllUserDataAD {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[Alias('DistinguishedName')]
		[array]$SearchBase
	)
	
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		foreach ($DN in $SearchBase) {
			$users += (Get-ADUser -SearchBase $DN -Filter 'employeeid -like "*"' -Properties AccountExpirationDate,`
								  Department,`
								  DisplayName,`
								  DistinguishedName,`
								  Division,`
								  MemberOf,`
								  EmailAddress,`
								  EmployeeID,`
								  personalTitle,`
								  ScriptPath,`
								  Office,`
								  Title,`
								  Surname,`
								  GivenName,`
								  Enabled,`
								  proxyAddresses,`
								  HomeDirectory,`
								  HomeDrive,`
								  Initials`
)
		}
		Write-Output $users
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Get-FileSCP {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$DLLPath,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$RemoteFilePath,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$LocalFilePath,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$HostName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Username,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Password,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$SshHostKeyFingerprint
	)
	
	#https://winscp.net/eng/docs/guide_protecting_credentials_for_automation
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		try {
			Add-Type -Path $DLLPath
			$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
				Protocol = [WinSCP.Protocol]::Sftp
				HostName = $HostName
				UserName = $Username
				Password = $Password
				SshHostKeyFingerprint = $SshHostKeyFingerprint
			}
			
			$session = New-Object WinSCP.Session
			
			try {
				$session.Open($sessionOptions)
				$transferOptions = New-Object WinSCP.TransferOptions
				$transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
				$transferResult = $session.GetFiles($RemoteFilePath, $LocalFilePath, $False, $transferOptions)
				$transferResult.Check()
			} finally {
				$session.Dispose()
			}
		} catch [Exception]
		{
			Write-Error ("Error: {0}" -f $_.Exception.Message)
		}
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Get-OrganizationalUnitPath {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Location,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[int]$GradYear,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Grade,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$OrganizationalUnit,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$District
	)
	
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		switch ($District) {
			NWRESD { $path = $OrganizationalUnit }
			default { $path = "OU=$GradYear", $OrganizationalUnit -join ',' }
		}
		$properties = @{
			Path = $path
		}
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Get-UserDataImport {
	[CmdletBinding()]
	param
	(
		$UserDataRaw
	)
	
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		$properties = @{
			AccountExpirationDate = $UserDataRaw.AccountExpirationDate
			ChangePasswordAtNextLogon = $false
			PasswordNeverExpires  = $UserDataRaw.PasswordNeverExpires
			CannotChangePassword  = $UserDataRaw.CannotChangePassword
			ScriptPath		      = $UserDataRaw.ScriptPath
			DataOfBirth		      = $UserDataRaw.BIRTHDATE
			Department		      = 'Student'
			Description		      = "Last import: $todaysDate"
			DisplayName		      = $UserDataRaw.FIRST_NAME, $UserDataRaw.LAST_NAME -join ' '
			DistinguishedName	  = $null
			Division			  = $UserDataRaw.HOMEROOM_TCH
			EmailAddress		  = $null
			EmployeeID		      = $UserDataRaw.SIS_NUMBER
			Enabled			      = $true
			GivenName			  = $UserDataRaw.FIRST_NAME
			Initials			  = $UserDataRaw.MIDDLE_INIITAL
			HomeDirectory		  = $null
			HomeDrive			  = $UserDataRaw.HomeDrive
			MemberOf			  = $null
			Name				  = $null
			Office			      = $UserDataRaw.LOCATION
			Password			  = $null
			PasswordCrypt		  = $null
			Path				  = $UserDataRaw.Path
			personalTitle		  = $UserDataRaw.CALCULATED_GRAD_YEAR
			proxyAddresses	      = $null
			SamAccountName	      = $null
			Surname			      = $UserDataRaw.LAST_NAME
			Title				  = $UserDataRaw.GRADE
			UserPrincipalName	  = $null
			Notify			      = $false
		}
		$userDataImport = New-Object -TypeName PSObject -Property $properties
		Write-Output $userDataImport
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Move-SuspendedAccounts {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$MoveFromOU,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$MoveToOU
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		foreach ($DN in $MoveFromOU) {
			if ($DN -contains 'Disabled') { continue }
			try {
				Search-ADAccount –AccountDisabled –UsersOnly –SearchBase $DN | Move-ADObject –TargetPath $MoveToOU
			} catch {
				Write-Error ("Error: {0}" -f $_.Exception.Message)
			}
		}
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function New-EmailAddress {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$GivenName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Surname,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$SamAccountName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$EmailSuffix,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$District
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		switch ($District) {
			default {
				$emailAddress = $SamAccountName, $EmailSuffix -join ''
			}
		}
		$properties = @{
			EmailAddress = $emailAddress
		}
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function New-GroupAD {
	[CmdletBinding()]
	param ()
	
	New-ADGroup -Name $physicalDeliveryOfficeNameRedux -GroupCategory Security -GroupScope Global -Path "OU=Meraki,OU=Groups,DC=intra,DC=parkrose,DC=k12,DC=or,DC=us"
}
function New-OrganizationalUnitPath {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$OrganizationalUnitDN,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$District
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		switch ($District) {
			default {
				$ouName, $ouPath = $OrganizationalUnitDN.TrimStart('OU=') -split ',', 2
				
				try {
					New-ADOrganizationalUnit -Name $ouName -Path $ouPath
				} catch [Exception]
				{
					Write-Error ("Error: {0}" -f $_.Exception.Message)
					continue
				}
			}
		}
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function New-Password {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.String]$Office,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.String]$DateOfBirth,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.String]$Grade,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.String]$District,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$Words,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$Numbers,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$SpecialCharacters,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.String]$DefaultPassword,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.Boolean]$UseDefaultPassword,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.String]$EmployeeID,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.Boolean]$DefaultPasswordIsStudentID,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.String]$DOBPasswordGrade,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.String]$DOBPasswordLocations
	)
	
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		switch ($District) {
			default {
				if (($UseDefaultPassword) -or ($DefaultPasswordIsStudentID)) {
					if ($DefaultPasswordIsStudentID) { $password = $EmployeeID } else { $password = $DefaultPassword }
				} else {
					if (($Grade -le $DOBPasswordGrade) -or ($Grade -match 'KG|PS|PK|TR') -or (($DOBPasswordLocations -replace '[^a-zA-Z0-9 ]', '') -contains ($Office -replace '[^a-zA-Z0-9 ]', ''))) {
						$dobMonth, $dobDay, [string]$dobYear = $DateOfBirth.split('\/', 3)
						$monthAbrv = @{ '01' = 'Jan'; '02' = 'Feb'; '03' = 'Mar'; '04' = 'Apr'; '05' = 'May'; '06' = 'Jun'; '07' = 'Jul'; '08' = 'Aug'; '09' = 'Sep'; '10' = 'Oct'; '11' = 'Nov'; '12' = 'Dec' }
						$password = $monthAbrv.Get_Item("$dobMonth"), $dobDay, '-', $dobYear.Substring(2) -join ''
					} else {
						$word1 = ($words[$ObjRandom.Next(0, $words.Count)])
						$word2 = ($words[$ObjRandom.Next(0, $words.Count)])
						$word1 = $word1.substring(0, 1).toupper() + $word1.substring(1).tolower()
						$Number = ($Numbers[$ObjRandom.Next(0, $Numbers.Count)])
						$SpecialCharacter = ($SpecialCharacters[$ObjRandom.Next(0, $SpecialCharacters.Count)])
						$password = $word1, $Number, $word2, $SpecialCharacter -join ''
					}
				}
			} # End of Default
		} # End of switch
		
		$properties = @{
			Password	  = $password
			PasswordCrypt = (ConvertTo-SecureString $($password) -AsPlainText -Force)
		}
		
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function New-SamAccountName {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$GivenName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$SurName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$GradYear,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$UPNSuffiix,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Grade,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$AllUsersAD,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$District
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		[int]$currentYear = Get-Date -Format 'yyyy'
		[int]$currentMonth = Get-Date -Format 'MM'
		switch ($District) {
			JEWELLSD {
				if (((($Grade -eq '12') -or ($Grade -eq 'TR')) -and ($currentMonth -ge '08') -and ($GradYear -le $currentYear)) -or ($GradYear -lt $currentYear)) { $GradYear = ($currentYear + 1) }
				if ($GivenName.Length -ge '18') { $GivenName = $GivenName.substring(0, 18) }
				$samAccountName = $GivenName, $SurName.substring(0, 1) -join ''
				$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
				if ($sAMAccountName.Length -ge '18') { $sAMAccountName = $sAMAccountName.substring(0, 18) }
				$samAccountName = $samAccountName, $GradYear.Substring($GradYear.get_Length() - 2) -join ''
				if (($AllUsersAD.SamAccountName.Contains($samAccountName)) -and (($SurName.Length -gt '1'))) {
					$i = 1
					Do {
						$i++
						Write-Verbose $i
						Write-Verbose $samAccountName
						if ($sAMAccountName.Length -ge '20') {
							if ($GivenName.Length -ge '19') { $GivenName = $GivenName.substring(0, 19) }
							$samAccountName = $GivenName.subString(0, $GivenName.get_Length() - $i), $SurName.substring(0, $i) -join ''
							$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
							if ($sAMAccountName.Length -gt '19') { $sAMAccountName = $sAMAccountName.substring(0, 19) }
							$samAccountName = $samAccountName, $GradYear.Substring($GradYear.get_Length() - 2) -join ''
						} else {
							$samAccountName = $GivenName, $SurName.substring(0, $i), $GradYear.Substring($GradYear.get_Length() - 2) -join ''
							$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
						}
					} while ($AllUsersAD.SamAccountName.Contains($samAccountName))
				}
			}
			GASTONSD {
				if (((($Grade -eq '12') -or ($Grade -eq 'TR')) -and ($currentMonth -ge '08') -and ($GradYear -le $currentYear)) -or ($GradYear -lt $currentYear)) { $GradYear = ($currentYear + 1) }
				if ($GivenName.Length -ge '18') { $GivenName = $GivenName.substring(0, 18) }
				$samAccountName = $GivenName, $SurName.substring(0, 1) -join ''
				$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
				if ($sAMAccountName.Length -ge '18') { $sAMAccountName = $sAMAccountName.substring(0, 18) }
				$samAccountName = $samAccountName, $GradYear.Substring($GradYear.get_Length() - 2) -join ''
				if (($AllUsersAD.SamAccountName.Contains($samAccountName)) -and (($SurName.Length -gt '1'))) {
					$i = 1
					Do {
						$i++
						if ($sAMAccountName.Length -ge '20') {
							if ($GivenName.Length -ge '19') { $GivenName = $GivenName.substring(0, 19) }
							$samAccountName = $GivenName.subString(0, $GivenName.get_Length() - $i), $SurName.substring(0, $i) -join ''
							$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
							if ($sAMAccountName.Length -gt '19') { $sAMAccountName = $sAMAccountName.substring(0, 19) }
							$samAccountName = $samAccountName, $GradYear.Substring($GradYear.get_Length() - 2) -join ''
						} else {
							$samAccountName = $GivenName, $SurName.substring(0, $i), $GradYear.Substring($GradYear.get_Length() - 2) -join ''
							$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
						}
					} while ($AllUsersAD.SamAccountName.Contains($samAccountName))
				}
			}
			SEASIDE {
				if (((($Grade -eq '12') -or ($Grade -eq 'TR')) -and ($currentMonth -ge '08') -and ($GradYear -le $currentYear)) -or ($GradYear -lt $currentYear)) { $GradYear = ($currentYear + 1) }
				if ($GivenName.Length -ge '18') { $GivenName = $GivenName.substring(0, 18) }
				$samAccountName = $GivenName, $SurName.substring(0, 1) -join ''
				$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
				if ($sAMAccountName.Length -ge '18') { $sAMAccountName = $sAMAccountName.substring(0, 18) }
				$samAccountName = $GradYear.Substring($GradYear.get_Length() - 2), $samAccountName -join ''
				if (($AllUsersAD.SamAccountName.Contains($samAccountName)) -and (($SurName.Length -gt '1'))) {
					$i = 1
					Do {
						$i++
						if ($sAMAccountName.Length -ge '20') {
							if ($GivenName.Length -ge '19') { $GivenName = $GivenName.substring(0, 19) }
							$samAccountName = $GivenName.subString(0, $GivenName.get_Length() - $i), $SurName.substring(0, $i) -join ''
							$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
							if ($sAMAccountName.Length -gt '19') { $sAMAccountName = $sAMAccountName.substring(0, 19) }
							$samAccountName = $GradYear.Substring($GradYear.get_Length() - 2), $samAccountName -join ''
						} else {
							$samAccountName = $GradYear.Substring($GradYear.get_Length() - 2), $GivenName, $SurName.substring(0, $i) -join ''
							$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
						}
					} while ($AllUsersAD.SamAccountName.Contains($samAccountName))
				}
			}
			default {
				if (((($Grade -eq '12') -or ($Grade -eq 'TR')) -and ($currentMonth -ge '08') -and ($GradYear -le $currentYear)) -or ($GradYear -lt $currentYear)) { $GradYear = ($currentYear + 1) }
				if ($GivenName.Length -ge '16') { $GivenName = $GivenName.substring(0, 16) }
				$samAccountName = $GivenName, $SurName.substring(0, 1) -join ''
				$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
				if ($sAMAccountName.Length -ge '16') { $sAMAccountName = $sAMAccountName.substring(0, 16) }
				$samAccountName = $samAccountName, $GradYear -join ''
				if (($AllUsersAD.SamAccountName.Contains($samAccountName)) -and (($SurName.Length -gt '1'))) {
					$i = 1
					Do {
						$i++
						if ($sAMAccountName.Length -ge '20') {
							if ($GivenName.Length -ge '17') { $GivenName = $GivenName.substring(0, 17) }
							$samAccountName = $GivenName.subString(0, $GivenName.get_Length() - $i), $SurName.substring(0, $i) -join ''
							$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
							if ($sAMAccountName.Length -gt '17') { $sAMAccountName = $sAMAccountName.substring(0, 17) }
							$samAccountName = $samAccountName, $GradYear -join ''
						} else {
							$samAccountName = $GivenName, $SurName.substring(0, $i), $GradYear -join ''
							$samAccountName = $samAccountName.ToLower() -replace '\s', '' -replace "'", "" -replace '`', '' -replace ',', '' -replace '\.', '' -replace '-', ''
						}
					} while ($AllUsersAD.SamAccountName.Contains($samAccountName))
				}
			}
		}
		$properties = @{
			SamAccountName    = $samAccountName
			UserPrincipalName = $samAccountName, $UPNSuffiix -join ''
			Name			  = $samAccountName
		}
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "New-SamAccountName $($samAccountName)"
		Write-Verbose "END $($MyInvocation.MyCommand)"
	}
}
function New-StudentUserAD {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$Student
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		try {
			New-ADUser `
					   -Name $Student.Name `
					   -AccountPassword $Student.PasswordCrypt `
					   -Department $Student.Department `
					   -Description "Last import: $todaysDate" `
					   -DisplayName $Student.DisplayName `
					   -Division $Student.Division `
					   -EmailAddress $Student.EmailAddress `
					   -EmployeeID $Student.EmployeeID `
					   -Enabled $Student.Enabled `
					   -GivenName $Student.GivenName `
					   -Initials $Student.Initials `
					   -Office $Student.Office `
					   -Path $Student.Path `
					   -SamAccountName $Student.SamAccountName `
					   -Surname $Student.Surname `
					   -Title $Student.Title `
					   -OtherAttributes @{ 'personalTitle' = "$($Student.personalTitle)" } `
					   -AccountExpirationDate $Student.AccountExpirationDate `
					   -ScriptPath $Student.ScriptPath `
					   -CannotChangePassword $Student.CannotChangePassword `
					   -PasswordNeverExpires $Student.PasswordNeverExpires `
					   -UserPrincipalName $Student.UserPrincipalName -ErrorAction Stop
			
			if ($Student.ChangePasswordAtNextLogon) { Set-ADUser -Identity $Student.SamAccountName -ChangePasswordAtLogon $true }
		} catch [Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException] {
			Write-Error ("Error: {0}" -f $_.Exception.Message)
			Write-LogEntry -Severity 3 -Value "The specified account already exists"
			Write-Verbose "Attempting to find Account and Associate with EmployeeID"
			Write-LogEntry -Severity 2 -Value "Attempting to find Account and Associate with EmployeeID"
			$ExistingStudent = Get-ADUser -Identity $Student.SamAccountName -Properties EmployeeID, Surname, GivenName
			if ($ExistingStudent.EmployeeID -eq $null) {
				Set-ADUser -Identity $Student.SamAccountName -EmployeeID $Student.EmployeeID
				Write-Verbose "Added EmployeeID to User."
				Write-LogEntry -Severity 2 -Value "Added EmployeeID to User"
				$script:userDataImport.Notify = $false
				continue
			} else {
				Write-Verbose "Users EmployeeID Attribute was not null. Didn't overwrite Attribute."
				Write-LogEntry -Severity 3 -Value "Users EmployeeID Attribute was not null. Didn't overwrite Attribute."
				$script:userDataImport.Notify = $false
				continue
			}
		} catch {
			Write-Error ("Error: {0}" -f $_.Exception.Message)
			$script:userDataImport.Notify = $false
			continue
		}
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function New-StudentUserGoogle {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$EXEPath,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Oauth2Path,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$EmailAddress,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$FirstName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$LastName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Password
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		$env:OAUTHFILE = $Oauth2Path
		try {
			if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {
				$p = (Start-Process -FilePath $EXEPath -ArgumentList "create user $emailAddress firstname $($firstName -replace ' ', '') lastname $($lastName -replace ' ', '') password $password" -NoNewWindow -Wait -PassThru)
			} else {
				$p = (Start-Process -FilePath $EXEPath -ArgumentList "create user $emailAddress firstname $($firstName -replace ' ', '') lastname $($lastName -replace ' ', '') password $password" -WindowStyle Hidden -Wait -PassThru)
			}
			if (($p.ExitCode -ne '0') -or ($p.ExitCode -ne '409')) { throw "GAM error exit $($p.exitcode)" }
		} catch {
			Write-Error ("Error: {0}" -f $_.Exception.Message)
		}
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function New-UserShare {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$SAMAccountName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$PathOnDrive,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$DriveLetter,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$NetBIOSDomainName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$HomeDirectoryServer,
		[Parameter(Mandatory = $true)]
		[string]$District
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		switch ($District) {
			default {
				$sharename = $SAMAccountName, '$' -join ''
				$shares = [WMICLASS]"\\$homeDirectoryServer\root\cimv2:WIN32_Share"
				$homeDirectory = '\\', $homeDirectoryServer, '\', $sharename -join ''
				$pathUNC = '\\', $homeDirectoryServer, '\', $driveLetter, '$\', $pathonDrive, '\', $SAMAccountName -join ''
				$pathOnHomeDirectoryServer = $driveLetter, ':\', $pathonDrive, '\', $SAMAccountName -join ''
				
				$trustee = ([wmiclass]'Win32_trustee').psbase.CreateInstance()
				$trustee.Domain = $NetBIOSDomainName
				$trustee.name = $SAMAccountName
				$fullcontrol = 2032127
				$change = 1245631
				$read = 1179785
				$ace = ([wmiclass]'Win32_ACE').psbase.CreateInstance()
				$ace.AccessMask = $fullcontrol
				$ace.AceFlags = 3
				$ace.AceType = 0
				$ace.Trustee = $trustee
				$sd = ([wmiclass]'Win32_SecurityDescriptor').psbase.CreateInstance()
				$sd.ControlFlags = 4
				$sd.DACL = $ace
				$sd.group = $trustee
				$sd.owner = $trustee
				
				Try {
					New-Item -path $pathUNC -ItemType directory -force | Out-Null
					icacls "$($pathUNC)" /grant "$($SAMAccountName):(OI)(CI)(M)" | Out-Null
					icacls "$($pathUNC)" /grant "Administrators:(OI)(CI)(F)" | Out-Null
					$shares.create($pathOnHomeDirectoryServer, $sharename, 0, 100, "", "", $sd) | Out-Null
				} catch {
					Write-Error ("Error: {0}" -f $_.Exception.Message)
					#log this
				}
			}
		}
		
		# DO NOT continue loop if home directory creation fails. The user already has been created and needs to be reported.
		
		$properties = @{
			HomeDirectory = $homeDirectory
		}
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Publish-FileSCP {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$DLLPath,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$RemoteFilePath,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$LocalFilePath,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$HostName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Username,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Password,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$SshHostKeyFingerprint
	)
	
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		try {
			Add-Type -Path $DLLPath
			$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
				Protocol = [WinSCP.Protocol]::Sftp
				HostName = $HostName
				UserName = $Username
				Password = $Password
				SshHostKeyFingerprint = $SshHostKeyFingerprint
			}
			
			$session = New-Object WinSCP.Session
			
			try {
				$session.Open($sessionOptions)
				$transferOptions = New-Object WinSCP.TransferOptions
				$transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
				$transferResult = $session.PutFiles($LocalFilePath, $RemoteFilePath, $False, $transferOptions)
				$transferResult.Check()
			} finally {
				$session.Dispose()
			}
		} catch [Exception]
		{
			Write-Host ("Error: {0}" -f $_.Exception.Message)
		}
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Read-ConfigXML {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$Path,
		[Parameter(Mandatory = $true)]
		[string]$District,
		[Parameter(Mandatory = $false)]
		[string]$Location,
		[switch]$Script,
		[switch]$ClearFile,
		[switch]$NewPassword,
		[switch]$SuspendExpiredAccounts,
		[switch]$MoveSuspendedAccounts,
		[switch]$SendReport,
		[switch]$NewUserShare,
		[switch]$Features,
		[switch]$PublishSynergyExport,
		[switch]$WriteSynergyExport,
		[switch]$NewStudentUserGoogle,
		[switch]$Notification,
		[switch]$GetSynergyImport
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		if (!($ConfigXMLObj)) { [System.Xml.XmlDocument]$script:ConfigXMLObj = Get-Content $Path }
		$Location = $Location.Split("\\\(\)\'./")[0]
		if ($Script) {
			$properties = @{
				UPNSuffix = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig" | Select-Object –ExpandProperty Node).UPNSuffix
				EmailSuffix = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig" | Select-Object –ExpandProperty Node).EmailSuffix
				studentsOUs = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node).Path | Get-Unique
				Locations = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node)
				SkipGrades = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Skip/Grades" | Select-Object –ExpandProperty Node).Grade
				SkipLocations = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Skip/Locations" | Select-Object –ExpandProperty Node).Location
				SkipStudents = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Skip/Students" | Select-Object –ExpandProperty Node).Student
				AccountExpirationDate = (get-date).AddDays(+ ([convert]::ToInt32((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig" | Select-Object –ExpandProperty Node).AccountExpirationDate)))
				PasswordNeverExpires = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password" | Select-Object –ExpandProperty Node).PasswordNeverExpires))
				CannotChangePassword = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password" | Select-Object –ExpandProperty Node).CannotChangePassword))
				ChangePasswordAtNextLogon = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password" | Select-Object –ExpandProperty Node).ChangePasswordAtNextLogon))
				ImportCSVPath = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Download" | Select-Object –ExpandProperty Node).PathLocal -join '\'
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($NewUserShare) {
			$properties = @{
				netBIOSDomainName = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig" | Select-Object –ExpandProperty Node).NetBIOSDomainName
				PathOnDrive	      = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name[contains(.,'$Location')]]/UserShare" | Select-Object –ExpandProperty Node).PathOnDrive
				DriveLetter	      = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name[contains(.,'$Location')]]/UserShare" | Select-Object –ExpandProperty Node).DriveLetter
				HomeDirectoryServer = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name[contains(.,'$location')]]/UserShare" | Select-Object –ExpandProperty Node).Server
				HomeDrive		  = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name[contains(.,'$location')]]" | Select-Object –ExpandProperty Node).HomeDrive
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($NewPassword) {
			$properties = @{
				Words = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password/Words" | Select-Object –ExpandProperty Node).Word
				SpecialCharacters = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password/SpecialCharacters" | Select-Object –ExpandProperty Node).Character
				Numbers = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password/Numbers" | Select-Object –ExpandProperty Node).Number
				UseDefaultPassword = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password" | Select-Object –ExpandProperty Node).UseDefaultPassword))
				DefaultPasswordIsStudentID = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password" | Select-Object –ExpandProperty Node).DefaultPasswordIsStudentID))
				DefaultPassword = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password" | Select-Object –ExpandProperty Node).DefaultPassword
				DOBPasswordGrade = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password" | Select-Object –ExpandProperty Node).DOBPasswordGrade
				DOBPasswordLocations = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password/DOBPasswordLocations" | Select-Object –ExpandProperty Node).Location
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($GetSynergyImport) {
			$properties = @{
				HostName = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Hostname
				UserName = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Username
				Password = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Password
				SshHostKeyFingerprint = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).HostKeyFingerprint
				RemoteFilePath = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Download" | Select-Object –ExpandProperty Node).PathRemote
				LocalFilePath = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Download" | Select-Object –ExpandProperty Node).PathLocal -join '\'
				DLLPath  = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP" | Select-Object –ExpandProperty Node).DLLPath -join '\'
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($PublishSynergyExport) {
			$properties = @{
				HostName = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Hostname
				UserName = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Username
				Password = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Password
				SshHostKeyFingerprint = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).HostKeyFingerprint
				RemoteFilePath = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Upload" | Select-Object –ExpandProperty Node).PathRemote
				LocalFilePath = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Upload" | Select-Object –ExpandProperty Node).PathLocal -join '\'
				DLLPath  = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP" | Select-Object –ExpandProperty Node).DLLPath -join '\'
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($NewStudentUserGoogle) {
			$properties = @{
				EXEPath = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/GAM" | Select-Object –ExpandProperty Node).EXEPath -join '\'
				Oauth2Path = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/GAM" | Select-Object –ExpandProperty Node).Oauth2Path -join '\'
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($ClearFile) {
			$properties = @{
				File = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/CleanupFiles" | Select-Object –ExpandProperty Node).File
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($SuspendExpiredAccounts) {
			$properties = @{
				OrganizationalUnit = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node).Path
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($MoveSuspendedAccounts) {
			$properties = @{
				MoveFromOU = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node).Path
				MoveToOU   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name=`'Disabled`']" | Select-Object –ExpandProperty Node).Path
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($SendReport) {
			$properties = @{
				SMTPServer = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Notifications" | Select-Object –ExpandProperty Node).SMTPServer
				From	   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Notifications" | Select-Object –ExpandProperty Node).From
				Body	   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Notifications" | Select-Object –ExpandProperty Node).Body
				Subject    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Notifications" | Select-Object –ExpandProperty Node).Subject
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($Notification) {
			$ConfigObj = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Notifications/Notification" | Select-Object –ExpandProperty Node)
			Write-Output $ConfigObj
		}
		if ($WriteSynergyExport) {
			$properties = @{
				OrganizationalUnit = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node).Path
				Path			   = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Upload" | Select-Object –ExpandProperty Node).PathLocal -join '\'
				LDAPAuth		   = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Features" | Select-Object –ExpandProperty Node).ExportSynergyLDAPAuth))
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($Features) {
			$properties = @{
				ExportSynergy = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Features" | Select-Object –ExpandProperty Node).ExportSynergy))
				GoogleAccount = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Features" | Select-Object –ExpandProperty Node).GoogleAccount))
				UserShare	  = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Features" | Select-Object –ExpandProperty Node).UserShare))
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Send-Report {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$StudentData,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$Recipient,
		[Parameter(Mandatory = $true)]
		[string]$SchoolName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$SMTPServer,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$From,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Subject,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Body
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		if (($StudentData.Office -replace '[^a-zA-Z0-9 ]', '') -like ('*', ($SchoolName -replace '[^a-zA-Z0-9 ]', ''), '*' -join '')) {
			$file = $PSScriptRoot, '\Files\', ($SchoolName -replace '[^a-zA-Z0-9 ]', ''), '.csv' -join ''
			$StudentData | Where-Object { ($_.Office -replace '[^a-zA-Z0-9 ]', '') -like ('*', ($SchoolName -replace '[^a-zA-Z0-9 ]', ''), '*' -join '') } | `
			Sort-Object -Descending -Property Office, Surname, GivenName | `
			Select-Object @{ Name = 'Student ID'; Expression = { $_.EmployeeID } }, `
						  @{ Name = 'First Name'; Expression = { $_.GivenName } }, `
						  @{ Name = 'Last Name'; Expression = { $_.Surname } }, `
						  @{ Name = 'School'; Expression = { $_.Office } }, `
						  @{ Name = 'Username'; Expression = { $_.SAMAccountName } }, `
						  @{ Name = 'Email Address'; Expression = { $_.EmailAddress } }, `
						  @{ Name = 'Grade'; Expression = { $_.Title } }, `
						  @{ Name = 'Grad Year'; Expression = { $_.personalTitle } }, `
						  @{ Name = 'Homeroom'; Expression = { $_.Division } }, `
						  Password | `
			Export-Csv -NoTypeInformation -Path $file
			foreach ($EmailAddress in $Recipient) { send-mailmessage -to $EmailAddress -from $From -subject $Subject -body $Body -Attachments $file -SmtpServer $SMTPServer }
		}
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Suspend-ExpiredAccounts {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$OrganizationalUnit
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		foreach ($DN in $OrganizationalUnit) {
			try {
				Search-ADAccount -SearchBase $DN -AccountExpired -UsersOnly | Where-Object { $_.Enabled } | Disable-ADAccount
			} catch {
				Write-Error ("Error: {0}" -f $_.Exception.Message)
			}
		}
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Test-OrganizationalUnitPath {
	[CmdletBinding(ConfirmImpact = 'None')]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$OrganizationalUnitDN
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		if (Get-ADOrganizationalUnit -Filter {
				distinguishedName -eq $OrganizationalUnitDN
			}) {
			$properties = @{
				Result = $true
			}
		} else {
			$properties = @{
				Result = $false
			}
		}
		
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Update-StudentUserAD {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[array]$DataAD,
		[Parameter(Mandatory = $true)]
		[array]$DataImport,
		[Parameter(Mandatory = $true)]
		[string]$District
	)
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		$notify = $false
		$script:userDataImport.EmailAddress = $DataAD.EmailAddress
		switch ($District) {
			SCAPPOOSE {
				if (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR')) {
					Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime ($DataImport.AccountExpirationDate).AddDays(+ 365)
				} else {
					Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime $DataImport.AccountExpirationDate
				}
				Set-ADUser -Identity $DataAD.SamAccountName -Title $DataImport.Title -Description $DataImport.Description -Department $DataImport.Department -Office $DataImport.Office
				if ($DataAD.ScriptPath -ne $DataImport.ScriptPath) { Set-ADUser -Identity $DataAD.SamAccountName -ScriptPath $DataImport.ScriptPath }
				if ($DataAD.HomeDirectory -ne $DataImport.HomeDirectory) { Set-ADUser -Identity $DataAD.SamAccountName -HomeDirectory $DataImport.HomeDirectory }
				if ($DataAD.HomeDrive -ne $DataImport.HomeDrive) { Set-ADUser -Identity $DataAD.SamAccountName -HomeDrive $DataImport.HomeDrive }
				if ($DataAD.personalTitle -ne $DataImport.personalTitle) {
					Set-ADUser -Identity $DataAD.SamAccountName -Clear personalTitle
					Set-ADUser -Identity $DataAD.SamAccountName -Add @{ 'personalTitle' = "$($DataImport.personalTitle)" }
				}
				if ($userDataAD.Enabled -eq $false) {
					$notify = $true
					try {
						Set-ADUser -Identity $DataAD.SamAccountName -Enabled $true
					} catch {
						Write-Error ("Error: {0}" -f $_.Exception.Message)
						#log this
						$notify = $false
						continue
					}
				}
				if ((($userDataAD.DistinguishedName -notlike ('*', $DataImport.Path -join '')) -and ($userDataAD.Office -ne $DataImport.Office)) -or ($userDataAD.DistinguishedName -like ('*Disabled*')) -or (($userDataAD.DistinguishedName -notlike ('*', $DataImport.Path -join '')) -and (($DataImport.SamAccountName.Substring($DataImport.SamAccountName.Length - 4) -ne $DataAD.SamAccountName.Substring($DataAD.SamAccountName.Length - 4)) -and (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR'))))) {
					$notify = $true
					if (!(Test-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path).Result) { New-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path -District $District }
					try { Move-ADObject -Identity $userDataAD.DistinguishedName -targetpath $DataImport.Path } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
				}
				if (($DataImport.SamAccountName.Substring($DataImport.SamAccountName.Length - 4) -ne $DataAD.SamAccountName.Substring($DataAD.SamAccountName.Length - 4)) -and (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR'))) {
					$notify = $true
					$script:userDataImport.EmailAddress = (New-EmailAddress -GivenName $DataImport.GivenName -Surname $DataImport.Surname -SamAccountName $DataImport.SamAccountName -EmailSuffix $ConfigScript.EmailSuffix -District $District).EmailAddress
					try { Set-ADUser -Identity $DataAD.SamAccountName -Add @{ 'proxyAddresses' = "$($DataAD.EmailAddress)" } } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
					try { Set-ADUser -Identity $DataAD.SamAccountName -EmailAddress $DataImport.EmailAddress -UserPrincipalName $DataImport.UserPrincipalName -HomeDirectory ('\\', $ConfigNUS.HomeDirectoryServer, '\', $DataImport.SamAccountName, '$' -join '') -SamAccountName $DataImport.SamAccountName -PassThru | Rename-ADObject -NewName $DataImport.SamAccountName } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
					try { Move-Item -Path ('\\', $ConfigNUS.HomeDirectoryServer, '\', $ConfigNUS.DriveLetter, '$\', $ConfigNUS.PathOnDrive, '\', $DataAD.SamAccountName -join '') -Destination ('\\', $ConfigNUS.HomeDirectoryServer, '\', $ConfigNUS.DriveLetter, '$\', $ConfigNUS.PathOnDrive, '\', $DataImport.SamAccountName -join '') } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
				}
			}
			NWRESD {
				#if ($DataImport.Office -Like 'Levi*') { Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime ($DataImport.AccountExpirationDate).AddDays(- 30) } elseif (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR')) { Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime ($DataImport.AccountExpirationDate).AddDays(+ 365) } else { Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime $DataImport.AccountExpirationDate }
				if (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR')) { Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime ($DataImport.AccountExpirationDate).AddDays(+ 365) } else { Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime $DataImport.AccountExpirationDate }
				Set-ADUser -Identity $DataAD.SamAccountName -Title $DataImport.Title -Description $DataImport.Description -Department $DataImport.Department -Office $DataImport.Office
				if (($DataImport.Initials) -and ($DataAD.Initials -ne $DataImport.Initials)) { Set-ADUser -Identity $DataAD.SamAccountName -Initials $DataImport.Initials }
				if ($DataAD.ScriptPath -ne $DataImport.ScriptPath) { Set-ADUser -Identity $DataAD.SamAccountName -ScriptPath $DataImport.ScriptPath }
				if ($DataAD.HomeDirectory -ne $DataImport.HomeDirectory) { Set-ADUser -Identity $DataAD.SamAccountName -HomeDirectory $DataImport.HomeDirectory }
				if ($DataAD.HomeDrive -ne $DataImport.HomeDrive) { Set-ADUser -Identity $DataAD.SamAccountName -HomeDrive $DataImport.HomeDrive }
				if ($DataAD.personalTitle -ne $DataImport.personalTitle) {
					Set-ADUser -Identity $DataAD.SamAccountName -Clear personalTitle
					Set-ADUser -Identity $DataAD.SamAccountName -Add @{ 'personalTitle' = "$($DataImport.personalTitle)" }
				}
				#if ($DataAD.DisplayName -ne $DataImport.DisplayName) { Set-ADUser -Identity $DataAD.SamAccountName -DisplayName $DataImport.DisplayName }
				if ($userDataAD.Enabled -eq $false) {
					$notify = $true
					try {
						Set-ADUser -Identity $DataAD.SamAccountName -Enabled $true
					} catch {
						Write-Error ("Error: {0}" -f $_.Exception.Message)
						$notify = $false
						continue
					}
				}
				if ((($userDataAD.DistinguishedName -notlike ('*', $DataImport.Path -join '')) -and ($userDataAD.Office -ne $DataImport.Office)) -or ($userDataAD.DistinguishedName -like ('*Disabled*')) -or (($userDataAD.DistinguishedName -notlike ('*', $DataImport.Path -join '')) -and (($DataImport.SamAccountName.Substring($DataImport.SamAccountName.Length - 4) -ne $DataAD.SamAccountName.Substring($DataAD.SamAccountName.Length - 4)) -and (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR'))))) {
					$notify = $true
					if (!(Test-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path).Result) { New-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path -District $District }
					try { Move-ADObject -Identity $userDataAD.DistinguishedName -targetpath $DataImport.Path } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
					try { Set-ADUser -Identity $DataAD.SamAccountName -HomeDirectory ('\\', $ConfigNUS.HomeDirectoryServer, '\', $DataImport.SamAccountName, '$' -join '') } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
					try { Move-Item -Path ('\\', $ConfigNUS.HomeDirectoryServer, '\', $ConfigNUS.DriveLetter, '$\', $ConfigNUS.PathOnDrive, '\', $DataAD.SamAccountName -join '') -Destination ('\\', $ConfigNUS.HomeDirectoryServer, '\', $ConfigNUS.DriveLetter, '$\', $ConfigNUS.PathOnDrive, '\', $DataImport.SamAccountName -join '') } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
				}
				if (($DataImport.SamAccountName.Substring($DataImport.SamAccountName.Length - 4) -ne $DataAD.SamAccountName.Substring($DataAD.SamAccountName.Length - 4)) -and (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR'))) {
					$notify = $true
					$script:userDataImport.EmailAddress = (New-EmailAddress -GivenName $DataImport.GivenName -Surname $DataImport.Surname -SamAccountName $DataImport.SamAccountName -EmailSuffix $ConfigScript.EmailSuffix -District $District).EmailAddress
					try { Set-ADUser -Identity $DataAD.SamAccountName -Add @{ 'proxyAddresses' = "$($DataAD.EmailAddress)" } } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
					try { Set-ADUser -Identity $DataAD.SamAccountName -EmailAddress $DataImport.EmailAddress -UserPrincipalName $DataImport.UserPrincipalName -HomeDirectory ('\\', $ConfigNUS.HomeDirectoryServer, '\', $DataImport.SamAccountName, '$' -join '') -SamAccountName $DataImport.SamAccountName -PassThru | Rename-ADObject -NewName $DataImport.SamAccountName } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
					try { Move-Item -Path ('\\', $ConfigNUS.HomeDirectoryServer, '\', $ConfigNUS.DriveLetter, '$\', $ConfigNUS.PathOnDrive, '\', $DataAD.SamAccountName -join '') -Destination ('\\', $ConfigNUS.HomeDirectoryServer, '\', $ConfigNUS.DriveLetter, '$\', $ConfigNUS.PathOnDrive, '\', $DataImport.SamAccountName -join '') } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
				}
			}
			SEASIDE {
				Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime $DataImport.AccountExpirationDate
				Set-ADUser -Identity $DataAD.SamAccountName -Title $DataImport.Title -Description $DataImport.Description -Department $DataImport.Department -Office $DataImport.Office
				if ($DataAD.Surname -ne $DataImport.Surname) { Set-ADUser -Identity $DataAD.SamAccountName -Surname $DataImport.Surname -DisplayName $DataImport.DisplayName }
				if ($DataAD.personalTitle -ne $DataImport.personalTitle) {
					Set-ADUser -Identity $DataAD.SamAccountName -Clear personalTitle
					Set-ADUser -Identity $DataAD.SamAccountName -Add @{ 'personalTitle' = "$($DataImport.personalTitle)" }
				}
				if ($userDataAD.Enabled -eq $false) {
					Write-LogEntry -Severity 1 -Value "Enabling Account"
					$notify = $true
					try {
						Set-ADUser -Identity $DataAD.SamAccountName -Enabled $true
					} catch {
						Write-Error ("Error: {0}" -f $_.Exception.Message)
						$notify = $false
						continue
					}
				}
				if (($userDataAD.DistinguishedName -like ('*Disabled*'))) {
					Write-LogEntry -Severity 2 -Value "Moving Account"
					$notify = $true
					if (!(Test-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path).Result) { New-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path -District $District }
					try { Move-ADObject -Identity $userDataAD.DistinguishedName -targetpath $DataImport.Path } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
				}
				if (($DataImport.SamAccountName.Substring(0, 2) -ne $DataAD.SamAccountName.Substring(0, 2)) -and (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR'))) {
					Write-LogEntry -Severity 2 -Value "Moving Account"
					$notify = $true
					$script:userDataImport.EmailAddress = (New-EmailAddress -GivenName $DataImport.GivenName -Surname $DataImport.Surname -SamAccountName $DataImport.SamAccountName -EmailSuffix $ConfigScript.EmailSuffix -District $District).EmailAddress
					try { Set-ADUser -Identity $DataAD.SamAccountName -Add @{ 'proxyAddresses' = "$($DataAD.EmailAddress)" } } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
					try { Set-ADUser -Identity $DataAD.SamAccountName -EmailAddress $DataImport.EmailAddress -UserPrincipalName $DataImport.UserPrincipalName -SamAccountName $DataImport.SamAccountName -PassThru | Rename-ADObject -NewName $DataImport.SamAccountName } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
				}
			}
			YAMHILL {
				if (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR')) {
					Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime ($DataImport.AccountExpirationDate).AddDays(+ 365)
				} else {
					Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime $DataImport.AccountExpirationDate
				}
				Set-ADUser -Identity $DataAD.SamAccountName -Title $DataImport.Title -Description $DataImport.Description -Department $DataImport.Department -Office $DataImport.Office
				if ($DataAD.Surname -ne $DataImport.Surname) { Set-ADUser -Identity $DataAD.SamAccountName -Surname $DataImport.Surname -DisplayName $DataImport.DisplayName }
				if ($DataAD.personalTitle -ne $DataImport.personalTitle) {
					Set-ADUser -Identity $DataAD.SamAccountName -Clear personalTitle
					Set-ADUser -Identity $DataAD.SamAccountName -Add @{ 'personalTitle' = "$($DataImport.personalTitle)" }
				}
				if ($userDataAD.Enabled -eq $false) {
					$notify = $true
					try {
						Set-ADUser -Identity $DataAD.SamAccountName -Enabled $true
					} catch {
						Write-Error ("Error: {0}" -f $_.Exception.Message)
						$notify = $false
						continue
					}
				}
				if ((($userDataAD.DistinguishedName -notlike ('*', $DataImport.Path -join '')) -and ($userDataAD.Office -ne $DataImport.Office)) -or ($userDataAD.DistinguishedName -like ('*Disabled*')) -or (($userDataAD.DistinguishedName -notlike ('*', $DataImport.Path -join '')) -and (($DataImport.SamAccountName.Substring($DataImport.SamAccountName.Length - 2) -ne $DataAD.SamAccountName.Substring($DataAD.SamAccountName.Length - 2)) -and (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR'))))) {
					$notify = $true
					if (!(Test-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path).Result) { New-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path -District $District }
					try { Move-ADObject -Identity $userDataAD.DistinguishedName -targetpath $DataImport.Path } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
				}
			}
			default {
				if (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR')) {
					Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime ($DataImport.AccountExpirationDate).AddDays(+ 365)
				} else {
					Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime $DataImport.AccountExpirationDate
				}
				Set-ADUser -Identity $DataAD.SamAccountName -Title $DataImport.Title -Description $DataImport.Description -Department $DataImport.Department -Office $DataImport.Office
				if ($DataAD.Surname -ne $DataImport.Surname) { Set-ADUser -Identity $DataAD.SamAccountName -Surname $DataImport.Surname -DisplayName $DataImport.DisplayName }
				if ($DataAD.personalTitle -ne $DataImport.personalTitle) {
					Set-ADUser -Identity $DataAD.SamAccountName -Clear personalTitle
					Set-ADUser -Identity $DataAD.SamAccountName -Add @{ 'personalTitle' = "$($DataImport.personalTitle)" }
				}
				if ($userDataAD.Enabled -eq $false) {
					$notify = $true
					try {
						Set-ADUser -Identity $DataAD.SamAccountName -Enabled $true
					} catch {
						Write-Error ("Error: {0}" -f $_.Exception.Message)
						$notify = $false
						continue
					}
				}
				if ((($userDataAD.DistinguishedName -notlike ('*', $DataImport.Path -join '')) -and ($userDataAD.Office -ne $DataImport.Office)) -or ($userDataAD.DistinguishedName -like ('*Disabled*')) -or (($userDataAD.DistinguishedName -notlike ('*', $DataImport.Path -join '')) -and (($DataImport.SamAccountName.Substring($DataImport.SamAccountName.Length - 2) -ne $DataAD.SamAccountName.Substring($DataAD.SamAccountName.Length - 2)) -and (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR'))))) {
					$notify = $true
					if (!(Test-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path).Result) { New-OrganizationalUnitPath -OrganizationalUnitDN $DataImport.Path -District $District }
					try { Move-ADObject -Identity $userDataAD.DistinguishedName -targetpath $DataImport.Path } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
				}
				if (($DataImport.SamAccountName.Substring($DataImport.SamAccountName.Length - 2) -ne $DataAD.SamAccountName.Substring($DataAD.SamAccountName.Length - 2)) -and (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR'))) {
					$notify = $true
					$script:userDataImport.EmailAddress = (New-EmailAddress -GivenName $DataImport.GivenName -Surname $DataImport.Surname -SamAccountName $DataImport.SamAccountName -EmailSuffix $ConfigScript.EmailSuffix -District $District).EmailAddress
					try { Set-ADUser -Identity $DataAD.SamAccountName -Add @{ 'proxyAddresses' = "$($DataAD.EmailAddress)" } } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
					try { Set-ADUser -Identity $DataAD.SamAccountName -EmailAddress $DataImport.EmailAddress -UserPrincipalName $DataImport.UserPrincipalName -SamAccountName $DataImport.SamAccountName -PassThru | Rename-ADObject -NewName $DataImport.SamAccountName } catch { Write-Error ("Error: {0}" -f $_.Exception.Message) }
				}
			}
		}
		$properties = @{
			Notify = $notify
		}
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	end {
		$FunctionStopWatch.Stop()
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
function Write-LogEntry {
	param
	(
		[Parameter(Mandatory = $true,
				   HelpMessage = 'Value added to the log file.')]
		[ValidateNotNullOrEmpty()]
		[string]$Value,
		[Parameter(Mandatory = $true,
				   HelpMessage = 'Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.')]
		[ValidateSet('1', '2', '3')]
		[ValidateNotNullOrEmpty()]
		[string]$Severity,
		[Parameter(Mandatory = $false,
				   HelpMessage = 'Name of the log file that the entry will written to.')]
		[ValidateNotNullOrEmpty()]
		[string]$FileName = "StudentImport.log"
	)
	
	# Determine log file location
	$global:LogFilePath = Join-Path -Path 'Script:\Files' -ChildPath $FileName
	
	# Construct time stamp for log entry
	$Time = -join @((Get-Date -Format "HH:mm:ss.fff"), "+", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
	
	# Construct date for log entry
	$Date = (Get-Date -Format "MM-dd-yyyy")
	
	# Construct context for log entry
	$Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
	
	# Construct final log entry
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""StudentImport.ps1"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
	
	# Add value to log file
	try {
		Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $global:LogFilePath -ErrorAction Stop
	} catch [System.Exception] {
		Write-Warning -Message "Unable to append log entry to $FileName file. Error message: $($_.Exception.Message)"
	}
}
function Write-SynergyExportFile {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$OrganizationalUnit,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$Path,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[boolean]$LDAPAuth,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$District
	)
	
	begin {
		Write-Verbose "Begin $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "Begin $($MyInvocation.MyCommand)"
		$FunctionStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	}
	process {
		
		switch ($District) {
			RIVERDALE {
				if ($LDAPAuth) {
					foreach ($DN in $OrganizationalUnit) {
						$users += Get-ADUser -SearchBase $DN -Filter { (Enabled -eq $true) } -Properties mail, employeeID, SamAccountName
					}
					$users | Where-Object { ($_.EmployeeID -match "^[\d\.]+$") -and ($_.Mail -ne $null) -and ($_.Title -match '[5-9]|10|11|12|TR') } | Select-Object employeeid, mail, SamAccountName | export-csv -delimiter "`t" -notypeinformation -append -path $Path
					(Get-Content $Path) | ForEach-Object { $_ -replace '"', '' } | out-file -FilePath $Path -Force -Encoding ascii
					(Get-Content $Path) | Select-Object -SkipLast 1 | Set-Content $Path
				} else {
					foreach ($DN in $OrganizationalUnit) {
						$users += Get-ADUser -SearchBase $DN -Filter { (Enabled -eq $true) } -Properties mail, employeeID, assistant
					}
					$users | Where-Object { ($_.EmployeeID -match "^[\d\.]+$") -and ($_.Mail -ne $null) -and ($_.Title -match '[5-9]|10|11|12|TR') } | Select-Object employeeid, mail, assistant | export-csv -delimiter "`t" -notypeinformation -append -path $Path
					(Get-Content $Path) | ForEach-Object { $_ -replace '"', '' } | out-file -FilePath $Path -Force -Encoding ascii
					(Get-Content $Path) | Select-Object -SkipLast 1 | Set-Content $Path
				}
			}
			default {
				if ($LDAPAuth) {
					foreach ($DN in $OrganizationalUnit | Sort-Object | Get-Unique ) {
						$users += Get-ADUser -SearchBase $DN -Filter { (Enabled -eq $true) } -Properties mail, employeeID, SamAccountName
					}
					$users | Where-Object { ($_.EmployeeID -match "^[\d\.]+$") -and ($_.Mail -ne $null) } | Select-Object employeeid, mail, SamAccountName | export-csv -delimiter "`t" -notypeinformation -append -path $Path
					(Get-Content $Path) | ForEach-Object { $_ -replace '"', '' } | out-file -FilePath $Path -Force -Encoding ascii
					(Get-Content $Path) | Select-Object -SkipLast 1 | Set-Content $Path
				} else {
					foreach ($DN in $OrganizationalUnit | Sort-Object | Get-Unique ) {
						$users += Get-ADUser -SearchBase $DN -Filter { (Enabled -eq $true) } -Properties mail, employeeID, assistant
					}
					$users | Where-Object { ($_.EmployeeID -match "^[\d\.]+$") -and ($_.Mail -ne $null) } | Select-Object employeeid, mail, assistant | export-csv -delimiter "`t" -notypeinformation -append -path $Path
					(Get-Content $Path) | ForEach-Object { $_ -replace '"', '' } | out-file -FilePath $Path -Force -Encoding ascii
					(Get-Content $Path) | Select-Object -SkipLast 1 | Set-Content $Path
				}
			}
		}
	}
	end {
		Write-Verbose "$($MyInvocation.MyCommand) Took $($FunctionStopWatch.Elapsed.TotalMilliseconds) Milliseconds"
		Write-Verbose "End $($MyInvocation.MyCommand)"
		Write-LogEntry -Severity 1 -Value "End $($MyInvocation.MyCommand)"
	}
}
#endregion Functions
#region Begin
try {
	$ScriptStopWatch = [system.diagnostics.stopwatch]::StartNew()
	Write-Verbose "Script Start"
	Read-ConfigXML -District $District -Path $ConfigXML -ClearFile | Clear-File
	Read-ConfigXML -District $District -Path $ConfigXML -GetSynergyImport | Get-FileSCP
	$ConfigScript = Read-ConfigXML -District $District -Path $ConfigXML -Script
	$ConfigF = Read-ConfigXML -District $District -Path $ConfigXML -Features
	$userdata = Import-Csv -Path $ConfigScript.ImportCSVPath | Sort-Object -Property SIS_NUMBER
	#endregion Begin
	#region Process
	foreach ($row in $userdata) {
		$rowStopWatch = [system.diagnostics.stopwatch]::StartNew()
		Write-Verbose "Begin $($row.SIS_NUMBER)"
		Write-LogEntry -Severity 1 -Value "Begin $($row.SIS_NUMBER)"
		$currentUsers = [Collections.Generic.List[Object]](Get-AllUserDataAD -SearchBase ($ConfigScript.studentsOUs))
		if (($ConfigScript.SkipGrades -eq $row.Grade) -or (($ConfigScript.SkipLocations -replace '[^a-zA-Z0-9 ]', '') -eq ($row.LOCATION -replace '[^a-zA-Z0-9 ]', '')) -or ($ConfigScript.SkipStudents -eq $row.SIS_NUMBER)) {
			Write-Verbose "SKIP $($row.SIS_NUMBER)"
			Write-LogEntry -Severity 2 -Value "SKIP $($row.SIS_NUMBER)"
			continue
		}
		$index = $currentUsers.FindIndex({ $args[0].EmployeeID -eq $row.SIS_NUMBER })
		$script:userDataAD = $currentUsers[$index]
		
		$ConfigNUS = Read-ConfigXML -District $District -Location $row.LOCATION -Path $ConfigXML -NewUserShare
		
		Add-Member -InputObject $row -NotePropertyName Path -NotePropertyValue (($ConfigScript.Locations | Where-Object { ($_.Name -replace '[^a-zA-Z0-9 ]', '') -match ($row.LOCATION -replace '[^a-zA-Z0-9 ]', '') }).Path)
		Add-Member -InputObject $row -NotePropertyName ScriptPath -NotePropertyValue (($ConfigScript.Locations | Where-Object { ($_.Name -replace '[^a-zA-Z0-9 ]', '') -match ($row.LOCATION -replace '[^a-zA-Z0-9 ]', '') }).ScriptPath)
		Add-Member -InputObject $row -NotePropertyName PasswordNeverExpires -NotePropertyValue $ConfigScript.PasswordNeverExpires
		Add-Member -InputObject $row -NotePropertyName AccountExpirationDate -NotePropertyValue $ConfigScript.AccountExpirationDate
		Add-Member -InputObject $row -NotePropertyName CannotChangePassword -NotePropertyValue $ConfigScript.CannotChangePassword
		Add-Member -InputObject $row -NotePropertyName ChangePasswordAtNextLogon -NotePropertyValue $ConfigScript.ChangePasswordAtNextLogon
		Add-Member -InputObject $row -NotePropertyName HomeDrive -NotePropertyValue $ConfigNUS.HomeDrive
		
		$userDataImport = Get-UserDataImport -UserDataRaw $row
		
		if ($index -eq -1) {
			Write-Verbose "New User $($row.SIS_NUMBER)"
			Write-LogEntry -Severity 1 -Value "New User $($row.SIS_NUMBER)"
			$userDataAD = $null
			$userDataImport.Notify = $true
			$userDataImport.Path = (Get-OrganizationalUnitPath -Location $userDataImport.Office -GradYear $userDataImport.personalTitle.Substring(2) -Grade $userDataImport.Title -OrganizationalUnit $userDataImport.Path -District $District).Path
			if (!(Test-OrganizationalUnitPath -OrganizationalUnitDN $userDataImport.Path).Result) { New-OrganizationalUnitPath -OrganizationalUnitDN $userDataImport.Path -District $District }
			New-SamAccountName -GivenName $userDataImport.GivenName -SurName $userDataImport.Surname -GradYear $userDataImport.personalTitle.Substring(2) -UPNSuffiix $ConfigScript.UPNSuffix -Grade $userDataImport.Title -AllUsersAD $currentUsers -District $District | Foreach-Object { $userDataImport.Name = $_.Name; $userDataImport.UserPrincipalName = $_.UserPrincipalName; $userDataImport.SamAccountName = $_.SamAccountName }
			Read-ConfigXML -District $District -Path $ConfigXML -NewPassword | New-Password -EmployeeID $userDataImport.EmployeeID -DateOfBirth $userDataImport.DataOfBirth -Grade $userDataImport.Title -Office $userDataImport.Office -District $District | Foreach-Object { $userDataImport.Password = $_.Password; $userDataImport.PasswordCrypt = $_.PasswordCrypt; }
			$userDataImport.EmailAddress = (New-EmailAddress -GivenName $userDataImport.GivenName -Surname $userDataImport.Surname -SamAccountName $userDataImport.SamAccountName -EmailSuffix $ConfigScript.EmailSuffix -District $District).EmailAddress
			New-StudentUserAD -Student $userDataImport
			if ($ConfigF.GoogleAccount) { Read-ConfigXML -District $District -Path $ConfigXML -NewStudentUserGoogle | New-StudentUserGoogle -EmailAddress $userDataImport.EmailAddress -FirstName $userDataImport.GivenName -LastName $userDataImport.Surname -Password $userDataImport.Password }
			if (($ConfigF.UserShare) -and ($ConfigNUS.HomeDirectoryServer)) { $ConfigNUS | New-UserShare -SAMAccountName $userDataImport.SamAccountName -District $District | ForEach-Object { $userDataImport.HomeDirectory = $_.HomeDirectory; Set-ADUser -Identity $userDataImport.SamAccountName -HomeDirectory $_.HomeDirectory; Set-ADUser -Identity $userDataImport.SamAccountName -HomeDrive $userDataImport.HomeDrive } }
		} else {
			Write-Verbose "Update Existing User $($row.SIS_NUMBER)"
			Write-LogEntry -Severity 1 -Value "Update Existing User $($row.SIS_NUMBER)"
			$userDataImport.HomeDirectory = '\\', $ConfigNUS.HomeDirectoryServer, '\', $userDataAD.SamAccountName, '$' -join ''
			$userDataImport.Path = (Get-OrganizationalUnitPath -Location $userDataImport.Office -GradYear $userDataImport.personalTitle.Substring(2) -Grade $userDataImport.Title -OrganizationalUnit $userDataImport.Path -District $District).Path
			New-SamAccountName -GivenName $userDataAD.GivenName -SurName $userDataAD.Surname -GradYear $userDataImport.personalTitle.Substring(2) -UPNSuffiix $ConfigScript.UPNSuffix -Grade $userDataImport.Title -AllUsersAD $currentUsers -District $District | Foreach-Object { $userDataImport.Name = $_.Name; $userDataImport.UserPrincipalName = $_.UserPrincipalName; $userDataImport.SamAccountName = $_.SamAccountName }
			$userDataImport.Notify = (Update-StudentUserAD -DataAD $userDataAD -DataImport $userDataImport -District $District).Notify
		}
		if ($userDataImport.Notify -eq $true) {
			$output += $userDataImport
			Write-Verbose "Notify $($row.SIS_NUMBER)"
			Write-LogEntry -Severity 1 -Value "Notify $($row.SIS_NUMBER)"
		}
		$rowStopWatch.Stop()
		Write-Verbose "$($row.SIS_NUMBER) took $($rowStopWatch.ElapsedMilliseconds) Milliseconds."
		Write-Verbose "End $($row.SIS_NUMBER)"
		Write-LogEntry -Severity 1 -Value "End $($row.SIS_NUMBER)"
	}
	#endregion Process
	#region End
	if ($output) {
		$Notification = (Read-ConfigXML -District $District -Path $ConfigXML -Notification)
		foreach ($location in $Notification.Name) {
			$Notification | Where-Object { $_.Name -eq $location } | ForEach-Object { $Name = $_.Name; $Recipient = $_.Recipient }
			Read-ConfigXML -District $District -Path $ConfigXML -SendReport | Send-Report -StudentData $Output -SchoolName $Name -Recipient $Recipient
		}
	}
	
	Read-ConfigXML -District $District -Path $ConfigXML -SuspendExpiredAccounts | Suspend-ExpiredAccounts
	Read-ConfigXML -District $District -Path $ConfigXML -MoveSuspendedAccounts | Move-SuspendedAccounts
	
	if ($ConfigF.ExportSynergy) {
		Read-ConfigXML -District $District -Path $ConfigXML -WriteSynergyExport | Write-SynergyExportFile -District $District
		Read-ConfigXML -District $District -Path $ConfigXML -PublishSynergyExport | Publish-FileSCP
		
	}
} finally {
	Pop-Location
	Remove-PSDrive -Name Script
	$ScriptStopWatch.Stop()
	Write-Verbose "StudentImport took $($global:ScriptStopWatch.Elapsed.TotalMinutes) Minutes."
}
#endregion End


# SIG # Begin signature block
# MIId7QYJKoZIhvcNAQcCoIId3jCCHdoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBvifFACy/8XJ9D
# 5I4g/noAVvZAGT8Z6ce23q/lqAktIKCCGK8wggPQMIICuKADAgECAhBWDzlgMpdv
# skWlu9PmVks2MA0GCSqGSIb3DQEBCwUAMFoxEzARBgoJkiaJk/IsZAEZFgNvcmcx
# GzAZBgoJkiaJk/IsZAEZFgtjYXNjYWRldGVjaDEVMBMGCgmSJomT8ixkARkWBWlu
# dHJhMQ8wDQYDVQQDEwZDVEEtQ0EwHhcNMTUwODE4MjIwMDI3WhcNMzUwODE4MjIy
# MDMwWjBaMRMwEQYKCZImiZPyLGQBGRYDb3JnMRswGQYKCZImiZPyLGQBGRYLY2Fz
# Y2FkZXRlY2gxFTATBgoJkiaJk/IsZAEZFgVpbnRyYTEPMA0GA1UEAxMGQ1RBLUNB
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAnwUnO1/doXUEwlU3eS8B
# koKLxPCNr54RnKXIS9abNqrUn+EzeDpOxsmjomYaErLsQGrSeO/k93kCLtBVhTx2
# OaeSt3klYk88PVS2jHmTkowwvkxl/Nk/L7941lgeBE5YNmugSjGjvcVRvzC5Hd2n
# GaNj1SLWwDw3rwDmifBY0rcs140RE7T3Ms4pquIejfzk4CYf9M3cEKEVLDwgnN7L
# yJVVd1Wj4M472nwdU9XHjMqLTAds0258iGqFooWa8cdNRGP1F57bLn4wSK9wJXfb
# ydpgnXWkFsjb8uEiagjxBkXaR6M9uldrGaDN0o0XpP4xMBLeQcNMuhEH0EB4Joto
# cwIDAQABo4GRMIGOMBMGCSsGAQQBgjcUAgQGHgQAQwBBMA4GA1UdDwEB/wQEAwIB
# hjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBRFtRVw0jTNG1WszSXHh+qWu8Il
# nzASBgkrBgEEAYI3FQEEBQIDAgADMCMGCSsGAQQBgjcVAgQWBBSnN6gbSu96Heb+
# 2FbsX4B6hpdPijANBgkqhkiG9w0BAQsFAAOCAQEAC4FQJ3s5HFFiLD1z8XV63o6J
# wx0NWWDeOwcw7EbXmp72bm0QQLoFGpQsM5GTLdiu+HtGxrtUj6uQNKWgpWUp+Koy
# PEy9JoCEencFaPIyrG9iYTc1UXgwz719RVZ0h1/QloSGFXzgNxGNXkLGw1FsBmzg
# +DnwSUEKHt7yICi7LOKHju6qauhHg9T2sNNShB0yZaoMvRekXPXMWw4k+ccDdgcW
# MOF9VkdAvBFXx5BVoE6GuUZRYRVcQMH2rc0eGddfvZ3ZpVmK4DdBssNU47CmNix2
# mCaDzLc6mGuzMtKYnkzsh3+G2hj3kOQ+2x0D9eoBzyPhE+rmFw2eIDTCML2N+TCC
# BBQwggL8oAMCAQICCwQAAAAAAS9O4VLXMA0GCSqGSIb3DQEBBQUAMFcxCzAJBgNV
# BAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMRAwDgYDVQQLEwdSb290
# IENBMRswGQYDVQQDExJHbG9iYWxTaWduIFJvb3QgQ0EwHhcNMTEwNDEzMTAwMDAw
# WhcNMjgwMTI4MTIwMDAwWjBSMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFs
# U2lnbiBudi1zYTEoMCYGA1UEAxMfR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0Eg
# LSBHMjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJTvZfi1V5+gUw00
# BusJH7dHGGrL8Fvk/yelNNH3iRq/nrHNEkFuZtSBoIWLZFpGL5mgjXex4rxc3SLX
# amfQu+jKdN6LTw2wUuWQW+tHDvHnn5wLkGU+F5YwRXJtOaEXNsq5oIwbTwgZ9oEx
# rWEWpGLmtECew/z7lfb7tS6VgZjg78Xr2AJZeHf3quNSa1CRKcX8982TZdJgYSLy
# Bvsy3RZR+g79ijDwFwmnu/MErquQ52zfeqn078RiJ19vmW04dKoRi9rfxxRM6YWy
# 7MJ9SiaP51a6puDPklOAdPQD7GiyYLyEIACDG6HutHQFwSmOYtBHsfrwU8wY+S47
# +XB+tCUCAwEAAaOB5TCB4jAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB
# /wIBADAdBgNVHQ4EFgQURtg+/9zjvv+D5vSFm7DdatYUqcEwRwYDVR0gBEAwPjA8
# BgRVHSAAMDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29t
# L3JlcG9zaXRvcnkvMDMGA1UdHwQsMCowKKAmoCSGImh0dHA6Ly9jcmwuZ2xvYmFs
# c2lnbi5uZXQvcm9vdC5jcmwwHwYDVR0jBBgwFoAUYHtmGkUNl8qJUC99BM00qP/8
# /UswDQYJKoZIhvcNAQEFBQADggEBAE5eVpAeRrTZSTHzuxc5KBvCFt39QdwJBQSb
# b7KimtaZLkCZAFW16j+lIHbThjTUF8xVOseC7u+ourzYBp8VUN/NFntSOgLXGRr9
# r/B4XOBLxRjfOiQe2qy4qVgEAgcw27ASXv4xvvAESPTwcPg6XlaDzz37Dbz0xe2X
# nbnU26UnhOM4m4unNYZEIKQ7baRqC6GD/Sjr2u8o9syIXfsKOwCr4CHr4i81bA+O
# NEWX66L3mTM1fsuairtFTec/n8LZivplsm7HfmX/6JLhLDGi97AnNkiPJm877k12
# H3nD5X+WNbwtDswBsI5//1GAgKeS1LNERmSMh08WYwcxS2Ow3/MwggSfMIIDh6AD
# AgECAhIRIdaZp2SXPvH4Qn7pGcxTQRQwDQYJKoZIhvcNAQEFBQAwUjELMAkGA1UE
# BhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExKDAmBgNVBAMTH0dsb2Jh
# bFNpZ24gVGltZXN0YW1waW5nIENBIC0gRzIwHhcNMTYwNTI0MDAwMDAwWhcNMjcw
# NjI0MDAwMDAwWjBgMQswCQYDVQQGEwJTRzEfMB0GA1UEChMWR01PIEdsb2JhbFNp
# Z24gUHRlIEx0ZDEwMC4GA1UEAxMnR2xvYmFsU2lnbiBUU0EgZm9yIE1TIEF1dGhl
# bnRpY29kZSAtIEcyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAsBeu
# otO2BDBWHlgPse1VpNZUy9j2czrsXV6rJf02pfqEw2FAxUa1WVI7QqIuXxNiEKlb
# 5nPWkiWxfSPjBrOHOg5D8NcAiVOiETFSKG5dQHI88gl3p0mSl9RskKB2p/243LOd
# 8gdgLE9YmABr0xVU4Prd/4AsXximmP/Uq+yhRVmyLm9iXeDZGayLV5yoJivZF6UQ
# 0kcIGnAsM4t/aIAqtaFda92NAgIpA6p8N7u7KU49U5OzpvqP0liTFUy5LauAo6Ml
# +6/3CGSwekQPXBDXX2E3qk5r09JTJZ2Cc/os+XKwqRk5KlD6qdA8OsroW+/1X1H0
# +QrZlzXeaoXmIwRCrwIDAQABo4IBXzCCAVswDgYDVR0PAQH/BAQDAgeAMEwGA1Ud
# IARFMEMwQQYJKwYBBAGgMgEeMDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmds
# b2JhbHNpZ24uY29tL3JlcG9zaXRvcnkvMAkGA1UdEwQCMAAwFgYDVR0lAQH/BAww
# CgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4YxaHR0cDovL2NybC5nbG9iYWxz
# aWduLmNvbS9ncy9nc3RpbWVzdGFtcGluZ2cyLmNybDBUBggrBgEFBQcBAQRIMEYw
# RAYIKwYBBQUHMAKGOGh0dHA6Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0
# L2dzdGltZXN0YW1waW5nZzIuY3J0MB0GA1UdDgQWBBTUooRKOFoYf7pPMFC9ndV6
# h9YJ9zAfBgNVHSMEGDAWgBRG2D7/3OO+/4Pm9IWbsN1q1hSpwTANBgkqhkiG9w0B
# AQUFAAOCAQEAj6kakW0EpjcgDoOW3iPTa24fbt1kPWghIrX4RzZpjuGlRcckoiK3
# KQnMVFquxrzNY46zPVBI5bTMrs2SjZ4oixNKEaq9o+/Tsjb8tKFyv22XY3mMRLxw
# L37zvN2CU6sa9uv6HJe8tjecpBwwvKu8LUc235IgA+hxxlj2dQWaNPALWVqCRDSq
# gOQvhPZHXZbJtsrKnbemuuRQ09Q3uLogDtDTkipbxFm7oW3bPM5EncE4Kq3jjb3N
# CXcaEL5nCgI2ZIi5sxsm7ueeYMRGqLxhM2zPTrmcuWrwnzf+tT1PmtNN/94gjk6X
# pv2fCbxNyhh2ybBNhVDygNIdBvVYBAexGDCCBbYwggSeoAMCAQICE10AAAJ8Xj/h
# LaszG+AAAwAAAnwwDQYJKoZIhvcNAQELBQAwWjETMBEGCgmSJomT8ixkARkWA29y
# ZzEbMBkGCgmSJomT8ixkARkWC2Nhc2NhZGV0ZWNoMRUwEwYKCZImiZPyLGQBGRYF
# aW50cmExDzANBgNVBAMTBkNUQS1DQTAeFw0xNzA5MjcyMTM4MDRaFw0yMjA5MjYy
# MTM4MDRaMF4xEzARBgoJkiaJk/IsZAEZFgNvcmcxGzAZBgoJkiaJk/IsZAEZFgtj
# YXNjYWRldGVjaDEVMBMGCgmSJomT8ixkARkWBWludHJhMRMwEQYDVQQDEwpDVEEt
# SU5ULUNBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAylW9Cd94uZun
# nWpoNn4UNKhhdiPZr1mHroZQoscxW8Tp4p98UMNu3uACWz+VkNfWMvZqytuaAUVb
# zEP8q1JVDZ5huNlC7aDrRL+76j2S47UIOYzdMDu9k/USeY/bDASfTWmX+0u4aGxM
# +QuYqx0RbSCSgr8AXr0VDJ3p2Lr+9HcmAGzR6Zt7afUjZbcS0tbdOhL2tySVQI5C
# FjOchncp8QrJI6tk8Wdg5fegqYALvIWfVrVxeNMpz747NI0b5P/JOXktMuKS7nBa
# 87yQIL7p/eQfKE0FWEISHGbeOyVEyUUxBlgTQ6W0HocVyzjJDguOOzOSIVcGB+Hr
# X8V4uELXMwIDAQABo4ICbzCCAmswEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYE
# FGa92DKMwOeyGR6PViTrCjEGqCF8MBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBB
# MA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFEW1
# FXDSNM0bVazNJceH6pa7wiWfMIIBEQYDVR0fBIIBCDCCAQQwggEAoIH9oIH6hoG/
# bGRhcDovLy9DTj1DVEEtQ0EoMiksQ049Q1RBLUNBLTAxLENOPUNEUCxDTj1QdWJs
# aWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9u
# LERDPWludHJhLERDPWNhc2NhZGV0ZWNoLERDPW9yZz9jZXJ0aWZpY2F0ZVJldm9j
# YXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9pbnSG
# Nmh0dHA6Ly9jdGFjcmwuY2FzY2FkZXRlY2gub3JnL0NlcnRFbnJvbGwvQ1RBLUNB
# KDIpLmNybDCBxQYIKwYBBQUHAQEEgbgwgbUwgbIGCCsGAQUFBzAChoGlbGRhcDov
# Ly9DTj1DVEEtQ0EsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENO
# PVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9aW50cmEsREM9Y2FzY2FkZXRl
# Y2gsREM9b3JnP2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZp
# Y2F0aW9uQXV0aG9yaXR5MA0GCSqGSIb3DQEBCwUAA4IBAQAlyqwL7Qoxf9jmAb3W
# 03Jiga2XKzz8Lzi4gy4N/DFHOfgYEiTb+hD245PxV9VXuWUpeMa28f7yCHHJ32CI
# 5881ab/r1/gwXQlxqg0Nof74P+hv7AdJ3eQjo92mEyUt64ElA7S/n3LGDAiJIDTq
# O/gUOj1n3NTOMGdAeKSId9WcgeTxh5j+X/aNnbpRy58s2rt46KDv4jQrXWLZqHp9
# dzIyCNtfnqotXA2jePXDWYbD8Zw8dsCBwQw+2C/ktXS8Z/GrLLDJB82CtPzmeTKN
# pclDXz1vttVeASkVc7w668OUsGgyRojKvRonLlcr00nRgcO8SOtsgsJcbQSiIvKF
# 70afMIIGYjCCBUqgAwIBAgITTQAACNTm6lyP5isnXwAAAAAI1DANBgkqhkiG9w0B
# AQsFADBeMRMwEQYKCZImiZPyLGQBGRYDb3JnMRswGQYKCZImiZPyLGQBGRYLY2Fz
# Y2FkZXRlY2gxFTATBgoJkiaJk/IsZAEZFgVpbnRyYTETMBEGA1UEAxMKQ1RBLUlO
# VC1DQTAeFw0xODAxMjEyMjE5MjJaFw0xOTAxMjEyMjE5MjJaMIGaMRMwEQYKCZIm
# iZPyLGQBGRYDb3JnMRswGQYKCZImiZPyLGQBGRYLY2FzY2FkZXRlY2gxFTATBgoJ
# kiaJk/IsZAEZFgVpbnRyYTENMAsGA1UECxMETUVTRDEUMBIGA1UEAxMLRWRlbiBO
# ZWxzb24xKjAoBgkqhkiG9w0BCQEWG2VkZW4ubmVsc29uQGNhc2NhZGV0ZWNoLm9y
# ZzCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALkQFRlY5uvISFbe6fUG
# i3nj805Vwr/LahJVsURbNK+Kd7Pu7Mh9x9CAhiq2tEjfVzaMh6tG4ByYM/DGchHw
# Pqoco1kqak9Wh7KdVdoNROozbcfe9PFrYLCMbbi1x/LBRaQwh26o4jt3AGHpNnqn
# DuN1DwlKQAI67TyiGa9zq4Rqxv+1txLaR/spVpcWRJwBwRx7I4UlucuVObBnGGia
# 0ysfn3iMy1r3C17b8T84tSjS2q3uNuWV3ZnpvpzYaMI0MK7cQt3/OpjZlHWxx8ju
# zCeN3xVjQnAA0HX98IUI+MIhjun8sJq482VDN+7M7N7iISXV7xyJVEX8FjSKeAUg
# kWECAwEAAaOCAtowggLWMBcGCSsGAQQBgjcUAgQKHggAVQBzAGUAcjApBgNVHSUE
# IjAgBgorBgEEAYI3CgMEBggrBgEFBQcDBAYIKwYBBQUHAwIwDgYDVR0PAQH/BAQD
# AgWgMEQGCSqGSIb3DQEJDwQ3MDUwDgYIKoZIhvcNAwICAgCAMA4GCCqGSIb3DQME
# AgIAgDAHBgUrDgMCBzAKBggqhkiG9w0DBzAdBgNVHQ4EFgQU20Y/FbHR+bpIsRwi
# JxRcmuFP99EwHwYDVR0jBBgwFoAUZr3YMozA57IZHo9WJOsKMQaoIXwwgdcGA1Ud
# HwSBzzCBzDCByaCBxqCBw4aBwGxkYXA6Ly8vQ049Q1RBLUlOVC1DQSxDTj1DVEEt
# Q0EtMDIsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZp
# Y2VzLENOPUNvbmZpZ3VyYXRpb24sREM9aW50cmEsREM9Y2FzY2FkZXRlY2gsREM9
# b3JnP2NlcnRpZmljYXRlUmV2b2NhdGlvbkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1j
# UkxEaXN0cmlidXRpb25Qb2ludDCByQYIKwYBBQUHAQEEgbwwgbkwgbYGCCsGAQUF
# BzAChoGpbGRhcDovLy9DTj1DVEEtSU5ULUNBLENOPUFJQSxDTj1QdWJsaWMlMjBL
# ZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPWlu
# dHJhLERDPWNhc2NhZGV0ZWNoLERDPW9yZz9jQUNlcnRpZmljYXRlP2Jhc2U/b2Jq
# ZWN0Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhvcml0eTBUBgNVHREETTBLoCwGCisG
# AQQBgjcUAgOgHgwcbmVsc29uQGludHJhLmNhc2NhZGV0ZWNoLm9yZ4EbZWRlbi5u
# ZWxzb25AY2FzY2FkZXRlY2gub3JnMA0GCSqGSIb3DQEBCwUAA4IBAQCbdFYwSLZQ
# nXftd/4H6im6gtKTGi6yZ7e0kju+u9GQ/NSnGuHXta49Ayyltyh1N0GC7Yke1Q9c
# f96wiyCyqoEgcas4La6nLdbL3Hv/Y1CLm0coPEPmMnTaC24HFJe2kaHEo8euiRCL
# ohjNjifnKx5gx3KwlRShCjvD75zF3G0TH2gBAjUWZVaKpxrnqbWdS4+4g4tbjiGq
# nrBGe0aItCTPGspkPCrrzvoRjHiTBAwD5cfABdBSeajq9Qt387JxKDbkyo/xPVyU
# huE/jh6nB9SED4nASvuxJwG40LO/KBnjQqlS91QR3akz0lzrmBxFsuuj7Qd/1GwB
# pA291YzWj2b+MYIElDCCBJACAQEwdTBeMRMwEQYKCZImiZPyLGQBGRYDb3JnMRsw
# GQYKCZImiZPyLGQBGRYLY2FzY2FkZXRlY2gxFTATBgoJkiaJk/IsZAEZFgVpbnRy
# YTETMBEGA1UEAxMKQ1RBLUlOVC1DQQITTQAACNTm6lyP5isnXwAAAAAI1DANBglg
# hkgBZQMEAgEFAKBMMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMC8GCSqGSIb3
# DQEJBDEiBCCk3/h12I86XbNsrcXK5YdbkdwhBvtGroblIEYekfjKZzANBgkqhkiG
# 9w0BAQEFAASCAQBMA+VziTaIxjbLrGlAYtthSmtePCCItxKKDgHfefHK7A/YnDeU
# bsH++7/0UvYvqfqjP1jE3QCLSBwISNNRI6zYs639CFxzmu+jOZYQBd0op9WFTHac
# e4U2a/RxSt9IFdRjzddSB1Zj/eldDCu/+GPBdR6+jcyt8HERiP5Ked7XjpeWjQqA
# anenC1W9BTPFrrKWeGUmWVpl/IuAq82R6GKat60SX7LH10yoXR1h+n4MGKFeKyzl
# eDmXDKIzhDnz9n779g+ggkKIg/Q2zHovZRcl9ZBStVJZeePf0BSNO/VaAX4ytTIe
# Ibv1mGWUjMRRMdN9DCvWr6X9tahRiexC+1dhoYICojCCAp4GCSqGSIb3DQEJBjGC
# Ao8wggKLAgEBMGgwUjELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24g
# bnYtc2ExKDAmBgNVBAMTH0dsb2JhbFNpZ24gVGltZXN0YW1waW5nIENBIC0gRzIC
# EhEh1pmnZJc+8fhCfukZzFNBFDAJBgUrDgMCGgUAoIH9MBgGCSqGSIb3DQEJAzEL
# BgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE4MTAwODE4MTIxNlowIwYJKoZI
# hvcNAQkEMRYEFCEVRNpcagoCZ9hRliPi4ROqPdaGMIGdBgsqhkiG9w0BCRACDDGB
# jTCBijCBhzCBhAQUY7gvq2H1g5CWlQULACScUCkz7HkwbDBWpFQwUjELMAkGA1UE
# BhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExKDAmBgNVBAMTH0dsb2Jh
# bFNpZ24gVGltZXN0YW1waW5nIENBIC0gRzICEhEh1pmnZJc+8fhCfukZzFNBFDAN
# BgkqhkiG9w0BAQEFAASCAQBeE5ZN6aymABRZuEeItouD6Wi7JqDcFLPcQJZfgPJn
# XMQS82O6RltC7SpM9Wgk13HJbAu6jPfnskm3lWq3MFrWZfAbFGuLYkZSy9/4S3uX
# Xoe6mGX0Z/tkNAkpwlq6uPAPsy/WmiYSp3UgypEBpe1puQvDZzqaWV4HA5lkJW2Q
# i5g9qOnduSMWrd39z8ipOaMhDLO2PxTwKMOFiwcEYNdWkGIqNcDk1ksCTB+EI4kt
# s6/Ljnnv6K+UunwiCQa9Tn5Hd1sFqFF01YCgdVfIew35W1jL7hXJlS7BVP+ltt6M
# Ez7f+SFiXznaMIF8w9aXcQ1/F6Ead1sZ6k1dLsF+d52O
# SIG # End signature block
