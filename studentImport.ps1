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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
		foreach ($filename in $File) { if (Test-Path -Path ($PSScriptRoot, '\Files\', $filename -join '')) { Remove-Item ($PSScriptRoot, '\Files\', $filename -join '') -Force -Confirm:$false } }
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
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
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	BEGIN {
		Write-Verbose "BEGIN $($MyInvocation.MyCommand)"
	}
	PROCESS {
		try {
			Add-Type -Path $DLLPath
			$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
				Protocol    = [WinSCP.Protocol]::Sftp
				HostName    = $HostName
				UserName    = $Username
				Password    = $Password
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
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
		switch ($District) {
			NWRESD { $path = $OrganizationalUnit }
			default { $path = "OU=$GradYear", $OrganizationalUnit -join ',' }
		}
		$properties = @{
			Path	 = $path
		}
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
}
function Get-UserDataImport {
	[CmdletBinding()]
	param
	(
		$UserDataRaw
	)
	
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
		$properties = @{
			AccountExpirationDate	  = $UserDataRaw.AccountExpirationDate
			ChangePasswordAtNextLogon = $false
			PasswordNeverExpires	  = $UserDataRaw.PasswordNeverExpires
			CannotChangePassword	  = $UserDataRaw.CannotChangePassword
			ScriptPath			      = $UserDataRaw.ScriptPath
			DataOfBirth			      = $UserDataRaw.BIRTHDATE
			Department			      = 'Student'
			Description			      = "Last import: $todaysDate"
			DisplayName			      = $UserDataRaw.FIRST_NAME, $UserDataRaw.LAST_NAME -join ' '
			DistinguishedName		  = $null
			Division				  = $UserDataRaw.HOMEROOM_TCH
			EmailAddress			  = $null
			EmployeeID			      = $UserDataRaw.SIS_NUMBER
			Enabled				      = $true
			GivenName				  = $UserDataRaw.FIRST_NAME
			Initials				  = $UserDataRaw.MIDDLE_INIITAL
			HomeDirectory			  = $null
			HomeDrive				  = $UserDataRaw.HomeDrive
			MemberOf				  = $null
			Name					  = $null
			Office				      = $UserDataRaw.LOCATION
			Password				  = $null
			PasswordCrypt			  = $null
			Path					  = $UserDataRaw.Path
			personalTitle			  = $UserDataRaw.CALCULATED_GRAD_YEAR
			proxyAddresses		      = $null
			SamAccountName		      = $null
			Surname				      = $UserDataRaw.LAST_NAME
			Title					  = $UserDataRaw.GRADE
			UserPrincipalName		  = $null
			Notify				      = $false
		}
		$userDataImport = New-Object -TypeName PSObject -Property $properties
		Write-Output $userDataImport
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
		foreach ($DN in $MoveFromOU) {
			if ($DN -contains 'Disabled') { continue }
			try {
				Search-ADAccount –AccountDisabled –UsersOnly –SearchBase $DN | Move-ADObject –TargetPath $MoveToOU
			} catch {
				Write-Error ("Error: {0}" -f $_.Exception.Message)
			}
		}
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
	switch ($District) {
		default {
			$emailAddress = $SamAccountName, $EmailSuffix -join ''
		}
	}
	$properties = @{
		EmailAddress    = $emailAddress
	}
	$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
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
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
		[string]$DateOfBirth,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.Int16]$Grade,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$District,
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
		[string]$DefaultPassword,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[boolean]$UseDefaultPassword,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$EmployeeID,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[boolean]$DefaultPasswordIsStudentID,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[System.Int16]$DOBPasswordGrade,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		$DOBPasswordLocations
	)
	
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
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
			Password		 = $password
			PasswordCrypt    = (ConvertTo-SecureString $($password) -AsPlainText -Force)
		}
		
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
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
			default {
				if (((($Grade -eq '12') -or ($Grade -eq 'TR')) -and ($currentMonth -ge '08') -and ($GradYear -le $currentYear)) -or ($GradYear -lt $currentYear)) { $GradYear = ($currentYear+1) }
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
		SamAccountName	     = $samAccountName
		UserPrincipalName    = $samAccountName, $UPNSuffiix -join ''
		Name				 = $samAccountName
	}
	$obj = New-Object -TypeName PSObject -Property $properties
	Write-Output $obj
	}
	END {
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
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
				   -UserPrincipalName $Student.UserPrincipalName
			
			if ($Student.ChangePasswordAtNextLogon) { Set-ADUser -Identity $Student.SamAccountName -ChangePasswordAtLogon $true }
		} catch [Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException] {
			Write-Error ("Error: {0}" -f $_.Exception.Message)
			Write-Verbose "Attempting to find Account and Associate with EmployeeID"
			$ExistingStudent = Get-ADUser -Identity $Student.SamAccountName -Properties AccountExpirationDate,`
										  Department,`
										  DisplayName,`
										  DistinguishedName,`
										  Division,`
										  MemberOf,`
										  EmailAddress,`
										  EmployeeID,`
										  personalTitle,`
										  Office,`
										  Title,`
										  Surname,`
										  GivenName,`
										  Enabled,`
										  proxyAddresses,`
										  HomeDirectory,`
										  Initials`
			
			if ($ExistingStudent.EmployeeID -eq $null) {
				if (($ExistingStudent.GivenName.substring(0, 1) -eq $Student.GivenName.substring(0, 1)) -and ($ExistingStudent.Surname -like $Student.Surname)) {
					Set-ADUser -Identity $Student.SamAccountName -EmployeeID $Student.EmployeeID
					Write-Verbose "Added EmployeeID to User"
				}
			} else {
				continue
			}
		} catch {
			Write-Error ("Error: {0}" -f $_.Exception.Message)
			$script:userDataImport.Notify = $true
			continue
		}
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
		$env:OAUTHFILE = $Oauth2Path
		try {
			if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {
				$p = (Start-Process -FilePath $EXEPath -ArgumentList "create user $emailAddress firstname $($firstName -replace ' ','') lastname $($lastName -replace ' ','') password $password" -NoNewWindow -Wait -PassThru)
			} else {
				$p = (Start-Process -FilePath $EXEPath -ArgumentList "create user $emailAddress firstname $($firstName -replace ' ','') lastname $($lastName -replace ' ','') password $password" -WindowStyle Hidden -Wait -PassThru)
			}
			if (($p.ExitCode -ne '0') -or ($p.ExitCode -ne '409')) { throw "GAM error exit $($p.exitcode)" }
		} catch {
			Write-Error ("Error: {0}" -f $_.Exception.Message)
		}
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
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
			HomeDirectory	  = $homeDirectory
		}
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
		try {
			Add-Type -Path $DLLPath
			$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
				Protocol    = [WinSCP.Protocol]::Sftp
				HostName    = $HostName
				UserName    = $Username
				Password    = $Password
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
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
		if (!($ConfigXMLObj)) { [System.Xml.XmlDocument]$script:ConfigXMLObj = Get-Content $Path }
		$Location = $Location.Split("\\\(\)\'./")[0]
		if ($Script) {
			$properties = @{
				UPNSuffix	  = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig" | Select-Object –ExpandProperty Node).UPNSuffix
				EmailSuffix   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig" | Select-Object –ExpandProperty Node).EmailSuffix
				studentsOUs   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node).Path
				Locations	  = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node)
				SkipGrades    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Skip/Grades" | Select-Object –ExpandProperty Node).Grade
				SkipLocations = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Skip/Locations" | Select-Object –ExpandProperty Node).Location
				SkipStudents  = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Skip/Students" | Select-Object –ExpandProperty Node).Student
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
				netBIOSDomainName   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig" | Select-Object –ExpandProperty Node).NetBIOSDomainName
				PathOnDrive		    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name[contains(.,'$Location')]]/UserShare" | Select-Object –ExpandProperty Node).PathOnDrive
				DriveLetter		    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name[contains(.,'$Location')]]/UserShare" | Select-Object –ExpandProperty Node).DriveLetter
				HomeDirectoryServer = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name[contains(.,'$location')]]/UserShare" | Select-Object –ExpandProperty Node).Server
				HomeDrive		    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name[contains(.,'$location')]]" | Select-Object –ExpandProperty Node).HomeDrive
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($NewPassword) {
			$properties = @{
				Words		  = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password/Words" | Select-Object –ExpandProperty Node).Word
				SpecialCharacters = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password/SpecialCharacters" | Select-Object –ExpandProperty Node).Character
				Numbers	      = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Password/Numbers" | Select-Object –ExpandProperty Node).Number
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
				HostName   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Hostname
				UserName   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Username
				Password   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Password
				SshHostKeyFingerprint = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).HostKeyFingerprint
				RemoteFilePath = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Download" | Select-Object –ExpandProperty Node).PathRemote
				LocalFilePath = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Download" | Select-Object –ExpandProperty Node).PathLocal -join '\'
				DLLPath    = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP" | Select-Object –ExpandProperty Node).DLLPath -join '\'
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($PublishSynergyExport) {
			$properties = @{
				HostName    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Hostname
				UserName    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Username
				Password    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).Password
				SshHostKeyFingerprint = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']" | Select-Object –ExpandProperty Node).HostKeyFingerprint
				RemoteFilePath = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Upload" | Select-Object –ExpandProperty Node).PathRemote
				LocalFilePath = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Upload" | Select-Object –ExpandProperty Node).PathLocal -join '\'
				DLLPath	    = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP" | Select-Object –ExpandProperty Node).DLLPath -join '\'
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($NewStudentUserGoogle) {
			$properties = @{
				EXEPath	    = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/GAM" | Select-Object –ExpandProperty Node).EXEPath -join '\'
				Oauth2Path  = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/GAM" | Select-Object –ExpandProperty Node).Oauth2Path -join '\'
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($ClearFile) {
			$properties = @{
				File    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/CleanupFiles" | Select-Object –ExpandProperty Node).File
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($SuspendExpiredAccounts) {
			$properties = @{
				OrganizationalUnit	   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node).Path
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($MoveSuspendedAccounts) {
			$properties = @{
				MoveFromOU	   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node).Path
				MoveToOU	   = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location[@Name=`'Disabled`']" | Select-Object –ExpandProperty Node).Path
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($SendReport) {
			$properties = @{
				SMTPServer	     = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Notifications" | Select-Object –ExpandProperty Node).SMTPServer
				From			 = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Notifications" | Select-Object –ExpandProperty Node).From
				Body			 = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Notifications" | Select-Object –ExpandProperty Node).Body
				Subject		     = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Notifications" | Select-Object –ExpandProperty Node).Subject
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
				OrganizationalUnit	    = (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/UserConfig/Locations/Location" | Select-Object –ExpandProperty Node).Path
				Path				    = $PSScriptRoot, (Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/WinSCP/Host[@Name=`'Synergy`']/Upload" | Select-Object –ExpandProperty Node).PathLocal -join '\'
				LDAPAuth			    = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Features" | Select-Object –ExpandProperty Node).ExportSynergyLDAPAuth))
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
		if ($Features) {
			$properties = @{
				ExportSynergy    = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Features" | Select-Object –ExpandProperty Node).ExportSynergy))
				GoogleAccount    = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Features" | Select-Object –ExpandProperty Node).GoogleAccount))
				UserShare	     = ([System.Convert]::ToBoolean((Select-Xml -Xml $ConfigXMLObj -XPath "/Districts/District[@Name='$District']/Features" | Select-Object –ExpandProperty Node).UserShare))
			}
			$ConfigObj = New-Object -TypeName PSObject -Property $properties
			Write-Output $ConfigObj
		}
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
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
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
}
function Suspend-ExpiredAccounts {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[array]$OrganizationalUnit
	)
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
		foreach ($DN in $OrganizationalUnit) {
			try {
				Search-ADAccount -SearchBase $DN -AccountExpired -UsersOnly | Where-Object { $_.Enabled } | Disable-ADAccount
			} catch {
				Write-Error ("Error: {0}" -f $_.Exception.Message)
			}
		}
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
}
function Test-OrganizationalUnitPath {
	[CmdletBinding(ConfirmImpact = 'None')]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$OrganizationalUnitDN
	)
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
	if (Get-ADOrganizationalUnit -Filter {
			distinguishedName -eq $OrganizationalUnitDN
		}) {
		$properties = @{
			Result    = $true
		}
	} else {
		$properties = @{
			Result    = $false
		}
	}
	
	$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
}
function Update-ExistingUser {
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
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
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
				if (($DataImport.Title -eq '12') -or ($DataImport.Title -eq 'TR')) { Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime ($DataImport.AccountExpirationDate).AddDays(+ 365) }
				else { Set-ADAccountExpiration -Identity $DataAD.SamAccountName -DateTime $DataImport.AccountExpirationDate }
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
			EXAMPLE {
				#	if ($DataAD.DisplayName -ne $DataImport.DisplayName) { Set-ADUser -Identity $DataAD.SamAccountName -DisplayName $DataImport.DisplayName ; $notify = $true }
				#	if ($DataAD.GivenName -ne $DataImport.GivenName) { Set-ADUser -Identity $DataAD.SamAccountName -GivenName $DataImport.GivenName ; $notify = $true }
				#	if ($DataAD.Name -ne $DataImport.Name) { Set-ADUser -Identity $DataAD.Name -GivenName $DataImport.Name }
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
			Notify    = $notify
		}
		$obj = New-Object -TypeName PSObject -Property $properties
		Write-Output $obj
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
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
	
	BEGIN { Write-Verbose "BEGIN $($MyInvocation.MyCommand)" }
	PROCESS {
		
		switch ($District) {
			RIVERDALE {
				if ($LDAPAuth) {
					foreach ($DN in $OrganizationalUnit) {
						$users += Get-ADUser -SearchBase $DN -Filter { (Enabled -eq $true) } -Properties mail, employeeID, SamAccountName
					}
					$users | Where-Object { ($_.EmployeeID -match "^[\d\.]+$") -and ($_.Mail -ne $null) -and ($_.Title -match '[5-9]|10|11|12|TR') } | Select-Object employeeid, mail, SamAccountName | export-csv -delimiter "`t" -notypeinformation -append -path $Path
					(Get-Content $Path) | ForEach-Object { $_ -replace '"', '' } | out-file -FilePath $Path -Force -Encoding ascii
					(Get-Content $Path) | Select-Object -Skip 1 | Set-Content $Path
				} else {
					foreach ($DN in $OrganizationalUnit) {
						$users += Get-ADUser -SearchBase $DN -Filter { (Enabled -eq $true) } -Properties mail, employeeID, assistant
					}
					$users | Where-Object { ($_.EmployeeID -match "^[\d\.]+$") -and ($_.Mail -ne $null) -and ($_.Title -match '[5-9]|10|11|12|TR') } | Select-Object employeeid, mail, assistant | export-csv -delimiter "`t" -notypeinformation -append -path $Path
					(Get-Content $Path) | ForEach-Object { $_ -replace '"', '' } | out-file -FilePath $Path -Force -Encoding ascii
					(Get-Content $Path) | Select-Object -Skip 1 | Set-Content $Path
				}
			}
			default {
				if ($LDAPAuth) {
					foreach ($DN in $OrganizationalUnit) {
						$users += Get-ADUser -SearchBase $DN -Filter { (Enabled -eq $true) } -Properties mail, employeeID, SamAccountName
					}
					$users | Where-Object { ($_.EmployeeID -match "^[\d\.]+$") -and ($_.Mail -ne $null) } | Select-Object employeeid, mail, SamAccountName | export-csv -delimiter "`t" -notypeinformation -append -path $Path
					(Get-Content $Path) | ForEach-Object { $_ -replace '"', '' } | out-file -FilePath $Path -Force -Encoding ascii
					(Get-Content $Path) | Select-Object -Skip 1 | Set-Content $Path
				} else {
					foreach ($DN in $OrganizationalUnit) {
						$users += Get-ADUser -SearchBase $DN -Filter { (Enabled -eq $true) } -Properties mail, employeeID, assistant
					}
					$users | Where-Object { ($_.EmployeeID -match "^[\d\.]+$") -and ($_.Mail -ne $null) } | Select-Object employeeid, mail, assistant | export-csv -delimiter "`t" -notypeinformation -append -path $Path
					(Get-Content $Path) | ForEach-Object { $_ -replace '"', '' } | out-file -FilePath $Path -Force -Encoding ascii
					(Get-Content $Path) | Select-Object -Skip 1 | Set-Content $Path
				}
			}
		}
	}
	END { Write-Verbose "END $($MyInvocation.MyCommand)" }
}
#endregion Functions
#region Begin
Read-ConfigXML -District $District -Path $ConfigXML -ClearFile | Clear-File
Read-ConfigXML -District $District -Path $ConfigXML -GetSynergyImport | Get-FileSCP
$ConfigScript = Read-ConfigXML -District $District -Path $ConfigXML -Script
$ConfigF = Read-ConfigXML -District $District -Path $ConfigXML -Features
$userdata = Import-Csv -Path $ConfigScript.ImportCSVPath
#endregion Begin
#region Process
foreach ($row in $userdata) {
	Write-Verbose "BEGIN $($row.SIS_NUMBER)"
	$currentUsers = [Collections.Generic.List[Object]](Get-AllUserDataAD -SearchBase ($ConfigScript.studentsOUs))
	if (($ConfigScript.SkipGrades -eq $row.Grade) -or (($ConfigScript.SkipLocations -replace '[^a-zA-Z0-9 ]', '') -eq ($row.LOCATION -replace '[^a-zA-Z0-9 ]', '')) -or ($ConfigScript.SkipStudents -eq $row.SIS_NUMBER)) {
		Write-Verbose "SKIP $($row.SIS_NUMBER)"
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
		Write-Verbose "NEW $($row.SIS_NUMBER)"
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
		Write-Verbose "UPDATE $($row.SIS_NUMBER)"
		$userDataImport.HomeDirectory = '\\', $ConfigNUS.HomeDirectoryServer, '\', $userDataAD.SamAccountName, '$' -join ''
		$userDataImport.Path = (Get-OrganizationalUnitPath -Location $userDataImport.Office -GradYear $userDataImport.personalTitle.Substring(2) -Grade $userDataImport.Title -OrganizationalUnit $userDataImport.Path -District $District).Path
		New-SamAccountName -GivenName $userDataAD.GivenName -SurName $userDataAD.Surname -GradYear $userDataImport.personalTitle.Substring(2) -UPNSuffiix $ConfigScript.UPNSuffix -Grade $userDataImport.Title -AllUsersAD $currentUsers -District $District | Foreach-Object { $userDataImport.Name = $_.Name; $userDataImport.UserPrincipalName = $_.UserPrincipalName; $userDataImport.SamAccountName = $_.SamAccountName }
		$userDataImport.Notify = (Update-ExistingUser -DataAD $userDataAD -DataImport $userDataImport -District $District).Notify
	}
	if ($userDataImport.Notify -eq $true) {
		$output += $userDataImport
		Write-Verbose "NOTIFY $($row.SIS_NUMBER)"
	}
	Write-Verbose "END $($row.SIS_NUMBER)"
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

Pop-Location
Remove-PSDrive -Name Script
#endregion End


# SIG # Begin signature block
# MIIQrgYJKoZIhvcNAQcCoIIQnzCCEJsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUrrjA6BLVnGZV6lAw5wFoCDFi
# pXOgggvFMIIDBjCCAe6gAwIBAgIQL5tLI0kB95tC0nRJzV7uKjANBgkqhkiG9w0B
# AQUFADAWMRQwEgYDVQQDDAtFZGVuIE5lbHNvbjAeFw0xNzA5MDgyMjI2NTFaFw0z
# NzA5MDgyMjM2NTFaMBYxFDASBgNVBAMMC0VkZW4gTmVsc29uMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEAt5TChlY9pEZX/3fWVpiaKiwxvoapE4Wkvkxf
# Loj8LJxZCBJI4yWjKCyFfuZpNTkL2cR2o7S4Ikp1VxJAlHa0moH27A4Cd6WDxAPF
# ksfyENY+aFpfZSMCLql8kQg9bzMDxyW+5Lr4r2tOX3Mx03HeCLo9l3ax3GnFJ7Ur
# FEMRFviaZMxDlFT4cUmHQfL2WaFviP6bfRT3+jLs3KJWyuEVgrsgg4fZAEiTo/nu
# C2RdfT00gsqVTzrg1FyPdnE43MTcojHEiRJ14GRRFEU/CX8PntGKvUf+qAlp2GYD
# Bg/YmBEJLtFxvCo1Vja7oDubKncPPWUKmA85WXtDUsHE5YohkQIDAQABo1AwTjAO
# BgNVHQ8BAf8EBAMCBaAwHQYDVR0lBBYwFAYIKwYBBQUHAwIGCCsGAQUFBwMBMB0G
# A1UdDgQWBBQQmjQ9M/4E7vwm7Jt1MUCG4f43uzANBgkqhkiG9w0BAQUFAAOCAQEA
# CHkEawdt83WaQ8pqAxOdhjdr/TeVLt3IJFM/ZO/PJ9y37kJ4QXQ0hl4tmYVF9pF8
# mrPxdLRNLhGMVm+jQ8APKmAUpewFi/oRANppG2uw496eikcipB7IjwGNAUru4Wcq
# WfKAtfWNrl2HFEMpnyeI4NrGZKXxC+CTOV102SqVq2O2xiKrvFSR+NI9xCYNXXBt
# ErEw7cqGMaHPdtc0TM7coGHvbS8muoJdM8tULayji7vkPKR7D/HXEnrLe0DS/MRG
# HRPdqgGuyz8M8IVUlNevC6GEYruuBQ/D/NGThwETeVNyakIw6rMeRSXmPV3RQ17S
# k1OeaULzyA/Vo5+g9jC6NzCCBBQwggL8oAMCAQICCwQAAAAAAS9O4VLXMA0GCSqG
# SIb3DQEBBQUAMFcxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52
# LXNhMRAwDgYDVQQLEwdSb290IENBMRswGQYDVQQDExJHbG9iYWxTaWduIFJvb3Qg
# Q0EwHhcNMTEwNDEzMTAwMDAwWhcNMjgwMTI4MTIwMDAwWjBSMQswCQYDVQQGEwJC
# RTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEoMCYGA1UEAxMfR2xvYmFsU2ln
# biBUaW1lc3RhbXBpbmcgQ0EgLSBHMjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCC
# AQoCggEBAJTvZfi1V5+gUw00BusJH7dHGGrL8Fvk/yelNNH3iRq/nrHNEkFuZtSB
# oIWLZFpGL5mgjXex4rxc3SLXamfQu+jKdN6LTw2wUuWQW+tHDvHnn5wLkGU+F5Yw
# RXJtOaEXNsq5oIwbTwgZ9oExrWEWpGLmtECew/z7lfb7tS6VgZjg78Xr2AJZeHf3
# quNSa1CRKcX8982TZdJgYSLyBvsy3RZR+g79ijDwFwmnu/MErquQ52zfeqn078Ri
# J19vmW04dKoRi9rfxxRM6YWy7MJ9SiaP51a6puDPklOAdPQD7GiyYLyEIACDG6Hu
# tHQFwSmOYtBHsfrwU8wY+S47+XB+tCUCAwEAAaOB5TCB4jAOBgNVHQ8BAf8EBAMC
# AQYwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQURtg+/9zjvv+D5vSFm7Dd
# atYUqcEwRwYDVR0gBEAwPjA8BgRVHSAAMDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8v
# d3d3Lmdsb2JhbHNpZ24uY29tL3JlcG9zaXRvcnkvMDMGA1UdHwQsMCowKKAmoCSG
# Imh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5uZXQvcm9vdC5jcmwwHwYDVR0jBBgwFoAU
# YHtmGkUNl8qJUC99BM00qP/8/UswDQYJKoZIhvcNAQEFBQADggEBAE5eVpAeRrTZ
# STHzuxc5KBvCFt39QdwJBQSbb7KimtaZLkCZAFW16j+lIHbThjTUF8xVOseC7u+o
# urzYBp8VUN/NFntSOgLXGRr9r/B4XOBLxRjfOiQe2qy4qVgEAgcw27ASXv4xvvAE
# SPTwcPg6XlaDzz37Dbz0xe2XnbnU26UnhOM4m4unNYZEIKQ7baRqC6GD/Sjr2u8o
# 9syIXfsKOwCr4CHr4i81bA+ONEWX66L3mTM1fsuairtFTec/n8LZivplsm7HfmX/
# 6JLhLDGi97AnNkiPJm877k12H3nD5X+WNbwtDswBsI5//1GAgKeS1LNERmSMh08W
# YwcxS2Ow3/MwggSfMIIDh6ADAgECAhIRIdaZp2SXPvH4Qn7pGcxTQRQwDQYJKoZI
# hvcNAQEFBQAwUjELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYt
# c2ExKDAmBgNVBAMTH0dsb2JhbFNpZ24gVGltZXN0YW1waW5nIENBIC0gRzIwHhcN
# MTYwNTI0MDAwMDAwWhcNMjcwNjI0MDAwMDAwWjBgMQswCQYDVQQGEwJTRzEfMB0G
# A1UEChMWR01PIEdsb2JhbFNpZ24gUHRlIEx0ZDEwMC4GA1UEAxMnR2xvYmFsU2ln
# biBUU0EgZm9yIE1TIEF1dGhlbnRpY29kZSAtIEcyMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEAsBeuotO2BDBWHlgPse1VpNZUy9j2czrsXV6rJf02pfqE
# w2FAxUa1WVI7QqIuXxNiEKlb5nPWkiWxfSPjBrOHOg5D8NcAiVOiETFSKG5dQHI8
# 8gl3p0mSl9RskKB2p/243LOd8gdgLE9YmABr0xVU4Prd/4AsXximmP/Uq+yhRVmy
# Lm9iXeDZGayLV5yoJivZF6UQ0kcIGnAsM4t/aIAqtaFda92NAgIpA6p8N7u7KU49
# U5OzpvqP0liTFUy5LauAo6Ml+6/3CGSwekQPXBDXX2E3qk5r09JTJZ2Cc/os+XKw
# qRk5KlD6qdA8OsroW+/1X1H0+QrZlzXeaoXmIwRCrwIDAQABo4IBXzCCAVswDgYD
# VR0PAQH/BAQDAgeAMEwGA1UdIARFMEMwQQYJKwYBBAGgMgEeMDQwMgYIKwYBBQUH
# AgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3JlcG9zaXRvcnkvMAkGA1Ud
# EwQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9ncy9nc3RpbWVzdGFtcGluZ2cyLmNy
# bDBUBggrBgEFBQcBAQRIMEYwRAYIKwYBBQUHMAKGOGh0dHA6Ly9zZWN1cmUuZ2xv
# YmFsc2lnbi5jb20vY2FjZXJ0L2dzdGltZXN0YW1waW5nZzIuY3J0MB0GA1UdDgQW
# BBTUooRKOFoYf7pPMFC9ndV6h9YJ9zAfBgNVHSMEGDAWgBRG2D7/3OO+/4Pm9IWb
# sN1q1hSpwTANBgkqhkiG9w0BAQUFAAOCAQEAj6kakW0EpjcgDoOW3iPTa24fbt1k
# PWghIrX4RzZpjuGlRcckoiK3KQnMVFquxrzNY46zPVBI5bTMrs2SjZ4oixNKEaq9
# o+/Tsjb8tKFyv22XY3mMRLxwL37zvN2CU6sa9uv6HJe8tjecpBwwvKu8LUc235Ig
# A+hxxlj2dQWaNPALWVqCRDSqgOQvhPZHXZbJtsrKnbemuuRQ09Q3uLogDtDTkipb
# xFm7oW3bPM5EncE4Kq3jjb3NCXcaEL5nCgI2ZIi5sxsm7ueeYMRGqLxhM2zPTrmc
# uWrwnzf+tT1PmtNN/94gjk6Xpv2fCbxNyhh2ybBNhVDygNIdBvVYBAexGDGCBFMw
# ggRPAgEBMCowFjEUMBIGA1UEAwwLRWRlbiBOZWxzb24CEC+bSyNJAfebQtJ0Sc1e
# 7iowCQYFKw4DAhoFAKBaMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZI
# hvcNAQkDMQwGCisGAQQBgjcCAQQwIwYJKoZIhvcNAQkEMRYEFGUgTekjLAIMz0dw
# XQJZAvIUCsxeMA0GCSqGSIb3DQEBAQUABIIBAG3onErn4MNRzKMoJDa6um58UpD1
# TyLzN8x2o3rheTj0pc4Gq0DRhizyF9DHoPYCqTCoarCBd1wmDtbyo8dneYPzPEE3
# fcsNx4Hz1fis1z9dlaqNFaCOgk5bv8B7LWgbNhvwqJKD+xVORAD15M35xCM7rTqr
# +1EGgaXDeGa5up7a+EY9knBQhihXFMl5N8Z2jy4Iy+g8IUsDtPKXxa/BdHShg/Mv
# Dhq8h24yOe3jGNsQeUpmwBLF33wec+Agy7avSNzmWh295GcykhQoo9wWTpmYJB1w
# HA/NyKCNIKwDdhuoe4WfN4pwGOzWuaRkjz80pZ1etr8vIzNtPfrc/ViHkDuhggKi
# MIICngYJKoZIhvcNAQkGMYICjzCCAosCAQEwaDBSMQswCQYDVQQGEwJCRTEZMBcG
# A1UEChMQR2xvYmFsU2lnbiBudi1zYTEoMCYGA1UEAxMfR2xvYmFsU2lnbiBUaW1l
# c3RhbXBpbmcgQ0EgLSBHMgISESHWmadklz7x+EJ+6RnMU0EUMAkGBSsOAwIaBQCg
# gf0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTcx
# MTE1MjExMzQxWjAjBgkqhkiG9w0BCQQxFgQUWzHf8E5jZL08Nm158KJyG/3T1oow
# gZ0GCyqGSIb3DQEJEAIMMYGNMIGKMIGHMIGEBBRjuC+rYfWDkJaVBQsAJJxQKTPs
# eTBsMFakVDBSMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1z
# YTEoMCYGA1UEAxMfR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0EgLSBHMgISESHW
# madklz7x+EJ+6RnMU0EUMA0GCSqGSIb3DQEBAQUABIIBACG2f9DRtfhUnH1RjRWF
# j+iHwBDLKi1022paJTyhYYkWX48+GlWXDAtQ2mOHJ6ZSYp/wg+fdITz9WTlaHt1A
# oDoh4kkc46LbWLMRTdKWp2bdm7urv/XddqY03bR8cHid0gJ3HumRLvF8m+0a9Inf
# tq0m9lxB6vsRfNJNVXkFT50mk5KHZ5MpPNZoGhiNvU8tHL+/OXxSbnSMKCqyoW1g
# OtTOC/zCZq7VrLUup9b0c9UqFrFftosWnXPEfvehrPIc14//sk4MrVLvNpgT4OVo
# l+vd0rtJwhfMHyIIJYRJq/DdBzwwamOXoMeoJ4laitbEqmivQBr2EWEZi/dO3LgK
# F88=
# SIG # End signature block
