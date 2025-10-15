#Set directory of powershell being run
$scriptBase = if ($PSCommandPath)
{
	Split-Path -Parent $PSCommandPath
}
else
{
	[System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}

#Global Variables:
$cmdExePath = "$env:windir\System32\cmd.exe"
$psexecPath = Join-Path $scriptBase "PSTools\PsExec.exe"
$script:Trace32Path = Join-Path $scriptBase "Trace32\Trace32.exe"

$script:timer = $null

#AdminToolkit Function
function AdminToolkit
{
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	# Create main form
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Desktop Admin Toolkit"
	$form.Size = New-Object System.Drawing.Size(1000, 800)
	$form.StartPosition = "CenterScreen"

	# Create tab control
	$tabs = New-Object System.Windows.Forms.TabControl
	$tabs.Size = New-Object System.Drawing.Size(970, 675)
	$tabs.Location = New-Object System.Drawing.Point(10, 80)
	$tabs.Multiline = $true

	#variables
	$Script:ConnectedPC = "localhost"
	$Script:ConnectedPCDNSName = "LocalHost"


	#region Dark Mode

	# Track theme state
	$global:isDarkMode = $false

	# Create the toggle button
	$darkModeButton = New-Object System.Windows.Forms.Button
	$darkModeButton.Text = "Toggle Dark Mode"
	$darkModeButton.Location = New-Object System.Drawing.Point(800, 10)
	$darkModeButton.Size = New-Object System.Drawing.Size(120, 30)

	# Function to apply theme
	function Set-Theme
	{
		param($control, $isDark)

		if ($isDark)
		{
			$control.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
			$control.ForeColor = [System.Drawing.Color]::White
		}
		else
		{
			$control.BackColor = [System.Drawing.Color]::White
			$control.ForeColor = [System.Drawing.Color]::Black
		}

		# Special handling for DataGridView
		if ($control -is [System.Windows.Forms.DataGridView])
		{
			if ($isDark)
			{
				$control.BackgroundColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
				$control.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
				$control.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
				$control.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
				$control.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
				$control.EnableHeadersVisualStyles = $false
			}
			else
			{
				$control.BackgroundColor = [System.Drawing.Color]::White
				$control.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
				$control.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
				$control.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::White
				$control.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
				$control.EnableHeadersVisualStyles = $true
			}
		}

		foreach ($child in $control.Controls)
		{
			Set-Theme -control $child -isDark $isDark
		}
	}


	# Button click event
	$darkModeButton.Add_Click({
			$global:isDarkMode = -not $global:isDarkMode
			Set-Theme -control $form -isDark $global:isDarkMode
		})

	$form.Controls.Add($darkModeButton)


	#endregion

	#region Form Labels and Actions
	<# Mode label and dropdown
	$modeLabel = New-Object System.Windows.Forms.Label
	$modeLabel.Text = "Mode:"
	$modeLabel.Width = 50
	$modeLabel.Location = New-Object System.Drawing.Point(10, 20)
	$form.Controls.Add($modeLabel) #>

	# GroupBox to contain the radio buttons (optional but recommended for clarity)
	$modeGroupBox = New-Object System.Windows.Forms.GroupBox
	$modeGroupBox.Text = "Mode"
	$modeGroupBox.Location = New-Object System.Drawing.Point(5, 2)
	$modeGroupBox.Size = New-Object System.Drawing.Size(160, 40)
	$form.Controls.Add($modeGroupBox)

	# Local Radio Button
	$radioLocal = New-Object System.Windows.Forms.RadioButton
	$radioLocal.Text = "Local"
	$radioLocal.Location = New-Object System.Drawing.Point(15, 14)
	$radioLocal.Width = 60
	$radioLocal.Checked = $true  # Default selection
	$modeGroupBox.Controls.Add($radioLocal)

	# Remote Radio Button
	$radioRemote = New-Object System.Windows.Forms.RadioButton
	$radioRemote.Text = "Remote"
	$radioRemote.Location = New-Object System.Drawing.Point(80, 14)
	$radioRemote.Width = 70
	$modeGroupBox.Controls.Add($radioRemote)


	# Computer name textbox
	$compNameBox = New-Object System.Windows.Forms.TextBox
	$compNameBox.Location = New-Object System.Drawing.Point(170, 18)
	$compNameBox.Width = 200
	$compNameBox.Text = "localhost"
	$compNameBox.ReadOnly = $true
	$form.Controls.Add($compNameBox)

	# Connect button
	$connectBtn = New-Object System.Windows.Forms.Button
	$connectBtn.Text = "Connect"
	$connectBtn.Location = New-Object System.Drawing.Point(390, 15)
	$form.Controls.Add($connectBtn)


	$compNameBox.Add_KeyDown({
			if ($_.KeyCode -eq 'Enter')
			{
				$connectBtn.PerformClick()
			}
		})


	# Export button
	$exportBtn = New-Object System.Windows.Forms.Button
	$exportBtn.Text = "Export Info"
	$exportBtn.Location = New-Object System.Drawing.Point(470, 15)
	$form.Controls.Add($exportBtn)

	$connectedPCLabel = New-Object System.Windows.Forms.Label
	$connectedPCLabel.Text = "Enter PC Name (Local or Remote) and click Connect"
	$connectedPCLabel.Width = 600
	$connectedPCLabel.Location = New-Object System.Drawing.Point(10, 50)
	$connectedPCLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
	$form.Controls.Add($connectedPCLabel)

	# Mode  change event
	$radioLocal.Add_CheckedChanged({
			if ($radioLocal.Checked)
			{
				$compNameBox.Text = "localhost"
				$compNameBox.ReadOnly = $true
			}
		})

	$radioRemote.Add_CheckedChanged({
			if ($radioRemote.Checked)
			{
				$compNameBox.Clear()
				$compNameBox.ReadOnly = $false
			}
		})



	#endregion

	#region PC Info Tab
	#####################################################################  PC Info Tab ###################################################################################################
	$pcInfoTab = New-Object System.Windows.Forms.TabPage
	$pcInfoTab.Text = "PC Info"

	# Computer Name
	$labelComputerName = New-Object System.Windows.Forms.Label
	$labelComputerName.Text = "Computer Name:"
	$labelComputerName.Location = New-Object System.Drawing.Point(10, 30)
	$labelComputerName.Width = 150
	$pcInfoTab.Controls.Add($labelComputerName)

	$pcinfoComputerName = New-Object System.Windows.Forms.TextBox
	$pcinfoComputerName.Location = New-Object System.Drawing.Point(170, 30)
	$pcinfoComputerName.Width = 650
	$pcinfoComputerName.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoComputerName)

	# Computer Model
	$labelComputerModel = New-Object System.Windows.Forms.Label
	$labelComputerModel.Text = "Computer Model:"
	$labelComputerModel.Location = New-Object System.Drawing.Point(10, 60)
	$labelComputerModel.Width = 150
	$pcInfoTab.Controls.Add($labelComputerModel)

	$pcinfoComputerModel = New-Object System.Windows.Forms.TextBox
	$pcinfoComputerModel.Location = New-Object System.Drawing.Point(170, 60)
	$pcinfoComputerModel.Width = 650
	$pcinfoComputerModel.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoComputerModel)

	# OS Version
	$labelOSVersion = New-Object System.Windows.Forms.Label
	$labelOSVersion.Text = "OS Version:"
	$labelOSVersion.Location = New-Object System.Drawing.Point(10, 90)
	$labelOSVersion.Width = 150
	$pcInfoTab.Controls.Add($labelOSVersion)

	$pcinfoOSVersion = New-Object System.Windows.Forms.TextBox
	$pcinfoOSVersion.Location = New-Object System.Drawing.Point(170, 90)
	$pcinfoOSVersion.Width = 650
	$pcinfoOSVersion.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoOSVersion)

	# Last Boot Time
	$labelLastBootTime = New-Object System.Windows.Forms.Label
	$labelLastBootTime.Text = "Last Boot Time:"
	$labelLastBootTime.Location = New-Object System.Drawing.Point(10, 120)
	$labelLastBootTime.Width = 150
	$pcInfoTab.Controls.Add($labelLastBootTime)

	$pcinfoLastBootTime = New-Object System.Windows.Forms.TextBox
	$pcinfoLastBootTime.Location = New-Object System.Drawing.Point(170, 120)
	$pcinfoLastBootTime.Width = 650
	$pcinfoLastBootTime.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoLastBootTime)

	# BIOS Version
	$labelBIOSVersion = New-Object System.Windows.Forms.Label
	$labelBIOSVersion.Text = "BIOS Version:"
	$labelBIOSVersion.Location = New-Object System.Drawing.Point(10, 150)
	$labelBIOSVersion.Width = 150
	$pcInfoTab.Controls.Add($labelBIOSVersion)

	$pcinfoBIOSVersion = New-Object System.Windows.Forms.TextBox
	$pcinfoBIOSVersion.Location = New-Object System.Drawing.Point(170, 150)
	$pcinfoBIOSVersion.Width = 650
	$pcinfoBIOSVersion.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoBIOSVersion)

	# Installed RAM
	$labelInstalledRAM = New-Object System.Windows.Forms.Label
	$labelInstalledRAM.Text = "Installed RAM:"
	$labelInstalledRAM.Location = New-Object System.Drawing.Point(10, 180)
	$labelInstalledRAM.Width = 150
	$pcInfoTab.Controls.Add($labelInstalledRAM)

	$pcinfoInstalledRAM = New-Object System.Windows.Forms.TextBox
	$pcinfoInstalledRAM.Location = New-Object System.Drawing.Point(170, 180)
	$pcinfoInstalledRAM.Width = 650
	$pcinfoInstalledRAM.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoInstalledRAM)

	# Available C Drive Space
	$labelDriveSpace = New-Object System.Windows.Forms.Label
	$labelDriveSpace.Text = "Available C Drive Space:"
	$labelDriveSpace.Location = New-Object System.Drawing.Point(10, 210)
	$labelDriveSpace.Width = 150
	$pcInfoTab.Controls.Add($labelDriveSpace)

	$pcinfoDriveSpace = New-Object System.Windows.Forms.TextBox
	$pcinfoDriveSpace.Location = New-Object System.Drawing.Point(170, 210)
	$pcinfoDriveSpace.Width = 650
	$pcinfoDriveSpace.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoDriveSpace)

	# Domain
	$labelDomain = New-Object System.Windows.Forms.Label
	$labelDomain.Text = "Domain:"
	$labelDomain.Location = New-Object System.Drawing.Point(10, 240)
	$labelDomain.Width = 150
	$pcInfoTab.Controls.Add($labelDomain)

	$pcinfoDomain = New-Object System.Windows.Forms.TextBox
	$pcinfoDomain.Location = New-Object System.Drawing.Point(170, 240)
	$pcinfoDomain.Width = 650
	$pcinfoDomain.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoDomain)

	# OU
	$labelOU = New-Object System.Windows.Forms.Label
	$labelOU.Text = "OU:"
	$labelOU.Location = New-Object System.Drawing.Point(10, 270)
	$labelOU.Width = 150
	$pcInfoTab.Controls.Add($labelOU)

	$pcinfoOU = New-Object System.Windows.Forms.TextBox
	$pcinfoOU.Location = New-Object System.Drawing.Point(170, 270)
	$pcinfoOU.Width = 650
	$pcinfoOU.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoOU)


	# IP Address (Multiline)
	$labelIPAddress = New-Object System.Windows.Forms.Label
	$labelIPAddress.Text = "IP Address:"
	$labelIPAddress.Location = New-Object System.Drawing.Point(10, 380)
	$labelIPAddress.Width = 150
	$pcInfoTab.Controls.Add($labelIPAddress)

	$pcinfoIPAddress = New-Object System.Windows.Forms.TextBox
	$pcinfoIPAddress.Location = New-Object System.Drawing.Point(170, 380)
	$pcinfoIPAddress.Width = 650
	$pcinfoIPAddress.Height = 60
	$pcinfoIPAddress.Multiline = $true
	$pcinfoIPAddress.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoIPAddress)

	# Logged-in Users (Multiline)
	$labelLoggedInUsers = New-Object System.Windows.Forms.Label
	$labelLoggedInUsers.Text = "Logged-in Users:"
	$labelLoggedInUsers.Location = New-Object System.Drawing.Point(10, 450)
	$labelLoggedInUsers.Width = 150
	$pcInfoTab.Controls.Add($labelLoggedInUsers)

	$pcinfoLoggedInUsers = New-Object System.Windows.Forms.TextBox
	$pcinfoLoggedInUsers.Location = New-Object System.Drawing.Point(170, 450)
	$pcinfoLoggedInUsers.Width = 650
	$pcinfoLoggedInUsers.Height = 80
	$pcinfoLoggedInUsers.Multiline = $true
	$pcinfoLoggedInUsers.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoLoggedInUsers)

	# System Time and Time Zone
	$labelSystemTime = New-Object System.Windows.Forms.Label
	$labelSystemTime.Text = "System Time:"
	$labelSystemTime.Location = New-Object System.Drawing.Point(10, 310)
	$labelSystemTime.Width = 150
	$pcInfoTab.Controls.Add($labelSystemTime)

	$pcinfoSystemTime = New-Object System.Windows.Forms.TextBox
	$pcinfoSystemTime.Location = New-Object System.Drawing.Point(170, 310)
	$pcinfoSystemTime.Width = 650
	$pcinfoSystemTime.ReadOnly = $true
	$pcInfoTab.Controls.Add($pcinfoSystemTime)


	function Clear-PCInfoFields
	{
		$connectedPCLabel.Text = ""
		$pcinfoComputerName.Text = ""
		$pcinfoComputerModel.Text = ""
		$pcinfoOSVersion.Text = ""
		$pcinfoLastBootTime.Text = ""
		$pcinfoBIOSVersion.Text = ""
		$pcinfoInstalledRAM.Text = ""
		$pcinfoDriveSpace.Text = ""
		$pcinfoDomain.Text = ""
		$pcinfoLoggedInUsers.Text = ""
		$pcinfoOU.Text = ""
		$pcinfoIPAddress.Text = ""
		$pcinfoSystemTime.Text = ""
		$pcinfoDriveSpace.BackColor = [System.Drawing.SystemColors]::Control

	}

	function Update-PCInfoFields
	{
		param (
			[string]$target
		)
		Clear-PCInfoFields
		$runspace = [runspacefactory]::CreateRunspace()
		$runspace.ApartmentState = "STA"
		$runspace.ThreadOptions = "ReuseThread"
		$runspace.Open()

		$ps = [PowerShell]::Create()
		$ps.Runspace = $runspace

		$ps.AddScript({
				param (
					$target,
					$connectedPCLabel,
					$pcinfoComputerName, $pcinfoComputerModel, $pcinfoOSVersion, $pcinfoLastBootTime,
					$pcinfoBIOSVersion, $pcinfoInstalledRAM, $pcinfoDriveSpace, $pcinfoDomain,
					$pcinfoLoggedInUsers, $pcinfoOU, $pcinfoIPAddress, $pcinfoSystemTime
				)

				try
				{
					$sysInfo = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $target
					$bios = Get-WmiObject -Class Win32_BIOS -ComputerName $target
					$cs = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $target
					$disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $target -Filter "DeviceID='C:'"
					$tz = Get-WmiObject -Class Win32_TimeZone -ComputerName $target

					$domain = $cs.Domain
					$ram = "{0:N2} GB" -f ($cs.TotalPhysicalMemory / 1GB)
					$freeSpace = "{0:N2} GB" -f ($disk.FreeSpace / 1GB)
					$bootTime = $sysInfo.ConvertToDateTime($sysInfo.LastBootUpTime)
					$timeZoneInfo = "$($tz.StandardName) (UTC$($tz.Bias / -60))"
					$users = query user /server:$target 2>&1

					$connectedPCLabel.Invoke([Action] { $connectedPCLabel.Text = "Connected to $($sysInfo.CSName)" })
					$pcinfoComputerName.Invoke([Action] { $pcinfoComputerName.Text = $sysInfo.CSName })
					$pcinfoComputerModel.Invoke([Action] { $pcinfoComputerModel.Text = "$($cs.Manufacturer) $($cs.Model)" })
					$pcinfoOSVersion.Invoke([Action] { $pcinfoOSVersion.Text = "$($sysInfo.Caption) [Version $($sysInfo.Version)]" })
					$pcinfoLastBootTime.Invoke([Action] { $pcinfoLastBootTime.Text = $bootTime })
					$pcinfoBIOSVersion.Invoke([Action] { $pcinfoBIOSVersion.Text = $bios.SMBIOSBIOSVersion })
					$pcinfoInstalledRAM.Invoke([Action] { $pcinfoInstalledRAM.Text = $ram })
					$pcinfoDomain.Invoke([Action] { $pcinfoDomain.Text = $domain })
					$pcinfoLoggedInUsers.Invoke([Action] { $pcinfoLoggedInUsers.Text = $users -join "`r`n" })

					# Set system time and time zone
					$remoteTime = $sysInfo.ConvertToDateTime($sysInfo.LocalDateTime)
					$timeZoneInfo = "$($tz.StandardName) (UTC$($tz.Bias / -60))"
					$formattedTime = $remoteTime.ToString("MM/dd/yyyy    hh:mm tt")
					$pcinfoSystemTime.Invoke([Action] { $pcinfoSystemTime.Text = "$formattedTime     $timeZoneInfo" })


					# Calculate utilization percentage and setting Drivespace
					$totalSpace = $disk.Size
					$usedSpace = $totalSpace - $disk.FreeSpace
					$utilization = ($usedSpace / $totalSpace) * 100

					$pcinfoDriveSpace.Invoke([Action] { $pcinfoDriveSpace.Text = "$freeSpace               The Drive is at {0:N0}% Utilization" -f $utilization })
					$pcinfoDriveSpace.Invoke([Action] { $pcinfoDriveSpace.ForeColor = 'Black' })


					# Set background color based on utilization
					if ($utilization -ge 95)
					{
						$pcinfoDriveSpace.BackColor = 'LightCoral'
					}
					elseif ($utilization -ge 90)
     {
						$pcinfoDriveSpace.BackColor = 'Moccasin'
					}
					elseif ($utilization -ge 80)
     {
						$pcinfoDriveSpace.BackColor = 'LightYellow'
					}
					else
     {
						$pcinfoDriveSpace.BackColor = 'LightGreen'
					}

					function Get-DomainFQDN
					{
						param ($netbiosName)
						try
						{
							$context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $netbiosName)
							$domainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
							return $domainObj.Name
						}
						catch
						{
							return $null
						}
					}

					$domainFQDN = Get-DomainFQDN -netbiosName $domain
					if ($domainFQDN)
					{
						$ldapPath = "LDAP://$domainFQDN"
						$entry = New-Object DirectoryServices.DirectoryEntry($ldapPath)
						$searcher = New-Object DirectoryServices.DirectorySearcher($entry)
						$searcher.Filter = "(&(objectClass=computer)(name=$($sysInfo.CSName)))"
						$searcher.ClientTimeout = [TimeSpan]::FromSeconds(5)
						$result = $searcher.FindOne()

						if ($result -ne $null -and $result.Properties["distinguishedname"])
						{
							$dn = $result.Properties["distinguishedname"][0]
							$ouParts = ($dn -split ',') | Where-Object { $_ -like "OU=*" } | ForEach-Object { $_.Substring(3) }
							[void][array]::Reverse($ouParts)
							$pcinfoOU.Invoke([Action] { $pcinfoOU.Text = $ouParts -join '\' })
						}
						else
						{
							$pcinfoOU.Invoke([Action] { $pcinfoOU.Text = "Computer object not found" })
						}
					}
					else
					{
						$pcinfoOU.Invoke([Action] { $pcinfoOU.Text = "Unable to resolve domain FQDN" })
					}

					$ipList = ""
					$adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $target | Where-Object { $_.IPAddress -ne $null }
					foreach ($adapter in $adapters)
					{
						$name = $adapter.Description
						foreach ($ip in $adapter.IPAddress)
						{
							$ipList += "${name}: ${ip}`r`n"
						}
					}
					$pcinfoIPAddress.Invoke([Action] { $pcinfoIPAddress.Text = $ipList })
				}
				catch
				{
					[System.Windows.Forms.MessageBox]::Show("Failed to retrieve system info from $target.`n$_", "Error", "OK", "Error")
				}
			}) | Out-Null

		$ps.AddArgument($target)
		$ps.AddArgument($connectedPCLabel)
		$ps.AddArgument($pcinfoComputerName)
		$ps.AddArgument($pcinfoComputerModel)
		$ps.AddArgument($pcinfoOSVersion)
		$ps.AddArgument($pcinfoLastBootTime)
		$ps.AddArgument($pcinfoBIOSVersion)
		$ps.AddArgument($pcinfoInstalledRAM)
		$ps.AddArgument($pcinfoDriveSpace)
		$ps.AddArgument($pcinfoDomain)
		$ps.AddArgument($pcinfoLoggedInUsers)
		$ps.AddArgument($pcinfoOU)
		$ps.AddArgument($pcinfoIPAddress)
		$ps.AddArgument($pcinfoSystemTime)

		$ps.BeginInvoke()
	}

	#endregion


	#region Timer Tab
	##########################timer for debugging

	<#
	$DebugTimerTab = New-Object System.Windows.Forms.TabPage
	$DebugTimerTab.Text = "Timer"



	# Create labels for each timer
	$labelTotalTime = New-Object System.Windows.Forms.Label
	$labelTotalTime.Text = "Total Time: "
	$labelTotalTime.AutoSize = $true
	$labelTotalTime.Location = New-Object System.Drawing.Point(0, 0)

	$labelMainTime = New-Object System.Windows.Forms.Label
	$labelMainTime.Text = "PC Info Tab Time: "
	$labelMainTime.AutoSize = $true
	$labelMainTime.Location = New-Object System.Drawing.Point(0, 20)

	$labelUserTime = New-Object System.Windows.Forms.Label
	$labelUserTime.Text = "User Tab Time: "
	$labelUserTime.AutoSize = $true
	$labelUserTime.Location = New-Object System.Drawing.Point(0, 40)

	$labelPathTime = New-Object System.Windows.Forms.Label
	$labelPathTime.Text = "User Path: "
	$labelPathTime.AutoSize = $true
	$labelPathTime.Location = New-Object System.Drawing.Point(0, 60)

	$labelBitLockerTime = New-Object System.Windows.Forms.Label
	$labelBitLockerTime.Text = "BitLocker: "
	$labelBitLockerTime.AutoSize = $true
	$labelBitLockerTime.Location = New-Object System.Drawing.Point(0, 80)

	$labelToolsTime = New-Object System.Windows.Forms.Label
	$labelToolsTime.Text = "BitTools: "
	$labelToolsTime.AutoSize = $true
	$labelToolsTime.Location = New-Object System.Drawing.Point(0, 100)

	# Add labels to the panel
	$DebugTimerTab.Controls.AddRange(@(
			$labelTotalTime, $labelMainTime, $labelUserTime, $labelPathTime, $labelBitLockerTime, $labelBitlockerTime, $labeltoolstime
		))

	# Add the panel to your main tab (e.g., $usersTab or $mainTab)
	$DebugTimerTab.Controls.Add($timerPanel)

#>

	#endregion

	#region Bitlocker Tab
	######################################################################################## Bitlocker Tab  ##################################################################################################################

	$bitlockerTab = New-Object System.Windows.Forms.TabPage
	$bitlockerTab.Text = "BitLocker"

	# BitLocker status output box (top)
	$bitlockerOutput = New-Object System.Windows.Forms.TextBox
	$bitlockerOutput.Multiline = $true
	$bitlockerOutput.ScrollBars = "Vertical"
	$bitlockerOutput.ReadOnly = $true
	$bitlockerOutput.Location = New-Object System.Drawing.Point(10, 40)
	$bitlockerOutput.Size = New-Object System.Drawing.Size(820, 300)

	# Label for command return output
	$cmdLabel = New-Object System.Windows.Forms.Label
	$cmdLabel.Text = "Values returned from Button Execution:"
	$cmdLabel.Location = New-Object System.Drawing.Point(10, 360)
	$cmdLabel.Width = 400

	# Command return output box (bottom)
	$bitlockerCommandOutput = New-Object System.Windows.Forms.TextBox
	$bitlockerCommandOutput.Multiline = $true
	$bitlockerCommandOutput.ScrollBars = "Vertical"
	$bitlockerCommandOutput.ReadOnly = $true
	$bitlockerCommandOutput.Location = New-Object System.Drawing.Point(10, 385)
	$bitlockerCommandOutput.Size = New-Object System.Drawing.Size(820, 225)

	# Suspend BitLocker button
	$suspendBtn = New-Object System.Windows.Forms.Button
	$suspendBtn.Text = "Suspend BitLocker (1 Reboot)"
	$suspendBtn.Location = New-Object System.Drawing.Point(10, 10)
	$suspendBtn.Width = 220

	# Enable BitLocker button
	$enableBtn = New-Object System.Windows.Forms.Button
	$enableBtn.Text = "Enable BitLocker Protection"
	$enableBtn.Location = New-Object System.Drawing.Point(240, 10)
	$enableBtn.Width = 220

	# Get BitLocker Key button
	$keyBtn = New-Object System.Windows.Forms.Button
	$keyBtn.Text = "Get BitLocker Key"
	$keyBtn.Location = New-Object System.Drawing.Point(470, 10)
	$keyBtn.Width = 180

	# Add controls to BitLocker tab
	$bitlockerTab.Controls.Add($bitlockerOutput)
	$bitlockerTab.Controls.Add($cmdLabel)
	$bitlockerTab.Controls.Add($bitlockerCommandOutput)
	$bitlockerTab.Controls.Add($suspendBtn)
	$bitlockerTab.Controls.Add($enableBtn)
	$bitlockerTab.Controls.Add($keyBtn)

	#Function for Bitlocker

	function Status-BitLocker
	{
		param (
			[string]$target
		)
		$bitlockerOutput.Text = ""
		$runspace = [runspacefactory]::CreateRunspace()
		$runspace.ApartmentState = "STA"
		$runspace.ThreadOptions = "ReuseThread"
		$runspace.Open()

		$ps = [PowerShell]::Create()
		$ps.Runspace = $runspace

		$ps.AddScript({
				param ($target, $bitlockerOutput)

				try
				{
					$status = & "$env:SystemRoot\System32\manage-bde.exe" -status -computername $target 2>&1
					$bitlockerOutput.Invoke([Action] { $bitlockerOutput.Text = $status -join "`r`n" })
				}
				catch
				{
					[System.Windows.Forms.MessageBox]::Show("Failed to retrieve BitLocker status from $target.`nError: $_", "Error", "OK", "Error")
				}
			}) | Out-Null

		$ps.AddArgument($target)
		$ps.AddArgument($bitlockerOutput)

		$ps.BeginInvoke()
	}



	function Suspend-BitLocker
	{
		param (
			[string]$target
		)

		try
		{
			$cmd = manage-bde -protectors -disable C: -RebootCount 1 -computername $target 2>&1
			$bitlockerCommandOutput.Text = $cmd -join "`r`n"

			Status-BitLocker -target $target
		}
		catch
		{
			[System.Windows.Forms.MessageBox]::Show("Failed to suspend BitLocker on $target.`nError: $_", "Error", "OK", "Error")
		}
	}

	function Enable-BitLocker
	{
		param (
			[string]$target
		)

		try
		{
			$cmd = manage-bde -protectors -enable C: -computername $target 2>&1
			$bitlockerCommandOutput.Text = $cmd -join "`r`n"

			Status-BitLocker -target $target
		}
		catch
		{
			[System.Windows.Forms.MessageBox]::Show("Failed to enable BitLocker on $target.`nError: $_", "Error", "OK", "Error")
		}
	}

	function Get-BitLockerKey
	{
		param (
			[string]$target
		)

		try
		{
			$cmd = manage-bde -protectors -get C: -computername $target 2>&1
			$bitlockerCommandOutput.Text = $cmd -join "`r`n"
		}
		catch
		{
			[System.Windows.Forms.MessageBox]::Show("Failed to retrieve BitLocker key from $target.`nError: $_", "Error", "OK", "Error")
		}
	}


	# Suspend BitLocker button click

	$suspendBtn.Add_Click({
			Suspend-BitLocker -target $Script:ConnectedPC
		})


	# Enable BitLocker button click

	$enableBtn.Add_Click({
			Enable-BitLocker -target $Script:ConnectedPC
		})


	# Get BitLocker Key button click

	$keyBtn.Add_Click({
			Get-BitLockerKey -target $Script:ConnectedPC
		})


	#endregion

	#region User Tab
	######################################################################################## User Tab  ###################################################################################################################
	# Create Users tab
	$usersTab = New-Object System.Windows.Forms.TabPage
	$usersTab.Text = "Users"


	# DataGridView for user profiles
	$script:userGrid = New-Object System.Windows.Forms.DataGridView
	$script:userGrid.Location = New-Object System.Drawing.Point(10, 70)
	$script:userGrid.Size = New-Object System.Drawing.Size(840, 300)
	$script:userGrid.ReadOnly = $true
	$script:userGrid.AllowUserToAddRows = $false
	$script:userGrid.AllowUserToDeleteRows = $false
	$script:userGrid.AutoSizeColumnsMode = 'AllCells'
	$script:userGrid.ScrollBars = "Both"

	# Add these columns BEFORE calling Monitor-UserProfileJob
	$script:userGrid.Columns.Add("UserName", "User Name") | Out-Null
	$script:userGrid.Columns.Add("LastModified", "Last Modified") | Out-Null
	$script:userGrid.Columns.Add("FullName", "Full Name") | Out-Null
	$script:userGrid.Columns.Add("LGLHold", "LGLHold") | Out-Null
	$script:userGrid.Columns.Add("SID", "SID") | Out-Null

	$usersTab.Controls.Add($script:userGrid)

	# LAR Label
	$larLabel = New-Object System.Windows.Forms.Label
	$larLabel.Text = "List of Local Admin Group (LAR):"
	$larLabel.Location = New-Object System.Drawing.Point(10, 380)
	$larLabel.Width = 300
	$larLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
	$usersTab.Controls.Add($larLabel)

	# LAR GridView for user profiles
	$script:larGrid = New-Object System.Windows.Forms.DataGridView
	$script:larGrid.Location = New-Object System.Drawing.Point(10, 410)
	$script:larGrid.Size = New-Object System.Drawing.Size(350, 200)
	$script:larGrid.ReadOnly = $true
	$script:larGrid.AllowUserToAddRows = $false
	$script:larGrid.AllowUserToDeleteRows = $false
	$script:larGrid.AutoSizeColumnsMode = 'Fill'
	$script:larGrid.Columns.Add("UserID", "User ID") | Out-Null
	$usersTab.Controls.Add($script:larGrid)

	#RDP List Label
	$rdplistLabel = New-Object System.Windows.Forms.Label
	$rdplistLabel.Text = "List of users in RDP Group:"
	$rdplistLabel.Location = New-Object System.Drawing.Point(500, 380)
	$rdplistLabel.Width = 300
	$rdplistLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
	$usersTab.Controls.Add($rdplistLabel)

	# RDP GridView for user profiles
	$script:rdplistGrid = New-Object System.Windows.Forms.DataGridView
	$script:rdplistGrid.Location = New-Object System.Drawing.Point(500, 410)
	$script:rdplistGrid.Size = New-Object System.Drawing.Size(350, 200)
	$script:rdplistGrid.ReadOnly = $true
	$script:rdplistGrid.AllowUserToAddRows = $false
	$script:rdplistGrid.AllowUserToDeleteRows = $false
	$script:rdplistGrid.AutoSizeColumnsMode = 'Fill'
	$script:rdplistGrid.Columns.Add("UserID", "User ID") | Out-Null
	$usersTab.Controls.Add($script:rdplistGrid)

	#LGLHold Disclaimer
	$lglholdlabel = New-Object System.Windows.Forms.Label
	$lglholdlabel.Text = "This does not substitute approved LGLHold Process"
	$lglholdlabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)
	$lglholdlabel.Size = New-Object System.Drawing.Size(300, 40)
	$lglholdlabel.Location = New-Object System.Drawing.Point(550, 20)
	$lglholdlabel.AutoSize = $false
	$lglholdlabel.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
	$lglholdlabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	$userstab.Controls.Add($lglholdlabel)



	function Start-UserProfileJob
	{
		param (
			[string]$ComputerName = "localhost"
		)

		# Clean up any previous jobs
		Get-Job | Where-Object { $_.Name -eq "UserProfileJob" } | Remove-Job -Force -ErrorAction SilentlyContinue

		# Start background job
		$script:job = Start-Job -Name "UserProfileJob" -ScriptBlock {
			param($ComputerName)

			$results = @{
				UserProfiles = @()
				Admins       = @()
				RDPUsers     = @()
				Error        = $null
			}

			$userFolderPath = "\\$ComputerName\C$\Users"

			try
			{
				$dirs = Get-ChildItem -Path $userFolderPath -Directory -ErrorAction Stop
			}
			catch
			{
				$results.Error = "Failed to access $userFolderPath.`nError: $_"
				return $results
			}

			try
			{
				$profiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $ComputerName -ErrorAction Stop
			}
			catch
			{
				$profiles = @()
			}

			$csvPath = "\\a70fpcrpnasv002\edis\list of custodians\custodians - all.csv"
			$custodianList = @()
			$csvAvailable = $false

			try
			{
				$stream = [System.IO.File]::Open($csvPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
				$stream.Close()
				$custodianList = Get-Content $csvPath
				$csvAvailable = $true
			}
			catch {}

			foreach ($dir in $dirs)
			{
				$userName = $dir.Name
				$lastModified = $dir.LastWriteTime
				$fullName = "N/A"
				$sid = "N/A"
				$lglHoldText = "Custodian.csv unreachable"
				$lglHoldColor = 'Gray'
				$isBold = $false

				$cleanID = $userName -replace '^A-', '' -replace '\\..*$', '' -replace '^.*?-', ''
				if ($cleanID -match '([a-zA-Z0-9]{4})')
				{
					$parsedID = $matches[1].ToUpper()
				}
				else
				{
					$parsedID = $userName
				}

				try
				{
					$url = "https://techutilities.bcbssc.com/personnel-api/person/$parsedID"
					$response = Invoke-RestMethod -Uri $url -Method Get -Headers @{accept = '*/*' }

					if ($response -and $response.firstName -and $response.lastName)
					{
						$fullName = "$($response.firstName) $($response.lastName)"
					}
					else
					{
						$fullName = "No Data from RACFer"
					}
				}
				catch
				{
					$fullName = "Not Found"
				}


				if ($csvAvailable)
				{
					$lglHoldText = "Not on LGLHOLD"
					$lglHoldColor = 'Green'
					if ($custodianList -match "(?i)\b$parsedID\b")
					{
						$lglHoldText = "USER ON LGLHOLD"
						$lglHoldColor = 'Red'
						$isBold = $true
					}
				}

				$profile = $profiles | Where-Object { $_.LocalPath -like "*\$userName" }
				if ($profile)
				{
					$sid = $profile.SID
				}

				$results.UserProfiles += [PSCustomObject]@{
					UserName     = $userName
					LastModified = $lastModified
					FullName     = $fullName
					LGLHold      = $lglHoldText
					SID          = $sid
					LGLColor     = $lglHoldColor
					IsBold       = $isBold
				}
			}

			function Get-GroupMembers
			{
				param ($ComputerName, $GroupName)
				$members = @()
				try
				{
					$group = [ADSI]"WinNT://$ComputerName/$GroupName,group"
					$group.Members() | ForEach-Object {
						try
						{
							$members += $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)
						}
						catch
						{
							$members += "Unknown"
						}
					}
				}
				catch
				{
					$members += "Error: $_"
				}
				return $members
			}

			$results.Admins = Get-GroupMembers -ComputerName $ComputerName -GroupName "Administrators"
			$results.RDPUsers = Get-GroupMembers -ComputerName $ComputerName -GroupName "Remote Desktop Users"

			return $results
		} -ArgumentList $ComputerName
	}

	function Monitor-UserProfileJob
	{
		param (
			[System.Windows.Forms.DataGridView]$userGrid,
			[System.Windows.Forms.DataGridView]$larGrid,
			[System.Windows.Forms.DataGridView]$rdplistGrid
		)

		if (-not $userGrid -or -not $larGrid -or -not $rdplistGrid)
		{
			Write-Error "One or more grid views are null. Ensure they are initialized before calling this function."
			return
		}


		if ($script:timer)
		{
			$script:timer.Stop()
			$script:timer.Dispose()
			$script:timer = $null
		}

		$script:timer = New-Object System.Windows.Forms.Timer
		$script:timer.Interval = 1000
		$script:timer.Add_Tick({
				if ($script:job -and (Get-Job -Id $script:job.Id -ErrorAction SilentlyContinue))
				{
					if ((Get-Job -Id $script:job.Id).State -eq 'Completed')
					{
						$script:timer.Stop()

						$refreshUsersBtn.Text = "Refresh"
						$refreshUsersBtn.Enabled = $true
						$exportUsersBtn.Enabled = $true
						$exportNoLGLHoldBtn.Enabled = $true

						$data = Receive-Job -Job $script:job
						Remove-Job -Job $script:job

						if ($data.Error)
						{
							[System.Windows.Forms.MessageBox]::Show($data.Error, "Access Error", "OK", "Error")
							return
						}

						$userGrid.Rows.Clear()
						foreach ($item in $data.UserProfiles)
						{
							try
							{
								$rowIndex = $userGrid.Rows.Add(
									$item.UserName,
									$item.LastModified,
									$item.FullName,
									$item.LGLHold,
									$item.SID
								)

								if ($null -ne $rowIndex)
								{
									$row = $userGrid.Rows[$rowIndex]
									$row.Cells["LGLHold"].Style.BackColor = $item.LGLColor
									if ($item.IsBold)
									{
										$row.Cells["LGLHold"].Style.Font = New-Object System.Drawing.Font($userGrid.Font, [System.Drawing.FontStyle]::Bold)
									}
								}
							}
							catch
							{
								Write-Warning "Failed to add row for user $($item.UserName): $_"
							}
						}

						$larGrid.Rows.Clear()
						foreach ($admin in $data.Admins)
						{
							$null = $larGrid.Rows.Add($admin)
						}

						$rdplistGrid.Rows.Clear()
						foreach ($rdpUser in $data.RDPUsers)
						{
							$null = $rdplistGrid.Rows.Add($rdpUser)
						}

						$userGrid.AutoSizeColumnsMode = 'AllCells'
						$null = $userGrid.PerformLayout()
						$columnWidths = @{}
						foreach ($col in $userGrid.Columns)
						{
							$columnWidths[$col.Name] = $col.Width
						}
						$userGrid.AutoSizeColumnsMode = 'None'
						foreach ($col in $userGrid.Columns)
						{
							$col.Width = $columnWidths[$col.Name] + 20
						}
					}
				}
				else
				{
					$script:timer.Stop()
				}
			})

		$refreshUsersBtn.Enabled = $false
		$exportUsersBtn.Enabled = $false
		$exportNoLGLHoldBtn.Enabled = $false
		$refreshUsersBtn.Text = "Running..."


		$script:timer.Start()
	}




	# User/Group Button for Users Tab
	$userGroupBtn = New-Object System.Windows.Forms.Button
	$userGroupBtn.Text = "Click to Open Users/Group Window"
	$userGroupBtn.Location = New-Object System.Drawing.Point(370, 410) 
	$userGroupBtn.Size = New-Object System.Drawing.Size(120, 60)
	$usersTab.Controls.Add($userGroupBtn)

	# Click event
	$userGroupBtn.Add_Click({
			$pc = $compNameBox.Text
			try
			{
				Start-Process "lusrmgr.msc" "/computer=$pc"
			}
			catch
			{
				[System.Windows.Forms.MessageBox]::Show("Failed to open Local Users and Groups on $pc.`nError: $_", "Error", "OK", "Error")
			}
		})




	# Refresh Button
	$refreshUsersBtn = New-Object System.Windows.Forms.Button
	$refreshUsersBtn.Text = "Refresh"
	$refreshUsersBtn.Location = New-Object System.Drawing.Point(10, 40)
	$refreshUsersBtn.Size = New-Object System.Drawing.Size(120, 25)
	$usersTab.Controls.Add($refreshUsersBtn)

	$refreshUsersBtn.Add_Click({

			$script:userGrid.Rows.Clear()
			$largrid.Rows.Clear()
			$rdplistgrid.Rows.Clear()

			$computerName = $compNameBox.Text.Trim()
			if ([string]::IsNullOrWhiteSpace($computerName))
			{
				[System.Windows.Forms.MessageBox]::Show("Please enter a computer name.")
				$connectedPCLabel.Text = "Enter PC Name (Local or Remote) and click Connect"
			}
			else
			{
				$target = if ($compNameBox.Text -eq "") { "localhost" } else { $compNameBox.Text }
				Start-UserProfileJob -ComputerName $target
				Monitor-UserProfileJob -userGrid $script:userGrid -larGrid $script:larGrid -rdplistGrid $script:rdplistGrid
			}
		})


	# Export to CSV Button
	$exportUsersBtn = New-Object System.Windows.Forms.Button
	$exportUsersBtn.Location = New-Object System.Drawing.Point(160, 40)
	$exportUsersBtn.Size = New-Object System.Drawing.Size(120, 25)
	$exportUsersBtn.Text = "Export Users to CSV"
	$usersTab.Controls.Add($exportUsersBtn)

	# Export button click event
	$exportUsersBtn.Add_Click({
			$computerName = if ($Script:ConnectedPC -eq "localhost")
			{
				$env:COMPUTERNAME
			}
			else
			{
				$compNameBox.Text
			}

			$csvDialog = New-Object System.Windows.Forms.SaveFileDialog
			$csvDialog.Filter = "CSV files (*.csv)|*.csv"
			$csvDialog.Title = "Export User Profiles"
			$csvDialog.FileName = "${computerName}_UserProfiles.csv"

			if ($csvDialog.ShowDialog() -eq "OK")
			{
				$exportData = @()
				foreach ($row in $script:userGrid.Rows)
				{
					if ($row -is [System.Windows.Forms.DataGridViewRow] -and -not $row.IsNewRow)
					{
						$exportData += [PSCustomObject]@{
							"User Name"     = $row.Cells[0].Value
							"Last Modified" = $row.Cells[1].Value
							"Full Name"     = $row.Cells[2].Value
							"LGLHold"       = $row.Cells[3].Value
							"SID"           = $row.Cells[4].Value
						}
					}
				}

				try
				{
					$exportData | Export-Csv -Path $csvDialog.FileName -NoTypeInformation -Encoding UTF8
					[System.Windows.Forms.MessageBox]::Show("Export successful!", "Export", "OK", "Information")
				}
				catch
				{
					[System.Windows.Forms.MessageBox]::Show("Failed to export file.`nError: $_", "Export Error", "OK", "Error")
				}
			}
		})



	# Export Users w/o LGLHold Button
	$exportNoLGLHoldBtn = New-Object System.Windows.Forms.Button
	$exportNoLGLHoldBtn.Location = New-Object System.Drawing.Point(300, 40)
	$exportNoLGLHoldBtn.Size = New-Object System.Drawing.Size(200, 25)
	$exportNoLGLHoldBtn.Text = "Export Users w/o LGLHold Data"
	$usersTab.Controls.Add($exportNoLGLHoldBtn)

	$exportNoLGLHoldBtn.Add_Click({
			$computerName = if ($script:ConnectedPC -eq "localhost")
			{
				$env:COMPUTERNAME
			}
			else
			{
				$script:ConnectedPC
			}

			$csvDialog = New-Object System.Windows.Forms.SaveFileDialog
			$csvDialog.Filter = "CSV files (*.csv)|*.csv"
			$csvDialog.Title = "Export Users w/o LGLHold or SID"
			$csvDialog.FileName = "${computerName}_Users.csv"

			if ($csvDialog.ShowDialog() -eq "OK")
			{
				$exportData = @()
				foreach ($row in $script:userGrid.Rows)
				{
					if ($row -is [System.Windows.Forms.DataGridViewRow] -and -not $row.IsNewRow)
					{
						$exportData += [PSCustomObject]@{
							"User Name"     = $row.Cells[0].Value
							"Last Modified" = $row.Cells[1].Value
							"Full Name"     = $row.Cells[2].Value
						}
					}
				}

				try
				{
					$exportData | Export-Csv -Path $csvDialog.FileName -NoTypeInformation -Encoding UTF8
					[System.Windows.Forms.MessageBox]::Show("Export successful!", "Export", "OK", "Information")
				}
				catch
				{
					[System.Windows.Forms.MessageBox]::Show("Failed to export file.`nError: $_", "Export Error", "OK", "Error")
				}
			}
		})


	#endregion

	#region GPO Tab
	########################################################################################  GPO  Tab  ####################################################################################################################

	# Create GPO tab
	$gpoTab = New-Object System.Windows.Forms.TabPage
	$gpoTab.Text = "Group Policy"

	# Output box
	$gpoOutput = New-Object System.Windows.Forms.TextBox
	$gpoOutput.Multiline = $true
	$gpoOutput.ScrollBars = "Vertical"
	$gpoOutput.ReadOnly = $true
	$gpoOutput.Location = New-Object System.Drawing.Point(10, 80)
	$gpoOutput.Size = New-Object System.Drawing.Size(820, 520)
	$gpoTab.Controls.Add($gpoOutput)

	# Run GPUpdate button
	$gpupdateBtn = New-Object System.Windows.Forms.Button
	$gpupdateBtn.Text = "Run GPUpdate"
	$gpupdateBtn.Location = New-Object System.Drawing.Point(10, 10)
	$gpupdateBtn.Size = New-Object System.Drawing.Size(120, 25)
	$gpoTab.Controls.Add($gpupdateBtn)

	# GPUpdate button Click
	$gpupdateBtn.Add_Click({
			$pc = $compNameBox.Text.Trim()
			$psexecPath = Join-Path $scriptBase "PSTools\PsExec.exe"

			$gpupdateBtn.Enabled = $false
			$gpupdateBtn.Text = "Running..."
			$gpoOutput.Clear()
			$gpoOutput.AppendText("Running GPUpdate /Force on $pc. Please wait...`r`n")

			$gpupdateCommand = 'echo Running gpupdate /force. Please standby.... && gpupdate /force'

			try
			{
				if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
				{
					Start-Process -FilePath $cmdExePath -ArgumentList "/k $gpupdateCommand"
					$gpoOutput.AppendText("Started local GPUpdate on $pc.`r`n")
				}
				elseif (Test-Path $psexecPath)
				{
					Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "cmd", "/k", $gpupdateCommand
					$gpoOutput.AppendText("Started PsExec GPUpdate on $pc.`r`n")
				}
				else
				{
					$gpoOutput.AppendText("PsExec.exe not found at $psexecPath`r`n")
				}
			}
			catch
			{
				$gpoOutput.AppendText("Failed to run GPUpdate on $($pc):`r`n$_`r`n")
			}

			$gpupdateBtn.Text = "Run GPUpdate"
			$gpupdateBtn.Enabled = $true
		})


	# Reset GPO button
	$resetGpoBtn = New-Object System.Windows.Forms.Button
	$resetGpoBtn.Text = "Reset GPO"
	$resetGpoBtn.Location = New-Object System.Drawing.Point(140, 10)
	$resetGpoBtn.Size = New-Object System.Drawing.Size(120, 25)
	$gpoTab.Controls.Add($resetGpoBtn)

	# Button click event
	$resetGpoBtn.Add_Click({
			$pc = $compNameBox.Text.Trim()
			$psexecPath = Join-Path $scriptBase "PSTools\PsExec.exe"

			$resetGpoBtn.Enabled = $false
			$resetGpoBtn.Text = "Resetting..."
			$gpoOutput.Clear()
			$gpoOutput.AppendText("Resetting Group Policy on $pc. Please wait...`r`n")



			$gpoCommand = 'echo Deleting C:\Windows\System32\GroupPolicy & rd /s /q C:\Windows\System32\GroupPolicy & echo Deleting C:\Windows\System32\GroupPolicyUsers & rd /s /q C:\Windows\System32\GroupPolicyUsers & echo Deleting C:\Windows\SysWOW64\GroupPolicy & rd /s /q C:\Windows\SysWOW64\GroupPolicy & echo Deleting C:\Windows\SysWOW64\GroupPolicyUsers & rd /s /q C:\Windows\SysWOW64\GroupPolicyUsers & echo Running GPUpdate /force. Please Standby.... & echo .............. & gpupdate /force'


			try
			{
				if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
				{
					Start-Process -FilePath $cmdExePath -ArgumentList "/k $gpoCommand"
					$gpoOutput.AppendText("Started local GPO reset on $pc.`r`n")
				}
				elseif (Test-Path $psexecPath)
				{
					Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "cmd", "/k", $gpocommand
					$gpoOutput.AppendText("Started PsExec GPO reset on $pc.`r`n")
				}
				else
				{
					$gpoOutput.AppendText("PsExec.exe not found at $psexecPath`r`n")
				}
			}
			catch
			{
				$gpoOutput.AppendText("Failed to reset GPO on $($pc):`r`n$_`r`n")
			}

			$resetGpoBtn.Text = "Reset GPO"
			$resetGpoBtn.Enabled = $true
		})


	# Checkbox to specify user
	$useUserCheckbox = New-Object System.Windows.Forms.CheckBox
	$useUserCheckbox.Text = "Specify user"
	$useUserCheckbox.Location = New-Object System.Drawing.Point(403, 10)
	$gpoTab.Controls.Add($useUserCheckbox)

	# Textbox for user input
	$userInputBox = New-Object System.Windows.Forms.TextBox
	$userInputBox.Location = New-Object System.Drawing.Point(510, 10)
	$userInputBox.Width = 150
	$userInputBox.Enabled = $false
	$gpoTab.Controls.Add($userInputBox)

	$useUserCheckbox.Add_CheckedChanged({
			$userInputBox.Enabled = $useUserCheckbox.Checked
		})


	# Generate GPResult button
	$gpresultHtmlBtn = New-Object System.Windows.Forms.Button
	$gpresultHtmlBtn.Text = "Generate GPResult"
	$gpresultHtmlBtn.Location = New-Object System.Drawing.Point(270, 10)
	$gpresultHtmlBtn.Size = New-Object System.Drawing.Size(130, 25)
	$gpoTab.Controls.Add($gpresultHtmlBtn)


	#GPResult Button Click
	$gpresultHtmlBtn.Add_Click({
			$gpoOutput.Clear()
			$gpoOutput.AppendText("Generating your GPResult report. Please standby...`r`n")
			$gpresultHtmlBtn.Enabled = $false
			$gpresultHtmlBtn.Text = "Generating..."

			$target = $compNameBox.Text.Trim()
			$psexecPath = Join-Path $scriptBase "PSTools\\PsExec.exe"

			# Generate timestamped filename
			$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
			$fileName = "GPResult_$timestamp.html"
			$remotePath = "C:\Temp\$fileName"
			$uncPath = "\\$target\C$\Temp\$fileName"

			# Determine scope or user switch
			$userSwitch = if ($useUserCheckbox.Checked -and $userInputBox.Text.Trim() -ne "")
			{
				"/USER:$($userInputBox.Text.Trim())"
			}
			else
			{
				"/SCOPE:computer"
			}

			try
			{
				$cmd = "gpresult $userSwitch /h `"$remotePath`""

				if ($target -eq "localhost" -or $target -eq "127.0.0.1" -or $target -eq $env:COMPUTERNAME)
				{
					Start-Process -FilePath $cmdExePath -ArgumentList "/c $cmd" -Wait
				}
				elseif (Test-Path $psexecPath)
				{
					Start-Process -FilePath $psexecPath -ArgumentList "\\$target", "cmd", "/c", $cmd -Wait
				}
				else
				{
					throw "PsExec.exe not found at $psexecPath"
				}

				if (Test-Path $uncPath)
				{
					Start-Process $uncPath
					$gpoOutput.Text = "GPResult HTML report generated and opened: `r`n$uncPath"
				}
				else
				{
					$gpoOutput.Text = "Failed to generate or access GPResult HTML report at $uncPath. User may not have RSOP data on machine. Try running GPReport without User."
				}
			}
			catch
			{
				$gpoOutput.Text = "Error generating GPResult: `r`n$_"
			}

			$gpresultHtmlBtn.Text = "Generate GPResult"
			$gpresultHtmlBtn.Enabled = $true
		})





	# Export GPO Output
	$exportGpoBtn = New-Object System.Windows.Forms.Button
	$exportGpoBtn.Text = "Export Output"
	$exportGpoBtn.Location = New-Object System.Drawing.Point(670, 10)
	$exportGpoBtn.Size = New-Object System.Drawing.Size(120, 25)
	$gpoTab.Controls.Add($exportGpoBtn)

	### Export Button Click ###
	$exportGpoBtn.Add_Click({

			$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
			$saveDialog.Filter = "Text Files (*.txt)|*.txt"
			$saveDialog.Title = "Save GPO Output"
			$saveDialog.FileName = "$($compNameBox.Text)_GPO_Output.txt"
			if ($saveDialog.ShowDialog() -eq "OK")
			{
				try
				{
					Set-Content -Path $saveDialog.FileName -Value $gpoOutput.Text -Encoding UTF8
					[System.Windows.Forms.MessageBox]::Show("Export successful!", "Export", "OK", "Information")
				}
				catch
				{
					[System.Windows.Forms.MessageBox]::Show("Failed to export file.`nError: $_", "Export Error", "OK", "Error")
				}
			}
		})


	# GPO Diagnostics
	$troubleshootBtn = New-Object System.Windows.Forms.Button
	$troubleshootBtn.Text = "Run Diagnostics"
	$troubleshootBtn.Location = New-Object System.Drawing.Point(10, 40)
	$troubleshootBtn.Size = New-Object System.Drawing.Size(120, 25)
	$gpoTab.Controls.Add($troubleshootBtn)


	function Format-DomainControllerOutput
	{
		param (
			[string[]]$rawOutput
		)

		$formatted = @()
		foreach ($line in $rawOutput)
		{
			if ($line -match "DC:\s+\\\\?(?<dc>[\w\.-]+)")
			{
				$formatted += "Domain Controller: $($matches['dc'])"
			}
			elseif ($line -match "Address:\s+\\\\?(?<ip>[\d\.]+)")
			{
				$formatted += "IP Address: $($matches['ip'])"
			}
			elseif ($line -match "Dom(ain)? Guid:\s+(?<guid>[a-f0-9\-]+)")
			{
				$formatted += "Domain GUID: $($matches['guid'])"
			}
			elseif ($line -match "Dom(ain)? Name:\s+(?<dom>[\w\.-]+)")
			{
				$formatted += "Domain Name: $($matches['dom'])"
			}
			elseif ($line -match "Forest Name:\s+(?<forest>[\w\.-]+)")
			{
				$formatted += "Forest Name: $($matches['forest'])"
			}
			elseif ($line -match "DC Site Name:\s+(?<dcsite>[\w\-]+)")
			{
				$formatted += "DC Site Name: $($matches['dcsite'])"
			}
			elseif ($line -match "Our Site Name:\s+(?<oursite>[\w\-]+)")
			{
				$formatted += "Client Site Name: $($matches['oursite'])"
			}
			elseif ($line -match "Flags:\s+(?<flags>.+)")
			{
				$flags = $matches['flags'] -split '\s+'
				$formatted += "Flags:`r`n  " + ($flags -join "`r`n  ")
			}
			elseif ($line -match "The command completed successfully")
			{
				$formatted += "Status: Success"
			}
		}

		# If nothing matched, return raw output
		if ($formatted.Count -eq 0)
		{
			return $rawOutput -join "`r`n"
		}

		return $formatted -join "`r`n"
	}



	### Troubleshoot Button Click
	$troubleshootBtn.Add_Click({
			$troubleshootBtn.Text = "Running..."
			$gpoOutput.Clear()
			$gpoOutput.AppendText("=== Group Policy Diagnostics. Please Standby..... ===`r`n")

			try
			{
				$output = @()
				$targetComputer = $compNameBox.Text

				# Domain Join Status
				$output += "`r`n[Domain Join Status]"
				$output += "Checks if the computer is joined to a domain."
				$output += "`r`n-----------------------------"
				$output += Get-WmiObject -Class Win32_ComputerSystem -ComputerName $targetComputer | Out-String

				# Domain Controller

				$output += "`r`n[Domain Controller]"
				$output += "`r`n-----------------------------"

				if ($targetComputer -ieq "localhost" -or $targetComputer -ieq $env:COMPUTERNAME)
				{
					$rawOutput = nltest /server:$targetComputer /dsgetdc:$env:USERDOMAIN 2>&1
				}
				else
				{
					try
					{
						$remoteDomain = (Get-WmiObject -Class Win32_ComputerSystem -ComputerName $targetComputer -ErrorAction Stop).Domain
						$rawOutput = nltest /server:$targetComputer /dsgetdc:$remoteDomain 2>&1
					}
					catch
					{
						$rawOutput = @("Failed to retrieve domain information from $($targetComputer): $_")
					}
				}

				$formattedOutput = Format-DomainControllerOutput -rawOutput $rawOutput
				$output += "`r`n$formattedOutput"




				# DNS Resolution
				$output += "`r`n[DNS Resolution]"
				$output += "Tests if the computer name can be resolved via DNS."
				$output += "`r`n-----------------------------"
				try
				{
					$output += "Resolving DNS for: $targetComputer`r`n"
					$output += Resolve-DnsName -Name $targetComputer -ErrorAction Stop | Out-String
				}
				catch
				{
					$output += "DNS resolution failed: $($_.Exception.Message)"
				}





				# Time Synchronization
				$output += "`r`n[Time Synchronization]"
				$output += "Checks the status of the Windows Time service."
				$output += "`r`n-----------------------------"
				try
				{
					$output += w32tm /query /computer:$targetComputer /status 2>&1
				}
				catch
				{
					$output += "Time service query failed: $($_.Exception.Message)"
				}

				# GPO Event Logs
				$output += "`r`n[Recent GPO Events]"
				$output += "Displays recent Group Policy-related events from the System event log."
				$output += "`r`n-----------------------------"
				try
				{
					$events = Get-WinEvent -ComputerName $targetComputer -LogName "System" -FilterXPath "*[System[Provider[@Name='GroupPolicy']]]" -MaxEvents 10 -ErrorAction Stop
					$output += $events | Select-Object TimeCreated, Id, Message | Format-List | Out-String
				}
				catch
				{
					$output += "Failed to retrieve GPO events: $($_.Exception.Message)"
				}

				# Group Policy File Timestamps via UNC
				$output += "`r`n[Group Policy File Timestamps]"
				$output += "Checks the last modified time of local Group Policy files to verify recent updates."
				$output += "`r`n-----------------------------"
				try
				{
					$paths = @(
						"\\$targetComputer\C$\Windows\System32\GroupPolicy\gpt.ini",
						"\\$targetComputer\C$\Windows\System32\GroupPolicy\Machine\Registry.pol"
					)
					foreach ($path in $paths)
					{
						if (Test-Path $path)
						{
							$timestamp = (Get-Item $path).LastWriteTime
							$output += "`r`n$path : $timestamp"
						}
						else
						{
							$output += "`r`n$path : Not found"
						}
					}
				}
				catch
				{
					$output += "Failed to retrieve Group Policy file timestamps: $($_.Exception.Message)"
				}

				# GPResult Summary
				$output += "`r`n[GPResult Summary]"
				$output += "Summarizes slow link detection and loopback processing from Group Policy Results."
				$output += "`r`n-----------------------------"
				try
				{
					$gpresult = gpresult /s $targetComputer /r 2>&1
					$slow = $gpresult | Select-String "slow link"
					if ($slow)
					{
						$output += "Slow link detected:`r`n$($slow.Line)"
					}
					else
					{
						$output += "No slow link detected."
					}

					$loopback = $gpresult | Select-String "Loopback"
					if ($loopback)
					{
						$output += "`r`nLoopback processing info:`r`n$($loopback.Line)"
					}
					else
					{
						$output += "`r`nNo loopback processing detected."
					}
				}
				catch
				{
					$output += "gpresult failed: $($_.Exception.Message)"
				}

				# Group Policy Client Service
				$output += "`r`n[Group Policy Client Service]"
				$output += "Checks if the Group Policy Client service is running."
				$output += "`r`n-----------------------------"
				try
				{
					$gpsvc = Get-WmiObject -Class Win32_Service -ComputerName $targetComputer -Filter "Name='gpsvc'" -ErrorAction Stop
					$output += $gpsvc | Format-List | Out-String
				}
				catch
				{
					$output += "Group Policy Client service not found or inaccessible on ${targetComputer}."
				}

				# WinRM Status
				$output += "`r`n[WinRM Service]"
				$output += "Checks if the Windows Remote Management (WinRM) service is running."
				$output += "`r`n-----------------------------"
				try
				{
					$winrm = Get-WmiObject -Class Win32_Service -ComputerName $targetComputer -Filter "Name='WinRM'" -ErrorAction Stop
					$output += $winrm | Format-List | Out-String
				}
				catch
				{
					$output += "WinRM service not found or inaccessible on $targetComputer."
				}

				# Output to GUI
				$output | ForEach-Object {
					if ($_ -ne $null)
					{
						$gpoOutput.AppendText("$_`r`n")
					}
				}

			}
			catch
			{
				$gpoOutput.AppendText("Diagnostics failed: `r`n$($_.Exception.Message)`r`n")
			}
			finally
			{
				$troubleshootBtn.Text = "Run Diagnostics"
			}
		})







	#GPEdit Button
	$gpediteBtn =
	$gpediteBtn = New-Object System.Windows.Forms.Button
	$gpediteBtn.Text = "GP Edit"
	$gpediteBtn.Location = New-Object System.Drawing.Point(140, 40)
	$gpediteBtn.Size = New-Object System.Drawing.Size(120, 25)
	$gpoTab.Controls.Add($gpediteBtn)

	$gpediteBtn.Add_Click({
			$pc = $compNameBox.Text.Trim()
			$gpoOutput.Clear()
			try
			{
				if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
				{
					gpedit.msc
					$gpoOutput.AppendText("Opened GPEdit on local machine.`r`n")
				}
				else
				{
					gpedit.msc /gpcomputer: $pc
					$gpoOutput.AppendText("Opened GPEdit for $pc.`r`n")
				}
			}
			catch
			{
				$gpoOutput.AppendText("Failed to open GPEdit for $($pc): $_`r`n")
			}
		})




	#endregion

	#region Software Tab
	############################################################################################ Software Tab #####################################################################################################################################

	# Create Software Inventory tab
	$softwareTab = New-Object System.Windows.Forms.TabPage
	$softwareTab.Text = "Software Inventory"

	# Software Inventory label
	$softwareLabel = New-Object System.Windows.Forms.Label
	$softwareLabel.Text = "Installed Software List:"
	$softwareLabel.Location = New-Object System.Drawing.Point(10, 50)
	$softwareLabel.Width = 300
	$softwareTab.Controls.Add($softwareLabel)

	# Gather Software List button
	$gatherSoftwareBtn = New-Object System.Windows.Forms.Button
	$gatherSoftwareBtn.Text = "Gather Software List"
	$gatherSoftwareBtn.Location = New-Object System.Drawing.Point(10, 10)
	$gatherSoftwareBtn.Size = New-Object System.Drawing.Size(160, 30)
	$softwareTab.Controls.Add($gatherSoftwareBtn)

	# DataGridView for software list
	$softwareGrid = New-Object System.Windows.Forms.DataGridView
	$softwareGrid.Location = New-Object System.Drawing.Point(10, 80)
	$softwareGrid.Size = New-Object System.Drawing.Size(820, 530)
	$softwareGrid.ReadOnly = $true
	$softwareGrid.AllowUserToAddRows = $false
	$softwareGrid.AllowUserToDeleteRows = $false
	$softwareGrid.AutoSizeColumnsMode = 'None'
	$softwareGrid.ScrollBars = 'Both' 
	$softwareGrid.ColumnHeadersHeightSizeMode = 'AutoSize'
	$softwareGrid.SelectionMode = 'FullRowSelect'
	$softwareGrid.MultiSelect = $false
	$softwareGrid.ColumnHeadersDefaultCellStyle.Alignment = 'MiddleLeft'
	$softwareGrid.DefaultCellStyle.Alignment = 'MiddleLeft'
	$softwareGrid.Columns.Add("Name", "Name") | Out-Null
	$softwareGrid.Columns.Add("Version", "Version") | Out-Null
	$softwareGrid.Columns.Add("InstallDate", "Install Date") | Out-Null
	$softwareGrid.Columns.Add("Vendor", "Vendor") | Out-Null





	$softwareTab.Controls.Add($softwareGrid)

	#Get software function
	function Get-InstalledSoftware
	{
		param (
			[string]$ComputerName,
			$softwareGrid,
			$gatherSoftwareBtn
		)

		# UI: Set loading state
		$softwareGrid.Rows.Clear()
		$softwareGrid.Refresh()
		$gatherSoftwareBtn.Text = "Loading..."
		$gatherSoftwareBtn.Enabled = $false

		$softwareList = @()
		$registryPaths = @(
			"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
			"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
		)

		try
		{
			if ($ComputerName -eq "localhost" -or $ComputerName -eq $env:COMPUTERNAME)
			{
				$baseKey = [Microsoft.Win32.Registry]::LocalMachine
			}
			else
			{
				$baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
			}

			foreach ($path in $registryPaths)
			{
				$uninstallKey = $baseKey.OpenSubKey($path)
				if ($uninstallKey)
				{
					foreach ($subKeyName in $uninstallKey.GetSubKeyNames())
					{
						$subKey = $uninstallKey.OpenSubKey($subKeyName)
						if ($subKey)
						{
							$name = $subKey.GetValue("DisplayName")
							if ($name)
							{
								$softwareList += [PSCustomObject]@{
									Name        = $name
									Version     = $subKey.GetValue("DisplayVersion")
									InstallDate = $subKey.GetValue("InstallDate")
									Vendor      = $subKey.GetValue("Publisher")
								}
							}
						}
					}
				}
			}

			$softwareList = $softwareList | Sort-Object {
				($_.Name -replace '\s+', '').ToLower() + ($_.Version -replace '\s+', '')
			} -Unique

			foreach ($app in $softwareList)
			{
				$null = $softwareGrid.Rows.Add($app.Name, $app.Version, $app.InstallDate, $app.Vendor)
			}
		}
		catch
		{
			[System.Windows.Forms.MessageBox]::Show("Failed to retrieve software list from $ComputerName.`nError: $_", "Error", "OK", "Error")
		}

		# UI: Reset button state
		$softwareGrid.AutoResizeColumns("AllCells")
		$gatherSoftwareBtn.Text = "Gather Software List"
		$gatherSoftwareBtn.Enabled = $true
	}

	# Gather Software List button click event
	$gatherSoftwareBtn.Add_Click({
			$computerName = $Script:ConnectedPC
			if ([string]::IsNullOrWhiteSpace($computerName))
			{
				[System.Windows.Forms.MessageBox]::Show("Please enter a computer name.")
			}
			else
			{
				Get-InstalledSoftware -ComputerName $computerName -softwareGrid $softwareGrid -gatherSoftwareBtn $gatherSoftwareBtn
			}
		})


	#export button
	$exportSoftwareBtn = New-Object System.Windows.Forms.Button
	$exportSoftwareBtn.Text = "Export to CSV"
	$exportSoftwareBtn.Location = New-Object System.Drawing.Point(180, 10)
	$exportSoftwareBtn.Size = New-Object System.Drawing.Size(120, 30)
	$softwareTab.Controls.Add($exportSoftwareBtn)

	$exportSoftwareBtn.Add_Click({

			$computerName = if ($compNameBox.Text -eq "" -or $compNameBox.Text -eq "localhost")
			{
				$env:COMPUTERNAME
			}
			else
			{
				$compNameBox.Text
			}


			$csvDialog = New-Object System.Windows.Forms.SaveFileDialog
			$csvDialog.Filter = "CSV files (*.csv)|*.csv"
			$csvDialog.Title = "Export Software Inventory"
			$csvDialog.FileName = "${computerName}_SoftwareInventory.csv"

			if ($csvDialog.ShowDialog() -eq "OK")
			{
				$exportData = @()
				foreach ($row in $softwareGrid.Rows)
				{
					if ($row -is [System.Windows.Forms.DataGridViewRow] -and -not $row.IsNewRow)
					{
						$exportData += [PSCustomObject]@{
							"Name"        = $row.Cells[0].Value
							"Version"     = $row.Cells[1].Value
							"InstallDate" = $row.Cells[2].Value
							"Vendor"      = $row.Cells[3].Value
						}
					}
				}

				try
				{
					$exportData | Export-Csv -Path $csvDialog.FileName -NoTypeInformation -Encoding UTF8
					[System.Windows.Forms.MessageBox]::Show("Export successful!", "Export", "OK", "Information")
				}
				catch
				{
					[System.Windows.Forms.MessageBox]::Show("Failed to export file.`nError: $_", "Export Error", "OK", "Error")
				}
			}
		})





	#endregion

	#region Tools Tab
	###############################################################################################   Tools Tab   #################################################################################################



	# Tools Tab
	$toolsTab = New-Object System.Windows.Forms.TabPage
	$toolsTab.Text = "Tools"

	$toolsPanel = New-Object System.Windows.Forms.Panel
	$toolsPanel.Dock = "Fill"
	$toolsTab.Controls.Add($toolsPanel)


	# Create the label for log box
	$logLabel = New-Object System.Windows.Forms.Label
	$logLabel.Text = "Output:"
	$logLabel.Location = New-Object System.Drawing.Point(10, 380)
	$logLabel.Size = New-Object System.Drawing.Size(100, 20)
	$toolsPanel.Controls.Add($logLabel)

	$logBox = New-Object System.Windows.Forms.TextBox
	$logBox.Multiline = $true
	$logBox.ScrollBars = "Vertical"
	$logBox.Location = New-Object System.Drawing.Point(10, 400)
	$logBox.Size = New-Object System.Drawing.Size(460, 150)
	$toolsPanel.Controls.Add($logBox)

	# Create the label for extra buttons
	$buttonLabel = New-Object System.Windows.Forms.Label
	$buttonLabel.Text = "Local Only Tools:"
	$buttonLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$buttonLabel.Location = New-Object System.Drawing.Point(10, 220)
	$buttonLabel.Size = New-Object System.Drawing.Size(120, 20)
	$toolsPanel.Controls.Add($buttonLabel)

	# Create the label for Bigfix buttons
	$bigfixbuttonLabel = New-Object System.Windows.Forms.Label
	$bigfixbuttonLabel.Text = "Bigfix Tools:"
	$bigfixbuttonlabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$bigfixbuttonLabel.Location = New-Object System.Drawing.Point(500, 220)
	$bigfixbuttonLabel.Size = New-Object System.Drawing.Size(100, 20)
	$toolsPanel.Controls.Add($bigfixbuttonLabel)

	# Create the label for PSEXEC buttons
	$psexecLabel = New-Object System.Windows.Forms.Label
	$psexecLabel.Text = "PSEXEC Tools:"
	$psexeclabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$psexecLabel.Location = New-Object System.Drawing.Point(10, 300)
	$psexecLabel.Size = New-Object System.Drawing.Size(100, 20)
	$toolsPanel.Controls.Add($psexecLabel)

	#RDP Status Window
	$rdpStatusLabel = New-Object System.Windows.Forms.Label
	$rdpStatusLabel.Location = New-Object System.Drawing.Point(500, 10)
	$rdpStatusLabel.Size = New-Object System.Drawing.Size(460, 30)
	$rdpStatusLabel.Text = "RDP Status: Unknown"
	$rdpStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
	$toolsPanel.Controls.Add($rdpStatusLabel)

	# PC Reboot Label
	$rebootTitleLabel = New-Object System.Windows.Forms.Label
	$rebootTitleLabel.Location = '500, 380'
	$rebootTitleLabel.Size = '200, 20'
	$rebootTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$rebootTitleLabel.Text = "PC Reboot Tools:"
	$toolsPanel.Controls.Add($rebootTitleLabel)


	<# Connected PC Name for Tools tab
	$connectedComputerLabel_Tools = New-Object System.Windows.Forms.Label
	$connectedComputerLabel_Tools.Location = New-Object System.Drawing.Point(10, 10)
	$connectedComputerLabel_Tools.Size = New-Object System.Drawing.Size(400, 25)
	$connectedComputerLabel_Tools.Text = $connectedComputerLabel.Text  # Sync text
	$connectedComputerLabel_Tools.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
	$toolsPanel.Controls.Add($connectedComputerLabel_Tools)
	#>

	#Function that pulls RDP status (Is used with Enable/Disable RDP button and Connect)
	function Update-RDPStatus
	{
		$pc = $compNameBox.Text
		try
		{
			if ($pc -eq "" -or $pc -eq "localhost" -or $pc -eq $env:COMPUTERNAME)
			{
				# Local machine
				$value = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections"
				$deny = $value.fDenyTSConnections
			}
			else
			{
				# Remote machine via remote registry
				$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $pc)
				$tsKey = $reg.OpenSubKey("System\CurrentControlSet\Control\Terminal Server")
				$deny = $tsKey.GetValue("fDenyTSConnections")
			}

			if ($deny -eq 0)
			{
				$rdpStatusLabel.Text = "RDP Status: Enabled"
				$rdpStatusLabel.BackColor = 'Green'
			}
			else
			{
				$rdpStatusLabel.Text = "RDP Status: Disabled"
				$rdpStatusLabel.BackColor = 'Red'
			}

			$rdpStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
		}
		catch
		{
			$rdpStatusLabel.Text = "RDP Status: Error retrieving status"
			$rdpStatusLabel.ForeColor = 'DarkOrange'
			$rdpStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
		}
	}




	function Add-ToolButton
	{
		param (
			[string]$text,
			[int]$x,
			[int]$y,
			[scriptblock]$action
		)
		$btn = New-Object System.Windows.Forms.Button
		$btn.Text = $text
		$btn.Size = New-Object System.Drawing.Size(150, 30)
		$btn.Location = New-Object System.Drawing.Point($x, $y)
		$btn.Add_Click($action)
		$toolsPanel.Controls.Add($btn)
	}


	Add-ToolButton "C$" 10 10 {
		$pc = $Script:ConnectedPC
		$sharePath = "\\$pc\c$"
		$logBox.Clear()

		try
		{
			Start-Process -FilePath "explorer.exe" -ArgumentList $sharePath -Verb RunAs
			$logBox.AppendText("Opened C$ share on $pc using elevated credentials`r`n")
		}
		catch
		{
			$logBox.AppendText("Failed to open C$ share on $($pc): $_`r`n")
		}
	}



	Add-ToolButton "Comp Mgmt" 160 10 {
		$pc = $Script:ConnectedPC
		Start-Process "compmgmt.msc" "/computer=$pc"
		$logBox.Clear()
		$logBox.AppendText("Opened Computer Management on $pc`r`n")
	}

	Add-ToolButton "Services" 310 10 {
		$pc = $Script:ConnectedPC
		Start-Process "services.msc" "/computer=$pc"
		$logBox.Clear()
		$logBox.AppendText("Opened Services on $pc`r`n")
	}

	Add-ToolButton "User/Group" 10 50 {
		$pc = $Script:ConnectedPC
		Start-Process "lusrmgr.msc" "/computer=$pc"
		$logBox.Clear()
		$logBox.AppendText("Opened Local Users and Groups on $pc`r`n")
	}

	Add-ToolButton "RegEdit" 160 50 {
		$pc = $Script:ConnectedPC
		Start-Process "regedit.exe"
		$logBox.Clear()
		$logBox.AppendText("Opened local RegEdit. To connect to $pc, go to File > Connect Network Registry...`r`n")
	}

	Add-ToolButton "PC Cert" 310 50 {
		$pc = $Script:ConnectedPC
		$logBox.Clear()
		Start-Process "certlm.msc"
		$logBox.AppendText("Opened Local Machine Certificates MMC. To connect to $pc, Right-Click on Certificates - Local Computer > Connect to another computer and enter $pc.`r`n")
	}



	Add-ToolButton "Check / Fix WMI" 10 90 {
		$pc = $Script:ConnectedPC
		$psexecPath = Join-Path $scriptBase "PSTools\\PsExec.exe"
		$logBox.Clear()

		$wmiCommand = 'winmgmt /verifyrepository & winmgmt /salvagerepository'

		if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
		{
			Start-Process -FilePath $cmdExePath -ArgumentList "/k $wmiCommand"
		}
		elseif (Test-Path $psexecPath)
		{
			Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "cmd", "/k", $wmiCommand
		}
		else
		{
			$logBox.AppendText("PsExec.exe not found at $psexecPath`r`n")
			return
		}

		$logBox.AppendText("Completed WMI repair on $pc`r`n")
	}

	Add-ToolButton "Reset WMI Repo" 160 90 {
		$pc = $Script:ConnectedPC
		$psexecPath = Join-Path $scriptBase "PSTools\\PsExec.exe"
		$logBox.Clear()


		$resetCommand = 'net stop winmgmt /y & winmgmt /resetrepository & cd /d %windir%\system32\wbem & for /f %s in (''dir /s /b *.mof *.mfl'') do mofcomp %s & net start winmgmt'



		if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
		{
			Start-Process -FilePath $cmdExePath -ArgumentList "/k $resetCommand"
		}
		elseif (Test-Path $psexecPath)
		{
			Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "cmd", "/k", $resetCommand
		}
		else
		{
			$logBox.AppendText("PsExec.exe not found at $psexecPath`r`n")
			return
		}

		$logBox.AppendText("Completed WMI repository reset on $pc`r`n")
	}





	Add-ToolButton "SFC /Scannow" 10 130 {
		$pc = $Script:ConnectedPC
		$psexecPath = Join-Path $scriptBase "PSTools\PsExec.exe"

		if ($pc -eq "localhost")
		{
			Start-Process -FilePath $cmdExePath -ArgumentList "/k sfc /scannow"
			$logBox.AppendText("Started local SFC scan on $pc`r`n")
		}
		elseif (Test-Path $psexecPath)
		{
			Start-Process -FilePath $psexecPath -ArgumentList "\\$pc cmd /k sfc /scannow"
			$logBox.AppendText("Started PsExec SFC scan on $pc`r`n")
		}
		else
		{
			$logBox.AppendText("PsExec.exe not found at $psexecPath`r`n")
		}
	}


	Add-ToolButton "DISM RestoreHealth" 160 130 {
		$pc = $Script:ConnectedPC
		$psexecPath = Join-Path $scriptBase "PSTools\\PsExec.exe"
		$logBox.Clear()

		$dismCommand = 'DISM /Online /Cleanup-Image /RestoreHealth'

		if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
		{
			Start-Process -FilePath $cmdExePath -ArgumentList "/k $dismCommand"
			$logBox.AppendText("Started local DISM RestoreHealth on $pc`r`n")
		}
		elseif (Test-Path $psexecPath)
		{
			Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "cmd", "/k", $dismCommand
			$logBox.AppendText("Started PsExec DISM RestoreHealth on $pc`r`n")
		}
		else
		{
			$logBox.AppendText("PsExec.exe not found at $psexecPath`r`n")
		}
	}




	Add-ToolButton "Clear C:\Win\Temp" 310 90 {
		$pc = $Script:ConnectedPC
		$psexecPath = Join-Path $scriptBase "PSTools\\PsExec.exe"
		$logBox.Clear()
		$logBox.AppendText("Starting cleanup of C:\Windows\Temp on $pc`r`n")

		$deleteCommand = 'for /d %d in (C:\Windows\Temp\*) do rd /s /q "%d"'

		if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
		{
			Start-Process -FilePath $cmdExePath -ArgumentList "/c $deleteCommand"
		}
		elseif (Test-Path $psexecPath)
		{
			Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "cmd", "/c", $deleteCommand
		}
		else
		{
			$logBox.AppendText("PsExec.exe not found at $psexecPath`r`n")
			return
		}

		$logBox.AppendText("Completed cleanup of C:\Windows\Temp on $pc`r`n")
	}

	#Fix Time Sync
	Add-ToolButton "Fix Time Sync" 310 130 {
		$pc = $Script:ConnectedPC
		$psexecPath = Join-Path $scriptBase "PSTools\\PsExec.exe"
		$logBox.Clear()

		$timeFixCommand = 'net stop w32time && w32tm /unregister && w32tm /register && net start w32time && w32tm /resync && w32tm /query /status && echo. && echo Current Time: %DATE% %TIME% && tzutil /g'

		if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
		{
			Start-Process -FilePath $cmdExePath -ArgumentList "/k $timeFixCommand"
			$logBox.AppendText("Ran time sync fix locally on $pc.`r`n")
		}
		elseif (Test-Path $psexecPath)
		{
			Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "cmd", "/k", $timeFixCommand
			$logBox.AppendText("Ran time sync fix remotely on $pc using PsExec.`r`n")
		}
		else
		{
			$logBox.AppendText("PsExec.exe not found at $psexecPath. Cannot run remote time sync fix.`r`n")
		}
	}

	#Clear Print Queue
	Add-ToolButton "Clear Print Queue" 10 170 {
		$pc = $Script:ConnectedPC
		$logBox.Clear()

		$clearPrintCommand = 'net stop spooler && del /Q /F %systemroot%\System32\spool\PRINTERS\* && net start spooler'

		if ($pc -eq "localhost" -or $pc -eq $env:COMPUTERNAME)
		{
			Start-Process -FilePath "cmd.exe" -ArgumentList "/c $clearPrintCommand"
			$logBox.AppendText("Cleared print queue on local machine.`r`n")
		}
		else
		{
			$psexecPath = Join-Path $scriptBase "PSTools\\PsExec.exe"
			if (Test-Path $psexecPath)
			{
				Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "cmd", "/c", $clearPrintCommand
				$logBox.AppendText("Cleared print queue on $pc using PsExec.`r`n")
			}
			else
			{
				$logBox.AppendText("PsExec.exe not found. Cannot clear print queue on $pc.`r`n")
			}
		}
	}



	######PSEXEC Buttons
	Add-ToolButton "PsExec CMD" 10 320 {
		$pc = $Script:ConnectedPC
		$psexecPath = Join-Path $scriptBase "PSTools\PsExec.exe"
		$logBox.Clear()
		if (Test-Path $psexecPath)
		{
			Start-Process -FilePath $psexecPath -ArgumentList "\\$pc cmd"
			$logBox.AppendText("Started PsExec CMD session on $pc`r`n")
		}
		else
		{
			$logBox.AppendText("PsExec.exe not found at $psexecPath`r`n")
		}
	}

	Add-ToolButton "Delete PSEXESVC.exe" 160 320 {
		$pc = $Script:ConnectedPC
		$taskName = "DeletePSEXESVC"
		$deleteCmd = '"cmd.exe /c sc stop PSEXESVC & del C:\Windows\PSEXESVC.exe"'
		$time = (Get-Date).AddMinutes(1).ToString("HH:mm")

		try
		{
			$logBox.Clear()
			$logBox.AppendText("Creating scheduled task on $pc...`r`n")

			$createOutput = schtasks /Create /S $pc /RU "SYSTEM" /SC ONCE /TN $taskName /TR $deleteCmd /ST $time
			$logBox.AppendText("Create Output:`r`n$createOutput`r`n")

			$runOutput = schtasks /Run /S $pc /TN $taskName
			$logBox.AppendText("Run Output:`r`n$runOutput`r`n")

			Start-Sleep -Seconds 3

			$deleteOutput = schtasks /Delete /S $pc /TN $taskName /F
			$logBox.AppendText("Delete Output:`r`n$deleteOutput`r`n")

			$logBox.AppendText("Scheduled task created and executed to delete PSEXESVC.exe on $pc`r`n")
		}
		catch
		{
			$logBox.AppendText("Failed to create or run scheduled task on $($pc): $_`r`n")
		}
	}



	########RDP Function
	function Set-RDPRegistryValue
	{
		param (
			[string]$ComputerName,
			[int]$Value  # 0 = Enable RDP, 1 = Disable RDP
		)

		$keyPath = "SYSTEM\CurrentControlSet\Control\Terminal Server"
		$logBox.Clear()

		try
		{
			$isLocal = $ComputerName -eq "localhost" -or $ComputerName -eq "127.0.0.1" -or $ComputerName -eq $env:COMPUTERNAME

			if ($isLocal)
			{
				$regKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($keyPath, $true)
			}
			else
			{
				$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName).OpenSubKey($keyPath, $true)
			}

			if ($regKey)
			{
				$regKey.SetValue("fDenyTSConnections", $Value, [Microsoft.Win32.RegistryValueKind]::DWord)
				$regKey.Close()

				$action = if ($Value -eq 0) { "Enabled" } else { "Disabled" }
				$logBox.AppendText("$action RDP on $ComputerName via registry.`r`n")
				Update-RDPStatus
			}
			else
			{
				$logBox.AppendText("Failed to open registry key on $ComputerName.`r`n")
			}
		}
		catch
		{
			$logBox.AppendText("Error modifying RDP setting on $($ComputerName): $_`r`n")
		}
	}


	###RDP Buttons
	Add-ToolButton "Launch RDP" 500 40 {
		$pc = $Script:ConnectedPC
		Start-Process "mstsc.exe" "/v:$pc"
		$logBox.Clear()
		$logBox.AppendText("Opened RDP to $pc`r`n")
	}

	Add-ToolButton "EnableRDP" 660 40 {
		Set-RDPRegistryValue -ComputerName $compNameBox.Text.Trim() -Value 0
	}


	Add-ToolButton "DisableRDP" 660 80 {
		Set-RDPRegistryValue -ComputerName $compNameBox.Text.Trim() -Value 1
	}


	######## Other Tools Buttons

	Add-ToolButton "Add/Remove Programs" 10 240 {
		$logBox.Clear()
		Start-Process "appwiz.cpl"
		$logBox.AppendText("Opening Add/Remove Programs.`r`n")
	}

	Add-ToolButton "Print Management Console"  160 240 {
		$logbox.clear()
		Start-Process "printmanagement.msc"
		$logbox.AppendText("Opening Print Management Console. To connect to printer servers, Right-Click on Print Management in left pane and Add/Remove Servers. Then connect to the desired Print Server. Example: A70TPCRPPRNT002 or A70TPCRPPRNT003 `r`n")

	}

	####BigFix Section

	#open Bigfix Logs
	Add-ToolButton "Open Today's BigFix Logs" 500 240 {
		$pc = $Script:ConnectedPC
		$today = Get-Date -Format "yyyyMMdd"
		$logpath = "\\$pc\C$\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs\$today.log"
		Start-Process -FilePath $script:Trace32Path -ArgumentList "`"$logpath`""
	}

	#restart BESClient (Bigfix) service
	Add-ToolButton "Restart BESClient Service" 660 240 {
		$pc = $Script:ConnectedPC
		$psexecPath = Join-Path $scriptBase "PSTools\\PsExec.exe"
		$logBox.Clear()

		$BESClientCommand = "Restart-Service -Name BESClient -Force"

		if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
		{
			Start-Process powershell -ArgumentList "$BESClientCommand" -Wait -NoNewWindow
			$logBox.AppendText("Restarted BESClient on $pc`r`n")
		}
		elseif (Test-Path $psexecPath)
		{
			Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "-nobanner", "powershell", "-Command", "`"$BESClientCommand`""
			$logBox.AppendText("Restarted BESClient on $pc`r`n")
		}
		else
		{
			$logBox.AppendText("PsExec.exe not found at $psexecPath`r`n")
		}
	}

	#####PC Reboot Section

	# Set global variables
	$global:restartButton = $null
	$global:cancelButton = $null

	# Countdown label
	$rebootLabel = New-Object System.Windows.Forms.Label
	$rebootLabel.Location = '500, 410'
	$rebootLabel.Size = '300, 30'
	$rebootLabel.Text = "Select delay and press 'Restart'"
	$toolsPanel.Controls.Add($rebootLabel)

	# Delay dropdown
	$delayDropdown = New-Object System.Windows.Forms.ComboBox
	$delayDropdown.Location = '500, 480'
	$delayDropdown.Size = '140, 20'
	$delayDropdown.DropDownStyle = 'DropDownList'
	$delayDropdown.Items.AddRange(@("0 seconds", "15 seconds", "30 seconds", "60 seconds", "3 minutes", "5 minutes"))
	$delayDropdown.SelectedIndex = 3  # Default to 60 seconds
	$toolsPanel.Controls.Add($delayDropdown)

	# Reboot timer
	$rebootTimer = New-Object System.Windows.Forms.Timer
	$rebootTimer.Interval = 1000
	$rebootTimer.Add_Tick({
			if ($global:countdown -gt 0)
			{
				$global:countdown--
				$rebootLabel.Text = "Restarting in $global:countdown seconds..."
			}
			else
			{
				$rebootTimer.Stop()
				$rebootLabel.Text = "Select delay and press 'Restart'"
				$logbox.AppendText("A restart on $Script:ConnectedPC has been completed.`r`n")
				$global:restartScheduled = $false
				$restartButton.Enabled = $true
				$cancelButton.Enabled = $false
				Restart-Computer -ComputerName $Script:ConnectedPC -Force
			}
		})

	# Restart Button
	$global:restartButton = New-Object System.Windows.Forms.Button
	$global:restartButton.Text = "Restart  $Script:ConnectedPC"
	$global:restartButton.Size = New-Object System.Drawing.Size(150, 30)
	$global:restartButton.Location = New-Object System.Drawing.Point(500, 440)
	$global:restartButton.Add_Click({
			if (-not $global:restartScheduled)
			{
				$selectedDelay = $delayDropdown.SelectedItem
				$global:countdown = switch -Wildcard ($selectedDelay)
				{
					"0 seconds" { 0 }
					"15 seconds" { 15 }
					"30 seconds" { 30 }
					"60 seconds" { 60 }
					"3 minutes" { 180 }
					"5 minutes" { 300 }
					default { 60 }
				}

				$logbox.Clear()
				$logbox.AppendText("A restart will begin on $Script:ConnectedPC in $global:countdown seconds. Click Cancel Restart to stop reboot command.`r`n")

				$global:restartScheduled = $true
				$global:restartButton.Enabled = $false
				$global:cancelButton.Enabled = $true
				$rebootTimer.Start()
				$rebootLabel.Text = "Restarting in $global:countdown seconds..."
			}
		})
	$toolsPanel.Controls.Add($global:restartButton)

	# Cancel Button
	$global:cancelButton = New-Object System.Windows.Forms.Button
	$global:cancelButton.Text = "Cancel Restart"
	$global:cancelButton.Size = New-Object System.Drawing.Size(150, 30)
	$global:cancelButton.Location = New-Object System.Drawing.Point(660, 440)
	$global:cancelButton.Add_Click({
			if ($global:restartScheduled)
			{
				$rebootTimer.Stop()
				$rebootLabel.Text = "Select delay and press 'Restart'"
				$logbox.Clear()
				$logbox.AppendText("Restart on $Script:ConnectedPC has been canceled.`r`n")
				$global:restartScheduled = $false
				$global:restartButton.Enabled = $true
				$global:cancelButton.Enabled = $false
			}
		})
	$toolsPanel.Controls.Add($global:cancelButton)





	#endregion

	#region Processes Tab


	# Create Processes Tab
	$processesTab = New-Object System.Windows.Forms.TabPage
	$processesTab.Text = "Processes"

	# Filter Type Dropdown
	$filterLabel = New-Object System.Windows.Forms.Label
	$filterLabel.Text = "Filter Type:"
	$filterLabel.Location = New-Object System.Drawing.Point(10, 10)
	$processesTab.Controls.Add($filterLabel)

	$filterDropdown = New-Object System.Windows.Forms.ComboBox
	$filterDropdown.Location = New-Object System.Drawing.Point(160, 8)
	$filterDropdown.Size = New-Object System.Drawing.Size(120, 20)
	$filterDropdown.Items.AddRange(@("All", "User", "Process Name", "PID"))
	$filterDropdown.SelectedIndex = 0
	$processesTab.Controls.Add($filterDropdown)

	# Filter Input TextBox
	$filterInput = New-Object System.Windows.Forms.TextBox
	$filterInput.Location = New-Object System.Drawing.Point(300, 8)
	$filterInput.Size = New-Object System.Drawing.Size(150, 20)
	$filterInput.Enabled = $false
	$processesTab.Controls.Add($filterInput)

	$filterDropdown.Add_SelectedIndexChanged({
			$filterInput.Enabled = ($filterDropdown.SelectedItem -ne "All")
		})

	# Get Processes Button
	$getProcessesButton = New-Object System.Windows.Forms.Button
	$getProcessesButton.Text = "Get Processes"
	$getProcessesButton.Location = New-Object System.Drawing.Point(500, 6)
	$processesTab.Controls.Add($getProcessesButton)

	# Kill Process Button
	$killProcessButton = New-Object System.Windows.Forms.Button
	$killProcessButton.Text = "Kill Selected Process"
	$killProcessButton.Location = New-Object System.Drawing.Point(750, 6)
	$processesTab.Controls.Add($killProcessButton)

	# DataGridView for process list
	$processGrid = New-Object System.Windows.Forms.DataGridView
	$processGrid.Location = New-Object System.Drawing.Point(10, 80)
	$processGrid.Size = New-Object System.Drawing.Size(940, 560)
	$processGrid.ReadOnly = $true
	$processGrid.SelectionMode = "FullRowSelect"
	$processGrid.AllowUserToAddRows = $false
	$processGrid.AutoSizeColumnsMode = 'None'
	$processGrid.ScrollBars = "Both"

	# Add Columns
	$processGrid.Columns.Add("Name", "Process Name") | Out-Null
	$processGrid.Columns.Add("PID", "PID") | Out-Null
	$processGrid.Columns.Add("RAM", "RAM (MB)") | Out-Null
	$processGrid.Columns.Add("Owner", "Owner") | Out-Null
	$processGrid.Columns.Add("Path", "Executable Path") | Out-Null
	$processGrid.Columns.Add("StartTime", "Start Time") | Out-Null
	$processGrid.Columns.Add("CommandLine", "Command Line") | Out-Null

	# Set Column Widths
	$processGrid.Columns["Name"].Width = 160
	$processGrid.Columns["PID"].Width = 60
	$processGrid.Columns["RAM"].Width = 100
	$processGrid.Columns["Owner"].Width = 160
	$processGrid.Columns["Path"].Width = 200
	$processGrid.Columns["StartTime"].Width = 140
	$processGrid.Columns["CommandLine"].Width = 400

	$processesTab.Controls.Add($processGrid)

	# Get Processes Logic
	$getProcessesButton.Add_Click({
			$getProcessesButton.Enabled = $false
			$getProcessesButton.Text = "Getting Processes..."

			$processGrid.Rows.Clear()
			$filter = $filterDropdown.SelectedItem
			$input = $filterInput.Text.Trim()
			$target = $compNameBox.Text.Trim()

			try
			{
				$processes = if ($target -eq "localhost" -or $target -eq "127.0.0.1" -or $target -eq $env:COMPUTERNAME)
				{
					Get-WmiObject -Class Win32_Process
				}
				else
				{
					Get-WmiObject -Class Win32_Process -ComputerName $target
				}

				if (-not $processes)
				{
					[System.Windows.Forms.MessageBox]::Show("No processes retrieved from $target.", "Info", "OK", "Information")
					return
				}

				foreach ($proc in $processes)
				{
					try
					{
						$owner = $proc.GetOwner()
						$user = if ($owner) { "$($owner.Domain)\$($owner.User)" } else { "N/A" }
					}
					catch
					{
						$user = "N/A"
					}

					try
					{
						$startTime = [Management.ManagementDateTimeConverter]::ToDateTime($proc.CreationDate)
					}
					catch
					{
						$startTime = "N/A"
					}

					$commandLine = if ($proc.CommandLine) { $proc.CommandLine } else { "N/A" }
					$ramMB = [System.Decimal]::Round(($proc.WorkingSetSize / 1MB), 2)


					$match = switch ($filter)
					{
						"All" { $true }
						"User" { $user -like "*$input*" }
						"Process Name" { $proc.Name -like "*$input*" }
						"PID" { "$($proc.ProcessId)" -eq $input }
					}

					if ($match)
					{
						$row = $processGrid.Rows.Add()
						$processGrid.Rows[$row].Cells["Name"].Value = $proc.Name
						$processGrid.Rows[$row].Cells["PID"].Value = $proc.ProcessId
						$processGrid.Rows[$row].Cells["RAM"].Value = $ramMB
						$processGrid.Rows[$row].Cells["Owner"].Value = $user
						$processGrid.Rows[$row].Cells["Path"].Value = $proc.ExecutablePath
						$processGrid.Rows[$row].Cells["StartTime"].Value = $startTime
						$processGrid.Rows[$row].Cells["CommandLine"].Value = $commandLine
					}
				}
			}
			catch
			{
				[System.Windows.Forms.MessageBox]::Show("Failed to retrieve processes from $target.`nError: $_", "Error", "OK", "Error")
			}
			finally
			{
				$getProcessesButton.Enabled = $true
				$getProcessesButton.Text = "Get Processes"
			}
		})

	# Kill Process Logic
	$killProcessButton.Add_Click({
			if ($processGrid.SelectedRows.Count -gt 0)
			{
				$selectedRow = $processGrid.SelectedRows[0]
				$procId = $selectedRow.Cells[1].Value
				$pc = $compNameBox.Text.Trim()
				$psexecPath = Join-Path $scriptBase "PSTools\\PsExec.exe"

				if ($pc -eq "localhost" -or $pc -eq "127.0.0.1" -or $pc -eq $env:COMPUTERNAME)
				{
					Start-Process -FilePath $cmdExePath -ArgumentList "/c taskkill /PID $procId /F"
					$logBox.AppendText("Killed process $procId on $pc`r`n")
				}
				elseif (Test-Path $psexecPath)
				{
					Start-Process -FilePath $psexecPath -ArgumentList "\\$pc", "cmd", "/c", "taskkill /PID $procId /F"
					$logBox.AppendText("Killed process $procId on $pc via PsExec`r`n")
				}
				else
				{
					$logBox.AppendText("PsExec.exe not found at $psexecPath`r`n")
				}
			}
		})

	#endregion

	#region Environment Variables Tab
	############################################################################# Environment Variable Tab ###############################################################################################

	# Create PATH Variables tab
	$pathTab = New-Object System.Windows.Forms.TabPage
	$pathTab.Text = "PATH Variables"

	# Label
	$pathLabel = New-Object System.Windows.Forms.Label
	$pathLabel.Text = "System PATH Entries (Ordered):"
	$pathLabel.Location = New-Object System.Drawing.Point(10, 10)
	$pathLabel.Width = 400
	$pathTab.Controls.Add($pathLabel)

	# DataGridView for PATH
	$pathGrid = New-Object System.Windows.Forms.DataGridView
	$pathGrid.Location = New-Object System.Drawing.Point(10, 40)
	$pathGrid.Size = New-Object System.Drawing.Size(840, 520)
	$pathGrid.ReadOnly = $true
	$pathGrid.AllowUserToAddRows = $false
	$pathGrid.AllowUserToDeleteRows = $false
	$pathGrid.AutoSizeColumnsMode = 'Fill'
	$pathGrid.ColumnHeadersHeightSizeMode = 'DisableResizing'
	$pathGrid.ColumnHeadersDefaultCellStyle.Alignment = 'MiddleLeft'
	$pathGrid.RowHeadersVisible = $false
	$pathGrid.SelectionMode = 'FullRowSelect'
	$pathGrid.MultiSelect = $false

	# Add column and disable sorting
	$col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col.Name = "PathEntry"
	$col.HeaderText = "PATH Entry"
	$col.SortMode = "NotSortable"
	$pathGrid.Columns.Add($col) | Out-Null

	$pathTab.Controls.Add($pathGrid)

	# Refresh Button
	$refreshPathBtn = New-Object System.Windows.Forms.Button
	$refreshPathBtn.Text = "Get System PATH"
	$refreshPathBtn.Location = New-Object System.Drawing.Point(500, 10)
	$refreshPathBtn.Size = New-Object System.Drawing.Size(120, 25)
	$pathTab.Controls.Add($refreshPathBtn)

	# Function to get PATH entries
	function Get-PathEntries
	{
		param (
			[string]$computerName
		)

		$refreshPathBtn.Text = "Running..."
		$refreshPathBtn.Enabled = $false
		$pathGrid.Rows.Clear()

		try
		{
			if ($computerName -eq "localhost" -or $computerName -eq $env:COMPUTERNAME)
			{
				$rawPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
			}
			else
			{
				$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computerName)
				$envKey = $reg.OpenSubKey("SYSTEM\CurrentControlSet\Control\Session Manager\Environment")
				$rawPath = $envKey.GetValue("Path")
			}

			$entries = $rawPath -split ";" | Where-Object { $_.Trim() -ne "" }

			foreach ($entry in $entries)
			{
				[void]$pathGrid.Rows.Add($entry)
			}
		}
		catch
		{
			[System.Windows.Forms.MessageBox]::Show("Failed to retrieve PATH from $computerName.`nError: $_", "Error", "OK", "Error")
		}

		$refreshPathBtn.Text = "Get System PATH"
		$refreshPathBtn.Enabled = $true
	}


	# Refresh Button Click Event
	$refreshPathBtn.Add_Click({
			Get-PathEntries -computerName $Script:connectedPC
		})







	#endregion

	#region App Control Tab
	############################################################################################# App Control Tab #########################################################################################
	# Create App Control tab
	$appControlTab = New-Object System.Windows.Forms.TabPage
	$appControlTab.Text = "App Control"

	# Label for current status
	$appStatusLabel = New-Object System.Windows.Forms.Label
	$appStatusLabel.Text = "Trellix App Control Status:"
	$appStatusLabel.Location = New-Object System.Drawing.Point(10, 10)
	$appStatusLabel.Width = 250
	$appControlTab.Controls.Add($appStatusLabel)

	# Button to check status
	$checkStatusBtn = New-Object System.Windows.Forms.Button
	$checkStatusBtn.Text = "Check Status"
	$checkStatusBtn.Location = New-Object System.Drawing.Point(270, 6)
	$checkStatusBtn.Width = 120
	$appControlTab.Controls.Add($checkStatusBtn)

	# Button to enable App Control
	$enableAppCtrlBtn = New-Object System.Windows.Forms.Button
	$enableAppCtrlBtn.Text = "Enable App Control"
	$enableAppCtrlBtn.Location = New-Object System.Drawing.Point(400, 6)
	$enableAppCtrlBtn.Width = 150
	$appControlTab.Controls.Add($enableAppCtrlBtn)

	# Button to update App Control
	$updateAppCtrlBtn = New-Object System.Windows.Forms.Button
	$updateAppCtrlBtn.Text = "Update App Control"
	$updateAppCtrlBtn.Location = New-Object System.Drawing.Point(560, 6)
	$updateAppCtrlBtn.Width = 150
	$appControlTab.Controls.Add($updateAppCtrlBtn)

	<#        # Button to disable App Control
        $disableAppCtrlBtn = New-Object System.Windows.Forms.Button
        $disableAppCtrlBtn.Text = "Disable App Control"
        $disableAppCtrlBtn.Location = New-Object System.Drawing.Point(720, 6)
        $disableAppCtrlBtn.Width = 150
        $appControlTab.Controls.Add($disableAppCtrlBtn)
#>
	# Button to recover
	$recoverBtn = New-Object System.Windows.Forms.Button
	$recoverBtn.Text = "Recover"
	$recoverBtn.Location = New-Object System.Drawing.Point(10, 45)
	$recoverBtn.Width = 120
	$appControlTab.Controls.Add($recoverBtn)

	# Button to lockdown
	$lockdownBtn = New-Object System.Windows.Forms.Button
	$lockdownBtn.Text = "Lockdown"
	$lockdownBtn.Location = New-Object System.Drawing.Point(140, 45)
	$lockdownBtn.Width = 120
	$appControlTab.Controls.Add($lockdownBtn)

	# Button to solidify C:\
	$solidifyCBtn = New-Object System.Windows.Forms.Button
	$solidifyCBtn.Text = "Solidify C:\"
	$solidifyCBtn.Location = New-Object System.Drawing.Point(270, 45)
	$solidifyCBtn.Width = 120
	$appControlTab.Controls.Add($solidifyCBtn)

	# Button to solidify file paths
	$solidifyPathBtn = New-Object System.Windows.Forms.Button
	$solidifyPathBtn.Text = "Solidify Filepath"
	$solidifyPathBtn.Location = New-Object System.Drawing.Point(400, 45)
	$solidifyPathBtn.Width = 150
	$appControlTab.Controls.Add($solidifyPathBtn)

	# Output TextBox
	$appControlOutput = New-Object System.Windows.Forms.TextBox
	$appControlOutput.Multiline = $true
	$appControlOutput.ScrollBars = "Vertical"
	$appControlOutput.ReadOnly = $true
	$appControlOutput.Location = New-Object System.Drawing.Point(10, 80)
	$appControlOutput.Size = New-Object System.Drawing.Size(860, 125)
	$appControlTab.Controls.Add($appControlOutput)


	# Button to check App Control logs
	$logCheckBtn = New-Object System.Windows.Forms.Button
	$logCheckBtn.Text = "App Control Logs Check"
	$logCheckBtn.Location = New-Object System.Drawing.Point(10, 215)
	$logCheckBtn.Width = 200
	$appControlTab.Controls.Add($logCheckBtn)

	# Grid for App Control Logs
	$appLogGrid = New-Object System.Windows.Forms.DataGridView
	$appLogGrid.Location = New-Object System.Drawing.Point(10, 240)
	$appLogGrid.Size = New-Object System.Drawing.Size(860, 360)
	$appLogGrid.ReadOnly = $true
	$appLogGrid.AllowUserToAddRows = $false
	$appLogGrid.AllowUserToDeleteRows = $false
	$appLogGrid.AutoSizeColumnsMode = 'None'
	$appLogGrid.ScrollBars = 'Both'
	$appLogGrid.ColumnHeadersHeightSizeMode = 'AutoSize'
	$appLogGrid.RowHeadersVisible = $false
	$appLogGrid.SelectionMode = 'FullRowSelect'
	$appLogGrid.MultiSelect = $false

	# Define columns
	$col1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col1.Name = "LogFile"
	$col1.HeaderText = "Log File"
	$col1.Width = 250
	$col1.SortMode = "NotSortable"

	$col2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col2.Name = "EventLine"
	$col2.HeaderText = "Denied Event"
	$col2.Width = 600
	$col2.SortMode = "NotSortable"

	$appLogGrid.Columns.Add($col1) | Out-Null
	$appLogGrid.Columns.Add($col2) | Out-Null

	$appControlTab.Controls.Add($appLogGrid)


	#Check App Control Log Button Click
	$logCheckBtn.Add_Click({
			$appLogGrid.Rows.Clear()
			$logCheckBtn.Text = "Reading Logs..."
			$logCheckBtn.Enabled = $false

			$target = $compNameBox.Text

			try
			{
				$logDir = "\\$target\C$\ProgramData\McAfee\Solidcore\Logs"
				$logFiles = @()

				if (Test-Path $logDir)
				{
					$logFiles = Get-ChildItem $logDir | Where-Object { $_.Name -match '^s3diag' }

					if ($logFiles.Count -eq 0)
					{
						$null = $appLogGrid.Rows.Add("No log files found", "")
					}
					else
					{
						foreach ($file in $logFiles)
						{
							$denied = Select-String -Path $file.FullName -Pattern '<EXECUTION_DENIED|<WRITE_DENIED'
							if ($denied)
							{
								foreach ($match in $denied)
								{
									$null = $appLogGrid.Rows.Add($file.Name, $match.Line)
								}
							}
							else
							{
								$null = $appLogGrid.Rows.Add($file.Name, "No denied events found.")
							}
						}
					}
				}
				else
				{
					$null = $appLogGrid.Rows.Add("Log directory not found", $logDir)
				}

				$appLogGrid.AutoResizeColumns("AllCells")
			}
			catch
			{
				$null = $appLogGrid.Rows.Add("Error", "Failed to retrieve logs from ${target}: $_")
			}
			finally
			{
				$logCheckBtn.Text = "App Control Logs Check"
				$logCheckBtn.Enabled = $true
			}
		})

	# Function to refresh status
	function Refresh-AppControlStatus
	{
		$checkStatusBtn.Enabled = $false
		$checkStatusBtn.Text = "Running..."

		$pc = $script:ConnectedPC
		$outputBox = $appControlOutput
		$psexec = $psexecPath
		$appControlOutput.Clear()
		$runspace = [runspacefactory]::CreateRunspace()
		$runspace.ApartmentState = "STA"
		$runspace.ThreadOptions = "ReuseThread"
		$runspace.Open()

		$ps = [PowerShell]::Create()
		$ps.Runspace = $runspace

		$ps.AddScript({
				param($pc, $outputBox, $psexec, $checkStatusBtn)

				try
				{
					& $psexec "\\$pc" cmd /c "sadmin status > C:\Windows\Temp\sadmin_status.txt" 2>$null
					Start-Sleep -Seconds 2
					$uncPath = "\\$pc\C$\Windows\Temp\sadmin_status.txt"

					if (Test-Path $uncPath)
					{
						$status = Get-Content $uncPath -Raw
						$outputBox.Invoke([Action] { $outputBox.Text = $status })
					}
					else
					{
						$outputBox.Invoke([Action] { $outputBox.Text = "Output file not found on $pc." })
					}
				}
				catch
				{
					$outputBox.Invoke([Action] { $outputBox.Text = "Failed to run sadmin status on $pc.`r`nError: $_" })
				}

				# Re-enable the button safely
				$checkStatusBtn.Invoke([Action] {
						$checkStatusBtn.Text = "Check Status"
						$checkStatusBtn.Enabled = $true
					})


			}) | Out-Null


		$ps.AddArgument($pc)
		$ps.AddArgument($outputBox)
		$ps.AddArgument($psexec)
		$ps.AddArgument($checkStatusBtn)

		$ps.BeginInvoke()
	}

	# Event handlers
	$checkStatusBtn.Add_Click({

			Refresh-AppControlStatus

		})

	$enableAppCtrlBtn.Add_Click({
			$enableAppCtrlBtn.Enabled = $false
			$enableAppCtrlBtn.Text = "Running..."
			& $psexecPath "\\$Script:ConnectedPC" cmd /c "sadmin eu"  2>$null
			Start-Sleep -Seconds 2
			Refresh-AppControlStatus
			$enableAppCtrlBtn.Enabled = $false
			$enableAppCtrlBtn.Text = "Enable App Control"
		})


	$updateAppCtrlBtn.Add_Click({
			$updateAppCtrlBtn.Enabled = $false
			$updateAppCtrlBtn.Text = "Running..."
			& $psexecPath "\\$Script:ConnectedPC" cmd /c "sadmin bu"  2>$null
			Start-Sleep -Seconds 2
			Refresh-AppControlStatus
			$updateAppCtrlBtn.Enabled = $true
			$updateAppCtrlBtn.Text = "Update App Control"
		})

	$recoverBtn.Add_Click({

			$recoverBtn.Text = "Running"
			$recoverBtn.Enabled = $false

			$form = New-Object System.Windows.Forms.Form
			$form.Text = "SADMIN Password"
			$form.Size = New-Object System.Drawing.Size(300, 150)
			$form.StartPosition = "CenterScreen"
			$form.TopMost = $true

			$label = New-Object System.Windows.Forms.Label
			$label.Text = "Please provide SADMIN password:"
			$label.AutoSize = $true
			$label.Location = New-Object System.Drawing.Point(10, 20)
			$form.Controls.Add($label)

			$textbox = New-Object System.Windows.Forms.TextBox
			$textbox.Location = New-Object System.Drawing.Point(10, 50)
			$textbox.Width = 260
			$textbox.UseSystemPasswordChar = $true
			$form.Controls.Add($textbox)

			$okButton = New-Object System.Windows.Forms.Button
			$okButton.Text = "OK"
			$okButton.Location = New-Object System.Drawing.Point(200, 80)
			$okButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
			$form.Controls.Add($okButton)

			$form.Add_Shown({ $form.Activate() })

			$textbox.Add_KeyDown({
					if ($_.KeyCode -eq 'Enter')
					{
						$okButton.PerformClick()
					}
				})



			if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
			{
				$securePassword = $textbox.Text
				& $psexecPath "\\$Script:ConnectedPC" cmd /c "sadmin recover -f -z $securePassword" 2>$null
				$securePassword = $null
				Refresh-AppControlStatus
			}
			$recoverBtn.Text = "Recover"
			$recoverBtn.Enabled = $true
		})



	$lockdownBtn.Add_Click({
			$lockdownBtn.Text = "Running"
			$lockdownBtn.Enabled = $false
			& $psexecPath "\\$Script:ConnectedPC" cmd /c "sadmin lockdown" 2>$null
			Refresh-AppControlStatus

			$lockdownBtn.Text = "Lockdown"
			$lockdownBtn.Enabled = $true
		})



	# Solidify C:\ event
	$solidifyCBtn.Add_Click({
			$target = $Script:ConnectedPC
			$SOCommand = "sadmin so c:"

			try
			{
				if ($target -eq "localhost" -or $target -eq "127.0.0.1" -or $target -eq $env:COMPUTERNAME)
				{
					Start-Process -FilePath $cmdExePath -ArgumentList "/k, $SOCommand"

				}
				elseif (Test-Path $psexecPath)
				{
					Start-Process -FilePath $psexecPath -ArgumentList "\\$target", "cmd", "/k", $SOcommand

				}
				else
				{

				}
			}
			catch
			{

			}
		})


	# Solidify Filepath event
	$solidifyPathBtn.Add_Click({
			$form = New-Object System.Windows.Forms.Form
			$form.Text = "Solidify File Paths"
			$form.Size = New-Object System.Drawing.Size(400, 150)
			$form.StartPosition = "CenterScreen"
			$form.TopMost = $true

			$label = New-Object System.Windows.Forms.Label
			$label.Text = "Please provide all paths separated by commas:"
			$label.AutoSize = $true
			$label.Location = New-Object System.Drawing.Point(10, 20)
			$form.Controls.Add($label)

			$textbox = New-Object System.Windows.Forms.TextBox
			$textbox.Location = New-Object System.Drawing.Point(10, 50)
			$textbox.Width = 360
			$form.Controls.Add($textbox)

			$okButton = New-Object System.Windows.Forms.Button
			$okButton.Text = "OK"
			$okButton.Location = New-Object System.Drawing.Point(290, 80)
			$okButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
			$form.Controls.Add($okButton)

			$form.Add_Shown({ $form.Activate() })

			if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
			{
				$paths = $textbox.Text.Trim()
				$target = $Script:ConnectedPC
				$PathSOCommand = "sadmin so $paths"

				try
				{
					if ($target -eq "localhost" -or $target -eq "127.0.0.1" -or $target -eq $env:COMPUTERNAME)
					{
						Start-Process -FilePath $cmdExePath -ArgumentList "/k, $PathSOCommand"
					}
					elseif (Test-Path $psexecPath)
					{
						Start-Process -FilePath $psexecPath -ArgumentList "\\$target", "cmd", "/k", $PathSOcommand
					}
					else
					{
					}
				}
				catch
				{
				}
			}
		})

	#clear CLI

	# Clear CLI Rules Button
	$clearCliBtn = New-Object System.Windows.Forms.Button
	$clearCliBtn.Text = "Clear CLI"
	$clearCliBtn.Size = New-Object System.Drawing.Size(150, 30)
	$clearCliBtn.Location = New-Object System.Drawing.Point(600, 50)
	$appControlTab.Controls.Add($clearCliBtn)

	$clearCliBtn.Add_Click({
			$target = $script:ConnectedPC
			$localComputer = $env:COMPUTERNAME
			$appControlOutput.Clear()

			$command = "sadmin cli --clear"

			try
			{
				if ($target -eq "localhost" -or $target -eq "127.0.0.1" -or $target -eq $localComputer)
				{
					Start-Process -FilePath "cmd.exe" -ArgumentList "/k", $command
					$appControlOutput.AppendText("Executed locally: $command`r`n")
				}
				elseif (Test-Path $psexecPath)
				{
					Start-Process -FilePath $psexecPath -ArgumentList "\\$target", "cmd.exe", "/k", $command
					$appControlOutput.AppendText("Executed remotely on $($target): $command`r`n")
				}
				else
				{
					$appControlOutput.AppendText("PsExec not found at $psexecPath`r`n")
				}
			}
			catch
			{
				$appControlOutput.AppendText("Error executing CLI clear command: $_`r`n")
			}
		})

	# Export App Control Log Button
	$exportAppLogBtn = New-Object System.Windows.Forms.Button
	$exportAppLogBtn.Location = New-Object System.Drawing.Point(250, 215)
	$exportAppLogBtn.Size = New-Object System.Drawing.Size(180, 24)
	$exportAppLogBtn.Text = "Export App Control Logs"
	$appControlTab.Controls.Add($exportAppLogBtn)

	# Export button click event
	$exportAppLogBtn.Add_Click({
			$computerName = if ($script:ConnectedPC -eq "localhost")
			{
				$env:COMPUTERNAME
			}
			else
			{
				$script:ConnectedPC
			}

			$csvDialog = New-Object System.Windows.Forms.SaveFileDialog
			$csvDialog.Filter = "CSV files (*.csv)|*.csv"
			$csvDialog.Title = "Export App Control Logs"
			$csvDialog.FileName = "${computerName}_AppControl_Logs.csv"

			if ($csvDialog.ShowDialog() -eq "OK")
			{
				$exportData = @()
				foreach ($row in $appLogGrid.Rows)
				{
					if ($row -is [System.Windows.Forms.DataGridViewRow] -and -not $row.IsNewRow)
					{
						$exportData += [PSCustomObject]@{
							"File Name" = $row.Cells[0].Value
							"Entry"     = $row.Cells[1].Value
						}
					}
				}

				try
				{
					$exportData | Export-Csv -Path $csvDialog.FileName -NoTypeInformation -Encoding UTF8
					[System.Windows.Forms.MessageBox]::Show("Export successful!", "Export", "OK", "Information")
				}
				catch
				{
					[System.Windows.Forms.MessageBox]::Show("Failed to export file.`nError: $_", "Export Error", "OK", "Error")
				}
			}
		})


	#endregion

	#############################Tabs################################################################################################


	$tabs.TabPages.Add($pcInfoTab)
	$tabs.TabPages.Add($usersTab)
	$tabs.TabPages.Add($softwareTab)
	$tabs.TabPages.Add($pathTab)
	$tabs.TabPages.Add($processesTab)
	$tabs.TabPages.Add($gpoTab)
	$tabs.TabPages.Add($bitlockerTab)
	$tabs.TabPages.Add($appControlTab)
	$tabs.TabPages.Add($toolsTab)
	#	$tabs.TabPages.Add($DebugTimerTab)

	$form.Controls.Add($tabs)


	#############################Tabs################################################################################################


	#region Connect Button code
	######################################################################################### Connect button click event  ####################################################################################################
	$connectBtn.Add_Click({

			#			$swtotal = [System.Diagnostics.Stopwatch]::StartNew()	#timer for all functions when clicking connect

			$target = $compNameBox.Text.Trim()
			$connectBtn.Enabled = $false
			$connectBtn.Text = "Connecting..."

			#Test if PC is reachable
			$sharePath = "\\$target\c$\applications"
			try
			{
				$exists = [System.IO.Directory]::Exists($sharePath)
				if (-not $exists)
				{
					throw "Share not accessible"
				}
			}
			catch
			{
				[System.Windows.Forms.MessageBox]::Show("Unable to connect to $target. The system may be offline or inaccessible.", "Connection Failed", "OK", "Error")
				$connectBtn.Enabled = $true
				$connectBtn.Text = "Connect"
				$compNameBox.Text = "$Script:ConnectedPC"
				return
			}

			#			$swmain = [System.Diagnostics.Stopwatch]::StartNew()	#timer for pcinfo tab
			try
			{
				#Sets connected PC name used in textbox (mainly used later if invalid pc name is entered to revert back to last connected pc)
				$Script:ConnectedPC = $compNameBox.Text.Trim()

				############## PC Info Tab Update
				Update-PCInfoFields -target $Script:ConnectedPC
				#				$swmain.Stop()
				#				$labelMainTime.Text = "PC Info Tab Time: $($swmain.Elapsed.TotalSeconds) sec"

				############# Bitlocker Screen Update on click
				#				$swbit = [System.Diagnostics.Stopwatch]::StartNew()
				Status-Bitlocker -target $Script:ConnectedPC
				$bitlockerCommandOutput.Clear()
				#				$swbit.Stop()
				#				$labelbitlockertime.Text = "Bitlocker Tab time: $($swbit.Elapsed.TotalSeconds) sec"

				############ Software Screen clear
				$softwareGrid.Rows.Clear()

				############ App Control Screen Clear
				$appControlOutput.Clear()
				Refresh-AppControlStatus
				$appLoggrid.rows.clear()

				############GPO Screen update on click
				$gpoOutput.Clear()

				#############Tool Screen Update on click
				#				$swtool = [System.Diagnostics.Stopwatch]::StartNew()
				Update-RDPStatus

				# Reset reboot state if active
				if ($global:restartScheduled)
				{
					$rebootTimer.Stop()
					$global:restartScheduled = $false
				}

				# Reset UI elements
				$rebootLabel.Text = "Select delay and press 'Restart'"
				$logbox.Clear()

				if ($global:restartButton -ne $null)
				{
					$global:restartButton.Enabled = $true
					$global:restartButton.Text = "Restart $Script:ConnectedPC"
				}
				if ($global:cancelButton -ne $null)
				{
					$global:cancelButton.Enabled = $false
				}


				#				$swtool.Stop()
				#				$labeltoolstime.Text = "Tools Time: $($swtool.Elapsed.TotalSeconds) sec"

				############Clear Processes tab
				$processGrid.rows.Clear()

				############ Clear PATH table
				#				$swpath = [System.Diagnostics.Stopwatch]::StartNew()
				Get-PathEntries -computerName $Script:connectedPC
				#				$swpath.Stop()
				#				$labelpathtime.Text = "Path Time: $($swpath.Elapsed.TotalSeconds) sec"

				############Populate Users tab
				#				$swuser = [System.Diagnostics.Stopwatch]::StartNew()
				$script:userGrid.Rows.Clear()
				$largrid.Rows.Clear()
				$rdplistgrid.Rows.Clear()
				$target = if ($compNameBox.Text -eq "") { "localhost" } else { $compNameBox.Text }
				Start-UserProfileJob -ComputerName $target
				Monitor-UserProfileJob -userGrid $script:userGrid -larGrid $script:larGrid -rdplistGrid $script:rdplistGrid
				#				$swuser.Stop()
				#				$labelusertime.Text = "User Function Elapsed: $($swuser.Elapsed.TotalSeconds) seconds"
				#				$swtotal.Stop()
				#				$labelTotalTime.Text = "Total Elapsed: $($swtotal.Elapsed.TotalSeconds) seconds"

				############# End of Connect Click Event
			}
			catch
			{
				[System.Windows.Forms.MessageBox]::Show("Failed to connect or retrieve data from $target.`nError: $_", "Connection Error", "OK", "Error")
			}
			$connectBtn.Enabled = $true
			$connectBtn.Text = "Connect"
		})

	#endregion

	#region Export Button Main Screen
	################################################################################ Export button click event  ###################################################################################
	$exportBtn.Add_Click({
			$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
			$saveDialog.Filter = "Text Files (*.txt)|*.txt"
			$saveDialog.Title = "Save All Output"
			$saveDialog.FileName = "$($ConnectedPC)_FullExport.txt"

			if ($saveDialog.ShowDialog() -eq "OK")
			{
				$exportContent = ""

				# Main Screen
				$exportContent += "-----------------------------`r`nPC Info`r`n-----------------------------`r`n"
				$exportContent += "Computer Name: $($pcinfoComputerName.Text)`r`n"
				$exportContent += "Computer Model: $($pcinfoComputerModel.Text)`r`n"
				$exportContent += "OS Version: $($pcinfoOSVersion.Text)`r`n"
				$exportContent += "Last Boot Time: $($pcinfoLastBootTime.Text)`r`n"
				$exportContent += "BIOS Version: $($pcinfoBIOSVersion.Text)`r`n"
				$exportContent += "Installed RAM: $($pcinfoInstalledRAM.Text)`r`n"
				$exportContent += "Available C Drive Space: $($pcinfoDriveSpace.Text)`r`n"
				$exportContent += "Domain: $($pcinfoDomain.Text)`r`n"
				$exportContent += "OU: $($pcinfoOU.Text)`r`n"
				$exportContent += "System Time: $($pcinfoSystemTime)"
				$exportContent += "IP Address:`r`n$($pcinfoIPAddress.Text)`r`n"
				$exportContent += "Logged-in Users:`r`n$($pcinfoLoggedInUsers.Text)`r`n`r`n"


				# BitLocker
				$exportContent += "-----------------------------`r`nBITLOCKER STATUS`r`n-----------------------------`r`n"
				$exportContent += $bitlockerOutput.Text + "`r`n`r`n"

				# Users
				$exportContent += "-----------------------------`r`nUSERS`r`n-----------------------------`r`n"
				foreach ($row in $Script:userGrid.Rows)
				{
					if (-not $row.IsNewRow)
					{
						$exportContent += "Username: $($row.Cells[0].Value), Last Modified: $($row.Cells[1].Value), Full Name: $($row.Cells[2].Value)`r`n"
					}
				}
				$exportContent += "`r`n"

				# Software Inventory
				$exportContent += "-----------------------------`r`nSOFTWARE INVENTORY`r`n-----------------------------`r`n"
				foreach ($row in $softwareGrid.Rows)
				{
					if (-not $row.IsNewRow)
					{
						$exportContent += "Name: $($row.Cells[0].Value), Version: $($row.Cells[1].Value), Publisher: $($row.Cells[2].Value)`r`n"
					}
				}
				$exportContent += "`r`n"

				#PATH
				$exportContent += "----------------------------------`r`nSYSTEM PATH VARIABLES`r`n-----------------------`r`n"
				foreach ($row in $pathGrid.Rows)
				{
					if (-not $row.IsNewRow)
					{
						$exportContent += "PATH $($row.Cells[0].Value)`r`n"
					}
				}
				$pathGrid
				# GPO
				$exportContent += "-----------------------------`r`nGROUP POLICY RESULTS`r`n-----------------------------`r`n"
				$exportContent += $gpoOutput.Text + "`r`n`r`n"

				# App Control
				$exportContent += "-----------------------------`r`nAPP CONTROL STATUS`r`n-----------------------------`r`n"
				$exportContent += $appControlOutput.Text + "`r`n`r`n"

				$exportContent += "-----------------------------`r`nApp Control Logs`r`n-----------------------------`r`n"
				foreach ($row in $appLogGrid.Rows)
				{
					if (-not $row.IsNewRow)
					{
						$exportContent += "File Name: $($row.Cells[0].Value), Entry: $($row.Cells[1].Value)`r`n"
					}
				}
				$exportContent += "`r`n"


				# Tools
				$exportContent += "-----------------------------`r`nTOOLS OUTPUT`r`n-----------------------------`r`n"
				$exportContent += $toolsOutput.Text + "`r`n`r`n"

				try
				{
					Set-Content -Path $saveDialog.FileName -Value $exportContent -Encoding UTF8
					[System.Windows.Forms.MessageBox]::Show("Export successful!", "Export", "OK", "Information")
				}
				catch
				{
					[System.Windows.Forms.MessageBox]::Show("Failed to export file.`nError: $_", "Export Error", "OK", "Error")
				}
			}
		})

	#endregion

	#####Form Settings
	$form.Topmost = $false
	$form.Add_Shown({ $form.Activate() })
	[void]$form.ShowDialog()

	#####closing bracket for AdminToolKit
}

AdminToolkit
