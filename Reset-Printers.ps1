<#
	.SYNOPSIS
	Remove and reinstall all network printers
	.DESCRIPTION
	This script removes (including drivers) all network printers then reinstalls
	them which should fix most printer related issues.
#>

[CmdletBinding()]
Param(
	
)

begin {
	$principal = New-Object System.Security.Principal.WindowsPrincipal(
		[System.Security.Principal.WindowsIdentity]::GetCurrent()
	)
	$role = [System.Security.Principal.WindowsBuiltInRole]::Administrator
	 
	if (!$principal.IsInRole($role)) {
		$command = $MyInvocation.Line.Replace(
			$MyInvocation.InvocationName,
			$MyInvocation.MyCommand.Definition
		)
		
		Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList "$command"
		Exit
	}

	$nw = New-Object -ComObject "WScript.Network"
}

process {
	# Get network printers
	$nwprinters = Get-WmiObject -Query "select Name,Default,ServerName,ShareName,DriverName from Win32_Printer where Network='TRUE'" | `
		Select-Object Name,Default,ServerName,ShareName,DriverName
		
	$nwprinters | Out-File "C:\Temp\printers.txt" -ErrorAction SilentlyContinue
	
	# Remove network printers
	foreach ($printer in $nwprinters) {
		$path = Join-Path $printer.ServerName $printer.ShareName
	
		try { $nw.RemovePrinterConnection($path, $true, $true) }
		catch {
			Write-Warning "Unable to remove: $path"
			$_
		}
	}

	Restart-Service -Name "Spooler" -Force
	
	# Delete all drivers that aren't in use
	Get-WMIObject -Class "Win32_PrinterDriver" | % {
		try { $_.Delete() }
		catch { Write-Warning "Unable to remove driver: $($_.Name)" }
	}
	
	# Add printers back
	foreach ($printer in $nwprinters) {
		$path = Join-Path $printer.ServerName $printer.ShareName
	
		try { $nw.AddWindowsPrinterConnection($path, $true, $true) }
		catch {
			Write-Warning "Unable to add: $path"
			$_
		}
	}
	
	# Set default printer if it was a network printer
	$default = $nwprinters | ? {$_.Default -eq $true}
	
	if ($default) {
		try { $nw.SetDefaultPrinter((Join-Path $default.ServerName $default.ShareName)) }
		catch {
			Write-Warning "Unable to set default printer."
			$_
		}
	}
}

end {
	
}
