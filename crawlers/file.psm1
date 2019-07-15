Function Search-File {
	Param
	(
		[Parameter(Mandatory=$true)][String]$FilePath
	)

	if(!(Test-Path -Path $FilePath -PathType leaf))
	{
		Show-Error "Can't open : $file"
		Return
	}

	$ExtensionFile = [System.IO.Path]::GetExtension($FilePath)
	$filename = "$($(Get-Location).Path)\$FilePath"

	Search-Extension $filename $ExtensionFile $filename
}

Export-ModuleMember -Function "Search-File"
