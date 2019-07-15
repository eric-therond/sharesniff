function Search-Folder {
	Param
	(
		[Parameter(Mandatory=$true)][String]$FolderPath
	)

	if(!(Test-Path $FolderPath))
	{
		Show-Error "Can't open : $FolderPath"
		Return
	}

	$total_files = $(Get-ChildItem $FolderPath -recurse -File).Count

	if($($global:config.dev)) {
		"Total files to analyze : $($total_files)"
	}

	$i = 0
	Get-ChildItem $FolderPath -recurse -File | Foreach-Object {
		Try {
			$ExtensionFile = $_.extension
			$FullNameFile = $_.FullName

			"analyze $FullNameFile"

			Write-Progress -Activity "Analyze folder" -Status "Progress:" -PercentComplete ($i/$total_files*100)

			Search-Extension $FullNameFile $ExtensionFile $FullNameFile
			$i += 1
		}
		Catch {
			if($($global:config.dev)) {
				Write-Host $_ -fore Red
			}
		}
	}
}

Export-ModuleMember -Function "Search-Folder"
