Function Search-Sharepoint {
	Param
	(
		[Parameter(Mandatory=$true)][String]$SharepointPath
	)

	if(!Test-Path $FolderPath)
	{
		Show-Error "Can't open : $FolderPath"
		Return
	}

	$total_files = $(Get-ChildItem $FolderPath -recurse).Count
	if($($global:config.dev)) {
		"Total files to analyze : $total_files"
	}

	Get-SPWeb $SharepointPath |
		Select -ExpandProperty Lists |
		Where { $_.GetType().Name -eq "SPDocumentLibrary" -and -not $_.Hidden } |
		Select -ExpandProperty Items |
		Where { $_.Name -Like "*.docx" } |
		Select Name, @{
			Name="URL";
			Expression={$_.ParentList.ParentWeb.Url + "/" + $_.Url}
		}

	$i = 0
	Get-ChildItem $FolderPath -recurse | Foreach-Object {
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

Export-ModuleMember -Function "Search-Sharepoint"
