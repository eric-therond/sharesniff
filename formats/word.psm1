Function Search-Word
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)][String]$WordPath,
		[Parameter(Mandatory=$true)][String]$DisplayName
	)
	Try {
		$config = $global:config

		# Open the target Word
		$msoTrue = [Microsoft.Office.Core.MsoTriState]::msoTrue
		$msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
		$config.word.Visible = $false
		$config.word.DisplayAlerts = 0
		$doc = $config.word.Documents.OpenNoRepairDialog($WordPath, $false, $false, $false, "", "", $true, "", "", 0, $null, $true, $true, 0, $false)
		# force quit without asking for saving document
		$doc.Saved = $true

		foreach($sh In $doc.InlineShapes) {
			$oleObj = $sh.OLEFormat
			Search-OLEObject $oleObj $WordPath
		}
		
		$randomnumber = Get-Random
		$tempdir = "$($(Get-Location).Path)\temp"
		$name = "$tempdir\wordtmp$randomnumber.txt"
		$doc.SaveAs2([ref]$name, [ref][Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatUnicodeText)
		$content = Get-Content -Path $name
		$artifacts = New-Artifacts
		
		foreach ($stringArtifact in $config.strings_artifacts) {
			$nbocc = Search-Artifacts $content $($stringArtifact.string)

			$nb_artifacts += $nbocc
			if($nbocc) {
				$artifacts.$($stringArtifact.name).nbartifacts += $nbocc
				$artifacts.$($stringArtifact.name).dvs = $stringArtifact.dvs
			}
		}

		if ($nb_artifacts -gt 0)
		{
			foreach ($kvp in $artifacts.GetEnumerator()) {
				$key = $kvp.Key
				$val = $kvp.Value

				if($val.nbartifacts -gt 0) {

					$obj = @{
						filename="$DisplayName";
						nbartifacts="$($val.nbartifacts)";
						artifact="<span data-toggle='tooltip' title='$($val.container)'>$key</span>";
						dvs="$($val.dvs)"
					}

					$global:globalResults.Add($obj) | Out-Null
				}
			}
		}

		$doc.Close([ref][Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
		Remove-Item -Path $name -Force
	} Catch {
		if($($global:config.dev)) {
			Throw $_
		}
	}
}

Export-ModuleMember -Function "Search-Word"
