Function Search-Excel
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)][String]$ExcelPath,
		[Parameter(Mandatory=$true)][String]$DisplayName
	)
	Try {
		$config = $global:config
		
		# Open the target excel
		$msoTrue = [Microsoft.Office.Core.MsoTriState]::msoTrue
		$msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
		Try{
			$workbook = $config.excel.Workbooks.Open($ExcelPath, 0, $msoTrue)
		} Catch {
			Show-Error "Could not open excel file : $ExcelPath"
		}
		$config.excel.Visible = $false
		# force quit without asking for saving document
		$workbook.Saved = $msoTrue

		# iterate all sheets
		$artifacts = New-Artifacts
		$tempdir = "$($(Get-Location).Path)\temp"

		$i = 0;
		While ($i -lt $workbook.Sheets.Count) {
			$i += 1
			"Sheet $i"

			$curSheet = $workbook.Sheets.Item($i)

			$name = "$tempdir\sheet$i.csv"
			# 6 = xlCSV
			$curSheet.SaveAs($name, 6)
			$content = Get-Content -Path $name

			foreach ($stringArtifact in $config.strings_artifacts) {
				$nbocc = Search-Artifacts $content $($stringArtifact.string)

				$nb_artifacts += $nbocc
				if($nbocc) {
					$artifacts.$($stringArtifact.name).container += "<br>Sheet: $($curSheet.Name)"
					$artifacts.$($stringArtifact.name).nbartifacts += $nbocc
					$artifacts.$($stringArtifact.name).dvs = $stringArtifact.dvs
				}
			}

			foreach ($OLEobj in $curSheet.OLEObjects()) {
				Search-OLEObject $OLEobj $ExcelPath
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

		# Release Excel Com Object resource
		$workbook.Close($false)

		Remove-Item "$tempdir\*.csv" -Force
	} Catch {
		Throw $_
	}
}

Export-ModuleMember -Function "Search-Excel"
