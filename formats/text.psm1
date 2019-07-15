Function Search-Text
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)][String]$TextPath,
		[Parameter(Mandatory=$true)][String]$DisplayName
	)

	$artifacts = New-Artifacts
	$config = $global:config
	
	$content = Get-Content $TextPath

	foreach ($stringArtifact in $config.strings_artifacts) {
		$nbocc = Search-Artifacts $content $($stringArtifact.string)
		$nb_artifacts += $nbocc

		if($nbocc) {
			$artifacts.$($stringArtifact.name).container = ""
			$artifacts.$($stringArtifact.name).nbartifacts += $nbocc
			$artifacts.$($stringArtifact.name).dvs = $stringArtifact.dvs
		}
	}

	if ($nb_artifacts -gt 0)
	{
		foreach ($kvp in $artifacts.GetEnumerator()) {
			$key = $kvp.Key
			$val = $kvp.Value

			if($val["nbartifacts"] -gt 0) {

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
}

Export-ModuleMember -Function "Search-Text"
