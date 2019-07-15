# thanks to https://gist.github.com/jongalloway/4343892

function Search-Powerpoint {
	param(
		[Parameter(Mandatory=$true)][string]$file,
		[Parameter(Mandatory=$true)][String]$DisplayName
	)

	$config = $global:config

	$config.powerpoint.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
	$msoTrue = [Microsoft.Office.Core.MsoTriState]::msoTrue
	$msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
	#$application.visible = $msoFalse

	Try{
		$presentation = $config.powerpoint.Presentations.Open($file, $msoTrue, $msoFalse, $msoFalse)
	} Catch {
		Show-Error "Could not open powerpoint file : $file"
	}

	[int]$nb_artifacts = 0;
	$artifacts = ""
	$slides = 0

	$artifacts = New-Artifacts

	foreach ($slide in $presentation.Slides) {
		$slides += 1
		"slide $slides"
		foreach ($shape in $slide.Shapes) {
			Switch -Exact ($shape.Type)
			{
				# [Microsoft.Office.Core.MsoShapeType]::msoGroup
				6 {
					foreach ($item in $shape.GroupItems) {
						if ($shape.HasTextFrame)
						{
							$textFrame = $shape.TextFrame
							$textRange = $textFrame.TextRange

							foreach ($stringArtifact in $config.strings_artifacts) {
								$nbocc = Search-Artifacts $textRange.Text $stringArtifact.string
								$nb_artifacts += $nbocc

								if($nbocc)
								{
									$artifacts.$($stringArtifact.name).container += "<br>Slide: $slides"
									$artifacts.$($stringArtifact.name).nbartifacts += $nbocc
									$artifacts.$($stringArtifact.name).dvs = $stringArtifact.dvs
								}
							}
						}
					}
					break
			  }
				# msoEmbeddedOLEObject
				7 {
					$oleObj = $shape.OLEFormat
					Search-OLEObject $oleObj $file
					break
				}
				Default {
					if ($shape.HasTextFrame)
			  	{
						$textFrame = $shape.TextFrame
						$textRange = $textFrame.TextRange

						foreach ($stringArtifact in $config.strings_artifacts) {
							$nbocc = Search-Artifacts $($textRange.Text) $($stringArtifact.string)
							$nb_artifacts += $nbocc

							if($nbocc)
							{
								$artifacts.$($stringArtifact.name).container += "<br>Slide: $slides"
								$artifacts.$($stringArtifact.name).nbartifacts += $nbocc
								$artifacts.$($stringArtifact.name).dvs = $stringArtifact.dvs
							}
						}
			 		}
				}
			}
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

	$presentation.Close()
}

Export-ModuleMember -Function "Search-Powerpoint"
