Function Activate-Object {
	Param
	(
		[Parameter(Mandatory=$true)][AllowNull()]$OLEObject
	)
	
	if($OLEObject)
	{
		if($OLEObject.ProgID -AND $OLEObject.ProgID -ne "Package") {
			$OLEObject.Activate() | Out-Null
		}
	}
}


function Search-OLEObject {
	Param
	(
		[Parameter(Mandatory=$true)][AllowNull()]$OLEObject,
		[Parameter(Mandatory=$true)][string]$FileSource
	)

	if($OLEObject) {
		Switch -Regex ($OLEObject.ProgID)
		{
			"Excel" { $extentionfile = ".xls"; Activate-Object $OLEObject; break }
			"Powerpoint" { $extentionfile = ".ppt"; Activate-Object $OLEObject; break }
			"Outlook" { $extentionfile = ".msg"; Activate-Object $OLEObject; break }
			"Word" { $extentionfile = ".doc"; Activate-Object $OLEObject; break }
			Default { $extentionfile = ".txt" }
		}

		# could be package type
		# https://social.msdn.microsoft.com/Forums/vstudio/en-US/c751c3ae-235d-4327-a26b-74fc297263b6/word-embedded-object-of-type-quotpackagequot?forum=vsto

		$obj = $OLEObject.Object
		$CurrentDir = $(get-location).Path;
		$name = "$CurrentDir\temp\objecttemp$extentionfile"

		try {
			$obj.SaveAs($name);
			$filename = [System.IO.Path]::GetFileName($name)
			$extension = [System.IO.Path]::GetExtension($name)
			$displayNameAttachedFile = "object $($OLEObject.Name) type of $($OLEObject.ProgID) embedded in $FileSource"
			Search-Extension $name $extension $displayNameAttachedFile
			Remove-Item -Path $name -Force
		}
		Catch {
			if($($global:config.dev)) {
				"Could not open embedded object"
			}
		}
	}
}

function Search-Extension {
	Param
	(
		[Parameter(Mandatory=$true)][String]$Fullname,
		[Parameter(Mandatory=$true)][AllowEmptyString()][String]$Extension,
		[Parameter(Mandatory=$true)][String]$DisplayName
	)

	Switch -Regex ($Extension)
	{
		".msg" { Search-Outlook $Fullname $DisplayName }
		".ppt" { Search-Powerpoint $Fullname $DisplayName ; break }
		".xls" { Search-Excel $Fullname $DisplayName ; break }
		".doc" { Search-Word $Fullname $DisplayName ; break }
		Default { Search-Text $Fullname $DisplayName }
	}
}

function New-Artifacts {
	$config = $global:config

	$artifacts = @{}
	foreach ($stringArtifact in $config.strings_artifacts) {
		if($stringArtifact.name)
		{
			$obj = @{
				"container" = "";
				"nbartifacts" = 0
			}

			if(!$artifacts.containsKey("$($stringArtifact.name)")) {
				$artifacts.add("$($stringArtifact.name)", $obj);
			}
		}
	}

	return $artifacts
}

Export-ModuleMember -Function "Search-OLEObject"
Export-ModuleMember -Function "Search-Extension"
Export-ModuleMember -Function "New-Artifacts"
