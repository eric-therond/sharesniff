Function Search-Outlook
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)][String]$OutlookPath,
		[Parameter(Mandatory=$true)][String]$DisplayName
	)
	Try {
		$config = $global:config

		# Open the target Outlook
		$namespace = $config.outlook.GetNamespace('MAPI')
		$mailItem = $NameSpace.OpenSharedItem($OutlookPath)

		$artifacts = New-Artifacts
		$nb_artifacts = 0

		$properties = @(
			"Attachments",
			"Subject",
			"HTMLBody",
			"ReceivedByName",
			"ReceivedOnBehalfOfName",
			"ReplyRecipientNames",
			"SenderName",
			"SentOnBehalfOfName",
			"To",
			"SenderEmailAddress"
		)

		$tmpDir = "$($(Get-Location).Path)\temp"


		foreach($attachment in $mailItem.Attachments) {
			$name = "$tmpDir\$($attachment.DisplayName)"

			$attachment.SaveAsFile($name);
			$filename = [System.IO.Path]::GetFileName($name)
			$extension = [System.IO.Path]::GetExtension($name)
			$displayNameAttachedFile = "$filename attached in $OutlookPath"
			Search-Extension $name $extension $displayNameAttachedFile
			Remove-Item -Path $name -Force
		}

		foreach($prop in $properties) {
			if(Get-Member -inputobject $mailItem -name "$prop" -Membertype Properties) {
				foreach ($stringArtifact in $config.strings_artifacts) {
					$nbocc = Search-Artifacts $mailItem[$prop] $stringArtifact.string
					$nb_artifacts += $nbocc

					if($nbocc)
					{
						$artifacts.$($stringArtifact.name).container += "<br>Mail property: $prop"
						$artifacts.$($stringArtifact.name).nbartifacts += $nbocc
						$artifacts.$($stringArtifact.name).dvs = $stringArtifact.dvs
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

		# Release Outlook Com Object resource
	} Catch {
		if($($global:config.dev)) {
			Throw $_
		}
	}
}

Export-ModuleMember -Function "Search-Outlook"
