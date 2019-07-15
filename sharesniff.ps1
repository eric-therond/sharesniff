Add-type -AssemblyName office, System.Runtime.Serialization
Add-Type -AssemblyName System.Web.Extensions
Add-Type -Path .\dlls\ChilkatDotNet47.dll

Import-Module .\helpers.psm1
Import-Module .\artifacts\common.psm1
Import-Module .\formats\powerpoint.psm1
Import-Module .\formats\excel.psm1
Import-Module .\formats\outlook.psm1
Import-Module .\formats\word.psm1
Import-Module .\formats\text.psm1
Import-Module .\crawlers\common.psm1
Import-Module .\crawlers\spiderhttp.psm1
Import-Module .\crawlers\folder.psm1
Import-Module .\crawlers\file.psm1

function Parse-Config
{
	$configFile = ".\config.json"
	if(Test-Path -Path $configFile -PathType leaf)
	{
		$contentFile = Get-Content $configFile

		Try{
			$jsonserial= New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
			$jsonserial.MaxJsonLength = [int]::MaxValue
			$global:config = $jsonserial.DeserializeObject($contentFile)
		}
		Catch{
			Show-Error "config.json is not a valid json file"
		}
	}

	$config = $global:config

	if(!$config.containsKey("report_name")){
		Show-Error "config : configuration must contain report_name key"
	}

	if(!$config["report_name"]){
		Show-Error "config : configuration must contain not null report_name value"
	}

	if(!$config["max_urls"]){
		Show-Error "config : configuration must contain not null max_urls value"
	}

	foreach ($stringArtifact in $config.strings_artifacts) {
		if(!$stringArtifact.containsKey("name")){
			Show-Error "config : each strings_artifacts must contain name key"
		}

		if(!$stringArtifact["name"]){
			Show-Error "config : each strings_artifacts must contain not null name value"
		}

		if(!$stringArtifact.containsKey("string")){
			Show-Error "config : each strings_artifacts must contain string key"
		}

		if(!$stringArtifact["string"]){
			Show-Error "config : each strings_artifacts must contain not null string value"
		}

		if(!$stringArtifact.containsKey("dvs")){
			Show-Error "config : each strings_artifacts must contain dvs key"
		}

		if(!$stringArtifact["dvs"]){
			Show-Error "config : each strings_artifacts must contain not null dvs value"
		}
	}
}

function Create-Report
{
	$count = $global:globalResults.Count

	$files = @{}
	foreach ($result in $global:globalResults)
	{
		$filename = $result["filename"]
		$nbartifacts = $result["nbartifacts"]
		$dvs = $result["dvs"]

		$totaldvs = [int]$dvs * [int]$nbartifacts

		if($files["$filename"]) {
			$files["$filename"] = [int]$files["$filename"] + [int]$totaldvs
		}
		else {

			$files["$filename"] = [int]$totaldvs
		}
	}

	$filesfinal = New-Object System.Collections.ArrayList
	foreach ($file in $files.GetEnumerator())
	{
		$obj = @{
			filename="$($file.Key)";
			dvs="$($file.Value)"
		}

		$filesfinal.Add($obj) | Out-Null
	}

	$jsonserial= New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
	$Obj2 = $jsonserial.Serialize($filesfinal)

	"var data2 = $Obj2;" | Out-File -filepath "./report/datatable.js"

	$Obj1 = $jsonserial.Serialize($global:globalResults)

	"var data1 = $Obj1;" | Out-File -append -filepath "./report/datatable.js"

	$config = $global:config
	"var report_name = '$($config.report_name)';" | Out-File -append -filepath "./report/datatable.js"

	$date = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

	"var date = '$date';" | Out-File -append -filepath "./report/datatable.js"
}

$global:globalResults = New-Object System.Collections.ArrayList

$nbargs = $($args.Count)
if($nbargs -lt 2)
{
	Show-Help
}
else
{
	Parse-Config

	Try {
		$global:config.word = New-Object -ComObject Word.Application
		$global:config.outlook = New-Object -ComObject Outlook.Application
		$global:config.powerpoint = New-Object -ComObject powerpoint.application
		$global:config.excel = New-Object -ComObject Excel.Application
	}
	Catch {
		Show-Error "Can't open Microsoft office"
	}

	$CurrentDir = $(Get-Location).Path;
	$tempdir = "$CurrentDir\temp\*"
	Remove-Item $tempdir -Force
	$tempdir = "$CurrentDir\cache\*"
	Remove-Item $tempdir -Force

	Switch ($args[0])  {
		"web" { Search-Web $args[1] }
		"sharepoint" { Search-Folder $args[1] }
		"filesystem" {
			if(Test-Path $args[0] -PathType leaf)
			{
				Search-File $args[1]
			}
			else
			{
				Search-Folder $args[1]
			}
		}
		default { Show-Help }
	}

	$global:config.excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($global:config.excel) | Out-Null
	$global:config.excel = $null

	$global:config.powerpoint.quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($global:config.powerpoint) | Out-Null
	$global:config.powerpoint = $null

	$global:config.outlook.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($global:config.outlook) | Out-Null
	$global:config.outlook = $null

	$global:config.word.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($global:config.word) | Out-Null
	$global:config.word = $null

	[gc]::collect()
	[gc]::WaitForPendingFinalizers()

	Create-Report
}
