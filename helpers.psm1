Function Show-Error
{
	Param
	(
		[Parameter(Mandatory=$true)][String]$errorstring
	)
	"******************************************"
	"Sharesniff : v1.0"
	"******************************************"
	""
	"Error : $errorstring"
	""
	
	Exit
}

function Show-Help
{
	"******************************************"
	"Sharesniff : v1.0"
	"******************************************"
	""
	"two arguments are required :"
	"	first argument : type of crawler (http|sharepoint|filesystem)"
	"	second argument : the repository to crawl :"
	"		for a website : an url like https://domain.com"
	"		for a sharepoint : the url of the site like https://sharepoint.domain.com/sites/repo"
	"		for a filesystem : the complete path of the file or folder like c:/folder/file.pptx"
	"	third optional argument : impersonate or not the crawl (default|impersonate)"
	""
}

Export-ModuleMember -Function "Show-Error"
Export-ModuleMember -Function "Show-Help"