function Search-Artifacts {
	Param
	(
		[Parameter(Mandatory=$true)][AllowNull()]$lookinto,
		[Parameter(Mandatory=$true)][AllowEmptyString()][string]$lookfor
	)

	if(!$lookinto) {
		Return 0
	}

	# https://www.regular-expressions.info/creditcard.html
	$re = ""
	Switch -Exact ($lookfor)
	{
		":::VISACREDITCARD:::" { $re = "^4[0-9]{12}(?:[0-9]{3})?$"; break }
		":::IBAN:::" { $re = "\b[A-Z]{2}[0-9]{2}(?:[ ]?[0-9]{4}){4}(?!(?:[ ]?[0-9]){3})(?:[ ]?[0-9]{1,2})?\b"; break }
		":::IPV4:::" { $re = "\b((((2[0-4][0-9])|(25[0-5])|[01][0-9][0-9])\.){3})((2[0-4][0-9])|(25[0-5])|[01][0-9][0-9])\b"; break }
		":::EMAIL:::" { $re = "\b[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\b"; break }
		
		Default { $re = $lookfor }
	}

	Return ([regex]::Matches($lookinto, $re)).count
}

Export-ModuleMember -Function "Search-Artifacts"
