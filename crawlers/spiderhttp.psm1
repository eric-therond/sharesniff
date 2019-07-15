function Search-Web {
	Param
	(
		[Parameter(Mandatory=$true)][String]$url
	)

	$spider = New-Object Chilkat.Spider
	$seenDomains = New-Object Chilkat.StringArray
	$seedUrls = New-Object Chilkat.StringArray

	$seenDomains.Unique = $true
	$seedUrls.Unique = $true

	#  You will need to change the start URL to something else...
	$seedUrls.Append($url)

	#  Use a cache so we don't have to re-fetch URLs previously fetched.
	$spider.CacheDir = "..\cache\"
	$spider.FetchFromCache = $true
	$spider.UpdateCache = $true
	$spider.ConnectTimeout = 3

	$tmpDir = "$($(Get-Location).Path)\temp";

	$config = $global:config

	while ($seedUrls.Count -gt 0) {
		$url = $seedUrls.Pop()
		$spider.Initialize($url)

		#  Spider 5 URLs of this domain.
		#  but first, save the base domain in seenDomains
		$domain = $spider.GetUrlDomain($url)
		$seenDomains.Append($spider.GetBaseDomain($domain))

		for ($i = 0; $i -le $config:max_urls; $i++) {
			if ($spider.CrawlNext() -ne $true) {
				break
			}
			"url crawled : $($spider.LastUrl)"

			#  Display the URL we just crawled.
			$extension = [System.IO.Path]::GetExtension($spider.LastUrl)
			if(!$extension) {
				$extension = ".html"
			}
			$name = "$tmpDir\webpagetmp$extension"

			Invoke-WebRequest -Uri $($spider.LastUrl) -OutFile $name
			Search-Extension $name $extension $($spider.LastUrl)
			Remove-Item -Path $name -Force

			#  If the last URL was retrieved from cache,
			#  we won't wait.  Otherwise we'll wait 1 second
			#  before fetching the next URL.
			if ($spider.LastFromCache -ne $true) {
				$spider.SleepMs(1000)
			}
		}
	}
}

Export-ModuleMember -Function "Search-Web"
