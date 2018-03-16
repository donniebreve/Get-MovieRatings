<#
.DESCRIPTION
Gets movie information and ratings from various websites and outputs in a tab delimited format.

.PARAMETER title
The movie title.

.EXAMPLE
Get-MovieRating -title "Memento"

.NOTES
#>
function Get-MovieRatings {

    param (
        [String]$title
    )

    Write-Host

    # get imdb/mc information, we use this as the primary information because it is usualy more correct
    $movie = Get-IMDbInformation $title
    # get rottentomatoes ratings
    $movie.RT = Get-RTInformation $movie.Title $movie.Year
    
    Write-Host "Movie"
    Write-Host "---------------"
    Write-Host "Title: $($movie.Title)"
    Write-Host "Year:  $($movie.Year)"
    Write-Host "Genre: $($movie.Genre)"
    Write-Host "IMDb:  $($movie.IMDb)"
    Write-Host "MC:    $($movie.MC)"
    Write-Host "RT:    $($movie.RT)"
    Write-Host

    Set-Clipboard -Value "$($movie.Title)`t$($movie.Year)`t$($movie.Genre)`t$($movie.IMDb)`t$($movie.MC)`t$($movie.RT[0])`t$($movie.RT[1])"
    Write-Host "Copied to clipboard." -ForegroundColor "DarkGray"
}

function Get-IMDbInformation {

    param (
        [Parameter(Mandatory=$true)]
        [String]$title
    )

    Write-Host "Retrieving IMDb information..."

    # encode the title
    $encodedTitle = $title -replace " ", "+"
    $encodedTitle = $encodedTitle -replace "&", "%26"

    # search IMDb
    $url = "http://www.imdb.com/find?q=$encodedTitle&s=all"
    write-host "imdb: $url" -ForegroundColor "DarkGray"
    $response = Invoke-WebRequest $url -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:58.0) Gecko/20100101 Firefox/58.0"
    
    # check for no results
    $elements = $response.ParsedHtml.getElementsByClassName('noresults')
    if ($elements.Length -gt 0) {
        return "No results found"
    }

    # check results
    $elements = $response.ParsedHtml.getElementsByClassName('result_text')
    $results = New-Object "System.Collections.ArrayList"
    Write-Host "Results:"
    for ($i = 0; $i -lt $elements.Length; $i++) {
        $match = [regex]::Match($elements[$i].innerHTML, '(?i)^<a href="(\/title\/[^\/]+)\/[^"]*">([^<]+)<\/a>.+\(([0-9]{4})\)')
        if ($match.Success) {
            $result = @{}
            $result.Url = "http://www.imdb.com$($match.Groups[1].value)"
            $result.Title = $match.Groups[2].value
            $result.Year = $match.Groups[3].value
            [void]$results.Add($result)
            Write-Host -NoNewline "$($i + 1): $($result.Title) ($($result.Year))"
            Write-Host " [$($result.Url)]" -ForegroundColor "DarkGray"
        }
    }
    if ($results.Count -gt 1) {
        $selection = Read-Host "Please select a movie"
        $result = $results[$selection - 1]
    }
    if ($results.Count -eq 1) {
        $result = $results[0]
    }

    # get movie details
    $response = Invoke-WebRequest $result.Url
    $html = [System.Net.WebUtility]::HtmlDecode($response.ParsedHtml.getElementsByClassName('title_bar_wrapper')[0].innerHTML)

    $movie = @{}

    # get the title
    $match = [regex]::Match($html, '(?i)itemprop="name">([^<]+)')
    if ($match.Success) {
        $movie.Title = $match.Groups[1].value.Trim()
    }
    # get the year
    $match = [regex]::Match($html, '(?i)id=titleYear>\(<a href="\/year\/([0-9]+)\/')
    if ($match.Success) {
        $movie.Year = $match.Groups[1].value.Trim()
    }
    # get the genre
    $matches = [regex]::Matches($html, '(?i)itemprop="genre">([^<]+)')
    if ($matches.Success) {
        for ($i = 0; $i -lt $matches.Groups.Length; $i+=2) {
            if ($i -eq 0) { $genre = $matches.Groups[$i+1].value.Trim() }
            else { $genre += ", $($matches.Groups[$i+1].value.Trim())" }
        }
    }
    $movie.Genre = $genre
    # get the rating
    $match = [regex]::Match($html, '(?i)itemprop="ratingValue">([^<]+)')
    if ($match.Success) {
        $movie.IMDb = [convert]::ToDecimal($match.Groups[1].value.Trim()) * 10
    }
    # get the metacritic rating
    $mcElements = $response.ParsedHtml.getElementsByClassName('metacriticScore')
    if ($mcElements -and $mcElements.Length -gt 0) {
        $movie.MC = $mcElements[0].childNodes[0].innerHTML
    }

    Write-Host

    return $movie
}

function Get-RTInformation {

    param (
        [Parameter(Mandatory=$true)]
        [String]$title,
        [Parameter(Mandatory=$true)]
        [Int32]$year
    )

    Write-Host "Retrieving Rottentomatoes information..."

    # encode the title
    $encodedTitle = $title -replace " ", "+"
    $encodedTitle = $encodedTitle -replace "&", "%26"

    # search RT
    $url = "https://www.rottentomatoes.com/search/?search=$encodedTitle"
    Write-Host "rt: $url" -ForegroundColor "DarkGray"
    $response = Invoke-WebRequest $url

    # check for results
    $match = [regex]::Match($response, '(?i){.*"movies":(\[.*\]),"tvCount".*}')
    if (-not $match.Success) {
        Write-Host "rt: no results found`n" -ForegroundColor "DarkGray"
        return "NA"
    }
    else {
        $url = $null
        $results = New-Object "System.Collections.ArrayList"
        $json = ConvertFrom-Json $match.Groups[1]
        for ($i = 0; $i -lt $json.Length; $i++) {
            $item = @{}
            $item.Title = $json[$i].name
            $item.Year = [convert]::ToInt32($json[$i].year)
            $item.Url = "https://www.rottentomatoes.com$($json[$i].url)"
            $item.Distance = LDCompare $title $item.Title $true
            [void]$results.Add($item)
            Write-Host "rt: comparing $($item.Title) $($item.Year)" -ForegroundColor "DarkGray"
            if (($item.distance/$title.Length -lt 0.2) -and ($item.year -eq $year -or $item.year -eq ($year + 1))) {
                Write-Host "rt: matched" -ForegroundColor "DarkGray"
                $url = $item.Url
                break
            }
        }
        if (-not $url) {
            Write-Host "rt: not matched" -ForegroundColor "DarkGray"
            Write-Host "Results:"
            Write-Host "0: Skip"
            for ($i = 0; $i -lt $results.Count; $i++) {
                Write-Host -NoNewline "$($i + 1): $($results[$i].Title) ($($results[$i].Year))"
                Write-Host " [$($results[$i].Url)]" -ForegroundColor "DarkGray"
            }
            $selection = Read-Host "Please select a movie"
            if ($selection -eq 0) {
                Write-Host
                return "NA"
            }
            else {
                $url = $results[$selection - 1].Url
            }
        }

        # get movie details
        $response = Invoke-WebRequest $url
        
        # get critics rating
        $criticsElement = $response.ParsedHtml.getElementsByClassName("meter critic-score");
        if ($criticsElement -and $criticsElement.Length -gt 0) {
            $meterElement = $criticsElement[0].getElementsByClassName("meter-value");
            if ($meterElement -and $meterElement.Length -gt 0) {
                $criticsRating = $meterElement[0].childNodes[0].innerHTML
            }
        }
        # get the users rating
        $usersElement = $response.ParsedHtml.getElementsByClassName("meter media");
        if ($usersElement -and $usersElement.Length -gt 0) {
            $meterElement = $usersElement[0].getElementsByClassName("meter-value");
            if ($meterElement -and $meterElement.Length -gt 0) {
                $usersRating = $meterElement[0].childNodes[0].innerHTML -replace "%",""
            }
        }
        
        Write-Host

        return @($criticsRating, $usersRating)
    }
}

# https://www.codeproject.com/Tips/102192/Levenshtein-Distance-in-Windows-PowerShell
function LDCompare {

    param(
        [String] $first,
        [String] $second,
        [Switch] $ignoreCase
    )

    $len1 = $first.length
    $len2 = $second.length

    # if either string has length of zero, the # of edits/distance between them is simply the length of the other string
    if ($len1 -eq 0) { return $len2 }
    if ($len2 -eq 0) { return $len1 }

    # make everything lowercase if ignoreCase flag is set
    if($ignoreCase -eq $true) {
        $first = $first.ToLower()
        $second = $second.ToLower()
    }

    # create 2d Array to store the "distances"
    $dist = new-object -type 'int[,]' -arg ($len1+1),($len2+1)

    # initialize the first row and first column which represent the 2 strings we're comparing
    for ($i = 0; $i -le $len1; $i++) { $dist[$i,0] = $i }
    for ($j = 0; $j -le $len2; $j++) { $dist[0,$j] = $j }

    # compare
    $cost = 0
    for($i = 1; $i -le $len1;$i++) {
        for($j = 1; $j -le $len2;$j++) {
            if ($second[$j-1] -ceq $first[$i-1]) { $cost = 0 }
            else { $cost = 1 }
    
            # The value going into the cell is the min of 3 possibilities:
            # 1. The cell immediately above plus 1
            # 2. The cell immediately to the left plus 1
            # 3. The cell diagonally above and to the left plus the 'cost'
            $tempmin = [System.Math]::Min(([int]$dist[($i - 1), $j] + 1), ([int]$dist[$i, ($j - 1)] + 1))
            $dist[$i, $j] = [System.Math]::Min($tempmin, ([int]$dist[($i - 1), ($j - 1)] + $cost))
        }
    }

    # the distance is stored in the bottom right cell
    return $dist[$len1, $len2];
}
