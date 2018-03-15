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

    Write-Host "Getting primary information"
    $movie = Get-IMDbInformation $title
    $movie.RT = Get-RTInformation $movie.Title $movie.Year
    
    Write-Host "Result:"
    $movie

    Set-Clipboard -Value "$($movie.Title)`t$($movie.Year)`t$($movie.Genre)`t$($movie.IMDb)`t$($movie.RT)"
    Write-Host "`ncopied to clipboard"
}

function Get-IMDbInformation {

    param (
        [Parameter(Mandatory=$true)]
        [String]$title
    )

    # encode the title
    $encodedTitle = $title -replace "-", " "
    $encodedTitle = $encodedTitle -replace " ", "+"
    $encodedTitle = $encodedTitle -replace "&amp;", "%26"

    # search IMDb
    $url = "http://www.imdb.com/find?q=$encodedTitle&s=all"
    $response = Invoke-WebRequest $url
    
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
            $result.Url = $match.Groups[1].value
            $result.Title = $match.Groups[2].value
            $result.Year = $match.Groups[3].value
            [void]$results.Add($result)
            Write-Host "$($i + 1): $($result.Title), $($result.Year)"
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
    $response = Invoke-WebRequest "http://www.imdb.com$($result.Url)"
    $html = [System.Net.WebUtility]::HtmlDecode($response.ParsedHtml.getElementsByClassName('title_bar_wrapper')[0].innerHTML)

    # get the title
    $match = [regex]::Match($html, '(?i)itemprop="name">([^<]+)')
    if ($match.Success) {
        $title = $match.Groups[1].value.Trim()
    }
    # get the year
    $match = [regex]::Match($html, '(?i)id=titleYear>\(<a href="\/year\/([0-9]+)\/')
    if ($match.Success) {
        $year = $match.Groups[1].value.Trim()
    }
    # get the genre
    $matches = [regex]::Matches($html, '(?i)itemprop="genre">([^<]+)')
    if ($matches.Success) {
        for ($i = 0; $i -lt $matches.Groups.Length; $i+=2) {
            if ($i -eq 0) { $genre = $matches.Groups[$i+1].value.Trim() }
            else { $genre += ", $($matches.Groups[$i+1].value.Trim())" }
        }
    }
    # get the rating
    $match = [regex]::Match($html, '(?i)itemprop="ratingValue">([^<]+)')
    if ($match.Success) {
        $rating = $match.Groups[1].value.Trim()
    }

    $movie = @{}
    $movie.Title = $title
    $movie.Year = $year
    $movie.Genre = $genre
    $movie.IMDb = $rating
    return $movie
}

function Get-RTInformation {

    param (
        [Parameter(Mandatory=$true)]
        [String]$title,
        [Parameter(Mandatory=$true)]
        [String]$year
    )

    # encode the title
    $encodedTitle = $title -replace "-", " "
    $encodedTitle = $encodedTitle -replace " ", "+"
    $encodedTitle = $encodedTitle -replace "&amp;", "%26"

    # search RT
    $url = "https://www.rottentomatoes.com/search/?search=$encodedTitle"
    Write-Host "rt: searching $url" -ForegroundColor "DarkGray"
    $response = Invoke-WebRequest $url

    # check for results
    $match = [regex]::Match($response, '(?i){.*"movies":(\[.*\]),"tvCount".*}')
    if (-not $match.Success) {
        Write-Host "rt: no results found" -ForegroundColor "DarkGray"
        return "No results found"
    }
    else {
        $url = $null
        $json = ConvertFrom-Json $match.Groups[1]
        for ($i = 0; $i -lt $json.Length; $i++) {
            $item = $json[$i]
            $y = [convert]::ToInt32($item.year)
            Write-Host "rt: checking against $($item.name) $($item.year)" -ForegroundColor "DarkGray"
            if (($item.name.Contains($title) -or $title.Contains($item.name)) -and ($y -eq $year -or $y + 1 -eq $year)) {
                Write-Host "rt: matched $($item.name) $($item.year)" -ForegroundColor "DarkGray"
                $url = "https://www.rottentomatoes.com$($json[$i].url)"
                break
            }
        }
        if (-not $url) {
            return "No results found"
        }

        # get movie details
        $response = Invoke-WebRequest $url
        
        # get critics rating
        $criticsElement = $response.ParsedHtml.getElementsByClassName("meter critic-score");
        if ($criticsElement) {
            $meterElement = $criticsElement[0].getElementsByClassName("meter-value");
            if ($meterElement) {
                $criticsRating = $meterElement[0].childNodes[0].innerHTML
            }
        }
        # get the users rating
        $usersElement = $response.ParsedHtml.getElementsByClassName("meter media");
        if ($usersElement) {
            $meterElement = $usersElement[0].getElementsByClassName("meter-value");
            if ($meterElement) {
                $usersRating = $meterElement[0].childNodes[0].innerHTML -replace "%",""
            }
        }

        $result = ""
        if ($criticsRating) { $result += $criticsRating }
        else { $result += "?" }
        if ($criticsRating) { $result += "/$usersRating" }
        else { $result += "/?" }
        return $result
    }
}
