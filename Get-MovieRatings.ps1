<#
.DESCRIPTION
Gets movie information and ratings from various websites and outputs in a tab delimited format.

.PARAMETER title
The movie title.

.EXAMPLE
Get-MovieRating -title "Memento"

.NOTES
#>
function Get-MovieRating {

    param (
        [String]$title
    )

    Write-Host "Getting primary information"
    $movie = Get-IMDbInformation $title


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
        $result = @{}
        $result.Data = "IMDb"
        $result.Title = $title
        $result.Rating = "No results found"
        return $result
    }

    # check results
    $elements = $response.ParsedHtml.getElementsByClassName('result_text')
    $results = New-Object "System.Collections.ArrayList"
    Write-Host "`nResults:"
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
        [Parameter(Mandatory=$false)]
        [String]$year
    )

    # encode the title
    $encodedTitle = $title -replace "-", " "
    $encodedTitle = $encodedTitle -replace " ", "+"
    $encodedTitle = $encodedTitle -replace "&amp;", "%26"

    # search RT
    $url = "http://www.rottentomatoes.com/search/?search=$encodedTitle"
    $response = Invoke-WebRequest $url -UserAgent "Mozilla/5.0 (iPhone; CPU iPhone OS 10_3 like Mac OS X) AppleWebKit/602.1.50 (KHTML, like Gecko) CriOS/56.0.2924.75 Mobile/14E5239e Safari/602.1"

    write-host $response


    if ([regex]::Match($html, '(?i)noresults')) {
        write-host "no results"
    }

    # check for no results
    $elements = $response.ParsedHtml.getElementsByClassName('noresults')
    write-host $elements.Length
    if ($elements.Length -gt 0) {
        $result = @{}
        $result.Data = "IMDb"
        $result.Title = $title
        $result.Rating = "No results found"
        return $result
    }
}
