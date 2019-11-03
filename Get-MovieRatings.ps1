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

    while ($true) {
        $movie = @{}

        # get imdb/mc information, we use this as the primary information because it is usualy more correct
        $result = Get-DuckDuckGoLink "imdb" $title
        if ($result) {
            $info = Get-IMDbInformation $result.Url
            $movie.Title = $info.Title
            $movie.Year = $info.Year
            $movie.Genre = $info.Genre
            $movie.IMDb = $info.IMDb
            $movie.MC = $info.MC
        }

        $result = Get-DuckDuckGoLink "rotten tomatoes" $title
        if ($result) {
            $info = Get-RottenTomatoesInformation $result.Url
            $movie.RT = $info
        }
        
        Write-Host
        Write-Host "Movie"
        Write-Host "---------------"
        Write-Host "Title: $($movie.Title)"
        Write-Host "Year:  $($movie.Year)"
        Write-Host "Genre: $($movie.Genre)"
        Write-Host "IMDb:  $($movie.IMDb)"
        Write-Host "MC:    $($movie.MC)"
        Write-Host "RT:    $($movie.RT)"

        Set-Clipboard -Value "$($movie.Title)`t$($movie.Year)`t$($movie.Genre)`t$($movie.IMDb)`t$($movie.MC)`t$($movie.RT[0])`t$($movie.RT[1])"
        Write-Host "Copied to clipboard." -ForegroundColor "DarkGray"

        $title = Read-Host "Title (or exit)"
        if ($title -eq "exit") { return }
    }
}

function Get-IMDbInformation {

    param (
        [Parameter(Mandatory=$true)]
        [String]$url
    )

    $info = @{}

    # get movie details
    $html = Get-Html $url

    # get the rating
    $match = [regex]::Match($html, '(?i)itemprop="ratingValue">([^<]+)')
    if ($match.Success) {
        $info.IMDb = [convert]::ToDecimal($match.Groups[1].value.Trim()) * 10
    }
    else {
        Write-Error "Could not match imdb rating"
    }

    # get the metacritic rating
    $match = [regex]::Match($html, '(?i)<div class="metacriticScore[^"]*">\s*<span>([0-9]+)<\/span>')
    if ($match.Success) {
        $info.MC = $match.Groups[1].value
    }
    else {
        Write-Error "Could not match metacritic rating"
    }

    # capture the title_wrapper element, this prevents additional genres matching below the desired content
    $match = [regex]::Match($html, '(?i)(<div class="title_wrapper">[\s\S]*)<\/div>\s*<\/div>\s*<\/div>\s*<\/div>\s*<\/div>\s*<div class="slate_wrapper">')
    if ($match.Success) {
        $html = $match.Groups[1].value
    }
    else {
        Write-Error "imdb: failed to parse"
        return
    }
    
    # get the title, (?i) is case insensitive mode
    $match = [regex]::Match($html, '(?i)class="title_wrapper">\s*<[^>]*>([^<]+)')
    if ($match.Success) {
        $info.Title = [System.Net.WebUtility]::HtmlDecode($match.Groups[1].value).Trim()
    }
    else {
        Write-Error "Could not match title"
    }

    # get the year
    $match = [regex]::Match($html, '(?i)id="titleYear">\(<a href="\/year\/([0-9]+)\/')
    if ($match.Success) {
        $info.Year = [System.Net.WebUtility]::HtmlDecode($match.Groups[1].value).Trim()
    }
    else {
        Write-Error "Could not match year"
    }

    # get the genre
    $matches = [regex]::Matches($html, '(?i)<a href="\/search\/title\?genres=[^>]+>([^<]+)')
    if ($matches.Success) {
        for ($i = 0; $i -lt $matches.Groups.Length; $i+=2) {
            if ($i -eq 0) { $genre = $matches.Groups[$i+1].value.Trim() }
            else { $genre += ", $($matches.Groups[$i+1].value.Trim())" }
        }
    }
    else {
        Write-Error "Could not match genres"
    }
    $info.Genre = $genre
    
    return $info
}

function Get-RottenTomatoesInformation {

    param (
        [Parameter(Mandatory=$true)]
        [String]$url
    )

    $html = Get-Html $url

    # get ratings
    #<a href="#contentReviews" class="unstyled articleLink mop-ratings-wrap__icon-link" id="tomato_meter_link">
    #            <span class="mop-ratings-wrap__icon meter-tomato icon big medium-xs certified_fresh"></span>
    #            <span class="mop-ratings-wrap__percentage">
    #                83%
    #            </span>
    $regex = '(?i)' # case insensitive
    $regex += '<a href="#([^"]+)"[^>]*>\s*' # critic/user identifier
    $regex += '<span[^>]*><\/span>\s*' # unused span
    $regex += '<span class="mop-ratings-wrap__percentage">\s*([0-9]+)%\s*<\/span>' # rating
    $regex = [regex]$regex
    $match = $regex.Match($html)
    while ($match.Success) {
        if ($match.Groups[1].value -eq "contentReviews") {
            $criticsRating = $match.Groups[2].value.Trim()
        }
        if ($match.Groups[1].value -eq "audience_reviews") {
            $usersRating = $match.Groups[2].value.Trim()
        }
        if ($criticsRating -and $usersRating) {
            break
        }
        $match = $match.NextMatch()
    }

    return @($criticsRating, $usersRating)
}

function Get-DuckDuckGoLink {
    
    param (
        [Parameter(Mandatory=$true)]
        [String]$website,
        [Parameter(Mandatory=$true)]
        [String]$title
    )

    # encode the query
    $query = "$website $title"
    $query = $query -replace " ", "+"
    $query = $query -replace "&", "%26"

    # send the request
    $html = Get-Html "https://duckduckgo.com/html?q=$query"

    # check results
    $regex = '(?i)' # case insensitive
    $regex += '<a rel="nofollow" class="result__a" href="([^"]+)">' # link to website
    $regex += '<b>([^<]+)</b>' # title
    $regex += '\s*\(([0-9]+)\)' # year
    $regex = [regex]$regex
    $match = $regex.Match($html)
    while ($match.Success -and $results.Length -lt 5) {
        $result = @{}
        $result.Url = $match.Groups[1].value
        $result.Title = $match.Groups[2].value.Trim()
        $result.Year = $match.Groups[3].value.Trim()
        $result.Distance = LDCompare $title $result.Title $true
        if ($result.Distance/$title.Length -lt 0.2) {
            return $result
        }
        $match = $match.NextMatch()
    }
}

function Get-Html {

    param (
        [Parameter(Mandatory=$true)]
        [String]$url
    )

    # make the web request, do not use normal parsing, some webpages never finish
    Write-Host "request: $url" -ForegroundColor "DarkGray"
    $response = Invoke-WebRequest $url -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:70.0) Gecko/20100101 Firefox/70.0"
    # debugging
    #  $response.Content | Out-File ./html.txt
    return $response.Content
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
