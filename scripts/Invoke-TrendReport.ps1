[CmdletBinding()]
param(
    [string]$ConfigPath = "..\config\sources.json",
    [string]$OutputDir,
    [datetime]$ReportDate = (Get-Date),
    [string]$EmailTo,
    [string]$EmailSubject = "Daily Social Trend Report",
    [ValidateSet("Outlook","Smtp")]
    [string]$EmailMethod = "Outlook",
    [string]$SmtpHost = "smtp.gmail.com",
    [int]$SmtpPort = 587,
    [string]$SmtpUser,
    [string]$SmtpPassword,
    [string]$SmtpCredentialPath,
    [string]$XBearerToken,
    [switch]$SendEmail,
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = `
    [Net.SecurityProtocolType]::Tls12 -bor `
    [Net.SecurityProtocolType]::Tls13 -bor `
    [Net.SecurityProtocolType]::SystemDefault

function Get-AbsolutePath {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$BasePath
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

function Get-DateWindow {
    param(
        [Parameter(Mandatory)]
        [datetime]$ReferenceDate
    )

    $today = $ReferenceDate.Date
    $yesterdayStart = $today.AddDays(-1)
    $yesterdayEnd = $today

    [pscustomobject]@{
        Start = $yesterdayStart
        End = $yesterdayEnd
        Label = $yesterdayStart.ToString("yyyy-MM-dd")
    }
}

function Get-PlainTextSecret {
    param(
        [Parameter(Mandatory)]
        [string]$SecretPath
    )

    if (-not (Test-Path -LiteralPath $SecretPath)) {
        throw "Secret file not found at $SecretPath"
    }

    $protectedValue = (Get-Content -LiteralPath $SecretPath -Raw).Trim()
    $secureValue = $protectedValue | ConvertTo-SecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureValue)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Get-EffectiveSetting {
    param(
        [AllowNull()]
        [string]$ExplicitValue,
        [AllowNull()]
        [string]$EnvironmentVariableName
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitValue)) {
        return $ExplicitValue
    }

    if (-not [string]::IsNullOrWhiteSpace($EnvironmentVariableName)) {
        return [Environment]::GetEnvironmentVariable($EnvironmentVariableName)
    }

    return $null
}

function Get-HttpContent {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $handler = [System.Net.Http.HttpClientHandler]::new()
    $client = [System.Net.Http.HttpClient]::new($handler)
    $client.Timeout = [TimeSpan]::FromSeconds(30)
    $client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
    $client.DefaultRequestHeaders.Accept.ParseAdd("text/html,application/xhtml+xml,application/json")

    try {
        $response = $client.GetAsync($Uri).GetAwaiter().GetResult()
        $response.EnsureSuccessStatusCode() | Out-Null
        return $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    }
    finally {
        $client.Dispose()
        $handler.Dispose()
    }
}

function Get-StopWords {
    [System.Collections.Generic.HashSet[string]]::new([string[]]@(
        "a","about","after","all","am","an","and","any","are","as","at","be","been","before","being",
        "but","by","can","could","did","do","does","for","from","get","getting","got","had","has","have",
        "he","her","here","hers","him","his","how","i","if","in","into","is","it","its","just","like",
        "me","more","most","my","new","no","not","now","of","on","one","or","our","out","over","people",
        "so","some","still","than","that","the","their","them","then","there","they","this","to","too",
        "up","us","was","we","were","what","when","which","who","why","will","with","you","your","yours",
        "viral","video","meme","memes","shorts","tiktok","youtube","reddit"
    ), [System.StringComparer]::OrdinalIgnoreCase)
}

function Get-KeywordMap {
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Items
    )

    $stopWords = Get-StopWords
    $weights = @{}

    foreach ($item in $Items) {
        $text = "$($item.Title) $($item.Domain)".ToLowerInvariant()
        foreach ($token in ([regex]::Matches($text, "[a-z0-9][a-z0-9'\-]{2,}"))) {
            $word = $token.Value.Trim("'")
            if ($stopWords.Contains($word)) {
                continue
            }

            if (-not $weights.ContainsKey($word)) {
                $weights[$word] = 0.0
            }

            $weights[$word] += [double]$item.Score
        }
    }

    return $weights
}

function Get-TopicLabel {
    param(
        [Parameter(Mandatory)]
        [string[]]$Keywords
    )

    $label = ($Keywords | Select-Object -First 3) -join " / "
    if ([string]::IsNullOrWhiteSpace($label)) {
        return "General social buzz"
    }

    return (Get-Culture).TextInfo.ToTitleCase($label)
}

function New-ItemRecord {
    param(
        [Parameter(Mandatory)]
        [string]$Platform,
        [Parameter(Mandatory)]
        [string]$Source,
        [Parameter(Mandatory)]
        [string]$Title,
        [Parameter(Mandatory)]
        [string]$Url,
        [datetime]$PublishedAt,
        [double]$Score,
        [string]$ExternalUrl,
        [string]$Domain,
        [string]$Author,
        [string]$Summary
    )

    $effectiveDomain = if ($Domain) { $Domain } else { ([uri]$Url).Host }
    $allLinks = @($Url)
    if ($ExternalUrl -and $ExternalUrl -ne $Url) {
        $allLinks += $ExternalUrl
    }

    [pscustomobject]@{
        Platform = $Platform
        Source = $Source
        Title = $Title
        Url = $Url
        ExternalUrl = $ExternalUrl
        PublishedAt = $PublishedAt
        Score = [math]::Round($Score, 2)
        Domain = $effectiveDomain
        Author = $Author
        Summary = $Summary
        SearchText = (($Title, $Summary, $effectiveDomain) -join " ").ToLowerInvariant()
        AllLinks = $allLinks
    }
}

function Get-RedditItems {
    param(
        [AllowEmptyCollection()]
        [object[]]$Sources,
        [Parameter(Mandatory)]
        [datetime]$Start,
        [Parameter(Mandatory)]
        [datetime]$End
    )

    $headers = @{
        "User-Agent" = "TrendScraperBot/1.0 (+https://example.local)"
    }

    $items = New-Object System.Collections.Generic.List[object]

    foreach ($source in $Sources) {
        $uri = "https://www.reddit.com/r/$($source.subreddit)/top.json?t=day&limit=$($source.limit)"
        try {
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -TimeoutSec 30
        }
        catch {
            Write-Warning "Reddit fetch failed for r/$($source.subreddit): $($_.Exception.Message)"
            continue
        }

        foreach ($child in $response.data.children) {
            $post = $child.data
            $published = [DateTimeOffset]::FromUnixTimeSeconds([int64]$post.created_utc).LocalDateTime
            if ($published -lt $Start -or $published -ge $End) {
                continue
            }

            $score = 1 + [math]::Max(0, [double]$post.score) + ([math]::Max(0, [double]$post.num_comments) * 0.35)
            $external = if ($post.PSObject.Properties["url_overridden_by_dest"] -and $post.url_overridden_by_dest) { $post.url_overridden_by_dest } else { $post.url }
            $items.Add((New-ItemRecord `
                -Platform "Reddit" `
                -Source "r/$($source.subreddit)" `
                -Title $post.title `
                -Url ("https://www.reddit.com" + $post.permalink) `
                -PublishedAt $published `
                -Score $score `
                -ExternalUrl $external `
                -Domain $(if ($post.PSObject.Properties["domain"]) { $post.domain } else { "reddit.com" }) `
                -Author $(if ($post.PSObject.Properties["author"]) { $post.author } else { $null }) `
                -Summary $(if ($post.PSObject.Properties["selftext"]) { $post.selftext } else { "" })))
        }
    }

    return $items.ToArray()
}

function Get-YouTubeItems {
    param(
        [AllowEmptyCollection()]
        [object[]]$Sources,
        [Parameter(Mandatory)]
        [datetime]$Start,
        [Parameter(Mandatory)]
        [datetime]$End
    )

    $items = New-Object System.Collections.Generic.List[object]

    foreach ($source in $Sources) {
        if ($source.feedUrl) {
            $uri = [string]$source.feedUrl
        }
        elseif ($source.channelId) {
            $uri = "https://www.youtube.com/feeds/videos.xml?channel_id=$($source.channelId)"
        }
        else {
            Write-Warning "Skipping YouTube source '$($source.name)' because it has no feedUrl or channelId."
            continue
        }

        try {
            $response = Invoke-WebRequest -Uri $uri -Method Get -TimeoutSec 30
            [xml]$feed = $response.Content
        }
        catch {
            Write-Warning "YouTube feed failed for $($source.name): $($_.Exception.Message)"
            continue
        }

        $entries = @($feed.feed.entry)
        if ($entries.Count -eq 0 -and $feed.rss.channel.item) {
            $entries = @($feed.rss.channel.item)
        }

        foreach ($entry in $entries) {
            $publishedText = if ($entry.published) { $entry.published } else { $entry.pubDate }
            if (-not $publishedText) {
                continue
            }

            $published = [datetime]$publishedText
            $localPublished = $published.ToLocalTime()
            if ($localPublished -lt $Start -or $localPublished -ge $End) {
                continue
            }

            $author = if ($entry.author.name) { $entry.author.name } else { $entry.author }
            $videoUrl = ($entry.link | Where-Object { $_.rel -eq "alternate" } | Select-Object -First 1).href
            if (-not $videoUrl) {
                $videoUrl = if ($entry.id) { $entry.id } else { $entry.link }
            }

            $score = 50
            if ($source.weight) {
                $score += [double]$source.weight
            }

            $items.Add((New-ItemRecord `
                -Platform "YouTube" `
                -Source $source.name `
                -Title $entry.title `
                -Url $videoUrl `
                -PublishedAt $localPublished `
                -Score $score `
                -ExternalUrl $videoUrl `
                -Domain "youtube.com" `
                -Author $author `
                -Summary ""))
        }
    }

    return $items.ToArray()
}

function Get-XItems {
    param(
        [Parameter(Mandatory)]
        [object]$Config,
        [Parameter(Mandatory)]
        [datetime]$Start,
        [Parameter(Mandatory)]
        [datetime]$End,
        [Parameter(Mandatory)]
        [string]$BasePath
    )

    if (-not $Config.enabled) {
        return @()
    }

    $bearerToken = Get-EffectiveSetting -ExplicitValue $XBearerToken -EnvironmentVariableName "X_BEARER_TOKEN"
    if ([string]::IsNullOrWhiteSpace($bearerToken)) {
        $tokenPath = Get-AbsolutePath -Path $Config.bearerTokenPath -BasePath $BasePath
        if (-not (Test-Path -LiteralPath $tokenPath)) {
            Write-Warning "X is enabled but no bearer token was provided. Skipping X collection."
            return @()
        }
        $bearerToken = Get-PlainTextSecret -SecretPath $tokenPath
    }

    $headers = @{
        Authorization = "Bearer $bearerToken"
    }

    $items = New-Object System.Collections.Generic.List[object]

    try {
        $trendUri = "https://api.x.com/2/trends/by/woeid/$($Config.woeid)?max_trends=$($Config.maxTrends)&trend.fields=trend_name,tweet_count"
        $trendResponse = Invoke-RestMethod -Uri $trendUri -Headers $headers -Method Get -TimeoutSec 30
    }
    catch {
        Write-Warning "X trend fetch failed: $($_.Exception.Message)"
        return @()
    }

    $trends = @($trendResponse.data | Select-Object -First ([int]$Config.maxTrends))
    foreach ($trend in $trends) {
        $trendName = [string]$trend.trend_name
        if ([string]::IsNullOrWhiteSpace($trendName)) {
            continue
        }

        $query = $trendName
        if ($Config.querySuffix) {
            $query = "$trendName $($Config.querySuffix)"
        }

        $searchQuery = [uri]::EscapeDataString($query)
        $maxResults = [math]::Min(100, [math]::Max(10, [int]$Config.postsPerTrend))
        $searchUri = "https://api.x.com/2/tweets/search/recent?query=$searchQuery&max_results=$maxResults&tweet.fields=created_at,public_metrics,author_id,lang&expansions=author_id&user.fields=username,name"

        try {
            $searchResponse = Invoke-RestMethod -Uri $searchUri -Headers $headers -Method Get -TimeoutSec 30
        }
        catch {
            Write-Warning "X search failed for trend '$trendName': $($_.Exception.Message)"
            continue
        }

        $usersById = @{}
        foreach ($user in @($searchResponse.includes.users)) {
            $usersById[$user.id] = $user
        }

        foreach ($tweet in @($searchResponse.data)) {
            $published = ([datetime]$tweet.created_at).ToLocalTime()
            if ($published -lt $Start -or $published -ge $End) {
                continue
            }

            $metrics = $tweet.public_metrics
            $score = 1 +
                ([double]$metrics.like_count) +
                ([double]$metrics.retweet_count * 2.0) +
                ([double]$metrics.reply_count * 1.5) +
                ([double]$metrics.quote_count * 2.0)

            $author = $null
            if ($usersById.ContainsKey($tweet.author_id)) {
                $author = "@$($usersById[$tweet.author_id].username)"
            }

            $tweetUrl = if ($author) {
                "https://x.com/$($author.TrimStart('@'))/status/$($tweet.id)"
            }
            else {
                "https://x.com/i/web/status/$($tweet.id)"
            }

            $items.Add((New-ItemRecord `
                -Platform "X" `
                -Source "Trend: $trendName" `
                -Title (($tweet.text -replace '\s+', ' ').Trim()) `
                -Url $tweetUrl `
                -PublishedAt $published `
                -Score $score `
                -ExternalUrl $tweetUrl `
                -Domain "x.com" `
                -Author $author `
                -Summary "Trend topic: $trendName"))
        }
    }

    return $items.ToArray()
}

function Get-TikTokVideoLinksFromHtml {
    param(
        [Parameter(Mandatory)]
        [string]$Html
    )

    $matches = [regex]::Matches($Html, 'https://www\.tiktok\.com/@[^"''\s<>]+/video/\d+')
    $links = New-Object System.Collections.Generic.HashSet[string]
    foreach ($match in $matches) {
        $links.Add($match.Value) | Out-Null
    }

    return @($links)
}

function Get-TikTokVideoMetadata {
    param(
        [Parameter(Mandatory)]
        [string]$VideoUrl
    )

    try {
        $html = Get-HttpContent -Uri $VideoUrl
    }
    catch {
        Write-Warning "TikTok video fetch failed for $VideoUrl : $($_.Exception.Message)"
        return $null
    }
    $jsonLdMatches = [regex]::Matches($html, '<script type="application/ld\+json">\s*(\{.*?\})\s*</script>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    foreach ($match in $jsonLdMatches) {
        try {
            $payload = $match.Groups[1].Value | ConvertFrom-Json
        }
        catch {
            continue
        }

        if ($payload.'@type' -eq 'VideoObject') {
            return [pscustomobject]@{
                Title = if ($payload.name) { [string]$payload.name } else { "TikTok video" }
                Author = if ($payload.author.name) { [string]$payload.author.name } else { $null }
                PublishedAt = if ($payload.uploadDate) { ([datetime]$payload.uploadDate).ToLocalTime() } else { $null }
                Description = [string]$payload.description
                Url = if ($payload.url) { [string]$payload.url } else { $VideoUrl }
                Thumbnail = [string]$payload.thumbnailUrl
            }
        }
    }

    $titleMatch = [regex]::Match($html, '<meta property="og:title" content="([^"]+)"')
    if ($titleMatch.Success) {
        return [pscustomobject]@{
            Title = $titleMatch.Groups[1].Value
            Author = $null
            PublishedAt = $null
            Description = $titleMatch.Groups[1].Value
            Url = $VideoUrl
            Thumbnail = $null
        }
    }

    return $null
}

function Get-TikTokItems {
    param(
        [Parameter(Mandatory)]
        [object]$Config,
        [Parameter(Mandatory)]
        [datetime]$Start,
        [Parameter(Mandatory)]
        [datetime]$End
    )

    if (-not $Config.enabled) {
        return @()
    }

    $items = New-Object System.Collections.Generic.List[object]
    $allVideoLinks = New-Object System.Collections.Generic.List[string]

    foreach ($seedUrl in @($Config.seedUrls)) {
        try {
            $seedContent = Get-HttpContent -Uri $seedUrl
        }
        catch {
            Write-Warning "TikTok seed fetch failed for $seedUrl : $($_.Exception.Message)"
            continue
        }

        foreach ($link in (Get-TikTokVideoLinksFromHtml -Html $seedContent)) {
            if ($allVideoLinks.Count -ge [int]$Config.maxVideos) {
                break
            }
            if ($allVideoLinks -notcontains $link) {
                $allVideoLinks.Add($link)
            }
        }
    }

    foreach ($videoUrl in $allVideoLinks) {
        $metadata = Get-TikTokVideoMetadata -VideoUrl $videoUrl
        if (-not $metadata) {
            continue
        }

        if ($metadata.PublishedAt -and ($metadata.PublishedAt -lt $Start -or $metadata.PublishedAt -ge $End)) {
            continue
        }

        $score = 40
        if ($metadata.PublishedAt) {
            $hoursOld = [math]::Max(1, ($End - $metadata.PublishedAt).TotalHours)
            $score += [math]::Round(24 / $hoursOld, 2)
        }

        $items.Add((New-ItemRecord `
            -Platform "TikTok" `
            -Source "Public scrape" `
            -Title $metadata.Title `
            -Url $videoUrl `
            -PublishedAt $(if ($metadata.PublishedAt) { $metadata.PublishedAt } else { $Start }) `
            -Score $score `
            -ExternalUrl $videoUrl `
            -Domain "tiktok.com" `
            -Author $metadata.Author `
            -Summary $metadata.Description))
    }

    return $items.ToArray()
}

function Get-Topics {
    param(
        [AllowEmptyCollection()]
        [object[]]$Items,
        [int]$TopCount = 5,
        [int]$ExamplesPerTopic = 3
    )

    $normalizedItems = @(
        $Items | Where-Object {
            $_ -and
            $_.PSObject.Properties["Title"] -and
            $_.PSObject.Properties["Score"] -and
            $_.PSObject.Properties["SearchText"] -and
            $_.PSObject.Properties["Url"]
        }
    )

    if (-not $normalizedItems -or $normalizedItems.Count -eq 0) {
        return @()
    }

    $keywordMap = Get-KeywordMap -Items $normalizedItems
    $keywords = $keywordMap.GetEnumerator() |
        Sort-Object Value -Descending |
        Select-Object -ExpandProperty Key -First ($TopCount * 6)

    $topics = New-Object System.Collections.Generic.List[object]
    $usedUrls = New-Object System.Collections.Generic.HashSet[string]
    $usedKeywords = New-Object System.Collections.Generic.HashSet[string]

    foreach ($keyword in $keywords) {
        if ($topics.Count -ge $TopCount) {
            break
        }

        if ($usedKeywords.Contains($keyword)) {
            continue
        }

        $matches = @(
            $normalizedItems |
                Where-Object {
                    $_.PSObject.Properties["SearchText"] -and
                    $_.SearchText -match "(^|[^a-z0-9])$([regex]::Escape($keyword))([^a-z0-9]|$)"
                } |
                Sort-Object Score -Descending
        )
        if ($matches.Count -lt 2) {
            continue
        }

        $topicKeywords = New-Object System.Collections.Generic.List[string]
        $topicKeywords.Add($keyword)
        foreach ($candidate in $keywords) {
            if ($candidate -eq $keyword -or $topicKeywords.Count -ge 3) {
                continue
            }

            $candidateMatches = @(
                $matches | Where-Object {
                    $_.PSObject.Properties["SearchText"] -and
                    $_.SearchText -match "(^|[^a-z0-9])$([regex]::Escape($candidate))([^a-z0-9]|$)"
                }
            )
            if ($candidateMatches.Count -ge 2) {
                $topicKeywords.Add($candidate)
            }
        }

        $examples = @()
        foreach ($match in $matches) {
            if ($examples.Count -ge $ExamplesPerTopic) {
                break
            }
            if (-not $match.PSObject.Properties["Url"]) {
                continue
            }
            if ($usedUrls.Contains($match.Url)) {
                continue
            }
            $usedUrls.Add($match.Url) | Out-Null
            $examples += $match
        }

        if ($examples.Count -eq 0) {
            continue
        }

        foreach ($topicKeyword in $topicKeywords) {
            $usedKeywords.Add($topicKeyword) | Out-Null
        }

        $score = ($examples | Measure-Object -Property Score -Sum).Sum
        $platforms = ($examples | Group-Object Platform | Sort-Object Count -Descending | ForEach-Object { $_.Name }) -join ", "
        $lead = $examples[0]
        $summary = "This cluster showed up across $platforms, led by '$($lead.Title)'."

        $topics.Add([pscustomobject]@{
            Label = Get-TopicLabel -Keywords $topicKeywords.ToArray()
            Keywords = $topicKeywords.ToArray()
            Score = [math]::Round($score, 2)
            Summary = $summary
            Examples = $examples
        })
    }

    if ($topics.Count -eq 0) {
        $fallbackExamples = $normalizedItems | Sort-Object Score -Descending | Select-Object -First $ExamplesPerTopic
        return @([pscustomobject]@{
            Label = "Top social posts"
            Keywords = @()
            Score = ($fallbackExamples | Measure-Object -Property Score -Sum).Sum
            Summary = "No clean repeated keywords emerged, so this section lists the strongest posts from yesterday."
            Examples = $fallbackExamples
        })
    }

    return $topics | Sort-Object Score -Descending
}

function Get-ExampleLine {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    $exampleUrl = if ($Item.ExternalUrl) { $Item.ExternalUrl } else { $Item.Url }
    $parts = @(
        "- [$($Item.Title)]($exampleUrl)"
        "($($Item.Platform) via $($Item.Source)"
    )

    if ($Item.Author) {
        $parts[-1] += ", by $($Item.Author)"
    }

    $parts[-1] += ", score $($Item.Score))"

    if ($Item.Url -and $Item.Url -ne $exampleUrl) {
        $parts += "[discussion]($($Item.Url))"
    }

    return ($parts -join " ")
}

function ConvertTo-HtmlEncoded {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Get-PlatformSummaryLines {
    param(
        [AllowEmptyCollection()]
        [object[]]$Items
    )

    return @(
        $Items |
            Group-Object Platform |
            Sort-Object Count -Descending |
            ForEach-Object { "- $($_.Name) items: $($_.Count)" }
    )
}

function New-HtmlReport {
    param(
        [Parameter(Mandatory)]
        [string]$DateLabel,
        [AllowEmptyCollection()]
        [object[]]$Items,
        [AllowEmptyCollection()]
        [object[]]$Topics
    )

    $safeTopics = @($Topics)
    $safeItems = @($Items | Where-Object { $_.PSObject.Properties["Url"] -and $_.PSObject.Properties["Title"] })
    $platformCounts = @(
        $safeItems |
            Group-Object Platform |
            Sort-Object Count -Descending |
            ForEach-Object {
                "<div style=""padding:12px 14px; background:#eff6ff; border-radius:12px; font:600 14px/1.4 Segoe UI, Arial, sans-serif;"">$($_.Name): $($_.Count)</div>"
            }
    )

    $topicSections = New-Object System.Collections.Generic.List[string]
    foreach ($topic in $safeTopics) {
        $exampleItems = @($topic.Examples)
        $exampleHtml = ($exampleItems | ForEach-Object {
            $link = if ($_.ExternalUrl) { $_.ExternalUrl } else { $_.Url }
            $discussion = if ($_.Url -and $_.Url -ne $link) {
                " <a href=""$($_.Url)"">Discussion</a>"
            }
            else {
                ""
            }

            "<li><a href=""$link"">$(ConvertTo-HtmlEncoded $_.Title)</a> <span style=""color:#666;"">($($_.Platform) via $(ConvertTo-HtmlEncoded $_.Source), score $($_.Score))</span>$discussion</li>"
        }) -join ""

        $topicSections.Add(@"
<section style="margin:0 0 24px;">
  <h2 style="font:600 20px/1.3 Segoe UI, Arial, sans-serif; margin:0 0 8px; color:#1f2937;">$(ConvertTo-HtmlEncoded $topic.Label)</h2>
  <p style="font:15px/1.6 Segoe UI, Arial, sans-serif; margin:0 0 10px; color:#374151;">$(ConvertTo-HtmlEncoded $topic.Summary)</p>
  <ul style="margin:0; padding-left:20px; font:15px/1.7 Segoe UI, Arial, sans-serif; color:#111827;">
    $exampleHtml
  </ul>
</section>
"@) | Out-Null
    }

    if ($topicSections.Count -eq 0) {
        $topicSections.Add(@"
<section style="margin:0 0 24px;">
  <h2 style="font:600 20px/1.3 Segoe UI, Arial, sans-serif; margin:0 0 8px; color:#1f2937;">No Items Collected</h2>
  <p style="font:15px/1.6 Segoe UI, Arial, sans-serif; margin:0; color:#374151;">No items were collected for this window. Check network access or adjust the configured sources.</p>
</section>
"@) | Out-Null
    }

    $rawItems = ($safeItems | Sort-Object PublishedAt, Score -Descending | Select-Object -First 12 | ForEach-Object {
        $link = if ($_.ExternalUrl) { $_.ExternalUrl } else { $_.Url }
        "<li><a href=""$link"">$(ConvertTo-HtmlEncoded $_.Title)</a> <span style=""color:#666;"">| $($_.Platform) | $(ConvertTo-HtmlEncoded $_.Source) | $($_.PublishedAt.ToString("yyyy-MM-dd HH:mm"))</span></li>"
    }) -join ""

    return @"
<!DOCTYPE html>
<html>
<body style="margin:0; padding:24px; background:#f3f4f6; color:#111827;">
  <div style="max-width:860px; margin:0 auto; background:#ffffff; border:1px solid #e5e7eb; border-radius:16px; overflow:hidden;">
    <div style="padding:28px 32px; background:linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%); color:#ffffff;">
      <div style="font:700 30px/1.1 Segoe UI, Arial, sans-serif; margin:0 0 8px;">Daily Social Trend Report</div>
      <div style="font:15px/1.6 Segoe UI, Arial, sans-serif; opacity:0.9;">Summary for $DateLabel with direct links to yesterday's strongest meme and video examples.</div>
    </div>
    <div style="padding:28px 32px;">
      <div style="display:flex; gap:12px; flex-wrap:wrap; margin:0 0 24px;">
        <div style="padding:12px 14px; background:#eff6ff; border-radius:12px; font:600 14px/1.4 Segoe UI, Arial, sans-serif;">Items collected: $($safeItems.Count)</div>
        $($platformCounts -join [Environment]::NewLine)
      </div>
      $($topicSections -join [Environment]::NewLine)
      <section style="margin-top:28px; border-top:1px solid #e5e7eb; padding-top:20px;">
        <h2 style="font:600 18px/1.3 Segoe UI, Arial, sans-serif; margin:0 0 10px; color:#1f2937;">Additional Source Links</h2>
        <ul style="margin:0; padding-left:20px; font:14px/1.7 Segoe UI, Arial, sans-serif; color:#111827;">
          $rawItems
        </ul>
      </section>
    </div>
  </div>
</body>
</html>
"@
}

function Send-OutlookReportEmail {
    param(
        [Parameter(Mandatory)]
        [string]$To,
        [Parameter(Mandatory)]
        [string]$Subject,
        [Parameter(Mandatory)]
        [string]$HtmlBody
    )

    try {
        $outlook = New-Object -ComObject Outlook.Application
    }
    catch {
        throw "Microsoft Outlook is not available on this machine. Install/configure Outlook or switch to an SMTP/API mail sender."
    }

    $mail = $outlook.CreateItem(0)
    $mail.To = $To
    $mail.Subject = $Subject
    $mail.HTMLBody = $HtmlBody
    $mail.Send()
}

function Get-SmtpCredential {
    param(
        [string]$CredentialPath,
        [string]$UserName,
        [string]$PlainTextPassword
    )

    $effectiveUser = Get-EffectiveSetting -ExplicitValue $UserName -EnvironmentVariableName "SMTP_USERNAME"
    $effectivePassword = Get-EffectiveSetting -ExplicitValue $PlainTextPassword -EnvironmentVariableName "SMTP_PASSWORD"
    if (-not [string]::IsNullOrWhiteSpace($effectiveUser) -and -not [string]::IsNullOrWhiteSpace($effectivePassword)) {
        $securePassword = ConvertTo-SecureString -String $effectivePassword -AsPlainText -Force
        return [pscredential]::new($effectiveUser, $securePassword)
    }

    if ([string]::IsNullOrWhiteSpace($CredentialPath)) {
        throw "SMTP credentials were not provided. Set SMTP_USERNAME and SMTP_PASSWORD, pass -SmtpUser and -SmtpPassword, or use -SmtpCredentialPath."
    }

    if (-not (Test-Path -LiteralPath $CredentialPath)) {
        throw "SMTP credential file not found at $CredentialPath"
    }

    $payload = Get-Content -LiteralPath $CredentialPath -Raw | ConvertFrom-Json
    $securePassword = $payload.password | ConvertTo-SecureString
    return [pscredential]::new($payload.username, $securePassword)
}

function Send-SmtpReportEmail {
    param(
        [Parameter(Mandatory)]
        [string]$To,
        [Parameter(Mandatory)]
        [string]$Subject,
        [Parameter(Mandatory)]
        [string]$HtmlBody,
        [Parameter(Mandatory)]
        [string]$SmtpServer,
        [Parameter(Mandatory)]
        [int]$Port,
        [Parameter(Mandatory)]
        [pscredential]$Credential
    )

    $message = [System.Net.Mail.MailMessage]::new()
    $message.From = $Credential.UserName
    $message.To.Add($To)
    $message.Subject = $Subject
    $message.Body = $HtmlBody
    $message.IsBodyHtml = $true

    $smtp = [System.Net.Mail.SmtpClient]::new($SmtpServer, $Port)
    $smtp.EnableSsl = $true
    $smtp.Credentials = $Credential.GetNetworkCredential()
    $smtp.Send($message)

    $message.Dispose()
    $smtp.Dispose()
}

function New-MarkdownReport {
    param(
        [Parameter(Mandatory)]
        [string]$DateLabel,
        [AllowEmptyCollection()]
        [object[]]$Items,
        [AllowEmptyCollection()]
        [object[]]$Topics
    )

    $topDomains = $Items |
        Group-Object Domain |
        Sort-Object Count -Descending |
        Select-Object -First 6 |
        ForEach-Object { "$($_.Name) ($($_.Count))" }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# Social Trend Report for $DateLabel")
    $lines.Add("")
    $lines.Add("Collected **$($Items.Count)** items from public feeds for the previous day.")
    $lines.Add("")
    $lines.Add("## Snapshot")
    $lines.Add("")
    foreach ($platformLine in (Get-PlatformSummaryLines -Items $Items)) {
        $lines.Add($platformLine)
    }
    if (@($topDomains).Count -gt 0) {
        $lines.Add("- Frequent linked domains: $($topDomains -join ', ')")
    }
    $lines.Add("")
    $lines.Add("## What Was Popular")
    $lines.Add("")

    if (@($Topics).Count -eq 0) {
        $lines.Add("")
        $lines.Add("")
        $lines.Add("No items were collected for this window. Check network access or adjust the configured sources.")
        $lines.Add("")
    }
    else {
        foreach ($topic in $Topics) {
            $lines.Add("### $($topic.Label)")
            $lines.Add("")
            $lines.Add($topic.Summary)
            $lines.Add("")
            foreach ($example in $topic.Examples) {
                $lines.Add((Get-ExampleLine -Item $example))
            }
            $lines.Add("")
        }
    }

    $lines.Add("## Raw Sources")
    $lines.Add("")
    foreach ($item in ($Items | Where-Object { $_.PSObject.Properties["Url"] -and $_.PSObject.Properties["Platform"] } | Sort-Object PublishedAt, Score -Descending)) {
        $displayUrl = if ($item.ExternalUrl) { $item.ExternalUrl } else { $item.Url }
        $lines.Add("- [$($item.Title)]($displayUrl) | $($item.Platform) | $($item.Source) | $($item.PublishedAt.ToString("yyyy-MM-dd HH:mm"))")
    }

    return ($lines -join [Environment]::NewLine)
}

$scriptBasePath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$configFullPath = Get-AbsolutePath -Path $ConfigPath -BasePath $scriptBasePath
if (-not (Test-Path -LiteralPath $configFullPath)) {
    throw "Config file not found: $configFullPath"
}

$configBasePath = Split-Path -Parent $configFullPath
$config = Get-Content -LiteralPath $configFullPath -Raw | ConvertFrom-Json
$dateWindow = Get-DateWindow -ReferenceDate $ReportDate

$resolvedOutputDir = if ($OutputDir) {
    Get-AbsolutePath -Path $OutputDir -BasePath $scriptBasePath
}
else {
    Get-AbsolutePath -Path $config.output.directory -BasePath $configBasePath
}

New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$items = New-Object System.Collections.Generic.List[object]
$redditItems = @(Get-RedditItems -Sources $config.reddit -Start $dateWindow.Start -End $dateWindow.End)
$youtubeItems = @(Get-YouTubeItems -Sources $config.youtube -Start $dateWindow.Start -End $dateWindow.End)
$xItems = @(Get-XItems -Config $config.x -Start $dateWindow.Start -End $dateWindow.End -BasePath $configBasePath)
$tiktokItems = @(Get-TikTokItems -Config $config.tiktok -Start $dateWindow.Start -End $dateWindow.End)
if ($redditItems.Count -gt 0) {
    $items.AddRange($redditItems)
}
if ($youtubeItems.Count -gt 0) {
    $items.AddRange($youtubeItems)
}
if ($xItems.Count -gt 0) {
    $items.AddRange($xItems)
}
if ($tiktokItems.Count -gt 0) {
    $items.AddRange($tiktokItems)
}

$rankedItems = @($items | Sort-Object Score -Descending)
$topics = @(Get-Topics -Items $rankedItems -TopCount ([int]$config.output.topTopics) -ExamplesPerTopic ([int]$config.output.examplesPerTopic))
$report = New-MarkdownReport -DateLabel $dateWindow.Label -Items $rankedItems -Topics $topics

$reportPath = Join-Path $resolvedOutputDir ("trend-report-{0}.md" -f $dateWindow.Label)
$report | Set-Content -LiteralPath $reportPath -Encoding UTF8

$htmlReport = New-HtmlReport -DateLabel $dateWindow.Label -Items $rankedItems -Topics $topics
$htmlPath = Join-Path $resolvedOutputDir ("trend-report-{0}.html" -f $dateWindow.Label)
$htmlReport | Set-Content -LiteralPath $htmlPath -Encoding UTF8

Write-Host "Trend report written to $reportPath"

if ($SendEmail) {
    if (-not $EmailTo) {
        throw "EmailTo is required when -SendEmail is used."
    }

    $subject = "{0} - {1}" -f $EmailSubject, $dateWindow.Label
    switch ($EmailMethod) {
        "Outlook" {
            Send-OutlookReportEmail -To $EmailTo -Subject $subject -HtmlBody $htmlReport
        }
        "Smtp" {
            if (-not $SmtpCredentialPath) {
                $envUser = Get-EffectiveSetting -ExplicitValue $SmtpUser -EnvironmentVariableName "SMTP_USERNAME"
                $envPassword = Get-EffectiveSetting -ExplicitValue $SmtpPassword -EnvironmentVariableName "SMTP_PASSWORD"
                if ([string]::IsNullOrWhiteSpace($envUser) -or [string]::IsNullOrWhiteSpace($envPassword)) {
                    throw "SmtpCredentialPath is required when -EmailMethod Smtp is used unless SMTP_USERNAME and SMTP_PASSWORD are available."
                }
            }

            $resolvedCredentialPath = if ($SmtpCredentialPath) {
                Get-AbsolutePath -Path $SmtpCredentialPath -BasePath $scriptBasePath
            }
            else {
                $null
            }
            $credential = Get-SmtpCredential -CredentialPath $resolvedCredentialPath -UserName $SmtpUser -PlainTextPassword $SmtpPassword
            Send-SmtpReportEmail `
                -To $EmailTo `
                -Subject $subject `
                -HtmlBody $htmlReport `
                -SmtpServer $SmtpHost `
                -Port $SmtpPort `
                -Credential $credential
        }
    }
    Write-Host "Report email sent to $EmailTo"
}

if ($PassThru) {
    [pscustomobject]@{
        ReportPath = $reportPath
        HtmlPath = $htmlPath
        ItemCount = $rankedItems.Count
        TopicCount = $topics.Count
        ReportDate = $dateWindow.Label
    }
}
