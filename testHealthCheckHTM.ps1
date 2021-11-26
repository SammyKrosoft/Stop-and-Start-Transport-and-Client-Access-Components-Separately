
[boolean]$trigger = $true
$URL = "https://E2016-01.canadadrey.ca/owa"
$testURL = "$URL/healthcheck.htm"


Write-Host "Testing $testURL " -NoNewline
While ($trigger) {
    $web = Invoke-WebRequest -Uri "$testURL" -UseBasicParsing -ErrorAction Stop
    if ($web.StatusCode -eq "200"){
        Write-Host -ForegroundColor Green "OK"
    } else {
        Write-Host -ForegroundColor Red $($Web.StatusCode)
        }
        $trigger = $false
}