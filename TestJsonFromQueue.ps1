$webUrl = "http://sharepointdev.com/enh/site01"
$listTitle = "E-signering - Kö"

$web = Get-SPWeb $webUrl

try {

    $list = $web.Lists[$listTitle]

    # Senast skapade köposten
    $item = $list.Items |
        Sort-Object ID -Descending |
        Select-Object -Last 1

    Write-Host "Kö-ID: $($item.ID)" -ForegroundColor Cyan

    $json = $item["ESignSignersJson"]

    Write-Host ""
    Write-Host "Råvärde från SharePoint:" -ForegroundColor Yellow
    Write-Host $json

    Write-Host ""
    Write-Host "Försöker läsa JSON..." -ForegroundColor Cyan

    $signers = $json | ConvertFrom-Json

    Write-Host ""
    Write-Host "JSON OK!" -ForegroundColor Green

    $signers |
        Select-Object `
            id,
            displayName,
            email,
            accountName,
            userName,
            userPrincipalName,
            source,
            identifier |
        Format-List
}
catch {

    Write-Host ""
    Write-Host "JSON kunde inte läsas." -ForegroundColor Red
    Write-Host $_.Exception.Message
}
finally {

    if ($null -ne $web) {
        $web.Dispose()
    }
}
