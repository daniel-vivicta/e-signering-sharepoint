param(
    [Parameter(Mandatory = $false)]
    [string]$WebUrl = "http://sharepointdev.com/enh/site01",

    [Parameter(Mandatory = $false)]
    [string]$SourceFolder = "C:\SPFx\ESignering"
)

$ErrorActionPreference = "Stop"

$bundleFileName = "ESignClassic.js"
$bundleFilePath = Join-Path $SourceFolder "dist\classic\$bundleFileName"

$iconFileName = "signing.png"
$iconFilePath = Join-Path $SourceFolder $iconFileName

$loaderActionName = "ESignClassicLoaderScript"
$oldBundleScriptActionName = "ESignClassicBundleScript"
$oldPocScriptActionName = "ESignClassicPocScript"
$ecbActionName = "ESignStartSigningClassic"
$oldPocFileName = "ESignClassicPoc.js"

$web = Get-SPWeb $WebUrl

try {
    foreach ($requiredFile in @($bundleFilePath, $iconFilePath)) {
        if (-not (Test-Path $requiredFile)) {
            throw "Filen hittades inte: $requiredFile"
        }
    }

    Write-Host "Installerar Classic e-signering med loader..." -ForegroundColor Cyan
    Write-Host "  Web:      $WebUrl"
    Write-Host "  Källa:    $SourceFolder"
    Write-Host "  Scope:    Site/Web"
    Write-Host "  Listtyp:  101 (Document Library)"
    Write-Host ""

    # ------------------------------------------------------------
    # Säkerställ Site Assets
    # ------------------------------------------------------------
    $siteAssets = $web.Lists.TryGetList("Site Assets")

    if ($null -eq $siteAssets) {
        Write-Host "Site Assets saknas. Skapar biblioteket..." -ForegroundColor Yellow

        $siteAssetsId = $web.Lists.Add(
            "Site Assets",
            "Filer för webbplatsen",
            [Microsoft.SharePoint.SPListTemplateType]::DocumentLibrary
        )
        
        $siteAssets = $web.Lists[$siteAssetsId]
    }

    $siteAssetsUrl = $siteAssets.RootFolder.ServerRelativeUrl.TrimEnd('/')

    function Upload-FileToSiteAssets {
        param(
            [string]$LocalPath,
            [string]$FileName
        )

        $bytes = [System.IO.File]::ReadAllBytes($LocalPath)
        $targetUrl = "$siteAssetsUrl/$FileName"

        $siteAssets.RootFolder.Files.Add(
            $targetUrl,
            $bytes,
            $true
        ) | Out-Null

        Write-Host "Uppladdad: $targetUrl" -ForegroundColor Green
        return $targetUrl
    }

    # ------------------------------------------------------------
    # Ladda upp bundle och ikon
    # ------------------------------------------------------------
    $bundleTargetUrl = Upload-FileToSiteAssets $bundleFilePath $bundleFileName
    $iconTargetUrl = Upload-FileToSiteAssets $iconFilePath $iconFileName

    # ------------------------------------------------------------
    # Ta bort gamla ScriptLinks och tidigare loader
    # ------------------------------------------------------------
    @(
        $web.UserCustomActions |
        Where-Object {
            $_.Name -eq $loaderActionName -or
            $_.Name -eq $oldBundleScriptActionName -or
            $_.Name -eq $oldPocScriptActionName
        }
    ) | ForEach-Object {
        Write-Host "Tar bort gammal script-action: $($_.Name)" -ForegroundColor DarkGray
        $_.Delete()
    }

    # ------------------------------------------------------------
    # Loader
    # ------------------------------------------------------------
    # Vi laddar INTE längre hela ESignClassic.js via ScriptSrc på varje sida.
    # I stället laddas en liten loader alltid. Den hämtar bundlen först när
    # användaren väljer Starta signering.
    #
    # Loadern köar även flera samtidiga open-anrop medan bundlen laddas.
    # ------------------------------------------------------------
    $bundleUrlForJavaScript = "$siteAssetsUrl/$bundleFileName"

    $loaderScript = @"
(function(w,d){
    if (w.ESignClassicLoader) { return; }

    var loading = false;
    var pending = [];
    var bundleUrl = '$bundleUrlForJavaScript';

    function runPending() {
        var calls = pending.slice(0);
        pending = [];

        for (var i = 0; i < calls.length; i++) {
            w.ESignClassic.open(
                calls[i].itemId,
                calls[i].itemUrl,
                calls[i].listId,
                calls[i].siteUrl
            );
        }
    }

    function failPending(message) {
        loading = false;
        pending = [];

        if (w.console && w.console.error) {
            w.console.error('[ESIGN][CLASSIC LOADER]', message);
        }

        alert('E-signering kunde inte laddas. Prova att ladda om sidan.');
    }

    function loadBundle() {
        if (w.ESignClassic && typeof w.ESignClassic.open === 'function') {
            runPending();
            return;
        }

        if (loading) {
            return;
        }

        loading = true;

        var script = d.createElement('script');
        script.type = 'text/javascript';
        script.src = bundleUrl + '?v=' + new Date().getTime();

        script.onload = function() {
            loading = false;

            if (w.ESignClassic && typeof w.ESignClassic.open === 'function') {
                runPending();
            }
            else {
                failPending('ESignClassic.js laddades men ESignClassic.open saknas.');
            }
        };

        script.onerror = function() {
            failPending('Kunde inte ladda ' + bundleUrl);
        };

        (d.head || d.getElementsByTagName('head')[0]).appendChild(script);
    }

    w.ESignClassicLoader = {
        open: function(itemId, itemUrl, listId, siteUrl) {
            if (w.ESignClassic && typeof w.ESignClassic.open === 'function') {
                w.ESignClassic.open(itemId, itemUrl, listId, siteUrl);
                return;
            }

            pending.push({
                itemId: itemId,
                itemUrl: itemUrl,
                listId: listId,
                siteUrl: siteUrl
            });

            loadBundle();
        }
    };
})(window, document);
"@

    $loaderAction = $web.UserCustomActions.Add()
    $loaderAction.Name = $loaderActionName
    $loaderAction.Title = "E-signering Classic Loader"
    $loaderAction.Location = "ScriptLink"
    $loaderAction.Sequence = 100
    $loaderAction.ScriptBlock = $loaderScript
    $loaderAction.Update()

    Write-Host "Classic loader registrerad." -ForegroundColor Green

    # ------------------------------------------------------------
    # Städa gammal site-level ECB-action
    # ------------------------------------------------------------
    @(
        $web.UserCustomActions |
        Where-Object { $_.Name -eq $ecbActionName }
    ) | ForEach-Object {
        Write-Host "Tar bort gammal site-level ECB-action: $($_.Name)" -ForegroundColor DarkGray
        $_.Delete()
    }

    # ------------------------------------------------------------
    # Städa gamla list-level actions från tidigare POC-installationer
    # ------------------------------------------------------------
    foreach ($list in $web.Lists) {
        @(
            $list.UserCustomActions |
            Where-Object { $_.Name -eq $ecbActionName }
        ) | ForEach-Object {
            Write-Host "Tar bort gammal list-level ECB-action från '$($list.Title)'." -ForegroundColor Yellow
            $_.Delete()
        }
    }

    # ------------------------------------------------------------
    # Site-level ECB-action för alla dokumentbibliotek
    # ------------------------------------------------------------
    $ecbAction = $web.UserCustomActions.Add()
    $ecbAction.Name = $ecbActionName
    $ecbAction.Title = "Starta signering"
    $ecbAction.Description = "Öppna formuläret för e-signering"
    $ecbAction.Location = "EditControlBlock"
    $ecbAction.Sequence = 510

    $ecbAction.RegistrationType =
        [Microsoft.SharePoint.SPUserCustomActionRegistrationType]::List

    $ecbAction.RegistrationId = "101"

    $ecbAction.Rights =
        [Microsoft.SharePoint.SPBasePermissions]::ViewListItems

    $ecbAction.Url =
        "javascript:ESignClassicLoader.open('{ItemId}','{ItemUrl}','{ListId}','{SiteUrl}');"

    $ecbAction.Update()

    Write-Host "Site-level ECB-kommandot använder nu ESignClassicLoader.open(...)." -ForegroundColor Green

    # ------------------------------------------------------------
    # Ta bort gamla POC-filen om den finns
    # ------------------------------------------------------------
    $oldPocUrl = "$siteAssetsUrl/$oldPocFileName"
    $oldPocFile = $web.GetFile($oldPocUrl)

    if ($null -ne $oldPocFile -and $oldPocFile.Exists) {
        Write-Host "Tar bort gammal POC-fil: $oldPocUrl" -ForegroundColor Yellow
        $oldPocFile.Delete()
    }

    # ------------------------------------------------------------
    # Slutkontroll
    # ------------------------------------------------------------
    Write-Host ""
    Write-Host "==============================================" -ForegroundColor Green
    Write-Host "Classic e-signering installerad med loader" -ForegroundColor Green
    Write-Host "==============================================" -ForegroundColor Green
    Write-Host "Web:        $WebUrl"
    Write-Host "Scope:      Site/Web"
    Write-Host "Listtyp:    101 (Document Library)"
    Write-Host "Bundle:     $bundleTargetUrl"
    Write-Host "Ikon:       $iconTargetUrl"
    Write-Host "Loader:     ESignClassicLoader"
    Write-Host "ECB:        Starta signering -> ESignClassicLoader.open(...)"
    Write-Host ""
    Write-Host "ESignClassic.js laddas nu först när användaren väljer Starta signering." -ForegroundColor Cyan
}
finally {
    if ($null -ne $web) {
        $web.Dispose()
    }
}