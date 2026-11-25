$webUrl = "http://sharepointdev.com/enh/site01"
$listTitle = "E-signering - Kö"

$web = Get-SPWeb $webUrl

try {

    #
    # Kontrollera om listan redan finns
    #
    $existingList = $web.Lists.TryGetList($listTitle)

    if ($null -ne $existingList) {

        Write-Host ""
        Write-Host "Listan '$listTitle' finns redan." `
            -ForegroundColor Yellow

        Write-Host "Ingen ändring gjord."
        return
    }

    Write-Host ""
    Write-Host "Skapar '$listTitle'..." `
        -ForegroundColor Cyan

    #
    # Skapa vanlig Custom List
    #
    $listId = $web.Lists.Add(
        $listTitle,
        "Kö för inkommande e-signeringsbegäranden",
        [Microsoft.SharePoint.SPListTemplateType]::GenericList
    )

    $list = $web.Lists[$listId]

    #
    # Title används som dokumentnamn
    #
    $titleField = $list.Fields.GetFieldByInternalName("Title")
    $titleField.Title = "Dokument"
    $titleField.Required = $true
    $titleField.Update()

    #
    # Hjälpfunktion för textfält
    #
    function Add-TextField {

        param(
            [string]$InternalName,
            [string]$DisplayName,
            [bool]$Required = $false,
            [bool]$Indexed = $false
        )

        $fieldXml = @"
<Field
    Type="Text"
    Name="$InternalName"
    StaticName="$InternalName"
    DisplayName="$DisplayName"
    Required="$($Required.ToString().ToUpper())"
/>
"@

        $list.Fields.AddFieldAsXml(
            $fieldXml,
            $true,
            [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldInternalNameHint
        ) | Out-Null

        $field = $list.Fields.GetFieldByInternalName($InternalName)

        if ($Indexed) {
            $field.Indexed = $true
            $field.Update()
        }
    }

    #
    # Hjälpfunktion för Note-fält
    #
    function Add-NoteField {

        param(
            [string]$InternalName,
            [string]$DisplayName,
            [bool]$Required = $false
        )

        $fieldXml = @"
<Field
    Type="Note"
    Name="$InternalName"
    StaticName="$InternalName"
    DisplayName="$DisplayName"
    Required="$($Required.ToString().ToUpper())"
    RichText="FALSE"
    NumLines="6"
/>
"@

        $list.Fields.AddFieldAsXml(
            $fieldXml,
            $true,
            [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldInternalNameHint
        ) | Out-Null
    }

    #
    # Correlation
    #
    Add-TextField `
        -InternalName "ESignCorrelationId" `
        -DisplayName "Correlation ID" `
        -Required $true `
        -Indexed $true

    #
    # Queue Status
    #
    $statusXml = @"
<Field
    Type="Choice"
    Name="ESignQueueStatus"
    StaticName="ESignQueueStatus"
    DisplayName="Köstatus"
    Required="TRUE">
    <CHOICES>
        <CHOICE>New</CHOICE>
        <CHOICE>Processing</CHOICE>
        <CHOICE>Transferred</CHOICE>
        <CHOICE>Error</CHOICE>
    </CHOICES>
    <Default>New</Default>
</Field>
"@

    $list.Fields.AddFieldAsXml(
        $statusXml,
        $true,
        [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldInternalNameHint
    ) | Out-Null

    $statusField =
        $list.Fields.GetFieldByInternalName("ESignQueueStatus")

    $statusField.Indexed = $true
    $statusField.Update()

    #
    # Dokumentidentitet
    #
    Add-TextField `
        "ESignSiteUrl" `
        "Site URL" `
        $true

    Add-TextField `
        "ESignListId" `
        "List ID" `
        $true

    #
    # Item ID som Number
    #
    $itemIdXml = @"
<Field
    Type="Number"
    Name="ESignItemId"
    StaticName="ESignItemId"
    DisplayName="Item ID"
    Required="TRUE"
    Decimals="0"
/>
"@

    $list.Fields.AddFieldAsXml(
        $itemIdXml,
        $true,
        [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldInternalNameHint
    ) | Out-Null

    Add-TextField `
        "ESignFileUniqueId" `
        "File Unique ID" `
        $true `
        $true

    Add-TextField `
        "ESignFileName" `
        "File Name" `
        $true `
        $true

    Add-TextField `
        "ESignFileRef" `
        "File Ref" `
        $true

    #
    # Undertecknare som JSON
    #
    Add-NoteField `
        "ESignSignersJson" `
        "Undertecknare JSON" `
        $true

    #
    # Meddelande
    #
    Add-NoteField `
        "ESignMessage" `
        "Meddelande" `
        $false

    #
    # Deadline
    #
    $deadlineXml = @"
<Field
    Type="DateTime"
    Name="ESignDeadline"
    StaticName="ESignDeadline"
    DisplayName="Deadline"
    Required="TRUE"
    Format="DateOnly"
/>
"@

    $list.Fields.AddFieldAsXml(
        $deadlineXml,
        $true,
        [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldInternalNameHint
    ) | Out-Null

    $deadlineField =
        $list.Fields.GetFieldByInternalName("ESignDeadline")

    $deadlineField.Indexed = $true
    $deadlineField.Update()

    #
    # Initiator
    #
    Add-TextField `
        "ESignInitiatorName" `
        "Initiator Name" `
        $true

    Add-TextField `
        "ESignInitiatorEmail" `
        "Initiator Email" `
        $false

    Add-TextField `
        "ESignInitiatorAccount" `
        "Initiator Account" `
        $true

    Add-TextField `
        "ESignInitiatorUserName" `
        "Initiator UserName" `
        $false

    Add-TextField `
        "ESignInitiatorUPN" `
        "Initiator UPN" `
        $false

    Add-TextField `
        "ESignInitiatorSource" `
        "Initiator Source" `
        $false

    Add-TextField `
        "ESignInitiatorIdentifier" `
        "Initiator Identifier" `
        $false

    #
    # Uppdatera listan
    #
    $list.Update()

    Write-Host ""
    Write-Host "Kölistan skapades." `
        -ForegroundColor Green

    Write-Host ""
    Write-Host "Web:   $webUrl"
    Write-Host "Lista: $listTitle"
    Write-Host "ID:    $($list.ID)"

    Write-Host ""
    Write-Host "Kolumner:" `
        -ForegroundColor Cyan

    $list.Fields |
        Where-Object {
            $_.InternalName -eq "Title" -or
            $_.InternalName -like "ESign*"
        } |
        Select-Object `
            Title,
            InternalName,
            StaticName,
            TypeAsString,
            Required,
            Indexed |
        Format-Table -AutoSize
}
finally {

    if ($null -ne $web) {
        $web.Dispose()
    }
}