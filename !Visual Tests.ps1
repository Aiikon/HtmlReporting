Import-Module $PSScriptRoot -Force

#$Global:2e128b9186234521b3ab5ce70cc83360_ForceLoadPowerShellCmdlets = $true
$someVariable = "abcdefg"

Get-HtmlFragment {
    
    h1 "Semi-Fancy Get-HtmlText Usage"
    p "Paragraph"
    p "Paragraph"
    h2 "Header 2"
    p "Paragraph"
    p "Paragraph"
    h3 "Header 3"
    p "Paragraph"
    p "Paragraph"
    h4 "Header 4"
    p "Paragraph"
    p "Paragraph"
    
    $header3 = "Header Three"
    h3 $header3
    
    p "Here's a sample paragraph" (br) "And a line break"
    p "Here's a sample paragraph with some" (strong strong) "text"
    p "Some" (em (strong "strong and emphasized")) "text"
    p "And making use of `$someVariable: $someVariable"
    p "Making use of some $(span -style 'color:red;' "red") text"

    p {
        "A sample link:"
        a -href 'test.html' 'Test'
    }
    ol {
        li "Item 1"
        li "Item 2"
    }
    ul {
        li "Item A"
        ol {
            li "Item A1"
            li "Item A2"
        }
        li "Item B"
    }

    hr

    h2 Indicators
    ul {
        li (Get-HtmlIndicatorText "Sample 1" -ColorName Green)
        li (Get-HtmlIndicatorText "Sample 2" -ColorName Red -BorderOnly)
        li (Get-HtmlIndicatorText "Sample 3" -ColorIndex 0 -Width 200px)
        li (Get-HtmlIndicatorText "Hover Over Me" -ColorRGB 0,0,0 -Title 'Title Usage')
    }

    hr

    h2 Table

    h3 Normal
    Get-Service | Select-Object Name, Status, DisplayName -First 10 | ConvertTo-HtmlTable

    h3 Narrow
    Get-Service | Select-Object Name, Status, DisplayName -First 10 -Skip 10 | ConvertTo-HtmlTable -Narrow

    h3 Custom Class
    "<style>table.HtmlReportingTable.Wide td { padding: 30px 10px 30px 10px; }</style>"
    Get-Service | Select-Object Name, Status, DisplayName -First 10 -Skip 20 | ConvertTo-HtmlTable -Narrow -Class Wide

    h3 Rowspan from Number
    "
        Rowspan,Value
        1,A
        2,B
        2,C
        1,D
    ".Trim() -replace '\A ' | ConvertFrom-Csv | ConvertTo-HtmlTable -CellRowspanScripts @{Value='Rowspan'}

    h2 No Content

    @() | ConvertTo-HtmlTable -NoContentHtml (p (em "No Content"))

    hr

    h2 "Convert PS Code to HTML"

    div {
        p {
            Convert-PSCodeToHtml -PsCode "Get-ChildItem C:\ | Out-GridView -Title `$var1"
        }
    }

    h2 ConvertTo-HtmlStrongText

    $modes = 'SameLine', 'MultiLine', 'MultiLineIndent', 'MultiLineHorizontal', 'MultiLineIndentHorizontal'
    $i = 0
    foreach ($mode in $modes)
    {
        h3 $mode

        Get-Service |
            Select-Object -First 1 -Skip $i -Property Name, DisplayName, Status |
            ConvertTo-HtmlStrongText -Mode $mode

        $i += 1
    }

    hr

    h2 ConvertTo-HtmlColorBlocks

    Get-Service | Select-Object -First 5 Name, DisplayName, Status, StartType |
        ConvertTo-HtmlColorBlocks -TocProperty Name -OutputScript { $_ | ConvertTo-HtmlStrongText }

    Get-Service | Select-Object -First 5 Name, DisplayName, Status, StartType |
        ConvertTo-HtmlColorBlocks -TocProperty Name -OutputScript { p $_.Name } -NarrowToc

    hr

    h2 Assorted HTML Table Tests

    Get-ChildItem C:\Windows |
        Select-Object Name, LastWriteTime |
        ConvertTo-HtmlTable -NoWrapProperty * -RightAlignProperty Name -Narrow -RowClassScript {
            if ($_.Name -match 'boot') { 'red' }
        } -RowStyleScript {
            if ($_.Name -match 'Branding') { 'font-family: monospace' }
        } -CellClassScripts @{
            PSParentPath = {
                if ($_.Name -eq 'Assembly') { 'blue' }
            }
        } -CellStyleScripts @{
            PSParentPath = { if ($_.Name -eq 'CSC') { 'font-family: monospace' } }
        } -CellColspanScripts @{
            PSPath = { if ($_.Name -eq 'addins') { 3 } }
        } -CellRowspanScripts @{
            PSParentPath = { if ($_.Name -eq 'debug') { 3 } }
        }

    h2 "HTML Table Insert Lines"

    [pscustomobject]@{A=1;B=2;C=3;D=4} |
        ConvertTo-HtmlTable -InsertSolidLine B -InsertDashedLine C -InsertDottedLine D

    h2 "Hourly Heatmap (One Set)"

    0..(14*24-1) |
        ForEach-Object {
            [pscustomobject]@{Timestamp=[DateTime]::Today.AddHours($_); Value = Get-Random -Minimum 0 -Maximum 100}
        } |
        ConvertTo-HtmlHourlyHeatmap -TimestampProperty Timestamp -ValueProperty Value -Columns 4 -IndicatorSize 20 -IndicatorPadding 2 -HeatmapColors @(
            Define-HtmlHeatmapColor -ColorName Green -Mode AtLeast -Value 50
            Define-HtmlHeatmapColor -ColorName Orange -Mode AtLeast -Value 25
            Define-HtmlHeatmapColor -ColorName Red
        )

    h2 "Hourly Heatmap (Multiple Sets)"

    'A', 'B', 'C' |
        ForEach-Object {
            $set = $_
            0..(14*24-1) |
                ForEach-Object {
                    [pscustomobject]@{Set = "Set $set"; Timestamp=[DateTime]::Today.AddHours($_); Value = Get-Random -Minimum 0 -Maximum 100}
                }
    } |
        ConvertTo-HtmlHourlyHeatmap -SetProperty Set -TimestampProperty Timestamp -ValueProperty Value -Columns 4 -IndicatorSize 20 -IndicatorPadding 2 -HeatmapColors @(
            Define-HtmlHeatmapColor -ColorName Green -Mode AtLeast -Value 50
            Define-HtmlHeatmapColor -ColorName Orange -Mode AtLeast -Value 25
            Define-HtmlHeatmapColor -ColorName Red
        )

    h2 "Monthly Heatmap (One Set)"

        0..(31*4) |
            ForEach-Object {
                [pscustomobject]@{Set = "Set $set"; Timestamp=[DateTime]::Today.AddDays($_); Value = Get-Random -Minimum 0 -Maximum 100}
            } |
        ConvertTo-HtmlMonthlyHeatmap -TimestampProperty Timestamp -ValueProperty Value -IndicatorSize 20 -IndicatorPadding 2 -HeatmapColors @(
            Define-HtmlHeatmapColor -ColorName Green -Mode AtLeast -Value 50
            Define-HtmlHeatmapColor -ColorName Orange -Mode AtLeast -Value 25
            Define-HtmlHeatmapColor -ColorName Red
        )

    h2 "Monthly Heatmap (Multiple Sets)"

    'A', 'B', 'C' |
        ForEach-Object {
            $set = $_
            0..(31*4) |
                ForEach-Object {
                    [pscustomobject]@{Set = "Set $set"; Timestamp=[DateTime]::Today.AddDays($_); Value = Get-Random -Minimum 0 -Maximum 100}
                }
    } |
        ConvertTo-HtmlMonthlyHeatmap -SetProperty Set -TimestampProperty Timestamp -ValueProperty Value -IndicatorSize 20 -IndicatorPadding 2 -HeatmapColors @(
            Define-HtmlHeatmapColor -ColorName Green -Mode AtLeast -Value 50
            Define-HtmlHeatmapColor -ColorName Orange -Mode AtLeast -Value 25
            Define-HtmlHeatmapColor -ColorName Red
        )

} |
    Out-HtmlFile -AddTimestamp

return

& {
    "<h1>Header 1</h1>"

    "<p>Alternatively just write HTML like this...</p>"
    
    Get-ChildItem C:\ | Select-Object Name, LastWriteTime, Length -First 10 | ConvertTo-HtmlTable
} |
    Out-HtmlFile ~\Desktop\Test.html -Open
