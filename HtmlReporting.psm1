try
{
    if ($Global:2e128b9186234521b3ab5ce70cc83360_ForceLoadPowerShellCmdlets -eq $true) { throw "Skipping HtmlReportingSharp Compilation" }
    $date = [System.IO.File]::GetLastWriteTime("$PSScriptRoot\HtmlReportingSharp\Helpers.cs").ToString("yyyyMMdd_HHmmss")
    $Script:OutputPath = "$Env:LOCALAPPDATA\Rhodium\Module\HtmlReportingSharp_$date\HtmlReportingSharp.dll"
    if (![System.IO.File]::Exists($outputPath))
    {
        [void][System.IO.Directory]::CreateDirectory("$Env:LOCALAPPDATA\Rhodium")
        [void][System.IO.Directory]::CreateDirectory("$Env:LOCALAPPDATA\Rhodium\Module")
        [void][System.IO.Directory]::CreateDirectory("$Env:LOCALAPPDATA\Rhodium\Module\HtmlReportingSharp_$date")
        $fileList = [System.IO.Directory]::GetFiles("$PSScriptRoot\HtmlReportingSharp", "*.cs")
        Add-Type -Path $fileList -OutputAssembly $Script:OutputPath -OutputType Library -ErrorAction Stop -ReferencedAssemblies System.Web
    }

    Import-Module -Name $Script:OutputPath -Force -ErrorAction Stop

    $Script:LoadedHtmlReportingSharp = $true
}
catch
{
    Write-Warning "Unable to compile C# cmdlets; falling back to regular cmdlets."
    Write-Host -ForegroundColor Red $_.Exception.Message
    $Script:LoadedHtmlReportingSharp = $false
}

Add-Type -AssemblyName 'System.Web'

Add-Type @"
using System;
using System.Management.Automation;
using System.Collections;
using System.Collections.Generic;

namespace Rhodium.HtmlReporting
{
    public static class HtmlReportingHelpers
    {
        public static IEnumerable<string> ExcludeLikeAny(string[] Values, string[] Filters)
        {
            if (Filters == null)
                foreach (string value in Values)
                    yield return value;
                
            var patterns = new WildcardPattern[Filters.Length];
            for (int i = 0; i < Filters.Length; i++)
                patterns[i] = new WildcardPattern(Filters[i], WildcardOptions.IgnoreCase);

            foreach (string value in Values)
            {
                bool matched = false;
                for (int j = 0; j < patterns.Length; j++)
                {
                    if (patterns[j].IsMatch(value))
                    {
                        matched = true;
                        break;
                    }
                }
                if (!matched)
                    yield return value;
            }
        }

        public static IEnumerable<string> GetStringsLike(string[] Values, string[] Like, string[] NotLike, bool LikeDefaultsToWildcard = true)
        {
            if (Values == null || Values.Length == 0)
                yield break;

            bool hasLike = Like != null && Like.Length > 0;
            bool hasNotLike = NotLike != null && NotLike.Length > 0;
            bool likeIsWildcard = hasLike && Like[0] == "*";

            if (!LikeDefaultsToWildcard && !hasLike)
                yield break;

            if (Like == null)
                Like = new string[0] {};
            if (NotLike == null)
                NotLike = new string[0] {};

            var likePatterns = new WildcardPattern[Like.Length];
            for (int i = 0; i < Like.Length; i++)
                likePatterns[i] = new WildcardPattern(Like[i], WildcardOptions.IgnoreCase);

            var notLikePatterns = new WildcardPattern[NotLike.Length];
            for (int i = 0; i < NotLike.Length; i++)
                notLikePatterns[i] = new WildcardPattern(NotLike[i], WildcardOptions.IgnoreCase);

            foreach (string value in Values)
            {
                if (hasNotLike)
                {
                    for (int j = 0; j < notLikePatterns.Length; j++)
                    {
                        if (notLikePatterns[j].IsMatch(value))
                        {
                            goto nextValue;
                        }
                    }
                }

                if (likeIsWildcard)
                {
                    yield return value;
                    continue;
                }

                if (hasLike)
                {
                    for (int j = 0; j < likePatterns.Length; j++)
                    {
                        if (likePatterns[j].IsMatch(value))
                        {
                            yield return value;
                            goto nextValue;
                        }
                    }
                    continue;
                }

                yield return value;

                nextValue:
                    continue;
            }
        }

        public static Hashtable GetStringsLikeHashtable(string[] Values, string[] Like, string[] NotLike, bool LikeDefaultsToWildcard = true)
        {
            var hashtable = new Hashtable(StringComparer.CurrentCultureIgnoreCase);
            foreach (string match in GetStringsLike(Values, Like, NotLike, LikeDefaultsToWildcard))
                hashtable[match] = true;
            return hashtable;
        }
    }
}
"@

$Script:HtmlStyle = @"
h1, h2, h3, h4, p, li, td, th {
    color: black;
    font-size: 11pt;
    font-family: Calibri, "Lucida Sans", Sans-Serif;
}
body, a {
    font-size: 11pt;
    font-family: Calibri, "Lucida Sans", Sans-Serif;
}
span.title {
    font-weight: bold;
    font-size: 18pt;
}
h1 {
    font-weight: bold;
    font-size: 22pt;
    margin: 0.6em 0 0.6em 0;
}
h2 {
    font-weight: bold;
    font-size: 15pt;
    margin: 0.8em 0 0.2em 0;
}
h3 {
    font-weight: bold;
    font-size: 13.5pt;
    margin: 0.6em 0 0.1em 0;
}
h4 {
    font-weight: bold;
    font-size: 12pt;
    margin: 0.5em 0 0.1em 0;
}
p {
    margin: 0.3em 0 0.5em 0;
}

table.HtmlReportingTable {
    border-collapse: collapse;
    border-spacing: 0;
}
table.HtmlReportingTable th {
    text-align: left;
    border-style: none none solid none;
    border-width: 0 0 2px 0;
    border-color: black;
    padding: 2px 10px 2px 10px;
    page-break-inside: avoid;
}
table.HtmlReportingTable td {
    border-style: solid none none none;
    border-width: 1px 0 0 0;
    border-color: black;
    padding: 6px 10px 6px 10px;
    page-break-inside: avoid;
}
table.HtmlReportingTable.Narrow td {
    padding: 1px 10px 1px 10px;
}
table.HtmlReportingTable td.nowrap {
    white-space: nowrap;
}
table.HtmlReportingTable th.nowrap {
    white-space: nowrap;
}
table.HtmlReportingTable .ralign {
    text-align: right;
}
table.HtmlReportingTable.Grid th {
    padding: 1px 2px 1px 2px;
    border-style: solid;
    border-width: 1px;
    border-color: black;
}
table.HtmlReportingTable.Grid td {
    padding: 1px 2px 1px 2px;
    border-style: solid;
    border-width: 1px;
    border-color: black;
}

.IndicatorText {
    color: white;
    padding: 1px 6px 1px 6px;
    border-radius: 5px;
    display: inline-block;
    white-space: nowrap;
    margin: 1px 1px 1px 1px;
}

.IndicatorTextBorder {
    padding: 0 5px 0 5px;
    border-style: solid;
    border-thickness: 1.5px;
    border-radius: 5px;
    display: inline-block;
    white-space: nowrap;
    margin: 1px 1px 1px 1px;
}

.red {
    color: red;
}
.green {
    color: green;
}
.yellow {
    color: yellow;
}
.orange {
    color: orange;
}
.blue {
    color: blue;
}
.cursorhand {
    cursor: hand;
}
.cursorhelp {
    cursor:help;
}
"@

Function ConvertTo-HtmlColorBlocks
{
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter(Position=0,Mandatory=$true)] [ScriptBlock] $OutputScript,
        [Parameter()] [string[]] $TocProperty,
        [Parameter()] [string[]] $HtmlTocProperty,
        [Parameter()] [string] $SectionProperty,
        [Parameter()] [switch] $NarrowToc
    )
    Begin
    {
        $inputObjectList = [System.Collections.Generic.List[object]]::new()
    }
    Process
    {
        if ($InputObject) { $inputObjectList.Add($InputObject) }
    }
    End
    {
        if ($TocProperty)
        {
            $propertyList = & { $TocProperty; 'Link' } | Select-Object -Unique
            $htmlPropertyList = & { $HtmlTocProperty; 'Link' } | Select-Object -Unique
            $idList = [System.Collections.Generic.List[string]]::new()
            $inputObjectList |
                ForEach-Object {
                    $id = [guid]::NewGuid().ToString('n')
                    if ($SectionProperty) { $id = $_.$SectionProperty }
                    $idList.Add($id)
                    $_.PSObject.Properties.Add([PSNoteProperty]::New('Link', "<a href='#$id'>Link</a>"))
                    $_
                } |
                ConvertTo-HtmlTable -Property $propertyList -HtmlProperty $htmlPropertyList -Narrow:$NarrowToc
            "<br /><br />"
        }

        $i = 0
        foreach ($InputObject in $inputObjectList)
        {
            $html = $OutputScript.InvokeWithContext($null, (New-Object PSVariable "_", @($InputObject)), $null)
            $color = Get-HtmlReportColor -Index $i -AsCssRgb
            if ($TocProperty)
            {
                $id = $idList[$i]
                "<a name='$id' />"
            }
            "<table style='margin-bottom:2em;'><tr><td style='background-color:$color; width:15px;'>&nbsp;</td>"
            "<td style='padding-left:10px; page-break-inside:avoid;'>"
            $html -join "`r`n"
            "</td></tr></table>"
            $i += 1
        }
    }
}

if (!$Script:LoadedHtmlReportingSharp) {
Function ConvertTo-HtmlTable
{
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter()] [string[]] $Property,
        [Parameter()] [string[]] $HtmlProperty,
        [Parameter()] [string[]] $Class,
        [Parameter()] [switch] $RowsOnly,
        [Parameter()] [scriptblock] $RowClassScript,
        [Parameter()] [scriptblock] $RowStyleScript,
        [Parameter()] [hashtable] $CellClassScripts = @{},
        [Parameter()] [hashtable] $CellStyleScripts = @{},
        [Parameter()] [hashtable] $CellColspanScripts = @{},
        [Parameter()] [hashtable] $CellRowspanScripts = @{},
        [Parameter()] [hashtable] $RenameHeader = @{},
        [Parameter()] [string[]] $RightAlignProperty,
        [Parameter()] [string[]] $NoWrapProperty,
        [Parameter()] [switch] $Narrow,
        [Parameter()] [switch] $AutoDetectHtml,
        [Parameter()] [switch] $AddDataColumnName,
        [Parameter()] [string] $NoContentHtml,
        [Parameter()] [string[]] $ExcludeProperty
    )
    Begin
    {
        $inputObjectList = [System.Collections.Generic.List[object]]::new()
    }
    Process
    {
        $inputObjectList.Add($InputObject)
    }
    End
    {
        $resultList = [System.Collections.Generic.List[string]]::new()
        $headerList = [System.Collections.Generic.List[string]]::new()
        if (-not $Property)
        {
            $Property = $inputObjectList[0].PSObject.Properties.Name
        }

        if ($ExcludeProperty)
        {
            $Property = [Rhodium.HtmlReporting.HtmlReportingHelpers]::ExcludeLikeAny($Property, $ExcludeProperty)
        }

        foreach ($p in $Property) { $headerList.Add($p) }

        if (-not $RowsOnly.IsPresent -and $inputObjectList.Count)
        {
            $classList = @('HtmlReportingTable')
            if ($Narrow) { $classList += 'Narrow' }
            if ($Class) { foreach ($c in $Class) { $classList += $c } }
            $resultList.Add("<table class='$($classList -join ' ')'>")
            $resultList.Add("<thead>")
            $resultList.Add("<tr class='header'>")
            foreach ($header in $headerList)
            {
                if ($RenameHeader[$header]) { $header = $RenameHeader[$header] }
                $attrDcn = if ($AddDataColumnName) { " data-column-name='$([System.Web.HttpUtility]::HtmlAttributeEncode($header))'" }
                $resultList.Add("<th$attrDcn>$header</th>")
            }
            $resultList.Add("</tr>")
            $resultList.Add("</thead>")
            $resultList.Add("<tbody>")
        }

        if (!$inputObjectList.Count -and $NoContentHtml) { $resultList.Add($NoContentHtml) }

        $rowspanCountHash = @{}
        $colspanRowHash = @{}
        $colspanCountHash = @{}

        foreach ($object in $inputObjectList)
        {
            $rowClassList = [System.Collections.Generic.List[string]]::new()
            $rowStyleList = [System.Collections.Generic.List[string]]::new()
            if ($RowClassScript)
            {
                $object | ForEach-Object $RowClassScript | ForEach-Object { $rowClassList.Add($_) }
            }
            if ($RowStyleScript)
            {
                $object | ForEach-Object $RowStyleScript | ForEach-Object { $rowStyleList.Add($_) }
            }

            $colspanCount = 0
            $resultList.Add('<tr>')

            $cellClassDict = [ordered]@{}
            foreach ($pair in $CellClassScripts.GetEnumerator())
            {
                $result = $object | ForEach-Object $pair.Value
                foreach ($header in $pair.Key)
                {
                    $cellClassDict[$header] = $result
                }
            }

            $cellStyleDict = [ordered]@{}
            foreach ($pair in $CellStyleScripts.GetEnumerator())
            {
                $result = $object | ForEach-Object $pair.Value
                foreach ($header in $pair.Key)
                {
                    $cellStyleDict[$header] = $result
                }
            }

            foreach ($header in $headerList)
            {
                $skipCell = $false
                if ($colspanRowHash[$header] -gt 0)
                {
                    $colspanRowHash[$header] -= 1
                    $colspanCount = $colspanCountHash[$header]
                    $skipCell = $true
                }
                if ($rowspanCountHash[$header] -gt 0)
                {
                    $rowspanCountHash[$header] -= 1
                    $skipCell = $true
                }
                if ($colspanCount -gt 0)
                {
                    $colspanCount -= 1
                    $skipCell = $true
                }

                if ($skipCell) { continue }

                $cellClassList = [System.Collections.Generic.List[string]]::new($rowClassList)
                $cellStyleList = [System.Collections.Generic.List[string]]::new($rowStyleList)

                foreach ($value in $cellClassDict[$header]) { $cellClassList.Add($value) }
                foreach ($value in $cellStyleDict[$header]) { $cellStyleList.Add($value) }

                $cellValue = "$($object.$header)"

                if ($header -notin $HtmlProperty -and !$AutoDetectHtml.IsPresent -and ($cellValue.Length -eq 0 -or $cellValue.Substring(0,1) -ne '<'))
                {
                    $cellValue = [System.Web.HttpUtility]::HtmlEncode($cellValue).Replace("`r`n", '<br />')
                }

                if ($header -in $RightAlignProperty) { $cellClassList.Add('ralign') }
                if ($header -in $NoWrapProperty) { $cellClassList.Add('nowrap') }

                $classHtml = ''
                $styleHtml = ''
                if ($cellClassList) { $classHtml = " class='$($cellClassList -join ' ')'" }
                if ($cellStyleList) { $styleHtml = " style='$($cellStyleList -join ' ')'" }
                
                $setColspan = $false
                $colspanHtml = ''
                if ($CellColspanScripts[$header])
                {
                    $colspanCount = $CellColspanScripts[$header]
                    if ($colspanCount -is [scriptblock]) { $colspanCount = $object | ForEach-Object $colspanCount | Select-Object -First 1 }
                    else { $colspanCount = [int]$object.$colspanCount }
                    if ($colspanCount -gt 1)
                    {
                        $colspanHtml = " colspan='$colspanCount'"
                        $colspanCount -= 1
                        $setColspan = $true
                    }
                    else
                    {
                        $colspanCount = 0
                    }
                }

                $setRowspan = $false
                $rowspanHtml = ''
                if ($CellRowspanScripts[$header])
                {
                    $rowspanCount = $CellRowspanScripts[$header]
                    if ($rowspanCount -is [scriptblock]) { $rowspanCount = $object | ForEach-Object $rowspanCount | Select-Object -First 1 }
                    else { $rowspanCount = [int]$object.$rowspanCount }
                    if ($rowspanCount -gt 1)
                    {
                        $rowspanHtml = " rowspan='$rowspanCount'"
                        $rowspanCountHash[$header] = $rowspanCount - 1
                        $setRowspan = $true
                    }
                }

                if ($setRowspan -and $setColspan)
                {
                    $colspanRowHash[$header] = $rowspanCount - 1
                    $colspanCountHash[$header] = $colspanCount + 1
                }

                $attrDcn = if ($AddDataColumnName) { " data-column-name='$([System.Web.HttpUtility]::HtmlAttributeEncode($header))'" }
                $resultList.Add("<td$colspanHtml$rowspanHtml$classHtml$styleHtml$attrDcn>$cellValue</td>")
            }

            $resultList.Add('</tr>')
        }       

        if (-not $RowsOnly.IsPresent -and $inputObjectList.Count)
        {
            $resultList.Add("</tbody>")
            $resultList.Add("</table>")
        }

        $resultList -join "`r`n"
    }
}
}

Function ConvertTo-HtmlStrongText
{
    Param
    (
        [Parameter(ValueFromPipeline=$true, Position=0)] [object] $InputObject,
        [Parameter()] [ValidateSet('SameLine', 'MultiLine', 'MultiLineIndent', 'MultiLineHorizontal', 'MultiLineIndentHorizontal')] [string] $Mode = 'MultiLine',
        [Parameter()] [switch] $NoEmptyValues,
        [Parameter()] [string[]] $Property,
        [Parameter()] [string[]] $ExcludeProperty,
        [Parameter()] [string[]] $HtmlProperty,
        [Parameter()] [string[]] $ExcludeHtmlProperty,
        [Parameter()] [switch] $AutoHtmlProperty
    )
    Process
    {
        if (!$InputObject) { return }
        $propertyHash = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLikeHashtable($InputObject.PSObject.Properties.Name, $Property, $ExcludeProperty)
        $htmlPropertyHash = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLikeHashtable($InputObject.PSObject.Properties.Name, $HtmlProperty, $ExcludeHtmlProperty, $false)
        $blockList = foreach ($psProperty in $InputObject.PSObject.Properties)
        {
            if (!$propertyHash[$psProperty.Name]) { continue }

            $value = $psProperty.Value
            if ($NoEmptyValues.IsPresent -and [String]::IsNullOrWhiteSpace($value)) { continue }
            if (!(($AutoHtmlProperty.IsPresent -and $value.Length -and $value.Substring(0,1) -eq '<') -or $htmlPropertyHash[$psProperty.Name]))
            {
                $value = Get-HtmlEncodedText $value -InsertBr
            }

            $name = Get-HtmlEncodedText $psProperty.Name

            if ($Mode -eq "SameLine")
            {
                $block = "<strong>${name}:</strong> $value"
            }
            elseif ($Mode -in "MultiLine", "MultiLineHorizontal")
            {
                $block = "<p><strong>$name</strong><br />$value</p>"
            }
            elseif ($Mode -in "MultiLineIndent", "MultiLineIndentHorizontal")
            {
                $block = "<p><strong>$name</strong><br /><span style='margin-left:1em;'>$value</span></p>"
            }

            if ($Mode -in "MultiLineHorizontal", "MultiLineIndentHorizontal")
            {
                "<div style='display:inline-block; margin-right:2em;'>$block</div>"
            }
            else { $block }
        }

        if ($Mode -eq "SameLine") { return "<p>$($blockList -join "<br />")</p>" }
        $blockList -join "`r`n"
    }
}

Function ConvertTo-HtmlHourlyHeatmap
{
    [CmdletBinding(PositionalBinding=$false)]
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter()] [string[]] $SetProperty,
        [Parameter()] [string] $KeyJoin = '|',
        [Parameter()] [string] $TimestampProperty = 'Timestamp',
        [Parameter()] [string] $ValueProperty = 'Value',
        [Parameter(Mandatory=$true)] [object[]] $HeatmapColors,
        [Parameter()] [int] $IndicatorSize = 10,
        [Parameter()] [int] $IndicatorPadding = 1,
        [Parameter()] [string[]] $Class = 'HtmlReportingTable',
        [Parameter()] [datetime] $StartDate,
        [Parameter()] [datetime] $EndDate,
        [Parameter()] [string] $DateHeaderFormat = 'M/d',
        [Parameter()] [string] $TooltipDateFormat = 'MM/dd @ h:mm tt',
        [Parameter()] [ValidateSet(1,2,3,4,6,8,12,24)] [int] $Columns = 3
    )
    Begin
    {
        $timestampDict = @{}
        $rawTimestampDict = @{}
        $setKeyValueDict = [ordered]@{}
    }
    Process
    {
        $timestamp = [datetime]$InputObject.$TimestampProperty
        if (!$timestamp) { Write-Error "InputObject does not have a $TimestampProperty value!"; return }
        $timestamp = $timestamp.Date.AddHours($timestamp.Hour)
        if ($SetProperty)
        {
            $setKey = $(foreach ($prop in $SetProperty) { $InputObject.$prop }) -join $KeyJoin
            $setKeyValueDict[$setKey] = $InputObject
            $fullKey = $setKey, $timestamp -join $KeyJoin
            $timestampDict[$fullKey] = $InputObject.$ValueProperty
        }
        else
        {
            $timestampDict[$timestamp] = $InputObject.$ValueProperty
        }
        $rawTimestampDict[$timestamp] = $true
    }
    End
    {
        trap { $PSCmdlet.ThrowTerminatingError($_) }
        $minMax = $rawTimestampDict.Keys | Measure-Object -Minimum -Maximum
        if (!$StartDate) { $StartDate = $minMax.Minimum }
        if (!$EndDate) { $EndDate = $minMax.Maximum }
        if ($EndDate -lt $StartDate) { throw "EndDate must be greater than StartDate." }

        $dateCount = ($EndDate.Date - $StartDate.Date).TotalDays
        $dateList = 0..$dateCount | ForEach-Object { $StartDate.Date.AddDays($_) }
        $classCss = if ($Class) { " class='$($Class -join ' ')'" }

        $heatmapColorList = foreach ($heatmapColor in $HeatmapColors)
        {
            if ($heatmapColor -is [scriptblock]) { & $heatmapColor }
            else { $heatmapColor }
        }

        "<table$classCss>"
        "<thead>"
        "<tr>"
        if ($SetProperty)
        {
            foreach ($prop in $SetProperty)
            {
                "<th>$prop</th>"
            }
        }
        foreach ($date in $dateList)
        {
            "<th>$($date.ToString($DateHeaderFormat))</th>"
        }
        "</tr>"
        "</thead>"
        "<tbody>"
        
        $rowCount = 24 / $Columns
        $width = $Columns * $IndicatorSize + ($Columns - 1) * $IndicatorPadding
        $height = $rowCount * $IndicatorSize + ($rowCount - 1) * $IndicatorPadding

        if (!$SetProperty)
        {
            $setKeyValueDict['NA'] = 'NA'
        }

        foreach ($setKey in $setKeyValueDict.Keys)
        {
            $setObject = $setKeyValueDict[$setKey]
            "<tr>"
            if ($SetProperty)
            {
                foreach ($prop in $SetProperty)
                {
                    "<td>$($setObject.$prop)</td>"
                }
            }
            foreach ($date in $dateList)
            {
                "<td>"
                "<svg width='$width' height='$height'>"
                $y = 0
                foreach ($row in 1..$rowCount)
                {
                    $x = 0
                    foreach ($col in 1..$Columns)
                    {
                        $hour = ($row - 1) * $Columns + $col - 1
                        $timestamp = $date.AddHours($hour)
                        $key = $timestamp
                        if ($SetProperty) { $key = $setKey, $timestamp -join $KeyJoin }
                        $value = $timestampDict[$key]
                        $fill = 'transparent'
                        if ($value -eq $null)
                        {
                            $value = 'No Value'
                        }
                        else
                        {
                            foreach ($heatmapColor in $heatmapColorList)
                            {
                                if ($heatmapColor.Mode -eq 'Default' -or
                                    ($heatmapColor.Mode -eq 'EqualTo' -and $value -eq $heatmapColor.Value) -or
                                    ($heatmapColor.Mode -eq 'AtLeast' -and $value -ge $heatmapColor.Value)
                                )
                                {
                                    $fill = $heatmapColor.ColorCss
                                    break
                                }
                            }
                        }
                        "<rect x='$x' y='$y' width='$IndicatorSize' height='$IndicatorSize' style='fill:$fill;'>"
                        "<title>$($timestamp.ToString($TooltipDateFormat)) : $([System.Web.HttpUtility]::HtmlEncode($value))</title>"
                        "</rect>"
                        $x = $x + $IndicatorSize + $IndicatorPadding
                    }
                    $y = $y + $IndicatorSize + $IndicatorPadding
                }
                "</td>"
                "</svg>"
            }
            "</tr>"
        }
        "</tbody>"
        "</table>"
    }
}

Function ConvertTo-HtmlMonthlyHeatmap
{
    [CmdletBinding(PositionalBinding=$false)]
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter()] [string[]] $SetProperty,
        [Parameter()] [string] $KeyJoin = '|',
        [Parameter()] [string] $TimestampProperty = 'Timestamp',
        [Parameter()] [string] $ValueProperty = 'Value',
        [Parameter(Mandatory=$true)] [object[]] $HeatmapColors,
        [Parameter()] [int] $IndicatorSize = 10,
        [Parameter()] [int] $IndicatorPadding = 1,
        [Parameter()] [string[]] $Class = 'HtmlReportingTable',
        [Parameter()] [datetime] $StartDate,
        [Parameter()] [datetime] $EndDate,
        [Parameter()] [string] $DateHeaderFormat = 'MMM yyyy',
        [Parameter()] [string] $TooltipDateFormat = 'MM/dd'
    )
    Begin
    {
        $timestampDict = @{}
        $rawTimestampDict = @{}
        $setKeyValueDict = [ordered]@{}
    }
    Process
    {
        $timestamp = [datetime]$InputObject.$TimestampProperty
        if (!$timestamp) { Write-Error "InputObject does not have a $TimestampProperty value!"; return }
        $timestamp = $timestamp.Date
        if ($SetProperty)
        {
            $setKey = $(foreach ($prop in $SetProperty) { $InputObject.$prop }) -join $KeyJoin
            $setKeyValueDict[$setKey] = $InputObject
            $fullKey = $setKey, $timestamp -join $KeyJoin
            $timestampDict[$fullKey] = $InputObject.$ValueProperty
        }
        else
        {
            $timestampDict[$timestamp] = $InputObject.$ValueProperty
        }
        $rawTimestampDict[$timestamp] = $true
    }
    End
    {
        trap { $PSCmdlet.ThrowTerminatingError($_) }
        $minMax = $rawTimestampDict.Keys | Measure-Object -Minimum -Maximum
        if (!$StartDate) { $StartDate = $minMax.Minimum }
        if (!$EndDate) { $EndDate = $minMax.Maximum }
        if ($EndDate -lt $StartDate) { throw "EndDate must be greater than StartDate." }

        $tempDate = $StartDate.AddDays(1-$StartDate.Day)
        $monthList = while ($tempDate -le $EndDate) { $tempDate; $tempDate = $tempDate.AddMonths(1) }
        $monthCount = $monthList.Count

        $classCss = if ($Class) { " class='$($Class -join ' ')'" }

        $heatmapColorList = foreach ($heatmapColor in $HeatmapColors)
        {
            if ($heatmapColor -is [scriptblock]) { & $heatmapColor }
            else { $heatmapColor }
        }

        "<table$classCss>"
        "<thead>"
        "<tr>"
        if ($SetProperty)
        {
            foreach ($prop in $SetProperty)
            {
                "<th>$prop</th>"
            }
        }
        foreach ($month in $monthList)
        {
            "<th>$($month.ToString($DateHeaderFormat))</th>"
        }
        "</tr>"
        "</thead>"
        "<tbody>"
        
        $rowCount = 6
        $width = 7 * $IndicatorSize + 6 * $IndicatorPadding
        $height = $rowCount * $IndicatorSize + ($rowCount - 1) * $IndicatorPadding

        if (!$SetProperty)
        {
            $setKeyValueDict['NA'] = 'NA'
        }

        foreach ($setKey in $setKeyValueDict.Keys)
        {
            $setObject = $setKeyValueDict[$setKey]
            "<tr>"
            if ($SetProperty)
            {
                foreach ($prop in $SetProperty)
                {
                    "<td>$($setObject.$prop)</td>"
                }
            }
            foreach ($month in $monthList)
            {
                "<td>"
                "<svg width='$width' height='$height'>"
                $y = 0
                $endOfMonth = $month.AddMonths(1)
                $timestamp = $month
                $dow = $timestamp.DayOfWeek.value__
                foreach ($row in 1..$rowCount)
                {
                    $x = $dow * $IndicatorSize + ($dow - 1) * $IndicatorPadding
                    while ($dow -le 6)
                    {
                        $key = $timestamp
                        if ($SetProperty) { $key = $setKey, $timestamp -join $KeyJoin }
                        $value = $timestampDict[$key]
                        $fill = 'transparent'
                        if ($timestamp -ge $StartDate)
                        {
                            if ($value -eq $null)
                            {
                                $value = 'No Value'
                            }
                            else
                            {
                                foreach ($heatmapColor in $heatmapColorList)
                                {
                                    if ($heatmapColor.Mode -eq 'Default' -or
                                        ($heatmapColor.Mode -eq 'EqualTo' -and $value -eq $heatmapColor.Value) -or
                                        ($heatmapColor.Mode -eq 'AtLeast' -and $value -ge $heatmapColor.Value)
                                    )
                                    {
                                        $fill = $heatmapColor.ColorCss
                                        break
                                    }
                                }
                            }
                            "<rect x='$x' y='$y' width='$IndicatorSize' height='$IndicatorSize' style='fill:$fill;'>"
                            "<title>$($timestamp.ToString($TooltipDateFormat)) : $([System.Web.HttpUtility]::HtmlEncode($value))</title>"
                            "</rect>"
                        }
                        $x = $x + $IndicatorSize + $IndicatorPadding
                        $dow += 1
                        $timestamp = $timestamp.AddDays(1)
                        if ($timestamp -ge $endOfMonth) { break }
                    }
                    $dow = 0
                    if ($timestamp -ge $endOfMonth) { break }
                    $y = $y + $IndicatorSize + $IndicatorPadding
                }
                "</td>"
                "</svg>"
            }
            "</tr>"
        }
        "</tbody>"
        "</table>"
    }
}

Function Define-HtmlHeatmapColor
{
    [CmdletBinding(PositionalBinding=$false,DefaultParameterSetName='ColorName')]
    Param
    (
        [Parameter(Position=0,Mandatory=$true,ParameterSetName='ColorName')]
            [ValidateSet('Blue', 'Orange', 'Red', 'Green', 'Purple', 'SeaBlue', 'SkyBlue',
                'Teal', 'DarkGreen', 'LightOrange', 'Salmon', 'DarkRed', 'LightPurple',
                'Brown', 'Tan')]
            [string] $ColorName,
        [Parameter(Mandatory=$true,ParameterSetName='ColorRgb')] [int[]] $ColorRgb,
        [Parameter(Position=1)] [ValidateSet('EqualTo', 'AtLeast', 'Default')] [string] $Mode = 'Default',
        [Parameter(Position=2)] [object] $Value
    )
    End
    {
        trap { $PSCmdlet.ThrowTerminatingError($_) }

        if ($ColorName) { $ColorRgb = Get-HtmlReportColor -Name $ColorName }
        
        $result = [ordered]@{}
        $result.ColorCss = "rgb($($ColorRgb -join ','))"
        $result.Mode = $Mode
        $result.Value = $Value
        [pscustomobject]$result
    }
}

Function Expand-XmlText
{
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $XmlObject
    )
    Process
    {
        if ($_.NodeType -eq 'Text')
        {
            $_.Value.Trim()
        }
        elseif ($_.NodeType -in 'Element', 'Document')
        {
            $_.ChildNodes | Expand-XmlText
        }
    }
}

Function ConvertFrom-HtmlTable
{
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] [string] $TableHtml,
        [Parameter(Position=0)] [string] $PreMatch,
        [Parameter()] [string[]] $Headers,
        [Parameter()] [string[]] $LinkColumn,
        [Parameter()] [string[]] $HtmlColumn
    )
    Process
    {
        $newTableHtml = $TableHtml -replace "(?s).*?$PreMatch.*?<table[^>]*>(.+?)</table>.*",'<table>$1</table>'
        $newTableHtml = $newTableHtml -replace '&nbsp;'
        $tableXml = [xml]$newTableHtml
        $rows = $tableXml.SelectNodes('//tr') | Where-Object { $_.InnerXml -notlike '*colspan*' }

        $selectArgs = @{}

        if (-not $Headers)
        {
            $selectArgs.Skip = 1

            Remove-Variable Headers
            $headers = [System.Collections.Generic.List[object]]::new()

            $rows[0].ChildNodes | ForEach-Object {
                $header = $_.'#text'.Trim()
                $headers.Add($header)
            }
        }

        $dataTemplate = [ordered]@{}
        $headers | ForEach-Object { $dataTemplate.$_ = $null }
        $dataTemplate = [pscustomobject]$dataTemplate

        $rows | Select-Object @selectArgs | ForEach-Object {
            $newData = $dataTemplate | Select-Object *
            $i = 0
            $_.ChildNodes | ForEach-Object {
                $header = $headers[$i]
                if ($header -in $HtmlColumn)
                {
                    $newData.$header = $_.InnerXml
                }
                elseif ($header -in $LinkColumn)
                {
                    if ($_.OuterXml -match "href=['""](.+?)['""]")
                    {
                        $newData.$header = $Matches[1]
                    }
                }
                else
                {
                    $textList = $_ | Expand-XmlText
                    $newData.$header = $textList -join ' '
                }
                $i += 1
            }
            $newData
        }
    }
}

Function Get-HtmlEncodedText
{
    Param
    (
        [Parameter(Position=0,ValueFromPipeline=$true)] [string] $Text,
        [Parameter()] [switch] $KeepWhitespace,
        [Parameter()] [switch] $InsertBr
    )
    Process
    {
        $Text = [System.Web.HttpUtility]::HtmlEncode($Text)
        if ($KeepWhitespace) { $Text = $Text.Replace(' ', '&nbsp;') }
        if ($InsertBr) { $Text = $Text.Replace("`r`n", "<br />") }
        $Text
    }
}

Function Get-HtmlReportColor
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0, ParameterSetName='ByIndex')] [int] $Index,
        [Parameter(Mandatory=$true, Position=0, ParameterSetName='ByColor')]
            [ValidateSet('Blue', 'Orange', 'Red', 'Green', 'Purple', 'SeaBlue', 'SkyBlue',
                'Teal', 'DarkGreen', 'LightOrange', 'Salmon', 'DarkRed', 'LightPurple',
                'Brown', 'Tan')]
            [string] $Name,
        [Parameter()] [switch] $AsCssRgb,
        [Parameter()] [double] $AsCssRgba
    )
    End
    {
        if ($Index -gt 14) { $Index = $Index % 15 }
        if (!$Script:ReportColorIndices)
        {
            $Script:ReportColorIndices = [Array]::CreateInstance([int[]], 15)
            $Script:ReportColorIndices[0] = 67, 134, 216
            $Script:ReportColorIndices[1] = 255, 154, 46
            $Script:ReportColorIndices[2] = 219, 68, 63
            $Script:ReportColorIndices[3] = 168, 212, 79
            $Script:ReportColorIndices[4] = 133, 96, 179
            $Script:ReportColorIndices[5] = 60, 191, 227
            $Script:ReportColorIndices[6] = 175, 216, 248
            $Script:ReportColorIndices[7] = 0, 142, 142
            $Script:ReportColorIndices[8] = 139, 186, 0
            $Script:ReportColorIndices[9] = 250, 189, 15
            $Script:ReportColorIndices[10] = 250, 110, 70
            $Script:ReportColorIndices[11] = 157, 8, 13
            $Script:ReportColorIndices[12] = 161, 134, 190
            $Script:ReportColorIndices[13] = 204, 102, 0
            $Script:ReportColorIndices[14] = 253, 198, 137

            $Script:ReportColorNameToIndex = @{}
            $i = 0
            foreach ($color in 'Blue', 'Orange', 'Red', 'Green', 'Purple', 'SeaBlue', 'SkyBlue',
                'Teal', 'DarkGreen', 'LightOrange', 'Salmon', 'DarkRed', 'LightPurple',
                'Brown', 'Tan')
            {
                $Script:ReportColorNameToIndex[$color] = $i++
            }
        }
        if ($Name) { $Index = $Script:ReportColorNameToIndex[$Name] }
        $rgb = $Script:ReportColorIndices[$Index]
        if ($AsCssRgb) { "rgb($($rgb -join ','))" }
        elseif ($AsCssRgba) { "rgba($($rgb -join ','),$AsCssRgba)" }
        else { $rgb }
    }
}

Function Get-HtmlImgTag
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0, ParameterSetName='FilePath')] [string] $FilePath,
        [Parameter(Mandatory=$true, ParameterSetName='Bytes')] [byte[]] $Bytes
    )
    End
    {
        if ($FilePath)
        {
            $extension = [System.IO.Path]::GetExtension($FilePath)
            $Bytes = [System.IO.File]::ReadAllBytes($FilePath)
        }
        $base64 = [Convert]::ToBase64String($Bytes)
        "<img srv='data:image/$extension;base64,$base64' />"
    }
}

Function Get-HtmlIndicatorText
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)] [string] $Text,
        [Parameter()] [switch] $TextIsHtml,
        [Parameter(Mandatory=$true, ParameterSetName='ByRGB')] [ValidateCount(3,3)] [int[]] $ColorRGB,
        [Parameter(Mandatory=$true, ParameterSetName='ByIndex')] [int] $ColorIndex,
        [Parameter(Mandatory=$true, ParameterSetName='ByColor')]
            [ValidateSet('Blue', 'Orange', 'Red', 'Green', 'Purple', 'SeaBlue', 'SkyBlue',
                'Teal', 'DarkGreen', 'LightOrange', 'Salmon', 'DarkRed', 'LightPurple',
                'Brown', 'Tan')]
            [string] $ColorName,
        [Parameter()] [switch] $BorderOnly,
        [Parameter()] [string] $Width,
        [Parameter()] [string] $Title,
        [Parameter()] [string[]] $Class
    )
    End
    {
        if ($ColorName) { $ColorRGB = Get-HtmlReportColor -Name $ColorName }
        elseif ($PSCmdlet.ParameterSetName -eq 'ByIndex') { $ColorRGB = Get-HtmlReportColor -Index $ColorIndex }
        $colorCss = "rgb($($ColorRGB -join ','))"
        $widthCss = ''
        if ($Width) { $widthCss = "width:$Width;text-align:center;" }
        if (!$TextIsHtml) { $Text = Get-HtmlEncodedText $Text -InsertBr }
        $titleHtml = if ($Title) { " title='$title'" }
        $classCss = if ($Class) { " $($Class -join ' ')" }
        if ($BorderOnly)
        {
            "<span class='IndicatorTextBorder$classCss' style='border-color:$colorCss;$widthCss'$titleHtml>$Text</span>"
        }
        else
        {
            "<span class='IndicatorText$classCss' style='background:$colorCss;$widthCss'$titleHtml>$Text</span>"
        }
    }
}

Function Get-HtmlSelect
{
    [CmdletBinding(PositionalBinding=$false,DefaultParameterSetName='Values')]
    Param
    (
        [Parameter(ParameterSetName='Pipeline',ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter(ParameterSetName='Pipeline',Mandatory=$true)] [string] $ValueProperty,
        [Parameter(ParameterSetName='Pipeline')] [string] $LabelProperty,
        [Parameter(ParameterSetName='Pipeline')] [string] $IsSelectedProperty,
        [Parameter(ParameterSetName='Values',Mandatory=$true)] [string[]] $Values,
        [Parameter()] [string[]] $SelectedValues,
        [Parameter(ParameterSetName='Dictionary',Mandatory=$true)] [System.Collections.IDictionary] $ValueLabelDictionary,
        [Parameter()] [switch] $Required,
        [Parameter()] [switch] $Multiple,
        [Parameter()] [string] $Id,
        [Parameter()] [string] $Name,
        [Parameter()] [string[]] $Class,
        [Parameter()] [string[]] $Style,
        [Parameter()] [int] $Size
    )
    Begin
    {
        trap { $PSCmdlet.ThrowTerminatingError($_) }
        if ($IsSelectedProperty -and $SelectedValues) { throw "IsSelectedProperty and SelectedValues can't be provided together." }
        $valueLabelDict = if ($PSCmdlet.ParameterSetName -eq 'Dictionary') { $ValueLabelDictionary } else { [ordered]@{} }
        $selectedValueDict = @{}

        if ($PSCmdlet.ParameterSetName -eq 'Values')
        {
            foreach ($v in $Values) { $valueLabelDict[$v] = $v }
        }
        if ($SelectedValues)
        {
            foreach ($v in $SelectedValues) { $selectedValueDict[$v] = $true }
        }
    }
    Process
    {
        if (!$InputObject -or $PSCmdlet.ParameterSetName -ne 'Pipeline') { return }
        $value = [string]$InputObject.$ValueProperty
        $label = if ($LabelProperty) { $InputObject.$LabelProperty } else { $value }
        if ($IsSelectedProperty -and $InputObject.$IsSelectedProperty) { $selectedValueDict[$value] = $true }
        $valueLabelDict[$value] = $label
    }
    End
    {
        $result = [System.Text.StringBuilder]::new()
        [void]$result.Append("<select")
        if ($Required) { [void]$result.Append(" required='required'") }
        if ($Multiple) { [void]$result.Append(" multiple='multiple'") }
        if ($Id) { [void]$result.AppendFormat(" id='{0}'", [System.Web.HttpUtility]::HtmlAttributeEncode($Id)) }
        if ($Name) { [void]$result.AppendFormat(" name='{0}'", [System.Web.HttpUtility]::HtmlAttributeEncode($Name)) }
        if ($Class) { [void]$result.AppendFormat(" class='{0}'", [System.Web.HttpUtility]::HtmlAttributeEncode($Class -join ' ')) }
        if ($Style) { [void]$result.AppendFormat(" style='{0}'", $Style.Replace("'", "''") -join ' ') }
        if ($PSBoundParameters.ContainsKey('Size') -and $Size -eq 0) { $Size = @($valueLabelDict.GetEnumerator()).Count }
        if ($Size) { [void]$result.Append(" size='$Size'") }
        [void]$result.Append(">")
        foreach ($pair in $valueLabelDict.GetEnumerator()) {
            $value = [string]$pair.Key
            $label = $pair.Value
            $selectedAttr = if ($selectedValueDict[$value]) { " selected='selected'" }
            [void]$result.Append("<option value='$([System.Web.HttpUtility]::HtmlAttributeEncode($value))'$selectedAttr>$([System.Web.HttpUtility]::HtmlEncode($label))</option>")
        }
        [void]$result.Append("</select>")
        $result.ToString()
    }
}

Function Convert-PSCodeToHtml
{
    [CmdletBinding()]
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [string] $PsCode
    )
    Begin
    {
        $tokenColors = @{
            'Attribute' = '#FFADD8E6'
            'Command' = '#FF0000FF'
            'CommandArgument' = '#FF8A2BE2'
            'CommandParameter' = '#FF000080'
            'Comment' = '#FF006400'
            'GroupEnd' = '#FF000000'
            'GroupStart' = '#FF000000'
            'Keyword' = '#FF00008B'
            'LineContinuation' = '#FF000000'
            'LoopLabel' = '#FF00008B'
            'Member' = '#FF000000'
            'NewLine' = '#FF000000'
            'Number' = '#FF800080'
            'Operator' = '#FFA9A9A9'
            'Position' = '#FF000000'
            'StatementSeparator' = '#FF000000'
            'String' = '#FF8B0000'
            'Type' = '#FF008080'
            'Unknown' = '#FF000000'
            'Variable' = '#FFFF4500'
        }

        [void][System.Reflection.Assembly]::Load("System.Web, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")

        $psCodeList = [System.Collections.Generic.List[string]]::new()
    }
    Process
    {
        $psCodeList.Add($PsCode)
    }
    End
    {
        $unformattedText = $psCodeList -join "`r`n"

        $currentLine = 1
        function Append-HtmlSpan ($block, $tokenColor)
        {
           if (($tokenColor -eq 'NewLine') -or ($tokenColor -eq 'LineContinuation'))
           {
              if($tokenColor -eq 'LineContinuation')
              {
                 $null = $codeBuilder.Append('`')
              }

              $null = $codeBuilder.Append("<br />`r`n")
           }
           else
           {
              $block = [System.Web.HttpUtility]::HtmlEncode($block)
              if (-not $block.Trim())
              {
                 $block = $block.Replace(' ', '&nbsp;')
              }

              $htmlColor = $tokenColors[$tokenColor].ToString().Replace('#FF', '#')

              if($tokenColor -eq 'String')
              {
                 $lines = $block -split "`r`n"
                 $block = ""

                 $multipleLines = $false
                 foreach($line in $lines)
                 {
                    if($multipleLines)
                    {
                       $block += "<BR />`r`n"
                    }

                    $newText = $line.TrimStart()
                    $newText = "&nbsp;" * ($line.Length - $newText.Length) + $newText
                    $block += $newText
                    $multipleLines = $true
                 }
              }

              $null = $codeBuilder.Append("<span style='color:$htmlColor'>$block</span>")
           }
        }

        trap { break }

        # Do syntax parsing.
        $errors = $null
        $tokens = [System.Management.Automation.PSParser]::Tokenize($unformattedText, [ref] $errors)

        # Initialize HTML builder.
        $codeBuilder = [System.Text.StringBuilder]::new()

        # Iterate over the tokens and set the colors appropriately.
        $position = 0
        foreach ($token in $tokens)
        {
            if ($position -lt $token.Start)
            {
                $block = $unformattedText.Substring($position, ($token.Start - $position))
                $tokenColor = 'Unknown'
                Append-HtmlSpan $block $tokenColor
            }

            $block = $unformattedText.Substring($token.Start, $token.Length)
            $tokenColor = $token.Type.ToString()
            Append-HtmlSpan $block $tokenColor

            $position = $token.Start + $token.Length
        }

        # Build the entire syntax-highlighted script
        $code = $codeBuilder.ToString()
   
        # Replace tabs with three blanks
        $code = $code.Replace("`t","&nbsp;&nbsp;&nbsp;")

        "<div style='font-family: ""Lucida Console"", Consolas; font-size: 9pt;'>$code</div>"
    }
}

Function GenerateHtmlTagFunctions
{
    $private:functionHash = @{}
        
    $private:other = @{}
    $other['a'] = 'name', 'href', 'target', 'onclick'
    $other['form'] = 'action', 'method'
    $other['input'] = 'type', 'name', 'value'
    $other['option'] = 'value'
    $other['textarea'] = 'name', 'rows', 'cols'
    $other['button'] = 'type', 'onclick', 'name', 'value'
    $other['label'] = 'for'
    $other['span'] = 'title'

    foreach ($private:t in 'html', 'h1', 'h2', 'h3', 'h4', 'ol', 'ul', 'li', 'p', 'span', 'div', 'strong', 'em', 'a',
        'form', 'input', 'button', 'option', 'textarea', 'button', 'label')
    {
        $functionHash[$t] = [ScriptBlock]::Create("
        [CmdletBinding(PositionalBinding=`$false)]
        Param
        (
            $(if ($t -eq 'html') {'[Parameter(Position=0,Mandatory=$true)] [string] $tag,'})
            [Parameter(ValueFromRemainingArguments=`$true)] [object[]] `$Definition,
            [Parameter()] [string[]] `$HtmlEncode,
            [Parameter()] [string[]] `$class,
            [Parameter()] [string] `$id,
            [Parameter()] [string[]] `$style,
            [Parameter()] [hashtable] `$Attributes$(foreach ($p in $other[$t]) { ",`r`n[Parameter()] [string] `$$p" } )
        )
        $(if ($t -ne 'html') { "`$private:tag = '$t'" })
        `$otherList = @()
        if (`$class) { `$otherList += ""class='`$(`$class -join ' ')'"" }
        if (`$style) { `$otherList += ""style='`$(`$style -join ' ')'"" }
        if (`$id) { `$otherList += ""id='`$id'"" }
        if (`$Attributes) { foreach (`$key in `$Attributes.Keys) { `$otherList += `"`$key='`$(`$Attributes[`$key])'`" } }
        $(foreach ($p in $other[$t]) { "if (`$$p) { `$otherList += """"$($p.ToLower())='`$$p'"""" }`r`n" } )
        `$otherCode = ''
        if (`$otherList) { `$otherCode = "" `$(`$otherList -join ' ')"" }
        `$text = @(
            foreach (`$item in `$Definition)
            {
                if (`$item -is [scriptblock]) { & `$item } else { `$item }
            }
            foreach (`$item in `$HtmlEncode) { [System.Web.HttpUtility]::HtmlEncode(`$item) }
        )
        ""<`$tag`$otherCode>"", (`$text -join ' '), ""</`$tag>"" -join ''")
    }
    $functionHash['br'] = { param($clear) "<br$(if ($clear) { " clear='$clear'" }) />" }
    $functionHash['hr'] = { "<hr />" }
    $functionHash['script'] = { param($private:script) "<script>$private:script</script>" }
    $functionHash['style'] = { param($private:style) "<style>$private:style</style>" }

    $Script:HtmlTagFunctionHash = $functionHash
}

Function Import-HtmlTagFunctions
{
    End
    {
        if (Get-Module -Name HtmlReportingTagFunctions) { return }
        if (!$Script:HtmlTagFunctionHash) { GenerateHtmlTagFunctions }

        $moduleText = foreach ($function in $Script:HtmlTagFunctionHash.Keys)
        {
            "Function $function { $($Script:HtmlTagFunctionHash[$function]) }"
        }

        New-Module -ScriptBlock ([ScriptBlock]::Create($moduleText -join "`r`n")) -Name HtmlReportingTagFunctions |
            Import-Module -Scope Global
    }
}

Function Get-HtmlFragment
{
    Param
    (
        [Parameter(Position=0)] [scriptblock] $Definiton
    )
    End
    {
        if (!$Script:HtmlTagFunctionHash) { GenerateHtmlTagFunctions }

        $Definiton.InvokeWithContext($Script:HtmlTagFunctionHash, $null, $null) -join ''
    }
}

Function Get-HtmlFullDocument
{
    Param
    (
        [Parameter(ValueFromPipeline=$true,Position=0)] [string] $Html,
        [Parameter()] [switch] $AddTimestamp
    )
    Begin
    {
        $lines = [System.Collections.Generic.List[string]]::new()
    }
    Process
    {
        $lines.Add($Html)
    }
    End
    {
        $htmlDoc = $lines -join "`r`n"
        $timestampHtml = ''

        if ($AddTimestamp.IsPresent)
        {
            $timestamp = [DateTime]::Now.ToString('g')
            $timestampHtml = "<div style='position: absolute; top: 5; right: 5;'>Created: $timestamp</div>"
        }

        "<html><head><style>$Script:HtmlStyle</style></head><body>$htmlDoc$timestampHtml</body></html>"
    }
}

Function Out-HtmlFile
{
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [string] $Html,
        [Parameter(Position=0)] [string] $FilePath,
        [Parameter()] [switch] $Open,
        [Parameter()] [switch] $AddTimestamp
    )
    Begin
    {
        $lines = [System.Collections.Generic.List[string]]::new()
    }
    Process
    {
        $lines.Add($Html)
    }
    End
    {
        $htmlText = Get-HtmlFullDocument -Html ($lines -join "`r`n") -AddTimestamp:$AddTimestamp

        if (-not $FilePath)
        {
            $tempDirectory = [System.IO.Path]::GetTempPath() + "HtmlReporting\"
            [void][System.IO.Directory]::CreateDirectory($tempDirectory)
            if (!$Script:CleanedUpTempFiles)
            {
                Get-ChildItem $tempDirectory |
                    Where-Object LastWriteTime -lt (Get-Date).AddDays(-2) |
                    Where-Object Name -Match "\A\d{4}.*\.html" |
                    ForEach-Object { [System.IO.File]::Delete($_.FullName) }
                $Script:CleanedUpTempFiles = $true
            }
            $FilePath = $tempDirectory + [DateTime]::Now.ToString("yyyy.MM.dd-HH.mm.ss-ffff") + ".html"
        }
        else
        {
            $FilePath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($FilePath)
        }

        [System.IO.File]::WriteAllLines($FilePath, @($htmlText))

        if ($Open -or !$PSBoundParameters['FilePath']) { & $FilePath }
    }
}

Function Out-Outlook
{
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [string] $Html,
        [Parameter()] [string[]] $To,
        [Parameter()] [string[]] $Cc,
        [Parameter()] [string] $Subject,
        [Parameter()] [switch] $SaveDraft,
        [Parameter()] [switch] $Send,
        [Parameter()] [switch] $AppendScript,
        [Parameter()] [string[]] $Attachment
    )
    Begin
    {
        $lines = [System.Collections.Generic.List[string]]::new()
    }
    Process
    {
        $lines.Add($Html)
    }
    End
    {
        trap { $PSCmdlet.ThrowTerminatingError($_) }

        if ($SaveDraft -and $Send) { throw "SaveDraft and Send cannot be provided together." }

        $outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
        $mail = $outlook.CreateItem(0)

        $html = Get-HtmlFullDocument -Html ($lines -join "`r`n")

        if ($AppendScript.IsPresent)
        {
            $scriptText = Get-PSCallStack |
                Select-Object -Last 1 |
                ForEach-Object { $_.InvocationInfo.MyCommand.Definition } |
                Out-String
            $scriptHtml = "", "", "# $('='*80)", "# Generating Script", "# $('='*80)", "", $scriptText |
                Convert-PSCodeToHtml

            $html = $html.Replace('</body>', "$scriptHtml</body>")
        }

        $imgRegex = [regex]"(<img.+?src=')data:image/png;base64,(.+?)(' ?(:?usemap='.+?')? ?/>)"
        $imgRegexMatchList = $imgRegex.Matches($html)

        $olByValue = [Microsoft.Office.Interop.Outlook.OlAttachmentType]::olByValue
        foreach ($imgRegexMatch in $imgRegexMatchList)
        {
            $imgBase64 = $imgRegexMatch.Captures[0].Groups[2].Value
            $imgBytes = [Convert]::FromBase64String($imgBase64)
            $tempFile = [System.IO.Path]::GetTempFileName()
            [System.IO.File]::Delete($tempFile)
            [System.IO.File]::WriteAllBytes("$tempFile.png", $imgBytes)
            $fileName = [System.IO.Path]::GetFileName("$tempFile.png")
            [void]$mail.Attachments.Add("$tempFile.png", $olByValue, 0)
            [System.IO.File]::Delete("$tempFile.png")
            $html = $imgRegex.Replace($html, "`$1cid:$fileName`$3", 1)
        }

        $mail.BodyFormat = 2
        $mail.HTMLBody = $html

        if ($Attachment)
        {
            $Attachment | ForEach-Object {
                [void]$mail.Attachments.Add($_)
            }
        }

        if ($To) { $mail.To = $To -join '; ' }
        if ($Cc) { $mail.CC = $Cc -join '; ' }
        if ($Subject) { $mail.Subject = $Subject }

        if ($Send)
        {
            $mail.Send()
        }
        elseif ($SaveDraft)
        {
            $mail.Save()
        }
        else
        {
            $mail.Display()
        }
    }
}
