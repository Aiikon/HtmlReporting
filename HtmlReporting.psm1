Add-Type -AssemblyName 'System.Web'

$Script:HtmlStyle = @"
h1, h2, h3, h4, p, li, td, th {
    color: black;
    font-size: 11pt;
    font-family: Calibri;
}
body, a {
    font-size: 11pt;
    font-family: Calibri;
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
"@

Function ConvertTo-HtmlColorBlocks
{
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter(Position=0)] [ScriptBlock] $OutputScript,
        [Parameter()] [string[]] $TocProperty,
        [Parameter()] [string[]] $HtmlTocProperty,
        [Parameter()] [string] $SectionProperty
    )
    Begin
    {
        $inputObjectList = New-Object System.Collections.Generic.List[Object]
    }
    Process
    {
        if ($InputObject) { $inputObjectList.Add($InputObject) }
    }
    End
    {
        if ($TocProperty)
        {
            $htmlPropertyList = & { $HtmlTocProperty; 'Link' } | Select-Object -Unique
            $idList = New-Object System.Collections.Generic.List[string]
            $inputObjectList |
                Select-Object $TocProperty |
                ForEach-Object {
                    $id = [guid]::NewGuid().ToString('n')
                    if ($SectionProperty) { $id = $InputObject.$SectionProperty }
                    $idList.Add($id)
                    $_.PSObject.Properties.Add([PSNoteProperty]::New('Link', "<a href='#$id'>Link</a>"))
                    $_
                } |
                ConvertTo-HtmlTable -HtmlProperty $htmlPropertyList
            "<br /><br />"
        }

        $i = 0
        foreach ($InputObject in $inputObjectList)
        {
            $html = $OutputScript.InvokeWithContext($null, (New-Object PSVariable "_", @($InputObject)), $null)
            $color = Get-HtmlReportColor -Index $i -AssCssRGB
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

Function ConvertTo-HtmlTable
{
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter()] [string[]] $Property,
        [Parameter()] [string[]] $HtmlProperty,
        [Parameter()] [switch] $RowsOnly,
        [Parameter()] [scriptblock] $RowClassScript,
        [Parameter()] [scriptblock] $RowStyleScript,
        [Parameter()] [hashtable] $CellClassScripts = @{},
        [Parameter()] [hashtable] $CellStyleScripts = @{},
        [Parameter()] [hashtable] $CellColspanScripts = @{},
        [Parameter()] [hashtable] $CellRowspanScripts = @{},
        [Parameter()] [string[]] $RightAlignProperty,
        [Parameter()] [string[]] $NoWrapProperty,
        [Parameter()] [switch] $Narrow
    )
    Begin
    {
        $inputObjectList = New-Object System.Collections.Generic.List[object]
    }
    Process
    {
        $inputObjectList.Add($InputObject)
    }
    End
    {
        $resultList = New-Object System.Collections.Generic.List[string]
        $headerList = New-Object System.Collections.Generic.List[string]
        if (-not $Property)
        {
            $Property = $inputObjectList[0].PSObject.Properties.Name
        }

        foreach ($p in $Property) { $headerList.Add($p) }

        if (-not $RowsOnly.IsPresent)
        {
            $narrowCss = if ($Narrow) { ' Narrow' }
            $resultList.Add("<table class='HtmlReportingTable$narrowCss'>")
            $resultList.Add("<thead>")
            $resultList.Add("<tr class='header'>")
            foreach ($header in $headerList)
            {
                $resultList.Add("<th>$header</th>")
            }
            $resultList.Add("</tr>")
            $resultList.Add("</thead>")
            $resultList.Add("<tbody>")
        }

        $rowspanCountHash = @{}

        foreach ($object in $inputObjectList)
        {
            $rowClassList = New-Object System.Collections.Generic.List[string]
            $rowStyleList = New-Object System.Collections.Generic.List[string]
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

            foreach ($header in $headerList)
            {
                $skipCell = $false
                if ($colspanCount -gt 0)
                {
                    $colspanCount -= 1
                    $skipCell = $true
                }
                if ($rowspanCountHash.$header -gt 0)
                {
                    $rowspanCountHash.$header -= 1
                    $skipCell = $true
                }
                if ($skipCell) { continue }

                $cellClassList = New-Object System.Collections.Generic.List[string] (,$rowClassList)
                $cellStyleList = New-Object System.Collections.Generic.List[string] (,$rowStyleList)

                $cellValue = $object.$header

                if ($CellClassScripts[$header])
                {
                    $object | ForEach-Object $CellClassScripts[$header] |
                        ForEach-Object { $cellClassList.Add($_) }
                }

                if ($CellStyleScripts[$header])
                {
                    $object | ForEach-Object $CellStyleScripts[$header] |
                        ForEach-Object { $cellStyleList.Add($_) }
                }

                if ($header -notin $HtmlProperty)
                {
                    $cellValue = [System.Web.HttpUtility]::HtmlEncode("$cellValue").Replace("`r`n", '<br />')
                }

                if ($header -in $RightAlignProperty) { $cellClassList.Add('ralign') }
                if ($header -in $NoWrapProperty) { $cellClassList.Add('nowrap') }

                $classHtml = ''
                $styleHtml = ''
                if ($cellClassList) { $classHtml = " class='$($cellClassList -join ' ')'" }
                if ($cellStyleList) { $styleHtml = " style='$($cellStyleList -join '; ')'" }
                
                $colspanHtml = ''
                if ($CellColspanScripts[$header])
                {
                    $colspanCount = $object | ForEach-Object $CellColspanScripts.$header | Select-Object -First 1
                    if ($colspanCount -gt 1)
                    {
                        $colspanHtml = " colspan='$colspanCount'"
                        $colspanCount -= 1
                    }
                    else
                    {
                        $colspanCount = 0
                    }
                }
                $rowspanHtml = ''
                if ($CellRowspanScripts[$header])
                {
                    $rowspanCount = $object | ForEach-Object $CellRowspanScripts.$header | Select-Object -First 1
                    if ($rowspanCount -gt 1)
                    {
                        $rowspanHtml = " rowspan='$rowspanCount'"
                        $rowspanCount -= 1
                        $rowspanCountHash.$header = $rowspanCount
                    }
                    else
                    {
                        $rowspanCount = 0
                    }
                }

                $resultList.Add("<td$colspanHtml$rowspanHtml$classHtml$styleHtml>$cellValue</td>")
            }

            $resultList.Add('</tr>')
        }       

        if (-not $RowsOnly.IsPresent)
        {
            $resultList.Add("</tbody>")
            $resultList.Add("</table>")
        }

        $resultList -join "`r`n"
    }
}

Function ConvertTo-HtmlStrongText
{
    Param
    (
        [Parameter(ValueFromPipeline=$true, Position=0)] [object] $InputObject,
        [Parameter()] [ValidateSet('SameLine', 'MultiLine', 'MultiLineIndent', 'MultiLineHorizontal', 'MultiLineIndentHorizontal')] [string] $Mode = 'MultiLine',
        [Parameter()] [switch] $NoEmptyValues,
        [Parameter()] [string[]] $HtmlProperty
    )
    Process
    {
        if (!$InputObject) { return }
        $blockList = foreach ($property in $InputObject.PSObject.Properties)
        {
            $name = Get-HtmlEncodedText $property.Name
            $value = $property.Value
            if ($NoEmptyValues.IsPresent -and [String]::IsNullOrWhiteSpace($value)) { continue }
            if ($name -notin $HtmlProperty) { $value = Get-HtmlEncodedText $value -InsertBr }
            
            if ($Mode -eq "SameLine")
            {
                $block = "<strong>${name}:</strong> $value"
            }
            elseif ($Mode -in "MultiLine", "MultiLineHorizontal")
            {
                $block = "<p><strong>$name</strong></br>$value</p>"
            }
            elseif ($Mode -in "MultiLineIndent", "MultiLineIndentHorizontal")
            {
                $block = "<p><strong>$name</strong></br><span style='margin-left:1em;'>$value</span></p>"
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
            $headers = New-Object System.Collections.Generic.List[object]

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
        [Parameter()] [switch] $AssCssRGB
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
        if ($AssCssRGB) { "rgb($($rgb -join ','))" }
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
        [Parameter()] [string] $Width
    )
    End
    {
        if ($ColorName) { $ColorRGB = Get-HtmlReportColor -Name $ColorName }
        elseif ($PSCmdlet.ParameterSetName -eq 'ByIndex') { $ColorRGB = Get-HtmlReportColor -Index $ColorIndex }
        $colorCss = "rgb($($ColorRGB -join ','))"
        $widthCss = ''
        if ($Width) { $widthCss = "width:$Width;text-align:center;" }
        if (!$TextIsHtml) { $Text = Get-HtmlEncodedText $Text -InsertBr }
        if ($BorderOnly)
        {
            "<span class='IndicatorTextBorder' style='border-color:$colorCss;$widthCss'>$Text</span>"
        }
        else
        {
            "<span class='IndicatorText' style='background:$colorCss;$widthCss'>$Text</span>"
        }
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

        $psCodeList = New-Object System.Collections.Generic.List[string]
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
        $codeBuilder = New-Object System.Text.StringBuilder

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

Function Get-HtmlFragment
{
    Param
    (
        [Parameter(Position=0)] [scriptblock] $Definiton
    )
    End
    {
        $functionHash = @{}
        
        $otherParameters = @{}
        $otherParameters['a'] = ",`r`n[Parameter()] [string] `$Name,`r`n[Parameter()] [string] `$Href"

        $otherLines = @{}
        $otherLines['a'] = "if (`$Name) { `$otherList += ""name='`$Name`'"" }", "if (`$Href) { `$otherList += ""href='`$Href`'"" }"

        foreach ($private:t in 'h1', 'h2', 'h3', 'h4', 'ol', 'ul', 'li', 'p', 'span', 'div', 'strong', 'em', 'a')
        {
            $functionHash[$t] = [ScriptBlock]::Create("
            [CmdletBinding(PositionalBinding=`$false)]
            Param
            (
                [Parameter(ValueFromRemainingArguments=`$true)] [object[]] `$Definition,
                [Parameter()] [string[]] `$Class,
                [Parameter()] [string[]] `$Style$($otherParameters[$t])
            )
            `$otherList = @()
            if (`$Class) { `$otherList += ""class='`$(`$Class -join ' ')'"" }
            if (`$Style) { `$otherList += ""style='`$(`$Style -join ' ')'"" }
            $($otherLines[$t] -join "`r`n")
            `$otherCode = ''
            if (`$otherList) { `$otherCode = "" `$(`$otherList -join ' ')"" }
            `$text = foreach (`$item in `$Definition)
            {
                if (`$item -is [scriptblock]) { & `$item } else { `$item }
            }
            ""<$t`$otherCode>"", (`$text -join ' '), ""</$t>`r`n"" -join ''")
        }
        $functionHash['br'] = { "<br />" }

        $Definiton.InvokeWithContext($functionHash, $null, $null) -join ''
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
        $lines = New-Object System.Collections.Generic.List[string]
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

        "<html><head><style>$Script:HtmlStyle</style></head><body>$htmlDoc$timestampHtml</body></head>"
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
        $lines = New-Object System.Collections.Generic.List[string]
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
        $lines = New-Object System.Collections.Generic.List[string]
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
