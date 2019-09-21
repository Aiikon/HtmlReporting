Import-Module C:\PSModule\HtmlReporting -Force

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

    h2 Indicators
    ul {
        li (Get-HtmlIndicatorText "Sample 1" -ColorName Green)
        li (Get-HtmlIndicatorText "Sample 2" -ColorName Red -BorderOnly)
        li (Get-HtmlIndicatorText "Sample 3" -ColorIndex 0 -Width 200px)
        li (Get-HtmlIndicatorText "Sample 4" -ColorRGB 0,0,0)
    }

    h2 Table

    Get-Service | Select-Object Name, Status, DisplayName -First 10 | ConvertTo-HtmlTable

    Get-Service | Select-Object Name, Status, DisplayName -First 10 -Skip 10 | ConvertTo-HtmlTable -Narrow

    "
        Rowspan,Value
        1,A
        2,B
        2,C
        1,D
    ".Trim() -replace '\A ' | ConvertFrom-Csv | ConvertTo-HtmlTable -CellRowspanScripts @{Value='Rowspan'}

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


    h2 ConvertTo-HtmlColorBlocks

    Get-Service | Select-Object -First 5 Name, DisplayName, Status, StartType |
        ConvertTo-HtmlColorBlocks -TocProperty Name -OutputScript { $_ | ConvertTo-HtmlStrongText }

} |
    Out-HtmlFile -AddTimestamp

return

& {
    "<h1>Header 1</h1>"

    "<p>Alternatively just write HTML like this...</p>"
    
    Get-ChildItem C:\ | Select-Object Name, LastWriteTime, Length -First 10 | ConvertTo-HtmlTable
} |
    Out-HtmlFile ~\Desktop\Test.html -Open