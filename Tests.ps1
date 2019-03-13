﻿Import-Module C:\PSModule\HtmlReporting -Force

$someVariable = "abcdefg"

Get-HtmlFragment {
    
    h1 "Semi-Fancy Get-HtmlText Usage"

    h2 "Header 2"
    h3 "Header 3"
    h4 "Header 4"
    
    $header3 = "Header Three"
    h3 $header3
    
    p "Here's a sample paragraph"
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

    h2 "Convert PS Code to HTML"

    div {
        p {
            Convert-PSCodeToHtml -PsCode "Get-ChildItem C:\ | Out-GridView -Title `$var1"
        }
    }

} |
    Out-HtmlFile -AddTimestamp

return

& {
    "<h1>Header 1</h1>"

    "<p>Alternatively just write HTML like this...</p>"
    
    Get-ChildItem C:\ | Select-Object Name, LastWriteTime, Length -First 10 | ConvertTo-HtmlTable
} |
    Out-HtmlFile ~\Desktop\Test.html -Open