Import-Module C:\PSModule\HtmlReporting -Force

$someVariable = "abcdefg"

Get-HtmlText {
    $header3 = "Header Three"
    h1 "Header 1"
    h2 "Header 2"
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

    Get-Service | Select-Object Name, Status, DisplayName -First 10 | ConvertTo-HtmlTable

    div {
        p {
            Convert-PSCodeToHtml -PsCode "Get-ChildItem C:\ | Out-GridView"
        }
    }

} |
    Out-HtmlFile

