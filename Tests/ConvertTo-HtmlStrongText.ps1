Import-Module $PSScriptRoot\.. -DisableNameChecking -Force

Describe "ConvertTo-HtmlStrongText" {
    Context "Basic Checks" {

        It "Basic MultiLine Test" {
            [pscustomobject]@{A='One';B='Two'} |
                ConvertTo-HtmlStrongText |
                Should Be "<p><strong>A</strong><br />One</p>`r`n<p><strong>B</strong><br />Two</p>"

            [pscustomobject]@{A='One';B='Two'} |
                ConvertTo-HtmlStrongText -Mode MultiLine |
                Should Be "<p><strong>A</strong><br />One</p>`r`n<p><strong>B</strong><br />Two</p>"
        }

        It "Basic SameLine Test" {
            [pscustomobject]@{A='One';B='Two'} |
                ConvertTo-HtmlStrongText -Mode SameLine |
                Should Be "<p><strong>A:</strong> One<br /><strong>B:</strong> Two</p>"
        }

        It "Basic MultiLineHorizontal Test" {
            [pscustomobject]@{A='One'; B='Two'} |
                ConvertTo-HtmlStrongText -Mode MultiLineHorizontal
        }

        It "Basic HtmlProperty Test" {
            [pscustomobject]@{A='<i>Test</i>'; B='<i>Test</i>'} |
                ConvertTo-HtmlStrongText -HtmlProperty B |
                Should Be "<p><strong>A</strong><br />&lt;i&gt;Test&lt;/i&gt;</p>`r`n<p><strong>B</strong><br /><i>Test</i></p>"
        }

        It "Wildcard HtmlProperty Test" {
            [pscustomobject]@{A='<i>Test</i>'; B='<i>Test</i>'} |
                ConvertTo-HtmlStrongText -HtmlProperty * |
                Should Be "<p><strong>A</strong><br /><i>Test</i></p>`r`n<p><strong>B</strong><br /><i>Test</i></p>"
        }

        It "AutoDetectHtml Test" {
            [pscustomobject]@{A='<i>Test</i>'; B='Test < Other'; C="Line`r`nBreak"} |
                ConvertTo-HtmlStrongText -AutoDetectHtml |
                Should Be "<p><strong>A</strong><br /><i>Test</i></p>`r`n<p><strong>B</strong><br />Test &lt; Other</p>`r`n<p><strong>C</strong><br />Line<br />Break</p>"
        }
    }
}