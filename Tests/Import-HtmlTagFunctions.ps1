Import-Module $PSScriptRoot\.. -DisableNameChecking -Force
Remove-Module -Name HtmlReportingTagFunctions -ErrorAction Ignore

Describe "Import-HtmlTagFunctions" {
    Context "Default" {
        Import-HtmlTagFunctions
        It "Works with simple tags" {
            h1 "a" | Should Be "<h1>a</h1>"
            h2 "b" | Should Be "<h2>b</h2>"
            h3 "c" | Should Be "<h3>c</h3>"
            h4 "d" | Should Be "<h4>d</h4>"
            p e | Should Be "<p>e</p>"
            ul (li c) | Should Be "<ul><li>c</li></ul>"
        }
        It "Appends multiple children with spaces" {
            p @(
                "Sentence One."
                "Sentence Two."
            ) | Should Be "<p>Sentence One. Sentence Two.</p>"
        }
        It "Works with scriptblocks" {
            ol {
                li One
                li Two
            } | Should Be "<ol><li>One</li> <li>Two</li></ol>"
        }
        It "Works with freeform tag" {
            html img "Test" | Should Be "<img>Test</img>"
        }
        It "Works with attributes" {
            div -Attributes @{aria='test'; data='one'} |
                Should Be "<div aria='test' data='one'></div>"
        }
        It "Works with style and class" {
            span Words -class abc, def -style "font-weight:bold;", "font-style:italic;" |
                Should Be "<span class='abc def' style='font-weight:bold; font-style:italic;'>Words</span>"
        }

        It "span has title" {
            span abc -title def |
                Should Be "<span title='def'>abc</span>"
        }

        It "br exists" {
            br | Should Be "<br />"
        }

        It "br supports clear" {
            br -clear all | Should Be "<br clear='all' />"
        }

        It "input supports checked" {
            input -type checkbox -checked $true | Should Be "<input type='checkbox' checked='checked'></input>"
            input -type checkbox -checked $false | Should Be "<input type='checkbox'></input>"
        }

        It "Generates a pre function" {
            pre -htmlencode "Test<>" | Should Be "<pre>Test&lt;&gt;</pre>"
        }

        It "Generates a small function" {
            small "text" | Should Be "<small>text</small>"
        }

        It "Generates a script function" {
            script "function test() { }" |
                Should Be "<script>function test() { }</script>"
        }

        It "Generates a style function" {
            style "class { }" |
                Should Be "<style>class { }</style>"
        }

        It "Generates tables" {
            table @(
                thead @(
                    tr @(
                        th "Col 1"
                        th "Col 2"
                    )
                )
                tbody @(
                    tr @(
                        td "Cell A"
                        td "Cell B"
                    )
                )
            ) |
                Should Be '<table><thead><tr><th>Col 1</th> <th>Col 2</th></tr></thead> <tbody><tr><td>Cell A</td> <td>Cell B</td></tr></tbody></table>'
            # ^ There will be spaces between thead and tbody and the tds because that's the default behavior, i.e p text1 text2
        }
    }
}
