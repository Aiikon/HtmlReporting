Import-Module $PSScriptRoot\.. -DisableNameChecking -Force

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
    }
}
