Import-Module $PSScriptRoot\.. -DisableNameChecking -Force

Describe "Get-HtmlDatalist" {
    Context "Default" {
        It "Works with Values" {
            $html = Get-HtmlDatalist
            $html | Should Be "<select><option value='A'>A</option><option value='B'>B</option></select>"
        }

        It "Works with Pipeline" {
            $html = Get-HtmlSelect -Values A, B -SelectedValues B
            $html | Should Be "<select><option value='A'>A</option><option value='B' selected='selected'>B</option></select>"
        }
    }
}
