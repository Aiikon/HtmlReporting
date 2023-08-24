Import-Module $PSScriptRoot\.. -DisableNameChecking -Force

Describe "Get-HtmlDataList" {
    Context "Default" {
        It "Works with Values" {
            $html = Get-HtmlDataList -Id list1 -Values A, B
            $html | Should Be "<datalist id='list1'><option value='A' /><option value='B' /></datalist>"
        }

        It "Works with Pipeline" {
            $html = [pscustomobject]@{Val='C'}, [pscustomobject]@{Val='D'} | Get-HtmlDataList -ValueProperty Val -Id list2
            $html | Should Be "<datalist id='list2'><option value='C' /><option value='D' /></datalist>"
        }
    }
}
