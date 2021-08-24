Import-Module $PSScriptRoot\.. -DisableNameChecking -Force

Describe "Get-HtmlSelect" {
    Context "Default" {
        It "Works with Values" {
            $html = Get-HtmlSelect -Values A, B
            $html | Should Be "<select><option value='A'>A</option><option value='B'>B</option></select>"
        }

        It "Works with Selected" {
            $html = Get-HtmlSelect -Values A, B -SelectedValues B
            $html | Should Be "<select><option value='A'>A</option><option value='B' selected='selected'>B</option></select>"
        }

        It "Works with Hashtable Ordered" {
            $html = Get-HtmlSelect -ValueLabelDictionary ([ordered]@{A='ehh'; C='see'; B='bee'})
            $html | Should Be "<select><option value='A'>ehh</option><option value='C'>see</option><option value='B'>bee</option></select>"
        }

        It "Works with Hashtable Unordered" {
            $html = Get-HtmlSelect -ValueLabelDictionary @{A='ehh'; B='bee'; C='see'}
            $html | Should Be "<select><option value='C'>see</option><option value='B'>bee</option><option value='A'>ehh</option></select>"
        }

        It "Works with InputObject" {
            $html = [pscustomobject]@{Name='C'; Value=1}, [pscustomobject]@{Name='D'; Value=2; Checked=$true} |
                Get-HtmlSelect -ValueProperty Value -LabelProperty Name -IsSelectedProperty Checked
            $html | Should Be "<select><option value='1'>C</option><option value='2' selected='selected'>D</option></select>"
        }

        It "Works with Required" {
            $html = Get-HtmlSelect -Values A, B -Required
            $html | Should Be "<select required='required'><option value='A'>A</option><option value='B'>B</option></select>"
        }

        It "Works with Multiple" {
            $html = Get-HtmlSelect -Values A, B -Multiple
            $html | Should Be "<select multiple='multiple'><option value='A'>A</option><option value='B'>B</option></select>"
        }

        It "Works with Id" {
            $html = Get-HtmlSelect -Values A, B -Id 12345
            $html | Should Be "<select id='12345'><option value='A'>A</option><option value='B'>B</option></select>"
        }

        It "Works with Name" {
            $html = Get-HtmlSelect -Values A, B -Name TheName
            $html | Should Be "<select name='TheName'><option value='A'>A</option><option value='B'>B</option></select>"
        }

        It "Works with Class" {
            $html = Get-HtmlSelect -Values A, B -Class ClassA, ClassB
            $html | Should Be "<select class='ClassA ClassB'><option value='A'>A</option><option value='B'>B</option></select>"
        }

        It "Works with Style" {
            $html = Get-HtmlSelect -Values A, B -Style 'size:150%;'
            $html | Should Be "<select style='size:150%;'><option value='A'>A</option><option value='B'>B</option></select>"
        }

        It "Works with Specific Size" {
            $html = Get-HtmlSelect -Values A, B -Size 1
            $html | Should Be "<select size='1'><option value='A'>A</option><option value='B'>B</option></select>"
        }

        It "Works with Zero Size" {
            $html = Get-HtmlSelect -Values A, B -Size 0
            $html | Should Be "<select size='2'><option value='A'>A</option><option value='B'>B</option></select>"
        }
    }
}
