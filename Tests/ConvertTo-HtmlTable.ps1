#Requires -Modules @{ ModuleName="Pester"; ModuleVersion="3.4.0" }
Import-Module Pester -MaximumVersion 3.4.0

foreach ($value in $true, $false)
{
    $Global:2e128b9186234521b3ab5ce70cc83360_ForceLoadPowerShellCmdlets = $value
    Import-Module $PSScriptRoot\.. -DisableNameChecking -Force

    Describe "ConvertTo-HtmlTable" {
        $SampleData1 = "ClusterId,ClusterName,ClusterType
        1,SQL001,SQL
        2,SQL002,SQL
        3,CAFile,File
        4,TXFile,File" -replace ' ' | ConvertFrom-Csv
        Context "Default - PowerShell: $value" {
            It "Works with no parameters" {
                $actualHtml = $SampleData1 | ConvertTo-HtmlTable
                $expectedHtml = "
                <table class='HtmlReportingTable'>
                <thead>
                <tr class='header'>
                <th>ClusterId</th>
                <th>ClusterName</th>
                <th>ClusterType</th>
                </tr>
                </thead>
                <tbody>
                <tr>
                <td>1</td>
                <td>SQL001</td>
                <td>SQL</td>
                </tr>
                <tr>
                <td>2</td>
                <td>SQL002</td>
                <td>SQL</td>
                </tr>
                <tr>
                <td>3</td>
                <td>CAFile</td>
                <td>File</td>
                </tr>
                <tr>
                <td>4</td>
                <td>TXFile</td>
                <td>File</td>
                </tr>
                </tbody>
                </table>
                ".Trim() -split "[`r`n]+" -replace "^ *" -join "`r`n"
                $actualHtml | Should Be $expectedHtml
            }

            It "Works with simple cell style scripts" {
                $result = $SampleData1 | ConvertTo-HtmlTable -CellStyleScripts @{ClusterName={if ($_.ClusterId -eq 2) { 'color:red;' }}}
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[2]/td[2]/@style').'#text' | Should Be "color:red;"
                $resultXml.SelectNodes('//tbody/tr[2]/td[3]/@style').'#text' | Should Be $null
            }

            It "Works with array cell style scripts" {
                $result = $SampleData1 | ConvertTo-HtmlTable -CellStyleScripts @{('ClusterName', 'ClusterType')={if ($_.ClusterId -eq 2) { 'color:red;' }}}
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[2]/td[2]/@style').'#text' | Should Be "color:red;"
                $resultXml.SelectNodes('//tbody/tr[2]/td[3]/@style').'#text' | Should Be "color:red;"
            }

            It "Works with multiple output cell style scripts" {
                $result = $SampleData1 | ConvertTo-HtmlTable -CellStyleScripts @{ClusterName={if ($_.ClusterId -eq 2) { 'color:red;'; 'font-weight:bold;' }}}
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[2]/td[2]/@style').'#text' | Should Be "color:red; font-weight:bold;"
                $resultXml.SelectNodes('//tbody/tr[2]/td[3]/@style').'#text' | Should Be $null
            }

            It "Works with cell style scripts referencing strings (properties)" {
                $result = [pscustomobject]@{A=1; B=2; C=3; D=4; Y='font-weight:bold;'; Z='color:red;'} |
                    ConvertTo-HtmlTable -Property A, B, C, D -CellStyleScripts @{A='Y'; B='Z'; C='Z'}

                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[1]/td[1]').OuterXml | Should Be '<td style="font-weight:bold;">1</td>'
                $resultXml.SelectNodes('//tbody/tr[1]/td[2]').OuterXml | Should Be '<td style="color:red;">2</td>'
                $resultXml.SelectNodes('//tbody/tr[1]/td[3]').OuterXml | Should Be '<td style="color:red;">3</td>'
                $resultXml.SelectNodes('//tbody/tr[1]/td[4]').OuterXml | Should Be '<td>4</td>'
            }

            It "Works with row style script" {
                $result = $SampleData1 | ConvertTo-HtmlTable -RowStyleScript {if ($_.ClusterId -eq 2) { 'color:red;'; 'font-weight:bold;' }}
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[2]/td[1]/@style').'#text' | Should Be "color:red; font-weight:bold;"
                $resultXml.SelectNodes('//tbody/tr[2]/td[2]/@style').'#text' | Should Be "color:red; font-weight:bold;"
                $resultXml.SelectNodes('//tbody/tr[2]/td[3]/@style').'#text' | Should Be "color:red; font-weight:bold;"
            }

            It "Works with simple cell class scripts" {
                $result = $SampleData1 | ConvertTo-HtmlTable -CellClassScripts @{ClusterName={if ($_.ClusterId -eq 2) { 'red' }}}
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[2]/td[2]/@class').'#text' | Should Be "red"
                $resultXml.SelectNodes('//tbody/tr[2]/td[3]/@class').'#text' | Should Be $null
            }

            It "Works with array cell class scripts" {
                $result = $SampleData1 | ConvertTo-HtmlTable -CellClassScripts @{('ClusterName', 'ClusterType')={if ($_.ClusterId -eq 2) { 'red' }}}
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[2]/td[2]/@class').'#text' | Should Be "red"
                $resultXml.SelectNodes('//tbody/tr[2]/td[3]/@class').'#text' | Should Be "red"
            }

            It "Works with multiple output cell class scripts" {
                $result = $SampleData1 | ConvertTo-HtmlTable -CellClassScripts @{ClusterName={if ($_.ClusterId -eq 2) { 'red';'bold' }}}
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[2]/td[2]/@class').'#text' | Should Be "red bold"
                $resultXml.SelectNodes('//tbody/tr[2]/td[3]/@class').'#text' | Should Be $null
            }

            It "Works with cell class scripts referencing strings (properties)" {
                $result = [pscustomobject]@{A=1; B=2; C=3; D=4; Y='red'; Z='green'} |
                    ConvertTo-HtmlTable -Property A, B, C, D -CellClassScripts @{A='Y'; C='Z'}

                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[1]/td[1]').OuterXml | Should Be '<td class="red">1</td>'
                $resultXml.SelectNodes('//tbody/tr[1]/td[2]').OuterXml | Should Be '<td>2</td>'
                $resultXml.SelectNodes('//tbody/tr[1]/td[3]').OuterXml | Should Be '<td class="green">3</td>'
                $resultXml.SelectNodes('//tbody/tr[1]/td[4]').OuterXml | Should Be '<td>4</td>'
            }

            It "Works with row cell class script" {
                $result = $SampleData1 | ConvertTo-HtmlTable -RowClassScript {if ($_.ClusterId -eq 2) { 'red';'bold' }}
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[2]/td[1]/@class').'#text' | Should Be "red bold"
                $resultXml.SelectNodes('//tbody/tr[2]/td[2]/@class').'#text' | Should Be "red bold"
                $resultXml.SelectNodes('//tbody/tr[2]/td[3]/@class').'#text' | Should Be "red bold"
            }

            It "Joins array values with spaces (OFS)" {
                $result = [pscustomobject]@{A=1,'Two'} | ConvertTo-HtmlTable
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[1]/td[1]').'#text' | Should Be "1 Two"
            }

            It "Works with ExcludeProperty" {
                $result = [pscustomobject]@{Bad=1; Good=2} | ConvertTo-HtmlTable -ExcludeProperty B*
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//thead/tr[1]/th[1]').'#text' | Should Be "Good"
            }

            It "Works with ExcludeProperty (No Input)" {
                $result = @() | ConvertTo-HtmlTable -ExcludeProperty B*
                1 | Should Be 1
            }

            It "Works with AutoDetectHtml" {
                $result = [pscustomobject]@{Col1="<strong>Text</strong>"; Col2="Basic < text"; Col3="Line`r`nBreak"} |
                    ConvertTo-HtmlTable -AutoDetectHtml
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr[1]/td[1]').innerXml | Should Be "<strong>Text</strong>"
                $resultXml.SelectNodes('//tbody/tr[1]/td[2]').innerXml | Should Be "Basic &lt; text"
                $resultXml.SelectNodes('//tbody/tr[1]/td[3]').innerXml | Should Be "Line<br />Break"
            }

            It "Works with AddDataColumnName" {
                $result = [pscustomobject]@{Col1="Text1"; Col2="Text2"} |
                    ConvertTo-HtmlTable -AddDataColumnName
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//thead/tr[1]/th[1]').'data-column-name' | Should Be 'Col1'
                $resultXml.SelectNodes('//thead/tr[1]/th[2]').'data-column-name' | Should Be 'Col2'
                $resultXml.SelectNodes('//tbody/tr[1]/td[1]').'data-column-name' | Should Be 'Col1'
                $resultXml.SelectNodes('//tbody/tr[1]/td[2]').'data-column-name' | Should Be 'Col2'
            }

            It "Works with Colspan" {
                $result = @(
                    [pscustomobject]@{A=1;B=2;C=3;D='-'}
                ) |
                    ConvertTo-HtmlTable -Class Grid -CellColspanScripts @{
                        A = { if ($_.A -eq 1) { 1 } }
                        B = { if ($_.B -eq 2) { 2 } }
                    }
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr').Count | Should Be 1
                $resultXml.SelectNodes('//tbody/tr[1]').InnerXml | Should Be '<td>1</td><td colspan="2">2</td><td>-</td>'
            }

            It "Works with Rowspan" {
                $result = @(
                    [pscustomobject]@{A=1;B=2;C=3;D='-'}
                    [pscustomobject]@{A=4;B=5;C=6;D='-'}
                ) |
                    ConvertTo-HtmlTable -Class Grid -CellRowspanScripts @{
                        A = { if ($_.A -eq 1) { 1 } }
                        B = { if ($_.B -eq 2) { 2 } }
                    }
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr').Count | Should Be 2
                $resultXml.SelectNodes('//tbody/tr[1]').InnerXml | Should Be '<td>1</td><td rowspan="2">2</td><td>3</td><td>-</td>'
                $resultXml.SelectNodes('//tbody/tr[2]').InnerXml | Should Be '<td>4</td><td>6</td><td>-</td>'
            }

            It "Works with Rowspan and Colspan" {
                $result = @(
                    [pscustomobject]@{A=1;B=2;C=3;D='-'}
                    [pscustomobject]@{A=4;B=5;C=6;D='-'}
                    [pscustomobject]@{A=7;B=8;C=9;D='-'}
                    [pscustomobject]@{A=10;B=11;C=12;D='-'}
                    [pscustomobject]@{A=13;B=14;C=15;D='-'}
                ) |
                    ConvertTo-HtmlTable -Class Grid -CellColspanScripts @{
                        B = { if ($_.B -eq 2) { 2 } elseif ($_.B -eq 5) { 2 } }
                    } -CellRowspanScripts @{
                        A = { if ($_.A -eq 4) { 2 } }
                        B = { if ($_.B -eq 5) { 2 } }
                        C = { if ($_.C -eq 12) { 2 } }
                    }
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//tbody/tr').Count | Should Be 5
                $resultXml.SelectNodes('//tbody/tr[1]').InnerXml | Should Be '<td>1</td><td colspan="2">2</td><td>-</td>'
                $resultXml.SelectNodes('//tbody/tr[2]').InnerXml | Should Be '<td rowspan="2">4</td><td colspan="2" rowspan="2">5</td><td>-</td>'
                $resultXml.SelectNodes('//tbody/tr[3]').InnerXml | Should Be '<td>-</td>'
                $resultXml.SelectNodes('//tbody/tr[4]').InnerXml | Should Be '<td>10</td><td>11</td><td rowspan="2">12</td><td>-</td>'
                $resultXml.SelectNodes('//tbody/tr[5]').InnerXml | Should Be '<td>13</td><td>14</td><td>-</td>'
            }
        }

        Context "Class" {
            It 'Plain' {
                $result = [pscustomobject]@{A=1} | ConvertTo-HtmlTable -Plain
                $resultXml = [xml]$result
                $resultXml.table.Attributes['class'].'#text' | Should Be ''
            }
        }

        Context "Id" {
            It 'Sets' {
                $result = [pscustomobject]@{A=1} | ConvertTo-HtmlTable -Id MyId
                $resultXml = [xml]$result
                $resultXml.table.Attributes['id'].'#text' | Should Be 'MyId'
            }

            It 'Stays off if unset' {
                $result = [pscustomobject]@{A=1} | ConvertTo-HtmlTable
                $resultXml = [xml]$result
                $resultXml.table.Attributes['id'] | Should Be $null
            }
        }

        Context "InsertLine" {
            It 'All three' {
                $result = [pscustomobject]@{A=1;B=2;C=3;D=4} |
                    ConvertTo-HtmlTable -InsertSolidLine B -InsertDashedLine C -InsertDottedLine D
                $resultXml = [xml]$result
                $resultXml.SelectNodes('//thead/tr[1]/th[1]').class | Should Be $null
                $resultXml.SelectNodes('//thead/tr[1]/th[2]').class | Should Be InsertSolidLine
                $resultXml.SelectNodes('//thead/tr[1]/th[3]').class | Should Be InsertDashedLine
                $resultXml.SelectNodes('//thead/tr[1]/th[4]').class | Should Be InsertDottedLine
                $resultXml.SelectNodes('//tbody/tr[1]/td[1]').class | Should Be $null
                $resultXml.SelectNodes('//tbody/tr[1]/td[2]').class | Should Be InsertSolidLine
                $resultXml.SelectNodes('//tbody/tr[1]/td[3]').class | Should Be InsertDashedLine
                $resultXml.SelectNodes('//tbody/tr[1]/td[4]').class | Should Be InsertDottedLine
            }
        }
    }
}
