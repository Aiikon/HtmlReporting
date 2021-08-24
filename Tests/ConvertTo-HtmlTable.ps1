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
        }
    }
}
