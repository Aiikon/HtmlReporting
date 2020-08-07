Import-Module $PSScriptRoot\.. -DisableNameChecking -Force

Describe "Get-HtmlReportColor" {
    Context "Default" {
        $index0 = 67, 134, 216
        It "Works with Index" {
            $result = Get-HtmlReportColor -Index 0
            $result -join "`0" | Should Be ($index0 -join "`0")
        }
        It "Works with Name" {
            $result = Get-HtmlReportColor -Name Blue
            $result -join "`0" | Should Be ($index0 -join "`0")
        }
        It "Works with AsCssRgb" {
            $result = Get-HtmlReportColor -Index 0 -AsCssRgb
            $result | Should Be "rgb($($index0 -join ','))"
        }
        It "Works with AsCssRgba" {
            $result = Get-HtmlReportColor -Index 0 -AsCssRgba 0.25
            $result | Should Be "rgba($($index0 -join ','),0.25)"
        }
    }
}
