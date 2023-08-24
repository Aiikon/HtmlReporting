Import-Module $PSScriptRoot\.. -DisableNameChecking -Force

Describe "Get-HtmlReportColorSet" {
    Context "Default" {
        It "Works" {
            $colorSet = Get-HtmlReportColorSet
            $colorSet -is [System.Collections.IDictionary] | Should Be $true
            $colorSet[0] -join '+' | Should Be '67+134+216'
            $colorSet.Blue -join '+' | Should Be '67+134+216'
        }

        It "Can override the first colors" {
            $colorSet = Get-HtmlReportColorSet -StartingColorNames Green, Red
            $colorSet[0] | Should Be $colorSet.Green
            $colorSet[1] | Should Be $colorSet.Red
            $colorSet[2] | Should Be $colorSet.Blue
        }

        It "Can exclude colors" {
            $colorSet = Get-HtmlReportColorSet -ExcludeColorNames Green, Red
            $colorSet[0] | Should Be $colorSet.Blue
            $colorSet[1] | Should Be $colorSet.Orange
            $colorSet[2] | Should Be $colorSet.Purple
        }

        It "Can take and modify a custom color set" {
            $customSet = [ordered]@{
                Red = 255,0,0
                Green = 0,255,0
                Blue = 0,0,255
            }

            $colorSet = Get-HtmlReportColorSet -ColorSet $customSet -StartingColorNames Green -ExcludeColorNames Blue
            $colorSet[0] | Should Be $customSet.Green
            $colorSet[1] | Should Be $customSet.Red
            $colorSet[2] | Should Be $null
        }
    }
}

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
        It "Can override the first colors" {
            Get-HtmlReportColor -Index 0 -StartingColorNames Green, Red -AsCssRgb |
                Should Be (Get-HtmlReportColor -Name Green -AsCssRgb)
        }
        It "Can exclude colors" {
            Get-HtmlReportColor -Index 0 -ExcludeColorNames Blue -AsCssRgb |
                Should Be (Get-HtmlReportColor -Name Orange -AsCssRgb)
        }
        It "Can take a custom color set" {
            $customSet = [ordered]@{
                Red = 255,0,0
                Green = 0,255,0
                Blue = 0,0,255
            }

            Get-HtmlReportColor -Index 1 -ColorSet $customSet -AsCssRgb |
                Should Be "rgb(0,255,0)"

            Get-HtmlReportColor -Name Red -ColorSet $customSet -AsCssRgb |
                Should Be "rgb(255,0,0)"
        }
    }
}
