Import-Module $PSScriptRoot\.. -DisableNameChecking -Force

Describe "Rhodium.HtmlReporting.HtmlReportingHelpers" {
    Context "GetStringsLike" {

        BeforeEach {
            $stringList = 'StartTime', 'StopTime', 'StartAction', 'StopAction'
        }

        It "Null value is empty" {
            $result = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLike($null, $null, $null)
            @($result).Count | Should Be 0
        }

        It "Empty array value is empty" {
            $result = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLike(@(), $null, $null)
            @($result).Count | Should Be 0
        }

        It "Null filters passes all values" {
            $result = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLike($stringList, $null, $null)
            $result -join '+' | Should Be 'StartTime+StopTime+StartAction+StopAction'
        }

        It "Empty filters passes all values" {
            $result = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLike($stringList, @(), @())
            $result -join '+' | Should Be 'StartTime+StopTime+StartAction+StopAction'
        }

        It "Wildcard like passes all values" {
            $result = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLike($stringList, '*', $null)
            $result -join '+' | Should Be 'StartTime+StopTime+StartAction+StopAction'
        }

        It "Wildcard like works" {
            $result = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLike($stringList, 'Stop*', $null)
            $result -join '+' | Should Be 'StopTime+StopAction'
        }

        It "Wildcard notlike works" {
            $result = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLike($stringList, $null, '*Time')
            $result -join '+' | Should Be 'StartAction+StopAction'
        }

        It "Wildcard like plus notlike works" {
            $result = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLike($stringList, 'Start*', '*Time')
            $result -join '+' | Should Be 'StartAction'
        }

        It "Like not defaulting to wildcard works" {
            $result = [Rhodium.HtmlReporting.HtmlReportingHelpers]::GetStringsLike($stringList, $null, $null, $false)
            $result -join '+' | Should Be ''
        }
    }
}