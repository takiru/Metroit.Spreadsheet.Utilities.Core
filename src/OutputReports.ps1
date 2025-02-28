using namespace System.Diagnostics

# テスト設定ファイルパス
Set-Variable -Name RunSettingsFilePath -Value "test.runsettings" -Option Constant

# ReportGeneratorファイルパス
Set-Variable -Name ReportGeneratorPath -Value "$env:userprofile\.nuget\packages\reportgenerator\5.4.4\tools\net47\ReportGenerator.exe" -Option Constant

# Trxerファイルパス
Set-Variable -Name TrxerConsolePath -Value "TrxerConsole.exe" -Option Constant

# カバレッジ情報とtrxファイルが生成されるルートディレクトリ
Set-Variable -Name TestResultsRootDirectory -Value "TestResults" -Option Constant

# レポート結果の出力ルートディレクトリ
Set-Variable -Name ReportRootDirectory -Value "TestReports" -Option Constant

# ReportGeneratorレポートの出力ディレクトリ
Set-Variable -Name CoverageReportDirectory -Value "$ReportRootDirectory\CoverageReport" -Option Constant

# 単体テストレポートの出力ディレクトリ
Set-Variable -Name TestResultReportDirectory -Value "$ReportRootDirectory" -Option Constant


<##
 # レポート生成の初期化を行う。
 #>
 function Initialize()
 {
     Remove-Item -Path "$ReportRootDirectory" -Recurse -Force
 }

<##
 # ソリューションの単体テストを実施し、カバレッジ情報とtrxファイルを生成し、カバレッジXMLファイルパスを求める。
 # @return カバレッジXMLファイルパス。
 #>
function ExecuteTest()
{
    [string[]]$TestOutput = dotnet test --settings $RunSettingsFilePath
    [string]$cavarageXmlFile = $TestOutput | Select-String coverage.cobertura.xml | % { $_.ToString().Trim() }

    return $cavarageXmlFile
}

<##
 # カバレッジXMLファイルパスから、ReportGeneratorレポートを作成する。
 # @param $coverageXmlFilePath カバレッジXMLファイルパス。
 # @param $coverageReportDirectory ReportGeneratorレポート出力ディレクトリ。
 # @return ReportGeneratorレポートのindex.htmlパス。
 #>
function CreateCoverage($coverageXmlFilePath, $coverageReportDirectory)
{
    [ProcessStartInfo]$pinfo = [ProcessStartInfo]::new()
    $pinfo.FileName = $ReportGeneratorPath
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = "-reports:$coverageXmlFilePath -targetdir:$coverageReportDirectory -reporttypes:Html"
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()
    $stdout = $p.StandardOutput.ReadToEnd()
    Write-Host "$stdout"

    [string]$absolutePath = Resolve-Path "$coverageReportDirectory"
    [string]$caverageReportIndexPath = Join-Path -Path "$absolutePath" -ChildPath "index.html"

    return $caverageReportIndexPath
}

<##
 # trxファイルから単体テストレポートを作成する。
 # @param $destDirectory 単体テストレポート出力ディレクトリ。
 # @return 単体テストレポートのファイルパス。
 #>
function CreateTestReport($destDirectory)
{
    [ProcessStartInfo]$pinfo = [ProcessStartInfo]::new()
    $pinfo.FileName = $TrxerConsolePath
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = "$TestResultsRootDirectory\TestResult.trx"
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()
    $stdout = $p.StandardOutput.ReadToEnd()
    Write-Host "$stdout"
    Move-Item -Path "$TestResultsRootDirectory\TestResult.trx.html" -Destination "$destDirectory\TestResult.html" -Force

    [string]$absolutePath = Resolve-Path "$destDirectory"
    [string]$testReportResultPath = Join-Path -Path "$absolutePath" -ChildPath "TestResult.html"

    return $testReportResultPath
}

<##
 # レポート生成過程で作成されたフォルダー／ファイルを削除する。
 #>
function Cleanup()
{
    Remove-Item -Path "$TestResultsRootDirectory" -Recurse -Force
}


# 初期化
Write-Host "1) Initialize"
Initialize

# テストビルド
Write-Host "2) Creating test result with coverage"
[string]$coverageXmlFilePath = ExecuteTest

# カバレッジレポート生成
Write-Host "3) Creating coverage report"
[string]$coverageReportIndexPath = CreateCoverage $coverageXmlFilePath $CoverageReportDirectory

# 単体テストレポート生成
Write-Host "4) Creating test report"
[string]$testReportFilePath = CreateTestReport $TestResultReportDirectory

# クリーンアップ
Write-Host "5) Cleanup"
Cleanup

Write-Host "`r`nReport Path:"
Write-Host "  $coverageReportIndexPath"
Write-Host "  $testReportFilePath"
