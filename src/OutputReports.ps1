using namespace System.Diagnostics
using namespace System.IO

# �e�X�g�ݒ�t�@�C���p�X
Set-Variable -Name RunSettingsFilePath -Value "test.runsettings" -Option Constant

# ReportGenerator�t�@�C���p�X
Set-Variable -Name ReportGeneratorPath -Value "$env:userprofile\.nuget\packages\reportgenerator\5.4.4\tools\net47\ReportGenerator.exe" -Option Constant

# Trxer�t�@�C���p�X
Set-Variable -Name TrxerConsolePath -Value "TrxerConsole.exe" -Option Constant

# �J�o���b�W����trx�t�@�C������������郋�[�g�f�B���N�g��
Set-Variable -Name TestResultsRootDirectory -Value "TestResults" -Option Constant

# ���|�[�g���ʂ̏o�̓��[�g�f�B���N�g��
Set-Variable -Name ReportRootDirectory -Value "TestReports" -Option Constant

# ReportGenerator���|�[�g�̏o�̓f�B���N�g��
Set-Variable -Name CoverageReportDirectory -Value "$ReportRootDirectory\CoverageReport" -Option Constant

# �P�̃e�X�g���|�[�g�̏o�̓f�B���N�g��
Set-Variable -Name TestResultReportDirectory -Value "$ReportRootDirectory" -Option Constant


<##
 # �������|�[�g���폜����B
 #>
 function RemoveReports()
 {
    if ([Directory]::Exists($ReportRootDirectory))
    {
        Remove-Item -Path "$ReportRootDirectory" -Recurse -Force
    }
 }

 <##
  # �\�����[�V�����̃N���[�����s���B
  #>
 function ExecuteCleanSolution()
 {
    dotnet clean --configuration Debug
 }

<##
 # �\�����[�V�����̒P�̃e�X�g�����{���A�J�o���b�W����trx�t�@�C���𐶐����A�J�o���b�WXML�t�@�C���p�X�����߂�B
 # @return �J�o���b�WXML�t�@�C���p�X�B
 #>
function ExecuteTestSolution()
{
    [string[]]$TestOutput = dotnet test --settings $RunSettingsFilePath
    [string]$cavarageXmlFile = $TestOutput | Select-String coverage.cobertura.xml | % { $_.ToString().Trim() }

    return $cavarageXmlFile
}

<##
 # �J�o���b�WXML�t�@�C���p�X����AReportGenerator���|�[�g���쐬����B
 # @param $coverageXmlFilePath �J�o���b�WXML�t�@�C���p�X�B
 # @param $coverageReportDirectory ReportGenerator���|�[�g�o�̓f�B���N�g���B
 # @return ReportGenerator���|�[�g��index.html�p�X�B
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
 # trx�t�@�C������P�̃e�X�g���|�[�g���쐬����B
 # @param $destDirectory �P�̃e�X�g���|�[�g�o�̓f�B���N�g���B
 # @return �P�̃e�X�g���|�[�g�̃t�@�C���p�X�B
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
 # ���|�[�g�����ߒ��ō쐬���ꂽ�t�H���_�[�^�t�@�C�����폜����B
 #>
function Cleanup()
{
    Remove-Item -Path "$TestResultsRootDirectory" -Recurse -Force
}


# �������|�[�g�̍폜
Write-Host "1) Remove existing reports"
RemoveReports

# �\�����[�V�����̃N���[��
Write-Host "2) Solution clean"
ExecuteCleanSolution

# �\�����[�V�����̃e�X�g�r���h
Write-Host "3) Creating test result with coverage"
[string]$coverageXmlFilePath = ExecuteTestSolution

# �J�o���b�W���|�[�g����
Write-Host "4) Creating coverage report"
[string]$coverageReportIndexPath = CreateCoverage $coverageXmlFilePath $CoverageReportDirectory

# �P�̃e�X�g���|�[�g����
Write-Host "5) Creating test report"
[string]$testReportFilePath = CreateTestReport $TestResultReportDirectory

# �N���[���A�b�v
Write-Host "6) Cleanup"
Cleanup

Write-Host "`r`nReport Path:"
Write-Host "  $coverageReportIndexPath"
Write-Host "  $testReportFilePath"
