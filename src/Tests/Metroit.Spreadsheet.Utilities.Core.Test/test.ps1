Write-Output "1) Creating test result with coverage"
$TestOutput = dotnet test --collect:"XPlat Code Coverage" --logger "trx;logfilename=testResults2.trx"

# Get file path
$TestReports = $TestOutput | Select-String coverage.cobertura.xml | % { $_.ToString().Trim() }

Write-Output "2) Creating coverage report"
$pinfo = New-Object System.Diagnostics.ProcessStartInfo
$pinfo.FileName = "$env:userprofile\.nuget\packages\reportgenerator\5.4.4\tools\net47\ReportGenerator.exe"
$pinfo.RedirectStandardError = $true
$pinfo.RedirectStandardOutput = $true
$pinfo.UseShellExecute = $false
$pinfo.Arguments = "-reports:$TestReports -targetdir:CoverageReport -reporttypes:Html"
$p = New-Object System.Diagnostics.Process
$p.StartInfo = $pinfo
$p.Start() | Out-Null
$p.WaitForExit()
$stdout = $p.StandardOutput.ReadToEnd()
Write-Host "$stdout"


# Start-Process -Wait -WindowStyle Hidden "$env:userprofile\.nuget\packages\reportgenerator\5.4.4\tools\net47\ReportGenerator.exe" -ArgumentList "-reports:$TestReports","-targetdir:CoverageReport","-reporttypes:Html"
