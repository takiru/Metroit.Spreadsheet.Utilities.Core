dotnet test --collect:"XPlat Code Coverage"

%UserProfile%\.nuget\packages\reportgenerator\5.4.4\tools\net47\ReportGenerator.exe -reports:"D:\App\GitHub\Metroit.Spreadsheet.Utilities.Core\src\Metroit.Spreadsheet.Utilities.Core.XUnitTests\TestResults\96ddbfc5-2901-45a5-9f5a-36d4d6ef0b44\coverage.cobertura.xml" -targetdir:"coveragereport" -reporttypes:Html
