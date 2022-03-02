$csv = import-csv .\WorkInProgress\PatchMyPC-SupportedProductsList.csv

$csv | export-csv .\WorkInProgress\clean.csv -Force

Get-Content .\WorkInProgress\clean.csv | Select-Object -Skip 1 | Out-GridView