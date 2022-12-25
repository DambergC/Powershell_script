$rootFolder = "$PSScriptRoot\result\"
$savefolder = "$psscriptroot"
$outfile    = Join-Path -Path $saveFolder -ChildPath 'Index.html'

$sw = New-Object System.IO.StreamWriter $outfile, $true  # $true is for Append
Get-ChildItem -Path $rootFolder -Filter '*.html' -File | ForEach-Object {
    Get-Content -Path $_.FullName -Encoding UTF8 | ForEach-Object {
        $sw.WriteLine($_)
    }
}
$sw.Dispose()