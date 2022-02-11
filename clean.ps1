# -------------
# Clean project
# -------------

dotnet clean
Remove-Item "Tethys.XlsxSupport\bin" -Recurse
Remove-Item "Tethys.XlsxSupport\obj" -Recurse
Remove-Item "Tethys.XlsxSupport.Demo\bin" -Recurse
Remove-Item "Tethys.XlsxSupport.Demo\obj" -Recurse
Remove-Item MySheet.xlsx
