# ---------------------------------------------
# Clean project
# SPDX-FileCopyrightText: (c) 2022-2023 T. Graf
# SPDX-License-Identifier: Apache-2.0
# ---------------------------------------------

dotnet clean
Remove-Item "Tethys.XlsxSupport\bin" -Recurse
Remove-Item "Tethys.XlsxSupport\obj" -Recurse
Remove-Item "Tethys.XlsxSupport.Demo\bin" -Recurse
Remove-Item "Tethys.XlsxSupport.Demo\obj" -Recurse
Remove-Item MySheet.xlsx
