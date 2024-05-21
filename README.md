<!-- 
SPDX-FileCopyrightText: (c) 2022-2024 T. Graf
SPDX-License-Identifier: Apache-2.0
-->

# Tethys.XlsxSupport

![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)
[![Build status](https://ci.appveyor.com/api/projects/status/6huida3wfgnklsrs?svg=true)](https://ci.appveyor.com/project/tngraf/tethys-xlsxsupport)
[![Nuget](https://img.shields.io/badge/nuget-1.0.0-brightgreen.svg)](https://www.nuget.org/packages/Tethys.XlsxSupport/1.0.0)
[![REUSE status](https://api.reuse.software/badge/git.fsfe.org/reuse/api)](https://api.reuse.software/info/git.fsfe.org/reuse/api)
[![SBOM](https://img.shields.io/badge/SBOM-CycloneDX-brightgreen)](https://github.com/tngraf/Tethys.XlsxSupport/blob/master/SBOM/sbom.cyclonedx.xml)

This library simplifies working with the [Open XML SDK for Office](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk?redirectedfrom=MSDN).
The **XlsX** format has an enormous number of features, but if you just want to create a simple document
there is a steep learning curve. Tethys.XlsxSupport simplifies a number of operations.

## Get Package

You can get Tethys.XlsxSupport by grabbing the latest NuGet packages from [here](https://www.nuget.org/packages/Tethys.XlsxSupport/1.1.0).

## Build

### Requisites

* Visual Studio 2019

### Build Solution

Just use the basic `dotnet` command:

```shell
dotnet build
```

Run the demo application:

```shell
dotnet run --project .\Tethys.XlsxSupport.Demo\Tethys.XlsxSupport.Demo.csproj
```

## License

Tethys.XlsxSupport is licensed under the Apache License, Version 2.0.
