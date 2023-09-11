![](https://github.com/heartace/OrderSheetConverter/actions/workflows/dotnet-desktop.yml/badge.svg)
![](https://img.shields.io/github/v/release/heartace/OrderSheetConverter.svg)

# OrderSheetConverter

Windows desktop app for converting the raw Excel workbook of orders to a more friendly one.

## Overview

This Windows app is a dedicated tool for converting a certain type of the Excel workbook which includes multiple stationary orders (mostly masking tapes) to a more friendly Excel workbook. The accepted file format is `xls` file which is generated from a WeChat mini program. After converting, a `xlsx` file would be generated.

The input file example ([File location](Examples/RawOrders.xls), unnecessary or sensitive data has been removed or updated):

<img width="1703" alt="image" src="https://github.com/heartace/OrderSheetConverterDesktop/assets/669206/5bc4fef7-da5a-45ee-8480-ca59daf0aa47">

The output file example ([File location](Examples/RawOrders-converted.xlsx)): 

<img width="1827" alt="image" src="https://github.com/heartace/OrderSheetConverterDesktop/assets/669206/8513186e-eda2-4677-b12e-8ad901cb97e5">

## Instructions

The binaries can be downloaded in the `Releases` section. Make sure [.NET Runtime 6.0]([https://github.com/heartace/OrderSheetConverterDesktop/releases](https://dotnet.microsoft.com/en-us/download/dotnet/6.0)) has been installed in the system.

When the app is opened, either click the Excel logo or drag the xls file to the Excel logo to import the file.

![app-open](https://github.com/heartace/OrderSheetConverterDesktop/assets/669206/16c804b4-5bb8-43f0-b06e-0601274aa7f9)

If the file is valid (the file extension is `.xls`, and a sheet named `接龙列表(行排不合并)` can be found), the basic orders info would be displayed on the right pane. And the export button would be enabled:

![app-file-imported](https://github.com/heartace/OrderSheetConverterDesktop/assets/669206/782f6642-23f4-4f5c-b2d5-98fa2643bdb0)

Then click the export button to generate a new `xlsx` file. The default file name is the original file name with `-converted` appended.

## Technical Details

- Programming language: C#
- Target framework: .NET framework 6.0
- IDE: Visual Studio Community 2022

By default, the project will be published to a folder instead of generating an installer. The publish config file can be found [here](OrderSheetConverter/Properties/PublishProfiles/FolderProfile.pubxml). To publish it via command line, run the following command:

```shell
dotnet publish OrderSheetConverter.sln /p:PublishProfile=OrderSheetConverter\Properties\PublishProfiles\FolderProfile.pubxml
```

### Third-party Libraries

- [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) for reading the `xls` file.
- [EPPlus 6](https://github.com/EPPlusSoftware/EPPlus) for generating the `xlsx` file.

## App Logo

The app logo, a black cat, is downloaded from [Noun Project](https://thenounproject.com/icon/black-cat-6064070/).

![blackcat-resized](https://github.com/heartace/OrderSheetConverterDesktop/assets/669206/a8731a5e-8bcb-418b-a63d-ef2d7e9098db)
