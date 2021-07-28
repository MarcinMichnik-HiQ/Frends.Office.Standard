# Frends.Office.Standard

frends Community Task for WriteExcelFile

[![Actions Status](https://github.com/CommunityHiQ/Frends.Office.Standard/workflows/PackAndPushAfterMerge/badge.svg)](https://github.com/CommunityHiQ/Frends.Office.Standard/actions) ![MyGet](https://img.shields.io/myget/frends-community/v/Frends.Office.Standard) [![License: UNLICENSED](https://img.shields.io/badge/License-UNLICENSED-yellow.svg)](https://opensource.org/licenses/UNLICENSED) 

- [Installing](#installing)
- [Tasks](#tasks)
     - [WriteExcelFileTask](#WriteExcelFileTask)
     - [WriteWordFileTask](#WriteWordFileTask)
     - [ExportFileToSharepointTask](#ExportFileToSharepointTask)
- [Building](#building)
- [Contributing](#contributing)
- [Change Log](#change-log)

# Installing

You can install the Task via frends UI Task View or you can find the NuGet package from the following NuGet feed
https://www.myget.org/F/frends-community/api/v3/index.json and in Gallery view in MyGet https://www.myget.org/feed/frends-community/package/nuget/Frends.Office.Standard

# Tasks

## WriteExcelFileTask

Reads csv string, converts it to a DataTable and creates an excel file.

### Properties

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| StringInput | `string` | Input csv string | `"one;two;three\r\n1;2;3"` |
| CellDelimiter | `char` | Determines what character will be used for splitting based on cell in csv | `';'` |
| LineDelimiter | `string` | Determines what string will be used for splitting lines | `"\r\n"` |
| TargetPath | `string` | Full path of the target file to be written. File format should be .xlsx | `@"c:\temp\file.xlsx"` |

### Returns

JToken with 'message' and 'filePath' keys.

## WriteWordFileTask

Reads string and creates a word file.

### Properties

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| StringInput | `string` | Input text string | `"Paragraph\r\nNewLine"` |
| LineDelimiter | `string` | Determines what string will be used for splitting lines | `"\r\n"` |
| TargetPath | `string` | Full path of the target file to be written. File format should be .docx | `@"c:\temp\file.docx"` |

### Returns

JToken with keys: message, savedTo.

## ExportFileToSharepointTask

Finds a file at given path and sends it to sharepoint via Microsoft Graph API.

### Properties

#### FileExportInput

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| SourceFilePath | `string` | Path to a local file | `"c:\temp\file.xlsx"` |
| targetFolderName | `string` | Target folder path | `"General/Folder/"` |

#### Authentication

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| clientID | `string` | Azure Active Directory Site ID | `"1ce3f5e1-fc04-3f24-2c0e-v76d5b44b13c"` |
| clientSecret | `string` | Azure Active Directory client secret password | `"_Sgx6Jdi2NC1N27Z4_plRm55L-DeCWJ.yq"` |
| tenantID | `string` | Azure Active Directory tenant id | `"3d426023-5x12-4s11-afae-159b1865eabc"` |
| siteID | `string` | Azure Active Directory Site ID | `"f7b1c426-4x3c-4a7e-2129-296ed8449b49"` |

### Returns

JToken with keys: FileSize, Path, FileName, TargetFolderName, ClientID, TenantID, SiteID, DriveID, UploadUrl.

# Building

Clone a copy of the repository

`git clone https://github.com/CommunityHiQ/Frends.Office.Standard.git`

Rebuild the project

`dotnet build`

Run tests

`dotnet test`

Create a NuGet package

`dotnet pack --configuration Release`

# Contributing
When contributing to this repository, please first discuss the change you wish to make via issue, email, or any other method with the owners of this repository before making a change.

1. Fork the repository on GitHub
2. Clone the project to your own machine
3. Commit changes to your own branch
4. Push your work back up to your fork
5. Submit a Pull request so that we can review your changes

NOTE: Be sure to merge the latest from "upstream" before making a pull request!

# Change Log

| Version | Changes |
| ------- | ------- |
| 1.0.4   | .net standard working version |
