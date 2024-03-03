# AzureDataFactoryTriggerExtractor

AzureDataFactoryTriggerExtractor is a command-line tool written in C# that extracts trigger recurrences from Azure Data Factory and exports them to an Excel file. It uses Azure SDK for .NET and OfficeOpenXml library to connect to Azure Data Factory, retrieve trigger information, and generate an Excel report.

## Features

- Extracts trigger recurrences from Azure Data Factory.
- Exports trigger information to an Excel file for analysis.
- Uses Azure SDK for .NET for Azure Data Factory integration.
- Supports authentication using client secret.

## Prerequisites

- .NET 5.0 SDK or higher
- Azure subscription
- Azure Data Factory instance

## Installation

1. Clone the repository: `git clone https://github.com/your-repo/AzureDataFactoryTriggerExtractor.git`
2. Navigate to the project directory: `cd AzureDataFactoryTriggerExtractor`

## Usage

1. Set the required environment variables: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `SUBSCRIPTION_ID`, `RESOURCE_GROUP_NAME`, `FACTORY_NAME`.
2. Build the project: `dotnet build`
3. Run the tool: `dotnet run`

The generated Excel file will contain the extracted trigger recurrences from Azure Data Factory.

## License

This project is licensed under the MIT License - see the LICENSE.md file for details.
