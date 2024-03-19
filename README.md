# JSON to Excel Formatter for C# Projects

## Project Overview

This project contains a C# application that reads a JSON file containing database table schemas and their respective properties, then processes this data and generates a formatted Excel spreadsheet. The Excel file groups the properties under their respective tables, providing a clear and organized view of the table structures defined in the JSON.

## Features

- **JSON Parsing**: The application parses JSON data using the Newtonsoft.Json library.
- **Excel File Creation**: Utilizing the EPPlus library, the application generates an Excel file from the JSON data.
- **Formatted Output**: Table names are presented in separate rows with an empty row above and below them for clear visual separation.

## Getting Started

### Prerequisites

- .NET Framework (version as per your development environment)
- EPPlus library (for Excel file manipulation)
- Newtonsoft.Json library (for JSON parsing)

### Installation

1. Clone the repository using Git:
  `git clone https://github.com/osama-abu-baker/JsonToExcelFormatter.git`

2. Navigate to the cloned repository.

3. Install the necessary NuGet packages:
   `dotnet add package EPPlus`
   `dotnet add package Newtonsoft.Json`

### Usage

1. Add your JSON file to the project's root directory.

2. Modify the `jsonFilePath` and `excelFilePath` variables in the `Program.cs` file to point to your JSON file and your desired output Excel file path.

3. Run the program:
`dotnet run`

## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.
