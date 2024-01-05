# Trx2Xlsx Converter

## Overview

Trx2Xlsx Converter is a simple command-line utility that converts Test Results XML (TRX) files to Excel (XLSX) format. It extracts relevant information from TRX files generated by Visual Studio test runs and organizes it into an Excel spreadsheet for better readability and analysis.

## Features

- Converts TRX files to XLSX format.
- Extracts information such as test name, test scenario, outcome, duration, start time, end time, and error messages.
- Applies color-coding to the "Outcome" column based on test results (Passed: Green, Failed: Red, Not Executed: Blue).
- Supports basic styling and formatting for improved presentation.

## Usage

### Prerequisites

- [.NET SDK](https://dotnet.microsoft.com/download) (8.0 or later)

### How to Use

1. **Clone the repository:**

   ```bash
   git clone https://github.com/olesme/trx2xlsx.git

2. **Navigate to the project directory:**

   ```bash
   cd Trx2Xlsx-Converter
   
3. **Build the project:**

   ```bash
   dotnet build
   
4. **Run the converter:**
   ```bash
   dotnet run -- <inputFileName.trx> <outputFileName.xlsx>
Replace <inputFileName.trx> with the path to your TRX file and <outputFileName.xlsx> with the desired name for the output Excel file.

## Acknowledgments
- This converter uses [EPPlus](https://github.com/JanKallman/EPPlus) for working with Excel files.

## Need some updates/improvements, want to use part of the code? Just fork it!

#### Authored by Oleksandr Menzerov (olesme)
