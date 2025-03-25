# BALANCE

**Bilateral Accounting Ledger for Analyzing Networked Couple Expenses**

## Overview

BALANCE is an Excel VBA application designed to help couples track, analyze, and balance shared expenses. The application provides a modern dashboard interface with data visualization, transaction management, and automated balance calculations to ensure fair expense sharing.

## Features

- **CSV Import**: Import transactions from various CSV formats with auto-detection
- **Dynamic Dashboard**: Visual overview of spending patterns and balance status
- **Transaction Management**: Add, edit, and categorize transactions
- **Balance Calculation**: Automatic calculation of who owes whom based on shared expenses
- **Data Visualization**: Charts for spending by category, monthly trends, and day-of-week analysis
- **Financial Insights**: Automated analysis of spending patterns with recommendations
- **Multi-User Support**: Track expenses for two users with customizable names
- **Modern UI**: Clean, card-based interface with intuitive controls

## Getting Started

### Prerequisites
- Microsoft Excel (2010 or newer)
- Macros must be enabled

### Installation
1. Download the BALANCE.xlsm workbook
2. Open the workbook in Excel
3. If prompted, click "Enable Macros"
4. Run the `InitializeBALANCE` macro to set up the application

### Initial Setup
1. Go to Settings to configure user names and default paths
2. Import transactions or create sample data for testing

## Usage

### Importing Transactions
1. Click "Import CSV" on the dashboard
2. Choose to import a single file or all CSVs from a folder
3. Select the expense owner
4. Review import results

### Managing Transactions
1. Click "Edit Transactions" to view and modify all transactions
2. Edit values directly in the sheet
3. Click "Save Changes" to update the data

### Viewing Insights
1. Navigate to the Insights sheet to see automated analysis
2. Review spending patterns, trends, and recommendations

## Project Structure

- **Models**: Core data structures (Transaction)
- **Repositories**: Data storage and retrieval (TransactionRepository)
- **Services**: Business logic (BalanceCalculator, CSVImportEngine, TransactionAnalyzer)
- **UI**: User interface components (DashboardManager, ChartFactory)
- **Utils**: Utility functions and services (AppSettings, ErrorLogger, Utilities)

## License

This project is available for personal use.

## Acknowledgments

Built with VBA and Excel's charting capabilities.