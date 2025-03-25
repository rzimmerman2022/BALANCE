# BALANCE Architecture

This document describes the architecture and design of the BALANCE application.

## System Overview

BALANCE follows a modular, object-oriented architecture that separates concerns into distinct components:

BALANCE
│
├── Models - Core data objects
│   └── Transaction
│
├── Repositories - Data access layer
│   └── TransactionRepository
│
├── Services - Business logic
│   ├── BalanceCalculator
│   ├── CSVImportEngine
│   └── TransactionAnalyzer
│
├── UI - Presentation layer
│   ├── DashboardManager
│   └── ChartFactory
│
└── Utils - Common utilities
├── AppSettings
├── ErrorLogger
└── Utilities

## Component Details

### Models

#### Transaction (Transaction.cls)
- **Purpose**: Represents a financial transaction with all its properties and behaviors
- **Key features**:
  - Properties: ID, Date, Merchant, Category, Amount, Account, Owner, IsShared, Notes, SourceFile
  - Methods for calculating splits between users
  - Methods for determining transaction type (expense/income)
  - Validation and data cleaning logic

### Repositories

#### TransactionRepository (TransactionRepository.vb)
- **Purpose**: Stores and retrieves transaction data
- **Key features**:
  - In-memory collection of Transaction objects
  - Persistence to worksheet
  - Transaction filtering by various criteria
  - Duplicate detection and resolution
  - CRUD operations (Create, Read, Update, Delete)

### Services

#### BalanceCalculator (BalanceCalculator.cls)
- **Purpose**: Calculates balances and analyzes spending patterns
- **Key features**:
  - Balance calculation between users
  - Category analysis with percentage breakdown
  - Monthly spending summaries
  - Day-of-week spending analysis
  - Performance-optimized sorting algorithms

#### CSVImportEngine (CSVImportEngine.cls)
- **Purpose**: Imports transactions from CSV files
- **Key features**:
  - CSV format auto-detection
  - Support for multiple file formats
  - CSV parsing with quotes and special character handling
  - Validation and error reporting
  - Bulk import capabilities

#### TransactionAnalyzer (TransactionAnalyzer.cls)
- **Purpose**: Provides advanced analytics and insights
- **Key features**:
  - Spending trend analysis
  - Outlier detection
  - Automated recommendations
  - Insight generation and visualization

### UI Components

#### DashboardManager (DashboardManager.cls)
- **Purpose**: Manages the main dashboard interface
- **Key features**:
  - Dashboard layout management
  - Summary card creation and updates
  - UI control management
  - Dashboard refresh logic

#### ChartFactory (ChartFactory.cls)
- **Purpose**: Creates charts for data visualization
- **Key features**:
  - Creates category pie charts
  - Creates monthly column charts
  - Creates day-of-week charts
  - Creates balance comparison charts
  - Handles chart formatting and styling

### Utilities

#### AppSettings (AppSettings.cls)
- **Purpose**: Manages application configuration and settings
- **Key features**:
  - User name storage and retrieval
  - Color theme definitions
  - Path configurations
  - Persistence to settings sheet

#### ErrorLogger (ErrorLogger.cls)
- **Purpose**: Provides centralized error logging and reporting
- **Key features**:
  - Logs errors, warnings, and info messages
  - Multi-level logging with filtering
  - Export capabilities
  - Error aggregation and filtering

#### Utilities (Utilities.cls)
- **Purpose**: Provides common utility functions
- **Key features**:
  - Sheet management functions
  - File system operations
  - UI helper functions
  - Date and string utilities

## Data Flow

1. **Transaction Import**:
   - CSVImportEngine parses CSV files into Transaction objects
   - Transactions are validated and added to TransactionRepository
   - DashboardManager updates UI to reflect new data

2. **Balance Calculation**:
   - User actions trigger balance calculation via BalanceCalculator
   - BalanceCalculator processes transactions and determines who owes whom
   - Results are displayed on dashboard via DashboardManager

3. **Data Visualization**:
   - ChartFactory requests data from BalanceCalculator
   - Data is processed into chart-friendly format
   - Charts are created and displayed on the dashboard

4. **Insights Generation**:
   - TransactionAnalyzer processes transaction data
   - Patterns and insights are identified
   - Results are formatted and displayed on the Insights sheet

## Design Patterns

1. **Singleton Pattern**: Used in service classes (BalanceCalculator, CSVImportEngine, etc.) to ensure only one instance exists
2. **Repository Pattern**: Used for data access abstraction in TransactionRepository
3. **Factory Pattern**: Used in ChartFactory for chart creation
4. **Type Object Pattern**: Used for returning structured data (BalanceSummary, CategorySummary, etc.)
5. **Service Layer Pattern**: Used to separate business logic from data access and presentation

## Error Handling Strategy

BALANCE implements a centralized error handling approach:
- All modules use structured error handling with On Error GoTo statements
- Errors are logged through the ErrorLogger class
- Critical errors are reported to users via message boxes
- Non-critical errors are logged for troubleshooting
- Each function includes error exits that maintain application stability

## Performance Considerations

- The application uses optimized algorithms (QuickSort vs. BubbleSort)
- Data is maintained in memory for faster access
- UI updates are batched to reduce Excel's recalculation overhead
- Sheet operations are minimized to improve performance