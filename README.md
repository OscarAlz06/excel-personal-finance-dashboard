# Excel Personal Finance Dashboard

## Overview

This project is a personal finance management and analysis system built in Microsoft Excel using VBA automation. The goal of the system is to provide a structured way to record financial transactions, analyze spending patterns, and monitor monthly budgets through an interactive dashboard.

The solution integrates automated data entry, data processing using pivot tables, and visual analytics to support better financial decision-making.

Monetary values shown in the repository images have been anonymized to protect personal financial information.

## Features

- Automated transaction entry using VBA macros
- Structured financial transaction log
- Monthly financial analysis dashboard
- Budget tracking with execution percentage indicators
- Data visualization using charts and conditional formatting
- Backend processing using pivot tables

## System Architecture

The workbook is organized into three main functional components.

### RECORD

This sheet is responsible for storing all financial transactions.

Data is entered through a VBA macro that prompts the user to provide the following information:

- Date
- Description
- Transaction type (Income or Expense)
- Amount

The macro validates the input and automatically inserts the information into a structured table.

### ANALYSIS

This sheet provides a dynamic financial analysis interface. The user can select a specific month to visualize financial activity.

The analysis includes:

- Tables summarizing income, fixed expenses, and variable expenses
- Proportional charts to show the distribution of spending categories
- In-cell data bars to visually compare expense magnitudes
- A time-series chart comparing total income and total expenses by month

The analysis layer is powered by pivot tables located in hidden sheets, which function as the backend data processing layer.

### BUDGET

The budget sheet is designed to track planned spending against actual expenses.

For each category the system calculates:

- Planned monthly budget
- Actual spending
- Budget execution percentage

Conditional formatting and visual indicators highlight deviations and help identify categories that exceed the defined budget.


## Technologies Used

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Pivot Tables
- Excel Charts
- Conditional Formatting


## Repository Structure
excel-personal-finance-dashboard

│

├── excel

│ └── finance_tracker.xlsm

│

├── video

│ └── demo.mp4

│

└── README.md

