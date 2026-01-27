# EIR Income Projection Tool

A React-based tool for calculating and projecting EIR (Effective Interest Rate) income for loan portfolios.

## Features

- Support for multiple product types (Other Products, Tractor)
- Multiple repayment frequencies (Monthly, Bimonthly, Quarterly, Halfyearly)
- Batch processing of multiple loan cases via Excel upload
- Detailed calculation steps (Step A, B, C, D)
- Excel export with consolidated and individual sheets

## Installation

```bash
npm install
```

## Running the Application

```bash
npm start
```

The application will open at `http://localhost:3000`

## Building for Production

```bash
npm run build
```

## Usage

1. Select Product Type (Other Products or Tractor)
2. Download the template Excel file
3. Fill in loan details in the template
4. Upload the filled Excel file
5. Click "Process File" to calculate EIR projections
6. Download the output Excel file with results

## Input Fields

- Agreement Number
- Product Type
- Repayment Frequency
- Amount Financed
- Tenure (Months)
- Disbursement Date
- Amort IRR (%)
- Advance EMI
- Upfront Income
- Upfront Expense

## Output

The tool generates an Excel file with:
- All Cases (consolidated sheet)
- Individual sheets for each loan case
- Columns: Agreement Number, Product Type, Amount Financed, Tenure, Repayment Frequency, No. of Installments, Disbursement, First EMI, Last EMI, Amort IRR, Advance EMI, Upfront Income, Upfront Expense, Month, Step C EIR Income, EIR Income
