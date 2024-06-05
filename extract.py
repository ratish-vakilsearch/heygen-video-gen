# Import libraries
import pandas as pd
import json

def get_excel_value(file_path, sheet_name, row, column, in_thousands=True):
    """
    Retrieve a value from an Excel sheet based on specified row and column indices.
    Optionally convert the value to thousands and round to three decimal places.

    :param file_path: str, path to the Excel file
    :param sheet_name: str, name of the sheet to retrieve the value from
    :param row: int, row index (starting from 0)
    :param column: str, column letter
    :param in_thousands: bool, whether to convert the value to thousands
    :return: value at the specified row and column
    """
    try:
        # Load the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Convert column letter to column index
        col_index = ord(column.upper()) - ord('A')

        # Retrieve the value
        value = df.iloc[row - 1, col_index]

        # Convert to thousands and round to three decimal places if requested
        if in_thousands:
            value = round(value / 1000, 3)

        return value

    except Exception as e:
        print(f"An error occurred when retrieving value from row {row}, column {column}: {e}")
        return None

def slide_2(file_path):
    sheet_name = '03 BS'

    non_current_assets = str(get_excel_value(file_path, sheet_name, 45, 'G'))
    current_assets = str(get_excel_value(file_path, sheet_name, 53, 'G'))
    total_equity = str(get_excel_value(file_path, sheet_name, 12, 'G'))

    non_current_liabilities = get_excel_value(file_path, sheet_name, 23, 'G')
    current_liabilities = get_excel_value(file_path, sheet_name, 31, 'G')

    if non_current_liabilities is None or current_liabilities is None:
        total_liabilities = None
    else:
        total_liabilities = str(round(non_current_liabilities + current_liabilities, 3))

    return {
            "title": "2024 Balance Sheet Overview-(Rs in Thousands)",
            "image_type": "Circle-chart",
            "data": [
                {"key": "Non-current assets", "value": non_current_assets},
                {"key": "Current assets", "value": current_assets},
                {"key": "Total equity", "value": total_equity},
                {"key": "Total liabilities", "value": total_liabilities}
            ]
        }

def slide_3(file_path):
    sheet_name = '03 BS'

    shareholders_funds = str(get_excel_value(file_path, sheet_name, 12, 'G'))

    non_current_liabilities = get_excel_value(file_path, sheet_name, 23, 'G')
    current_liabilities = get_excel_value(file_path, sheet_name, 31, 'G')

    if non_current_liabilities is None or current_liabilities is None:
        gross_liabilities = None
    else:
        gross_liabilities = str(non_current_liabilities + current_liabilities)

    return {
            "title": "2024 Balance Sheet Overview - (Rs in Thousands)",
            "image_type": "None",
            "data": [
                {"key": "Shareholders' funds", "value": shareholders_funds, 
                "percent":round((eval(shareholders_funds)/(eval(shareholders_funds)+eval(gross_liabilities)))*100,2)
                },
                {"key": "Gross liabilities", 
                "value": gross_liabilities, "percent":round((eval(gross_liabilities)/(eval(shareholders_funds)+eval(gross_liabilities)))*100,2)
                }
            ]
        }

def slide_4(file_path):
    trade_payables = get_excel_value(file_path, '03 BS', 28, 'G')
    statutory_payables = get_excel_value(file_path, '06 Sch 2-8', 92, 'C')
    salary_payable = get_excel_value(file_path, '06 Sch 2-8', 95, 'C')
    other_payables = get_excel_value(file_path, '03 BS', 22, 'G') + get_excel_value(file_path, '06 Sch 2-8', 89, 'C')

    return {
            "title": "2024 Balance Sheet Overview - (Rs in Thousands)",
            "image_type":"circle-chart",
            "note":"Trade payables+Statutory payables+salary payable = Total current liabilities + other payables = Total liabilities",
            "data": [
                {"key": "Trade payables", "value": str(trade_payables)},
                {"key": "Statutory payables", "value": str(statutory_payables)},
                {"key": "Salary payable", "value": str(salary_payable)},
                {"key": "Other payables", "value": str(other_payables)}
            ]
        }

def slide_5(file_path):
    sheet_name = '03 BS'

    current_assets = get_excel_value(file_path, sheet_name, 53, 'G')
    current_liabilities = get_excel_value(file_path, sheet_name, 31, 'G')

    working_capital = current_assets - current_liabilities

    return {
            "title": "2024 Balance Sheet Overview - (Rs in Thousands)",
            "image_type":"pie-chart",
            "data": [
                {"key": "item1", "name": "Current assets", "value": str(current_assets)},
                {"key": "item2", "name": "Current liabilities", "value": str(current_liabilities)},
                {"key": "item3", "name": "Working capital", "value": str(working_capital)}
            ]
        }

def slide_6(file_path):
    sheet_name = '04 PL'

    revenue_22_23 = get_excel_value(file_path, sheet_name, 13, 'F')
    revenue_23_24 = get_excel_value(file_path, sheet_name, 13, 'E')

    expenses_22_23 = get_excel_value(file_path, sheet_name, 23, 'F')
    expenses_23_24 = get_excel_value(file_path, sheet_name, 23, 'E')

    pbt_22_23 = abs(get_excel_value(file_path, sheet_name, 25, 'F'))
    pbt_23_24 = get_excel_value(file_path, sheet_name, 25, 'E')

    pat_22_23 = abs(get_excel_value(file_path, sheet_name, 34, 'F'))
    pat_23_24 = get_excel_value(file_path, sheet_name, 34, 'E')

    return {
            "title": "Income statement summary - Here's a summary of our performance improvements in 2024",
            "image_type":"bar-chart",
            "data": [
                {"key": "Revenue", "FY 22-23": str(revenue_22_23), "FY 23-24": str(revenue_23_24)},
                {"key": "Expenses", "FY 22-23": str(expenses_22_23), "FY 23-24": str(expenses_23_24)},
                {"key": "PBT", "FY 22-23": str(pbt_22_23), "FY 23-24": str(pbt_23_24)},
                {"key": "PAT", "FY 22-23": str(pat_22_23), "FY 23-24": str(pat_23_24)}
            ]
        }

def slide_7(file_path):
    sheet_name = '04 PL'

    # Define columns for different metrics
    revenue_col = 'E'
    cost_of_services_col = 'E'
    employee_benefit_expenses_col = 'E'
    other_expenses_col = 'E'

    # Define row indices for different metrics
    revenue_row = 13
    cost_of_services_row1 = 16
    cost_of_services_row2 = 17
    employee_benefit_expenses_row = 18
    other_expenses_row1 = 19
    other_expenses_row2 = 20
    other_expenses_row3 = 21

    # Retrieve values for different metrics
    revenue = get_excel_value(file_path, sheet_name, revenue_row, revenue_col)

    # Calculate cost of services
    cost_of_services = get_excel_value(file_path, sheet_name, cost_of_services_row1, cost_of_services_col) + \
                       get_excel_value(file_path, sheet_name, cost_of_services_row2, cost_of_services_col)

    employee_benefit_expenses = get_excel_value(file_path, sheet_name, employee_benefit_expenses_row, employee_benefit_expenses_col)

    # Calculate other expenses
    other_expenses = get_excel_value(file_path, sheet_name, other_expenses_row1, other_expenses_col) + \
                     get_excel_value(file_path, sheet_name, other_expenses_row2, other_expenses_col) + \
                     get_excel_value(file_path, sheet_name, other_expenses_row3, other_expenses_col)

    return {
            "title": "Statement of profit and loss for FY 2023-24",
            "image_type":"column chart",
            "data": [
                {"key": "item1", "name": "Revenue", "value": str(revenue)},
                {"key": "item2", "value": ""},
                {"key": "item3", "name": "Cost of services", "value": str(cost_of_services)},
                {"key": "item4", "name": "Employee benefit expenses", "value": str(employee_benefit_expenses)},
                {"key": "item5", "name": "Other expenses", "value": str(other_expenses)}
            ]
        }

def get_data(file_path):
  # file_path = '/content/PISCIS Network Pvt Ltd Financials March 24 final (1).xlsx'  # Update the path to the local path of your Excel file

  # Call all the functions and store their outputs
  slide_2_output = slide_2(file_path)
  slide_3_output = slide_3(file_path)
  slide_4_output = slide_4(file_path)
  slide_5_output = slide_5(file_path)
  slide_6_output = slide_6(file_path)
  slide_7_output = slide_7(file_path)

  combined_output = {
      "balance_sheet_1": slide_2_output,
      "balance_sheet_2": slide_3_output,
      "balance_sheet_3": slide_4_output,
      "balance_sheet_4": slide_5_output,
      "balance_sheet_5": slide_6_output,
      "balance_sheet_6": slide_7_output
  }


  return json.dumps(combined_output, indent=4)

'''op = Main('/content/PISCIS Network Pvt Ltd Financials March 24 final (1).xlsx')
print(op)'''