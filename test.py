import os
import pandas as pd
from script import process_excel_file

def test_process_excel_file():
    input_file = 'Analytics Template for Exercise.xlsx'
    output_file = 'test_output.xlsx'

    # Call the function
    process_excel_file(input_file, output_file)

    # Assert that the output file was created
    assert os.path.exists(output_file)

    # Assert that the output file has the expected columns
    expected_columns = ['Day of Month', 'Date', 'Site ID', 'Page Views', 'Unique Visitors', 'Total Time Spent', 'Visits', 'Average Time Spent on Site']
    df = pd.read_excel(output_file)
    assert list(df.columns) == expected_columns

    # Assert that the output file has the expected number of rows
    assert len(df) == 3100

    # Clean up - delete the output file
    os.remove(output_file)


def test_process_excel_file_value_error():
    input_file = 'Invalid_File.xlsx'
    output_file = 'test_output.xlsx'

    # Call the function
    captured_output=process_excel_file(input_file, output_file)

    # Assert that the function return None
    assert not captured_output

    # Assert that the output file was not created
    assert not os.path.exists(output_file)