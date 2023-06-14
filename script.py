import datetime

import pandas as pd


def process_excel_file(input_file, output_file):
    try:
        df = pd.read_excel(input_file, sheet_name="input_refresh_template")
    except (ValueError,FileNotFoundError) as e:
        print(f"ERROR:{e}")
        return

    # Create a new DataFrame with the desired columns
    new_df = pd.DataFrame(columns=['Day of Month', 'Date', 'Site ID', 'Page Views', 'Unique Visitors', 'Total Time Spent', 'Visits', 'Average Time Spent on Site'])

    if not len(df):
        print("Error: No data found..")
        return
    
    if not isinstance(df.loc[0][0],datetime.datetime):
        print("Error: Wrong formatting..")
        return

    # Iterate over the rows of the original DataFrame
    for index in range(2,len(df),3):
        site_index = index
        day_to_increment = df.loc[0][0].day
        site_id = df.loc[site_index][0]    

        if not isinstance(site_id,str):
            print("Error: Wrong formatting..")
            return
        max_col = day_to_increment*5
        for col_index in range(day_to_increment):
            col = col_index+1
            date = df.loc[site_index-1][col]

            page_view  = df.loc[site_index][col]

            col = col+ day_to_increment
            unique_visitors  = df.loc[site_index][col]

            col = col+day_to_increment
            total_time_spent  = df.loc[site_index][col]

            col = col + day_to_increment
            visits  = df.loc[site_index][col]

            col = col + day_to_increment

            col = min(col, max_col)  # Limit the column index to avoid going beyond the DataFrame size
            avg_time_spent  = df.loc[site_index][col]

            new_df = new_df._append({
                    'Day of Month': date.day,
                    'Date': str(date.date()),
                    'Site ID': site_id,
                    'Page Views': page_view,
                    'Unique Visitors': unique_visitors,
                    'Total Time Spent': total_time_spent,
                    'Visits': visits,
                    'Average Time Spent on Site': avg_time_spent
                }, ignore_index=True)
        

    # Save the new DataFrame to a new Excel file
    new_df.to_excel(output_file, index=False)
    print(f"Script Run Successfully output file:{output_file}")

if __name__ == '__main__':
    input_file = 'Analytics Template for Exercise.xlsx'
    output_file = 'output.xlsx'
    process_excel_file(input_file, output_file)