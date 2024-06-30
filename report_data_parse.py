# import pandas as pd
#
# # Load the new datasets
# new_data_file_path = 'CSVs/fch2024_calculated_scores_data.csv'
# new_demographic_file_path = 'CSVs/FCH May 2024 wb².csv'
#
# new_data = pd.read_csv(new_data_file_path)
# new_demographics = pd.read_csv(new_demographic_file_path)
#
# # Redefine metrics_columns to include only numeric columns
# numeric_metrics_columns = new_data.select_dtypes(include=['float64', 'int64']).columns.tolist()
# demographic_column = 'Demographic 1'
# department_column = 'groupName'
#
# # Calculate the total members and completed members per department
# department_counts = new_demographics.groupby(department_column)['userId'].count().reset_index(name='total_members')
# completed_counts = new_data.groupby(demographic_column)['userId'].count().reset_index(name='completed_members')
#
# completion_rate = pd.merge(department_counts, completed_counts, left_on=department_column, right_on=demographic_column, how='left')
# completion_rate['completed_members'] = completion_rate['completed_members'].fillna(0)
# completion_rate['completion_rate'] = (completion_rate['completed_members'] / completion_rate['total_members']) * 100
#
# # Calculate overall completion rate
# total_members = new_demographics['userId'].count()
# completed_members = new_data['userId'].count()
# overall_completion_rate = (completed_members / total_members) * 100
#
# # Calculate average scores for each demographic group
# average_scores_by_demographic = new_data.groupby(demographic_column)[numeric_metrics_columns].mean().reset_index()
#
# # Calculate overall average scores when demographic is not present
# overall_average_scores = new_data[numeric_metrics_columns].mean().reset_index().rename(columns={0: 'Metric', 1: 'Score'})
#
# # Add a row for the overall average scores
# overall_average_scores = overall_average_scores.T
# overall_average_scores.columns = overall_average_scores.iloc[0]
# overall_average_scores = overall_average_scores[1:].reset_index(drop=True)
# overall_average_scores[demographic_column] = 'Org Total'
#
# # Append the overall average scores to the average scores by demographic
# average_scores_by_demographic = pd.concat([average_scores_by_demographic, overall_average_scores], ignore_index=True)
#
# # Identify outliers using Z-score
# z_scores = (new_data[numeric_metrics_columns] - new_data[numeric_metrics_columns].mean()) / new_data[numeric_metrics_columns].std()
# outliers = z_scores[(z_scores > 3) | (z_scores < -3)].dropna(how='all').reset_index()
#
# # Identify the 3 lowest scores for each demographic group
# def find_lowest_scores(row):
#     scores = row[numeric_metrics_columns].astype(float)
#     lowest_scores = scores.nsmallest(3)
#     return pd.Series([f"{metric}: {score}" for metric, score in lowest_scores.items()])
#
# lowest_scores_by_demographic = average_scores_by_demographic.apply(find_lowest_scores, axis=1)
# lowest_scores_by_demographic[demographic_column] = average_scores_by_demographic[demographic_column]
#
# # Save initial report to Excel file
# initial_report_path = 'CSVs/Report.xlsx'
# with pd.ExcelWriter(initial_report_path) as writer:
#     average_scores_by_demographic.to_excel(writer, sheet_name='Average Scores by Demographic', index=False)
#     completion_rate.to_excel(writer, sheet_name='Completion Rate by Department', index=False)
#     outliers.to_excel(writer, sheet_name='Outliers', index=False)
#
# # Reopen the Excel file and round all numeric values
# xlsx = pd.ExcelFile(initial_report_path)
# average_scores_by_demographic = pd.read_excel(xlsx, sheet_name='Average Scores by Demographic')
# completion_rate = pd.read_excel(xlsx, sheet_name='Completion Rate by Department')
# outliers = pd.read_excel(xlsx, sheet_name='Outliers')
#
# # Round all numeric values
# average_scores_by_demographic = average_scores_by_demographic.round(1)
# completion_rate = completion_rate.round(1)
# outliers = outliers.round(1)
#
# # Save the final rounded report to a new Excel file
# final_report_path = 'CSVs/Report.xlsx'
# with pd.ExcelWriter(final_report_path) as writer:
#     average_scores_by_demographic.to_excel(writer, sheet_name='Average Scores by Demographic', index=False)
#     completion_rate.to_excel(writer, sheet_name='Completion Rate by Department', index=False)
#     outliers.to_excel(writer, sheet_name='Outliers', index=False)
#
# # Print overall completion rate
# overall_completion_rate = round(overall_completion_rate, 0)
# print("Overall Completion Rate:", overall_completion_rate)
# print("Generated report file:", final_report_path)
#
# # Load the file
# xls = pd.ExcelFile(final_report_path)
#
# # Load the first sheet
# df = xls.parse(sheet_name=0)
#
# # Convert all columns to numeric, forcing errors to NaN
# df_numeric = df.apply(pd.to_numeric, errors='coerce')
#
# # Create a new DataFrame to hold the results
# results = pd.DataFrame(columns=['Department', 'Lowest Influencer', 'Second Lowest Influencer', 'Third Lowest Influencer'])
#
# # Process each row to find the lowest 3 scores and their column names
# for index, row in df_numeric.iterrows():
#     lowest_scores = row.nsmallest(3)
#     results = pd.concat([results, pd.DataFrame({
#         'Department': [df.iloc[index, 0]],  # Assuming the first column contains the department names
#         'Lowest Influencer': [f"{lowest_scores.index[0]}: {lowest_scores.iloc[0]}"],
#         'Second Lowest Influencer': [f"{lowest_scores.index[1]}: {lowest_scores.iloc[1]}"],
#         'Third Lowest Influencer': [f"{lowest_scores.index[2]}: {lowest_scores.iloc[2]}"]
#     })], ignore_index=True)
#
# # Append the new sheet to the xlsx file
# with pd.ExcelWriter(final_report_path, mode='a', engine='openpyxl') as writer:
#     results.to_excel(writer, sheet_name='Lowest Scores', index=False)
#
# results.head()
#

import pandas as pd

def create_xlsx_report(demographics_df, raw_data_df, output_file_path):
    # Redefine metrics_columns to include only numeric columns
    numeric_metrics_columns = raw_data_df.select_dtypes(include=['float64', 'int64']).columns.tolist()
    demographic_column = 'Demographic 1'
    department_column = 'groupName'

    # Calculate the total members and completed members per department
    department_counts = demographics_df.groupby(department_column)['userId'].count().reset_index(name='total_members')
    completed_counts = raw_data_df.groupby(demographic_column)['userId'].count().reset_index(name='completed_members')

    completion_rate = pd.merge(department_counts, completed_counts, left_on=department_column, right_on=demographic_column, how='left')
    completion_rate['completed_members'] = completion_rate['completed_members'].fillna(0)
    completion_rate['completion_rate'] = (completion_rate['completed_members'] / completion_rate['total_members']) * 100

    # Calculate overall completion rate
    total_members = demographics_df['userId'].count()
    completed_members = raw_data_df['userId'].count()
    overall_completion_rate = (completed_members / total_members) * 100

    # Calculate average scores for each demographic group
    average_scores_by_demographic = raw_data_df.groupby(demographic_column)[numeric_metrics_columns].mean().reset_index()

    # Calculate overall average scores when demographic is not present
    overall_average_scores = raw_data_df[numeric_metrics_columns].mean().reset_index().rename(columns={0: 'Metric', 1: 'Score'})

    # Add a row for the overall average scores
    overall_average_scores = overall_average_scores.T
    overall_average_scores.columns = overall_average_scores.iloc[0]
    overall_average_scores = overall_average_scores[1:].reset_index(drop=True)
    overall_average_scores[demographic_column] = 'Org Total'

    # Append the overall average scores to the average scores by demographic
    average_scores_by_demographic = pd.concat([average_scores_by_demographic, overall_average_scores], ignore_index=True)

    # Identify outliers using Z-score
    z_scores = (raw_data_df[numeric_metrics_columns] - raw_data_df[numeric_metrics_columns].mean()) / raw_data_df[numeric_metrics_columns].std()
    outliers = z_scores[(z_scores > 3) | (z_scores < -3)].dropna(how='all').reset_index()

    # Identify the 3 lowest scores for each demographic group
    def find_lowest_scores(row):
        scores = row[numeric_metrics_columns].astype(float)
        lowest_scores = scores.nsmallest(3)
        return pd.Series([f"{metric}: {score}" for metric, score in lowest_scores.items()])

    lowest_scores_by_demographic = average_scores_by_demographic.apply(find_lowest_scores, axis=1)
    lowest_scores_by_demographic[demographic_column] = average_scores_by_demographic[demographic_column]

    # Save initial report to Excel file
    with pd.ExcelWriter(output_file_path) as writer:
        average_scores_by_demographic.to_excel(writer, sheet_name='Average Scores by Demographic', index=False)
        completion_rate.to_excel(writer, sheet_name='Completion Rate by Department', index=False)
        outliers.to_excel(writer, sheet_name='Outliers', index=False)

    with pd.ExcelFile(output_file_path) as xlsx:
        average_scores_by_demographic = pd.read_excel(xlsx, sheet_name='Average Scores by Demographic').round(1)
        completion_rate = pd.read_excel(xlsx, sheet_name='Completion Rate by Department').round(1)
        outliers = pd.read_excel(xlsx, sheet_name='Outliers').round(1)

    # Save the final rounded report to a new Excel file
    with pd.ExcelWriter(output_file_path, mode='w') as writer:
        average_scores_by_demographic.to_excel(writer, sheet_name='Average Scores by Demographic', index=False)
        completion_rate.to_excel(writer, sheet_name='Completion Rate by Department', index=False)
        outliers.to_excel(writer, sheet_name='Outliers', index=False)

    # Print overall completion rate
    overall_completion_rate = round(overall_completion_rate, 0)
    print("Overall Completion Rate:", overall_completion_rate)
    print("Generated report file:", output_file_path)

    # Load the file
    xls = pd.ExcelFile(output_file_path)

    # Load the first sheet
    df = xls.parse(sheet_name=0)

    # Convert all columns to numeric, forcing errors to NaN
    df_numeric = df.apply(pd.to_numeric, errors='coerce')

    # Create a new DataFrame to hold the results
    results = pd.DataFrame(columns=['Department', 'Lowest Influencer', 'Second Lowest Influencer', 'Third Lowest Influencer'])

    # Process each row to find the lowest 3 scores and their column names
    for index, row in df_numeric.iterrows():
        lowest_scores = row.nsmallest(3)
        results = pd.concat([results, pd.DataFrame({
            'Department': [df.iloc[index, 0]],  # Assuming the first column contains the department names
            'Lowest Influencer': [f"{lowest_scores.index[0]}: {lowest_scores.iloc[0]}"],
            'Second Lowest Influencer': [f"{lowest_scores.index[1]}: {lowest_scores.iloc[1]}"],
            'Third Lowest Influencer': [f"{lowest_scores.index[2]}: {lowest_scores.iloc[2]}"]
        })], ignore_index=True)

    # Append the new sheet to the xlsx file
    with pd.ExcelWriter(output_file_path, mode='a', engine='openpyxl') as writer:
        results.to_excel(writer, sheet_name='Lowest Scores', index=False)
