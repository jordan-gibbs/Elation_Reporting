import pandas as pd
import xlsxwriter

def create_xlsx_report(demographics_df, total_demo_df, raw_data_df, demographic_name):
    # Redefine metrics_columns to include only numeric columns
    numeric_metrics_columns = raw_data_df.select_dtypes(include=['float64', 'int64']).columns.tolist()

    # Assuming demographic_column is a Series from demographics_df
    demographic_column = demographics_df[demographic_name]

    # Initialize variables to track the column with the most matches
    best_match_column = None
    max_matches = 0

    best_match_column = None
    max_matches = 0

    # Iterate through the columns of total_demo_df
    for col in total_demo_df.columns:
        # Check if the column is of type 'object' and has less unique values than half the number of rows
        if total_demo_df[col].dtype == 'object' and len(total_demo_df[col].unique()) < len(total_demo_df) / 2:
            # Count the number of matching elements between demographic_column and the current column
            matches = sum(total_demo_df[col].isin(demographic_column))

            # Update the best match column if the current column has more matches
            if matches > max_matches:
                best_match_column = col
                max_matches = matches

    # Ensure that best_match_column is found
    if best_match_column is None:
        raise ValueError("No matching column found in total_demo_df.")

    # Now use best_match_column as demographic_column
    demographic_column = total_demo_df[best_match_column]

    # Example usage in your function, ensuring proper checks
    if not demographic_column.empty and demographic_column.any():
        # Proceed with your logic
        print(f"Best matching column: {best_match_column}")
    else:
        raise ValueError("Either demographic_column or department_column is empty or invalid.")

    department_column = best_match_column

    # Calculate the total members and completed members per department
    department_counts = total_demo_df.groupby(department_column)['userId'].count().reset_index(name='total_members')
    completed_counts = demographics_df[demographics_df['response'] == 'Y'].groupby(demographic_column)[
        'userId'].count().reset_index(
        name='completed_members')

    # Merge the results to get a complete view
    final_counts = pd.merge(department_counts, completed_counts, on=department_column, how='left')
    print(final_counts)


    completion_rate = pd.merge(department_counts, completed_counts, left_on=department_column,
                               right_on=department_column, how='left')
    completion_rate['completed_members'] = completion_rate['completed_members'].fillna(0)
    completion_rate['completion_rate'] = (completion_rate['completed_members'] / completion_rate['total_members']) * 100
    completion_rate = completion_rate.round(1)
    completion_rate.rename(columns={department_column: f'{demographic_name}'}, inplace=True)

    total_members = total_demo_df['userId'].count()
    completed_members = total_demo_df[total_demo_df['response'] == 'Y']['userId'].count()
    overall_completion_rate = (completed_members / total_members) * 100

    org_total_row = pd.DataFrame({
        demographic_name: ['Org Total'],
        'total_members': [total_members],
        'completed_members': [completed_members],
        'completion_rate': [round(overall_completion_rate, 1)]
    })

    completion_rate = pd.concat([completion_rate, org_total_row], ignore_index=True)
    completion_rate['completion_rate'] = completion_rate['completion_rate'].apply(lambda x: f"{x}%")

    # Calculate overall completion rate
    total_members = demographics_df['userId'].count()
    completed_members = raw_data_df['userId'].count()
    overall_completion_rate = (completed_members / total_members) * 100

    # Calculate average scores for each demographic group
    average_scores_by_demographic = raw_data_df.groupby(demographic_name)[numeric_metrics_columns].mean().reset_index()
    average_scores_by_demographic = average_scores_by_demographic.round(0)

    # Calculate overall average scores when demographic is not present
    overall_average_scores = raw_data_df[numeric_metrics_columns].mean().reset_index().rename(
        columns={0: 'Metric', 1: 'Score'})

    # Add a row for the overall average scores
    overall_average_scores = overall_average_scores.round(0)
    overall_average_scores = overall_average_scores.T
    overall_average_scores.columns = overall_average_scores.iloc[0]
    overall_average_scores = overall_average_scores[1:].reset_index(drop=True)
    overall_average_scores[demographic_name] = 'Org Total'

    # Append the overall average scores to the average scores by demographic
    average_scores_by_demographic = pd.concat([average_scores_by_demographic, overall_average_scores],
                                              ignore_index=True)

    # Identify outliers using Z-score
    z_scores = (raw_data_df[numeric_metrics_columns] - raw_data_df[numeric_metrics_columns].mean()) / raw_data_df[
        numeric_metrics_columns].std()
    outliers = z_scores[(z_scores > 3) | (z_scores < -3)].dropna(how='all').reset_index()
    outliers = outliers.round(0)

    # Identify the 3 lowest scores for each demographic group
    def find_lowest_scores(row):
        scores = row[numeric_metrics_columns].astype(int)
        lowest_scores = scores.nsmallest(3)
        lowest_scores = lowest_scores.round(0)
        return pd.Series([f"{metric}: {score}" for metric, score in lowest_scores.items()])

    lowest_scores_by_demographic = average_scores_by_demographic.apply(find_lowest_scores, axis=1)
    lowest_scores_by_demographic[demographic_name] = average_scores_by_demographic[demographic_name]

    # Save the report to a BytesIO object
    from io import BytesIO
    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        average_scores_by_demographic.to_excel(writer, sheet_name=f'Avg. Scores for {demographic_name}', index=False)
        completion_rate.to_excel(writer, sheet_name=f'Completion Rate', index=False)
        outliers.to_excel(writer, sheet_name='Outliers', index=False)

        # Process each row to find the lowest 3 scores and their column names
        results = pd.DataFrame(
            columns=[demographic_name, 'Demographic Size', 'Lowest Influencer', 'Second Lowest Influencer'])

        # Predefined list of columns to examine
        org_influencers = [
            'Belonging in Organization',
            'Healthy Workplace Relationships',
            'Inclusive Leadership',
            'Job Autonomy',
            'Job Crafting',
            'Job Security',
            'Job Significance',
            'Leadership Feedback Style',
            'Opportunities for Advancement',
            'Reward or Compensation Satisfaction',
            'Scheduling Control',
            'Social Support at Work',
            'Time for Leisure',
            'Value Alignment',
            'Work Knowledge Acquisition',
            'Work-Life Balance',
            'Workload'
        ]

        # Create demographic_totals Series
        demographic_totals = department_counts.set_index(department_column)['total_members']

        # Compute lowest scores and update results DataFrame
        for index, row in average_scores_by_demographic.iterrows():
            numeric_row = row[numeric_metrics_columns].astype(float)
            filtered_numeric_row = numeric_row[numeric_row.index.intersection(org_influencers)]
            lowest_scores = filtered_numeric_row.nsmallest(2)

            results = pd.concat([results, pd.DataFrame({
                demographic_name: [row[demographic_name]],
                'Lowest Influencer': [f"{lowest_scores.index[0]}: {lowest_scores.iloc[0]}"],
                'Second Lowest Influencer': [f"{lowest_scores.index[1]}: {lowest_scores.iloc[1]}"],
            })], ignore_index=True)

        # Add Demographic Size column to results DataFrame
        results['Demographic Size'] = results[demographic_name].map(demographic_totals)

        results.to_excel(writer, sheet_name='Lowest Scores', index=False)

    output.seek(0)
    return output
