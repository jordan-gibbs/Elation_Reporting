import streamlit as st
import pandas as pd
import report_data_parse
import doc_creator3

# Define function to extract scoring instructions
def extract_scoring_instructions(scoring_df):
    scoring_instructions = {}
    for _, row in scoring_df.iterrows():
        variable_name = row['Variable Name']
        question = row['Question']
        try:
            weight = float(row['Scoring'])
        except ValueError:
            # st.warning(f"Skipping row due to invalid weight: {row['Scoring']}")
            continue
        if variable_name not in scoring_instructions:
            scoring_instructions[variable_name] = []
        scoring_instructions[variable_name].append((question, weight))
    return scoring_instructions

# Define function to calculate scores
def calculate_scores(merged_df, scoring_instructions):
    for variable, instructions in scoring_instructions.items():
        total_score = sum([merged_df[question] * weight for question, weight in instructions])
        merged_df[variable] = total_score
    return merged_df

# Define function to back convert agree/disagree columns
def back_convert_agree_disagree(merged_df, columns, mapping):
    for column in columns:
        if column in merged_df.columns:
            merged_df[column] = merged_df[column].map(mapping)
    return merged_df

# Define function to extract final layout
def extract_final_layout(merged_df, final_layout_columns):
    final_df = merged_df[final_layout_columns]
    return final_df

# Streamlit app
st.title("Elation Data Processor")
org = st.text_input("Please Input the Organization Name")

st.header("Upload files")
demographics_file = st.file_uploader("Upload Demographics CSV", type="csv")
raw_data_file = st.file_uploader("Upload Raw Data CSV", type="csv")
scoring_instructions_df = pd.read_csv('CSVs/ER App - Scoring - 1. Assessment.csv')
final_layout_df = pd.read_csv('CSVs/ER App - Ideal Data Output Format.csv')

if demographics_file and raw_data_file:
    demographics_df = pd.read_csv(demographics_file)
    raw_data_df = pd.read_csv(raw_data_file)

    # Ensure the columns for merging are aligned
    demographics_df.rename(columns={'id': 'respondentId'}, inplace=True)

    # Merge data on respondentId
    merged_df = pd.merge(demographics_df, raw_data_df, on="respondentId", how="inner")

    # Fill Demographic 1 with groupName
    merged_df['Demographic 1'] = merged_df['groupName']

    # Rename durationSeconds to Duration
    merged_df.rename(columns={'durationSeconds': 'Duration'}, inplace=True)

    # Extract scoring instructions
    scoring_instructions = extract_scoring_instructions(scoring_instructions_df)

    # Calculate scores
    merged_df = calculate_scores(merged_df, scoring_instructions)

    # Transform categorical data
    categorical_mapping = {
        100: 'Strongly agree',
        75: 'Agree',
        50: 'Neutral',
        25: 'Disagree',
        0: 'Strongly disagree'
    }

    agree_disagree_columns = [
        'Our leaders treat staff with respect.',
        'Staff treat each other with respect.',
        'My co-workers trust in me and each other.',
        'In my area, employees are treated fairly.',
        'I am empowered to investigate problems and explore new ideas at work.'
    ]

    merged_df = back_convert_agree_disagree(merged_df, agree_disagree_columns, categorical_mapping)

    # Replace the userId column with respondentId (renamed to userId)
    merged_df['userId'] = merged_df['respondentId']

    # Ensure Duration is in the final layout columns and handle capitalization
    final_layout_columns = [col for col in final_layout_df.columns if
                            col not in ['Demographic 2', 'Demographic 3', 'Demographic 4']]
    if 'Duration' not in final_layout_columns:
        final_layout_columns.append('Duration')

    # Insert the new userId at the beginning of the final layout columns
    if 'userId' not in final_layout_columns:
        final_layout_columns.insert(0, 'userId')

    final_df = extract_final_layout(merged_df, final_layout_columns)

    # Round all values to the nearest whole number
    final_df = final_df.round()

    # Download link for the final dataframe
    csv = final_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Processed Data (Raw)",
        data=csv,
        file_name=f'{org}_processed_data.csv',
        mime='text/csv',
    )

    final_xlsx_path = f"{org}_Insights.xlsx"
    report_data_parse.create_xlsx_report(demographics_df, final_df, final_xlsx_path)

    # Download link for the xlsx report
    with open(final_xlsx_path, "rb") as file:
        btn = st.download_button(
            label="Download Insights Report",
            data=file,
            file_name=f"{org}_insights_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # st.success("Dataset created successfully")

    # Generate the PDF report
    pdf_path = f"{org}_report.pdf"
    logo_path = "elation_logo.png"  # Update this path to where your logo file is located
    doc_creator3.create_pdf_with_header_and_recommendations(final_xlsx_path, pdf_path, org, logo_path)

    # Download link for the PDF report
    with open(pdf_path, "rb") as file:
        btn = st.download_button(
            label="Download PDF Report",
            data=file,
            file_name=f"{org}_report.pdf",
            mime="application/pdf"
        )

