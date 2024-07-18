import streamlit as st
import numpy as np
import pandas as pd
from report_data_parse import create_xlsx_report
import doc_creator3
import tempfile
from workplace_culture import create_culture_report_with_header, create_glossary_pdf, subgroup_table, merge_pdfs
from Data_Reliability import calculate_response_reliability_index, calculate_statistical_deviation_score

if 'demo' not in st.session_state:
    st.session_state['demo'] = None

if 'subgroup' not in st.session_state:
    st.session_state['subgroup'] = None


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
st.title("Elation Reporting App")
org = st.text_input("Please Input the Organization Name")

st.header("Upload files")
demographics_file = st.file_uploader("Upload Demographics CSV", type="csv")
raw_data_file = st.file_uploader("Upload Raw Data CSV", type="csv")
scoring_instructions_df = pd.read_csv('CSVs/ER App - Scoring - 1. Assessment.csv')
final_layout_df = pd.read_csv('CSVs/ER App - Ideal Data Output Format.csv')

if demographics_file and raw_data_file:
    demographics_df = pd.read_csv(demographics_file)
    total_demo_df=demographics_df
    raw_data_df = pd.read_csv(raw_data_file)

    # Ensure the columns for merging are aligned
    demographics_df.rename(columns={'id': 'respondentId'}, inplace=True)

    # Merge data on respondentId
    merged_df = pd.merge(demographics_df, raw_data_df, on="respondentId", how="inner")

    # Check and fill Demographic 1 with groupName if groupName exists
    if 'groupName' in merged_df.columns:
        merged_df['Demographic 1'] = merged_df['groupName']
        print("'Demographic 1' assigned from 'groupName'")

    # Check if there are any columns ending with '_x'
    x_columns = [col for col in merged_df.columns if col.endswith('_x') and not col.startswith('score_')]

    if x_columns:
        demographic_columns = ['Demographic 2', 'Demographic 3', 'Demographic 4']
        for demo_col, x_col in zip(demographic_columns, x_columns):
            merged_df[demo_col] = merged_df[x_col]
            print(f"'{demo_col}' assigned from '{x_col}'")
            merged_df.drop(columns=[x_col], inplace=True)

    ### INSERT DEMOGRAPHIC 4 HERE IF NEEDED

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

    xlsx_df = merged_df
    merged_df = back_convert_agree_disagree(merged_df, agree_disagree_columns, categorical_mapping)

    # Replace the userId column with respondentId (renamed to userId)
    merged_df['userId'] = merged_df['respondentId']

    # Create the list of final layout columns, including all columns from final_layout_df
    final_layout_columns = list(final_layout_df.columns)

    # Ensure Duration is in the final layout columns and handle capitalization
    if 'Duration' not in final_layout_columns:
        final_layout_columns.append('Duration')

    # Insert the new userId at the beginning of the final layout columns
    if 'userId' not in final_layout_columns:
        final_layout_columns.insert(0, 'userId')

    # Ensure final_layout_columns only includes columns that exist in merged_df
    final_layout_columns = [col for col in final_layout_columns if col in merged_df.columns]

    # Proceed with extracting the final layout
    final_df = merged_df[final_layout_columns]

    raw_data_df.rename(columns={'respondentId': 'userId'}, inplace=True)

    # Calculate the Response Reliability Index using raw_data_df
    raw_data_df = calculate_response_reliability_index(raw_data_df)

    raw_data_df = calculate_statistical_deviation_score(raw_data_df)

    # Append the new columns to final_df
    final_df = final_df.merge(raw_data_df[['userId', 'Response Reliability Index', 'Social Desirability Score', 'Absolute Z-score', 'Above 95% threshold']],
                              on='userId', how='left')

    final_df['Valid Response'] = np.where(
        ((final_df['Response Reliability Index'] <= 6).astype(int) +
         (final_df['Social Desirability Score'] <= 50).astype(int) +
         (final_df['Above 95% threshold'] == 'Yes').astype(int)) >= 2,
        'No', 'Yes'
    )

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

    # Define the demographic terms to search for
    demographic_terms = ['Demographic 1', 'Demographic 2', 'Demographic 3', 'Demographic 4']

    # Filter the dataframe to include only the demographic columns
    demographics_df = final_df[[col for col in final_df.columns if col in demographic_terms]]

    # Display buttons horizontally and process the selected option
    if not demographics_df.empty:
        # Select the first 3 rows
        display_df = demographics_df.head(4).copy()

        # Create a new row with '...' for each column
        ellipsis_row = pd.DataFrame({col: ['...'] for col in display_df.columns})

        # Concatenate the ellipsis row
        display_df = pd.concat([display_df, ellipsis_row], ignore_index=True)

        # Add 1 to each index
        display_df.index = display_df.index + 1

        st.table(display_df)

        demographic_options = demographics_df.columns.tolist()

        # Display buttons horizontally
        st.write("Please choose the Demographic data you want to export")
        cols = st.columns(len(demographic_options))
        clicked_button = None
        for i, col in enumerate(cols):
            if col.button(demographic_options[i]):
                st.session_state['demo'] = demographic_options[i]
                clicked_button = demographic_options[i]

        if st.session_state['demo']:
            demo = st.session_state['demo']
            st.markdown("""---""")

            # Call the function to create the Excel report
            output_df = create_xlsx_report(xlsx_df, total_demo_df, final_df, demographic_name=demo)

            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(output_df.getvalue())
                tmp_path = tmp.name

                # Generate the PDF report
                pdf_path = f"{org}_report.pdf"
                logo_path = "elation_logo.png"  # Update this path to where your logo file is located
                doc_creator3.create_pdf_with_header_and_recommendations(tmp_path, pdf_path, org, logo_path, raw_data_df, demo)

                col1, col2 = st.columns(2)

                with col1:
                    st.download_button(
                        label=f"Download {demo} Insights Data",
                        data=output_df,
                        file_name=f"{org}_insights_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                with open(pdf_path, "rb") as file:
                    with col2:
                        st.download_button(
                            label=f"Download {demo} PDF Report",
                            data=file,
                            file_name=f"{org}_report.pdf",
                            mime="application/pdf"
                        )

                st.markdown("""---""")

                unique_items = xlsx_df[demo].unique()

                st.write("Select a subgroup to generate a report on:")
                subgroup = st.selectbox('Subgroup:', ['Select...'] + list(unique_items))

                if subgroup != 'Select...':
                    st.session_state['subgroup'] = subgroup

                if st.session_state['subgroup']:
                    subgroup = st.session_state['subgroup']
                    # Generate glossary PDF
                    glossary_pdf = create_glossary_pdf(org,logo_path)

                    # Generate culture report PDF
                    buf = create_culture_report_with_header(demo, subgroup, final_df,org,logo_path)
                    output_df = pd.read_excel(tmp_path, sheet_name=None)
                    buf2 = subgroup_table(raw_data_df, org, subgroup, output_df, demo,org,logo_path)

                    # Merge PDFs
                    combined_pdf = merge_pdfs(buf.getvalue(), glossary_pdf, buf2, org, logo_path)

                    # Provide a download button for the combined PDF
                    st.download_button(
                        label="Download Subgroup Report",
                        data=combined_pdf,
                        file_name=f'{subgroup}_Subgroup_Report.pdf',
                        mime="application/pdf"
                    )







