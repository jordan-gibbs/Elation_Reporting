from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from datetime import datetime
import pandas as pd
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from Data_Reliability import create_reliability_report
from reportlab.platypus import Image
import os


def subgroup_table(raw_df, org_name, demo, output_df, final_df, excel_file):
    font_path = "fonts/Lexend-Bold.ttf"
    pdfmetrics.registerFont(TTFont('Lexend-Bold', font_path))
    font_path = "fonts/Lexend-Regular.ttf"
    pdfmetrics.registerFont(TTFont('Lexend', font_path))
    font_path = "fonts/Lexend-Thin.ttf"
    pdfmetrics.registerFont(TTFont('Lexend-Light', font_path))
    raw_df['reportedAt'] = pd.to_datetime(raw_df['reportedAt'])
    latest_time = raw_df['reportedAt'].max()
    month = latest_time.strftime('%b %Y')

    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle(
        name='Normal',
        parent=styles['Normal'],
        fontName="Lexend",
        fontSize=9
    )

    first_sheet_name = list(output_df.keys())[0]
    scores_df = output_df[first_sheet_name]
    demographic_column = scores_df.columns[scores_df.columns.str.contains("Demographic", case=False)].tolist()[0]
    org_total_row = scores_df[scores_df[demographic_column].str.contains(demo, case=False, na=False)]
    total_row = scores_df[scores_df[demographic_column].str.contains("Org Total", case=False, na=False)]
    # List of metrics to filter columns
    metrics = [
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

    # Filter the columns based on the metrics list
    influencer_columns = [col for col in total_row.columns if any(metric in col for metric in metrics)]

    # Select only the numeric columns for the filtered metrics
    numeric_total_row = total_row[influencer_columns].select_dtypes(include='number')

    # Calculate bad scores (scores less than 40)
    bad_scores = (numeric_total_row < 40).sum().sum()

    # demographics_df = pd.read_csv(demographics_file)

    # completed_counts = demographics_df[demographics_df['response'] == 'Y'].groupby(demographic_column)[
    #     'userId'].count().reset_index(
    #     name='completed_members')
    #
    # print(completed_counts)

    # # Merge the results to get a complete view
    # final_counts = pd.merge(department_counts, completed_counts, on=department_column, how='left')
    # print(final_counts)

    # # Function to check if the completed members are above 5 for a given row
    # def has_sufficient_members(row):
    #     return row['completed_members'] > 5  # Replace 'Completed Members' with the actual column name
    #
    # # Filter rows with sufficient members
    # sufficient_members_df = scores_df[scores_df.apply(has_sufficient_members, axis=1)]

    # print(scores_df)

    sheet_name = 'Completion Rate'
    lowest_scores_sheet = 'Lowest Scores'

    # Read the data from the sheets
    data = pd.read_excel(excel_file, sheet_name=sheet_name)
    lowest_scores = pd.read_excel(excel_file, sheet_name=0)

    print(data)
    print(lowest_scores)

    # Define the common column for the merge as the first column of each DataFrame
    common_column_data = data.columns[0]
    common_column_lowest_scores = lowest_scores.columns[0]

    # Merge the two DataFrames based on the first column
    merged_df = pd.merge(data, lowest_scores, left_on=common_column_data, right_on=common_column_lowest_scores,
                         how='inner')

    # Function to check if the completed members are above 5 for a given row
    def has_sufficient_members(row):
        return row['completed_members'] > 5  # Replace 'Completed Members' with the actual column name

    # Filter rows with sufficient members
    sufficient_members_df = merged_df[merged_df.apply(has_sufficient_members, axis=1)]

    print(sufficient_members_df)

    # Extract the column name containing the wellbeing potential scores
    wellbeing_potential_column = \
    sufficient_members_df.columns[sufficient_members_df.columns.str.contains("Wellbeing/Performance Potential", case=False)].tolist()[0]

    # Check if there are rows with sufficient members
    if not sufficient_members_df.empty:
        # Find the row with the minimum wellbeing potential score among rows with sufficient members
        min_score_row = sufficient_members_df.loc[sufficient_members_df[wellbeing_potential_column].idxmin()]
    else:
        # Handle the case where no rows have sufficient members (e.g., fallback to a default or raise an error)
        min_score_row = scores_df.loc[scores_df[wellbeing_potential_column].idxmin()]

    # Extract the minimum score and its respective demographic label
    lowest_score = min_score_row[wellbeing_potential_column]
    low_demo = min_score_row[demographic_column]

    metrics = [
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

    influencer_columns = [col for col in scores_df.columns if any(metric in col for metric in metrics)]

    # Find the bottom three scores across all influencer columns
    melted_total_row = total_row.melt(value_vars=influencer_columns, var_name='Metric', value_name='Score')
    bottom_three_rows = melted_total_row.nsmallest(3, 'Score')

    # Extract the scores and their respective metrics
    bottom_scores = bottom_three_rows['Score'].tolist()
    bottom_metrics = bottom_three_rows['Metric'].tolist()

    elements = []

    # Elation norm values
    elation_norm = [62, 69, 65, 67]

    # Find the relevant rows and extract their values
    metrics = ["Wellbeing/Performance Potential", "Job Satisfaction", "Job Engagement", "Intent to Stay"]

    group_total = scores_df[scores_df[demographic_column].str.contains("Org Total", case=False, na=False)]

    score_data = {}
    for metric in metrics:
        score_data[metric] = float(org_total_row[metric].values[0])

    # Calculate deltas for elation norm and store in a new dictionary
    delta_elation_data = {}
    for metric, elation_value in zip(metrics, elation_norm):
        subgroup_value = float(group_total[metric].values[0])
        delta = subgroup_value - elation_value
        delta_elation_data[metric] = delta

    # Sort the score_data by values in descending order
    sorted_score_data = sorted(score_data.items(), key=lambda item: item[1], reverse=True)

    # Create the structured data for PDF
    outcomes = [
        ["Influencers", f"{month}", "vs Elation Norm"],
    ]

    for metric, value in sorted_score_data:
        delta_elation = delta_elation_data[metric]
        delta_elation_str = f"{'+' if delta_elation > 0 else ''}{int(delta_elation)}"
        outcomes.append([metric, str(int(value)), delta_elation_str])


    sheet_name = 'Completion Rate'
    completion_rate_df = output_df[sheet_name]

    def get_demographic_data(df, demo):
        # Filter the data for the specified demographic
        filtered_data = df[df.iloc[:, 0] == demo]

        if not filtered_data.empty:
            # Extract values for total_members, completed_members, and completion_data
            total_members = filtered_data['total_members'].values[0]
            completed_members = filtered_data['completed_members'].values[0]
            completion_data = filtered_data['completion_rate'].values[0]

            return total_members, completed_members, completion_data
        else:
            return None, None, None

    total_members, completed_members, completion_data = get_demographic_data(completion_rate_df, demo)
    not_completed = total_members-completed_members
    datetime_obj = datetime.today()
    formatted_date = datetime_obj.strftime("%m/%d/%Y")
    date = formatted_date.replace("/0", "/").replace(" 0", " ")

    # Sample style
    styles = getSampleStyleSheet()

    # Data
    org_date = [
        ["Organization", f"{org_name}"],
        ["Assessment Date", f"{month}"]
    ]

    participation = [
        [f"Assessment", "Completed", "Not completed", "Total", "Participation Rate"],
        [f"ElationAWPv4", f"{completed_members}", f"{not_completed}", f"{total_members}", f"{completion_data}"],
    ]

    demographic_column = scores_df.columns[scores_df.columns.str.contains("Demographic", case=False)].tolist()[0]
    org_total_row = scores_df[scores_df[demographic_column].str.contains(demo, case=False, na=False)]

    # Find the relevant rows and extract their values
    metrics = [
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

    score_data = {}
    for metric in metrics:
        score_data[metric] = str(org_total_row[metric].values[0])

    # Create the structured data for PDF
    workplace_indicators = [
        ["Influencers", f"{month}"],
    ]

    for metric, value in score_data.items():
        workplace_indicators.append([metric, value])

    # print(workplace_indicators)

    metrics = [
             'Charity',
             'Cognitive Empathy',
             'Dietary Quality',
             'Emotional Empathy',
             'Gratitude',
             'Healthy Personal Relationships',
             'Knowledge Desire and Curiosity',
             'Mindfulness Practice',
             'Regular Meal Energy',
             'Sleep Pattern Consistency',
             'Sleep Quality',
             'Workout Balance',
             'Workout Frequency and length']

    score_data = {}
    for metric in metrics:
        score_data[metric] = str(org_total_row[metric].values[0])

    personal_indicators = [
        ["Influencers", f"{month}"],
    ]

    for metric, value in score_data.items():
        personal_indicators.append([metric, value])

    # print(personal_indicators)

    # Separate header from body
    header_data = participation[0]
    header_data2 = outcomes[0]

    body_data = participation[1:]
    body_data2 = outcomes[1:]

    # Wrap the body data in Paragraphs
    body_data_wrapped = [[Paragraph(cell, normal_style) for cell in row] for row in body_data]
    body_data_wrapped2 = [[Paragraph(cell, normal_style) for cell in row] for row in body_data2]

    # Define the body table styles
    body_style = TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
        ('TOPPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
    ])

    nogrid_style = TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
        ('TOPPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
        # ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ])

    nogrid_style_bold = TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
        ('TOPPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
        # ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ])

    # Define a style for the header
    header_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Heading2'],
        alignment=0,
        fontSize=9,
        fontName="Lexend-Bold",
        textColor=colors.black
    )

    # Define a style for the header
    title_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Heading1'],
        alignment=1,
        fontSize=20,
        valign="TOP",
        fontName="Lexend-Bold",
        textColor=colors.black
    )

    # Define a style for the header
    subtitle_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Heading1'],
        alignment=1,
        fontSize=14,
        fontName="Lexend-Bold",
        textColor=colors.black
    )

    title_table_data = [[Paragraph(f"Summary", title_style)]]
    title_table = Table(title_table_data, colWidths=[550])
    title_table.setStyle(TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.gray),
        # ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        # ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    elements.append(title_table)
    elements.append(Spacer(1, 0.2 * inch))

    # Create and style the body table
    body_table = Table(org_date, colWidths=[130,420])
    body_table.setStyle(body_style)
    elements.append(body_table)
    elements.append(Spacer(1, 0.2 * inch))

    subtitle_table_data = [[Paragraph("Participation", subtitle_style)]]
    subtitle_table = Table(subtitle_table_data, colWidths=[550])
    subtitle_table.setStyle(TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.black),
        # ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        # ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    elements.append(subtitle_table)

    # Create and add the header as a table to control spacing
    header_table_data = [[Paragraph(header_data[0], header_style), Paragraph(header_data[1], header_style),
                          Paragraph(header_data[2], header_style), Paragraph(header_data[3], header_style),
                          Paragraph(header_data[4], header_style)]]
    header_table = Table(header_table_data, colWidths=[130,105,105,105,105])
    header_table.setStyle(TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    elements.append(header_table)

    # Create and style the body table
    body_table = Table(body_data_wrapped, colWidths=[130,105,105,105,105])
    body_table.setStyle(body_style)
    elements.append(body_table)

    elements.append(Spacer(1, 0.2 * inch))

    subtitle_table_data = [[Paragraph("Outcomes", subtitle_style)]]
    subtitle_table = Table(subtitle_table_data, colWidths=[550])
    subtitle_table.setStyle(TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.black),
        # ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        # ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    elements.append(subtitle_table)


    # Create and add the header as a table to control spacing
    header_table_data = [[Paragraph(header_data2[0], header_style), Paragraph(header_data2[1], header_style), Paragraph(header_data2[2], header_style),]]
    header_table = Table(header_table_data, colWidths=[235,105,210])
    header_table.setStyle(TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]))

    elements.append(header_table)

    table = Table(body_data_wrapped2, colWidths=[235,105,210])
    style = TableStyle([
        # ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ])
    # Create and style the body table
    # Add colors for deltas
    for i in range(len(body_data2)):
        if int(body_data2[i][2]) >= 0:
            style.add('BACKGROUND', (2, i), (2, i), colors.HexColor("#A6CF5C"))
        else:
            style.add('BACKGROUND', (2, i), (2, i), colors.HexColor("#DE9C95"))

    # Apply style to table
    table.setStyle(style)
    elements.append(table)
    elements.append(Spacer(1, 0.2 * inch))

    subtitle_table_data = [[Paragraph(f"Data Validity", subtitle_style)]]
    subtitle_table = Table(subtitle_table_data, colWidths=[550])
    subtitle_table.setStyle(TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.black),
        # ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        # ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    elements.append(subtitle_table)

    reliability_image = create_reliability_report(final_df, total_members)

    if reliability_image:
        img = Image(reliability_image)
        img.drawHeight = 1.4 * inch  # Adjust height as necessary
        img.drawWidth = 7.8 * inch  # Adjust width as necessary
        elements.append(img)
        elements.append(Spacer(1, 0.2 * inch))


    if bad_scores == 0:
        assessment_highlight = [["For the OVERALL group, there were 0 influencers scoring lower than " \
                               "40. This reflects that no one influencer within the organization comprehensively " \
                               "undermines the Wellbeing and Performance Potential of the cohort."]]
    elif bad_scores == 1:
        assessment_highlight = [["For the OVERALL group, there was 1 influencer scoring lower than " \
                               "40. This reflects that this influencer systemically limits the Wellbeing and Performance Potential of the cohort."]]
    else:
        assessment_highlight = [[f"For the OVERALL group, there were {str(bad_scores)} influencers scoring lower than " \
                               "40. This reflects that these influencers systemically limit the Wellbeing and Performance " \
                                   "Potential of the organization."]]

    subtitle_table_data = [[Paragraph("Assessment Highlights", subtitle_style)]]
    subtitle_table = Table(subtitle_table_data, colWidths=[550])
    subtitle_table.setStyle(TableStyle([
        # ('BACKGROUND', (0, 0), (-1, -1), colors.black),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        # ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    elements.append(subtitle_table)

    assessment_highlight = [[Paragraph(cell, normal_style) for cell in row] for row in assessment_highlight]
    body_table = Table(assessment_highlight, colWidths=[550])
    body_table.setStyle(nogrid_style)
    elements.append(body_table)

    assessment_highlight2 = [[f"{low_demo} with a score of {str(lowest_score)} had the lowest overall wellbeing and performance potential score of the subgroups."]]
    assessment_highlight2 = [[Paragraph(cell, normal_style) for cell in row] for row in assessment_highlight2]
    body_table = Table(assessment_highlight2, colWidths=[550])
    body_table.setStyle(nogrid_style)
    elements.append(Spacer(1, 0.2 * inch))
    elements.append(body_table)

    elements.append(Spacer(1, 0.2 * inch))

    descriptions = []
    for i, (metric, score) in enumerate(zip(bottom_metrics, bottom_scores)):
        order = ["lowest", "2nd lowest", "3rd lowest"]
        descriptions.append(f"• \"{metric}\" was the {order[i]} scoring Influencer for the organization ({score}).")

    # Get a sample style for the PDF
    styles = getSampleStyleSheet()
    style = normal_style

    # Create the table data
    assessment_highlight = [[Paragraph(desc, style)] for desc in descriptions]

    # Create the table with no grid style
    nogrid_style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Light'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ])

    # Create the table
    body_table = Table(assessment_highlight, colWidths=[550])
    body_table.setStyle(nogrid_style)

    elements.append(body_table)

    # # Build the PDF
    # pdf.build(elements)

    return elements

#
# file_path = 'CSVs/Test_insights_report.xlsx'
# output_df = pd.read_excel(file_path, sheet_name=None)
# pdf_path = "test.pdf"
# demo_df = pd.read_csv('CSVs/TESTwb².csv')
# raw_df = pd.read_csv('CSVs/TestRaw.csv')
# org_name = 'Testy'
# demo = 'Org Total'
#
# subgroup_table(pdf_path, raw_df, org_name, demo, output_df)
