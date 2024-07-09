from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from datetime import datetime
import pandas as pd


def subgroup_table(raw_df, org_name, demo, output_df):
    raw_df['reportedAt'] = pd.to_datetime(raw_df['reportedAt'])
    latest_time = raw_df['reportedAt'].max()
    month = latest_time.strftime('%b %Y')

    elements = []

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

    # Extract the column name containing the wellbeing potential scores
    wellbeing_potential_column = \
    scores_df.columns[scores_df.columns.str.contains("Wellbeing/Performance Potential", case=False)].tolist()[0]

    # Find the row with the minimum wellbeing potential score
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

    # Find the relevant rows and extract their values
    metrics = ["Wellbeing/Performance Potential", "Job Satisfaction", "Job Engagement", "Intent to Stay"]
    norms = [62, 69, 65, 67]

    score_data = {}
    for metric in metrics:
        score_data[metric] = int(org_total_row[metric].values[0])

    # Create the structured data for PDF
    outcomes = [
        ["Influencers", "Score", "vs Elation Norm"]
    ]

    for metric, norm in zip(metrics, norms):
        value = score_data[metric]
        difference = value - norm
        outcomes.append([metric, str(value), str(difference)])

    # print(outcomes)

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
    body_data_wrapped = [[Paragraph(cell, styles['Normal']) for cell in row] for row in body_data]
    body_data_wrapped2 = [[Paragraph(cell, styles['Normal']) for cell in row] for row in body_data2]

    # Define the body table styles
    body_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
        ('TOPPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
    ])

    nogrid_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
        ('TOPPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
    ])

    nogrid_style_bold = TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
        ('TOPPADDING', (0, 0), (-1, -1), 2),  # Adjust padding to reduce row height
    ])

    # Define a style for the header
    header_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Heading2'],
        alignment=0,
        fontSize=10,
        textColor=colors.black
    )

    # Define a style for the header
    title_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Heading1'],
        alignment=1,
        fontSize=14,
        textColor=colors.white
    )

    # Define a style for the header
    subtitle_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Heading1'],
        alignment=1,
        fontSize=14,
        textColor=colors.white
    )

    title_table_data = [[Paragraph(f"Summary", title_style)]]
    title_table = Table(title_table_data, colWidths=[550])
    title_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.gray),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]))

    elements.append(title_table)

    # Create and style the body table
    body_table = Table(org_date, colWidths=[130,420])
    body_table.setStyle(body_style)
    elements.append(body_table)

    subtitle_table_data = [[Paragraph("Participation", subtitle_style)]]
    subtitle_table = Table(subtitle_table_data, colWidths=[550])
    subtitle_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.darkgray),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]))

    elements.append(subtitle_table)

    # Create and add the header as a table to control spacing
    header_table_data = [[Paragraph(header_data[0], header_style), Paragraph(header_data[1], header_style),
                          Paragraph(header_data[2], header_style), Paragraph(header_data[3], header_style),
                          Paragraph(header_data[4], header_style)]]
    header_table = Table(header_table_data, colWidths=[130,105,105,105,105])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(header_table)

    # Create and style the body table
    body_table = Table(body_data_wrapped, colWidths=[130,105,105,105,105])
    body_table.setStyle(body_style)
    elements.append(body_table)

    subtitle_table_data = [[Paragraph("Outcomes", subtitle_style)]]
    subtitle_table = Table(subtitle_table_data, colWidths=[550])
    subtitle_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.darkgray),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]))

    elements.append(subtitle_table)

    # Create and add the header as a table to control spacing
    header_table_data = [[Paragraph(header_data2[0], header_style), Paragraph(header_data2[1], header_style), Paragraph(header_data2[2], header_style),]]
    header_table = Table(header_table_data, colWidths=[130,105,315])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]))

    elements.append(header_table)

    # Create and style the body table
    body_table = Table(body_data_wrapped2, colWidths=[130,105,315])
    body_table.setStyle(body_style)
    elements.append(body_table)

    if bad_scores == 0:
        assessment_highlight = [["For the OVERALL group, there were 0 influencers scoring lower than " \
                               "40. This reflects that no one influencer within the organization comprehensively " \
                               "undermines the Wellbeing and Performance Potential of the cohort."]]
    else:
        assessment_highlight = [[f"For the OVERALL group, there were {str(bad_scores)} influencers scoring lower than " \
                               "40. This reflects that these influencers systemically limit the Wellbeing and Performance " \
                                   "Potential of the organization."]]

    subtitle_table_data = [[Paragraph("Assessment Highlights", subtitle_style)]]
    subtitle_table = Table(subtitle_table_data, colWidths=[550])
    subtitle_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.darkgray),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]))

    elements.append(Spacer(1, 0.2 * inch))
    elements.append(subtitle_table)

    assessment_highlight = [[Paragraph(cell, styles['Normal']) for cell in row] for row in assessment_highlight]
    body_table = Table(assessment_highlight, colWidths=[550])
    body_table.setStyle(nogrid_style)
    elements.append(body_table)

    assessment_highlight2 = [[f"{low_demo} with a score of {str(lowest_score)} had the lowest overall wellbeing and performance potential score of the subgroups."]]
    assessment_highlight2 = [[Paragraph(cell, styles['Normal']) for cell in row] for row in assessment_highlight2]
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
    style = styles['Normal']

    # Create the table data
    assessment_highlight = [[Paragraph(desc, style)] for desc in descriptions]

    # Create the table with no grid style
    nogrid_style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
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