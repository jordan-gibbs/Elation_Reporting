# Required imports
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import textwrap
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import tempfile
import os
from PIL import Image as PILImage
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from datetime import datetime
import pandas as pd
from reportlab.platypus import Table, TableStyle, Paragraph, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image
from PyPDF2 import PdfReader, PdfWriter
from matplotlib import font_manager
from reportlab.platypus import Flowable


def culture_header(canvas, doc, company_name, logo_path, scale_factor=0.14):
    font_path = "fonts/Lexend-Bold.ttf"
    pdfmetrics.registerFont(TTFont('Lexend-Bold', font_path))

    canvas.saveState()
    image_path = 'blue.png'
    background_image = ImageReader(image_path)
    canvas.drawImage(background_image, 0, 0, width=letter[0], height=letter[1])

    canvas.setFont('Lexend-Bold', 12)
    canvas.setFillColor(colors.grey)
    header_text = f"Wb^2 Report – {company_name}, ALL"
    canvas.drawString(.42*inch, letter[1] - .6 * inch, header_text)

    # Load the logo using the Image class
    logo = Image(logo_path)

    # Calculate scaled dimensions
    scaled_width = logo.imageWidth * scale_factor
    scaled_height = logo.imageHeight * scale_factor

    # Draw the logo at the top right corner with scaling
    logo.drawWidth = scaled_width
    logo.drawHeight = scaled_height
    logo.drawOn(canvas, letter[0] - scaled_width - .35 * inch, letter[1] - scaled_height - 0.4 * inch)

    canvas.restoreState()


# Function to create glossary PDF
def create_glossary_pdf(company_name, logo_path):
    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle(
        name='Normal',
        parent=styles['Normal'],
        fontName="Lexend",
        fontSize=9
    )

    font_path = "fonts/Lexend-Bold.ttf"
    pdfmetrics.registerFont(TTFont('Lexend-Bold', font_path))
    font_path = "fonts/Lexend-Regular.ttf"
    pdfmetrics.registerFont(TTFont('Lexend', font_path))
    font_path = "fonts/Lexend-Thin.ttf"
    pdfmetrics.registerFont(TTFont('Lexend-Light', font_path))

    # Define the PDF path
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_file:
        temp_file_path = temp_file.name

    # Define the PDF path
    pdf_path = temp_file_path
    pdf = SimpleDocTemplate(pdf_path, pagesize=letter)
    elements = []

    # Sample style
    styles = getSampleStyleSheet()

    # Glossary content
    glossary_data = [
        ["Outcomes", "Definitions"],
        ["Wellbeing/Performance Potential", "This concept is grounded in the understanding that the holistic wellbeing of an employee - encompassing physical mental emotional financial social spiritual and intellectual dimensions - is the foundation of their performance potential at work. It asserts that wellbeing characterized by a balanced life low stress high motivation resilience and a profound sense of purpose is not merely about the absence of illness but is essential for an individual to flourish and achieve professional excellence. In this framework wellbeing is the precursor and key driver of an employee's capacity to perform and realize their full potential in the workplace."],
        ["Job Satisfaction", "Job Satisfaction refers to the level of contentment or fulfillment an employee experiences at their job."],
        ["Intent to Stay", "Intent to Stay refers to an employee's intention or desire to remain with their current job or organization for a specific period."],
        ["Job Engagement", "Job Engagement refers to the employee's level of enthusiasm commitment and involvement in their work and organization."]
    ]

    glossary_data2 = [
        ["Workplace Influencers", "Definitions"],
        ["Belonging in Organization", "Belonging within an organization encompasses feelings of security support acceptance inclusion and identity of employees."],
        ["Healthy Personal Relationships", "The quality of interpersonal relationships between employees working directly or indirectly with each other"],
        ["Job Autonomy", "Employees’ perception of having the freedom to work in the way that suits them best."],
        ["Job Crafting", "An active process where employees alter their job design to better suit their personal interests skills and passions."],
        ["Job Security", "Employees’ perception and feeling that their job is secure for the indefinite future."],
        ["Job Significance", "Employees’ perception that the tasks and roles they are performing are adding value to the organization."],
        ["Opportunities for Advancement", "The perception of accessible and fair advancement/promotion opportunities within the organization."],
        ["Participative Leadership", "The perception that one’s leader considers and incorporates the opinions and expertise of the employees they lead when making decisions."],
        ["Psychological Empowerment", "Employees’ satisfaction with a given leader’s feedback (positive/negative) style."],
        ["Reward or Compensation Satisfaction", "Employees' satisfaction of pay/rewards considering their job type tenure skillset and education."],
        ["Scheduling Control", "The level of satisfaction employees feel to adapt their work schedule (time flexibility)."],
        ["Social Support at Work", "The perceived existence and quality of social workplace relationships."],
        ["Time for Leisure", "The extent to which employees feel they have enough time off of work to relax."],
        ["Value Alignment", "The extent to which employees have the same values (personally) as their immediate group or organization"],
        ["Work Knowledge Acquisition", "The degree to which your employees engage in voluntary learning activities that improve their job knowledge and skills."],
        ["Work-Life Balance", "Employees’ perceived satisfaction with time spent on personal wants/needs vs. time spent at work."],
        ["Workload", "The number of work responsibilities or productivity expectations of employees’ jobs and their ability/capacity to meet these expectations."],
    ]

    glossary_data3 = [
        ["Personal Influencers", "Defined"],
        ["Charity", "Refers to the group's shared inclination to prioritize the well-being and needs of those less fortunate."],
        ["Cognitive Empathy", "Cognitive Empathy is defined as the collective ability of a team to recognize and comprehend the emotions and viewpoints of others."],
        ["Dietary Quality", "Dietary Quality represents the collective assessment within a team concerning their nutritional balance and intake."],
        ["Emotional Empathy", "Emotional Empathy refers to the collective capacity within a team to resonate with and genuinely feel the emotions of others."],
        ["Gratitude", "Gratitude is the group sentiment or quality of being thankful and appreciative."],
        ["Healthy Personal Relationships", "Healthy Personal Relationships refer to the quality and strength of interpersonal connections that individuals maintain outside the workplace."],
        ["Knowledge Desire and Curiosity", "Knowledge Desire and Curiosity refer to the inherent motivation and eagerness of team members to pursue learning understand new concepts and grow both professionally and personally."],
        ["Mindfulness Practice", "Mindfulness Practice encompasses the intentional effort made by team members to stay present in the moment consciously observing thoughts emotions and sensations without judgment."],
        ["Regular Meal Energy", "Regular Meal Energy is a measure of the consistency and frequency of meal intake among your team members."],
        ["Sleep Pattern Consistency", "Sleep Pattern Consistency refers to the regularity and predictability of an individual's sleep schedule."],
        ["Sleep Quality", "Sleep Quality refers to the depth continuity and restfulness of an individual's sleep rather than its mere duration."],
        ["Workout Balance", "Workout Balance delves into the team's perception of the variety and comprehensiveness in their exercise routines."],
        ["Workout Frequency and length", "Workout Frequency and Length pertain to the regularity with which team members engage in physical exercise and the typical duration of these sessions."]
    ]

    # Separate header from body
    header_data = glossary_data[0]
    body_data = glossary_data[1:]

    header_data2 = glossary_data2[0]
    body_data2 = glossary_data2[1:]

    header_data3 = glossary_data3[0]
    body_data3 = glossary_data3[1:]

    # Wrap the body data in Paragraphs
    body_data_wrapped = [[Paragraph(cell, normal_style) for cell in row] for row in body_data]
    body_data_wrapped2 = [[Paragraph(cell, normal_style) for cell in row] for row in body_data2]
    body_data_wrapped3 = [[Paragraph(cell, normal_style) for cell in row] for row in body_data3]

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

    title_table_data = [[Paragraph("Glossary", title_style)]]
    title_table = Table(title_table_data, colWidths=[550])
    title_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]))

    elements.append(title_table)
    elements.append(Spacer(1, 0.2 * inch))

    # Create and add the header as a table to control spacing
    header_table_data = [[Paragraph(header_data[0], header_style), Paragraph(header_data[1], header_style)]]
    header_table = Table(header_table_data, colWidths=[180, 370])
    header_table.setStyle(TableStyle([
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
    body_table = Table(body_data_wrapped, colWidths=[180, 370])
    body_table.setStyle(TableStyle([
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    elements.append(body_table)

    header_table_data = [[Paragraph(header_data2[0], header_style), Paragraph(header_data2[1], header_style)]]
    header_table = Table(header_table_data, colWidths=[180, 370])
    header_table.setStyle(TableStyle([
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
    body_table = Table(body_data_wrapped2, colWidths=[180, 370])
    body_table.setStyle(body_style)
    elements.append(body_table)

    header_table_data = [[Paragraph(header_data3[0], header_style), Paragraph(header_data3[1], header_style)]]
    header_table = Table(header_table_data, colWidths=[180, 370])
    header_table.setStyle(TableStyle([
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
    body_table = Table(body_data_wrapped3, colWidths=[180, 370])
    body_table.setStyle(body_style)
    elements.append(body_table)

    # Build the PDF with headers
    pdf.build(elements, onFirstPage=lambda c, d: culture_header(c, d, company_name, logo_path),
              onLaterPages=lambda c, d: culture_header(c, d, company_name, logo_path))

    with open(temp_file_path, 'rb') as pdf_file:
        pdf_data = pdf_file.read()

    return pdf_data

# Function to create culture report
def create_culture_report_with_header(subgroup, subgroup_value, data, company_name, logo_path):

    # Create a PDF buffer
    pdf_buf = BytesIO()

    # Define the PDF path
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_file:
        temp_file_path = temp_file.name

    pdf = SimpleDocTemplate(
        pdf_buf,
        pagesize=letter,
        rightMargin=0.35 * inch,
        leftMargin=0.35 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch
    )
    elements = []

    # Your PDF content generation logic
    # Convert the data to a DataFrame if it's not already
    data = pd.DataFrame(data)

    # Filter the data for the specified subgroup
    filtered_data = data[data[subgroup] == subgroup_value]

    # Extract relevant columns
    columns_of_interest = [
        "Our leaders treat staff with respect.",
        "Staff treat each other with respect.",
        "My co-workers trust in me and each other.",
        "In my area, employees are treated fairly.",
        "I am empowered to investigate problems and explore new ideas at work."
    ]

    survey_data = filtered_data[columns_of_interest]

    # Calculate percentages for each response type
    response_types = ["Strongly disagree", "Disagree", "Neutral", "Agree", "Strongly agree"]
    percentages = {col: survey_data[col].value_counts(normalize=True).reindex(response_types, fill_value=0) * 100 for col in columns_of_interest}

    data_values = np.array([
        percentages[columns_of_interest[0]].values,
        percentages[columns_of_interest[1]].values,
        percentages[columns_of_interest[2]].values,
        percentages[columns_of_interest[3]].values,
        percentages[columns_of_interest[4]].values
    ])

    # Colors for each category
    colorz = ['#F24837', '#FC8F3E', '#FCCE48', '#C8DA51', '#5FBE78']
    categories = ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree']

    # plt.rcParams['font.family'] = 'Lexend-Regular'

    fig, ax = plt.subplots(figsize=(14, 8))  # Increase figure width for wider labels

    # Plotting the data
    cumulative_data = np.cumsum(data_values, axis=1)
    for i, (color, category) in enumerate(zip(colorz, categories)):
        ax.barh(columns_of_interest, data_values[:, i], left=cumulative_data[:, i] - data_values[:, i], color=color, label=category, edgecolor='white', height=0.5)  # Increase bar height

    # Adding percentage labels within the blocks
    for j in range(data_values.shape[0]):
        for i in range(data_values.shape[1]):
            if data_values[j, i] >= 5:
                ax.text(cumulative_data[j, i] - data_values[j, i] / 2, j, f"{data_values[j, i]:.0f}%", va='center', ha='center', color='white', fontsize=16, fontweight='bold')

    # Wrap y-tick labels and hide y-ticks
    wrap_width = 30  # Set your desired wrap width here
    ax.set_yticklabels([textwrap.fill(label, wrap_width) for label in columns_of_interest], fontsize=14, fontname='Lexend')
    ax.yaxis.set_ticks_position('none')

    # Customizing the plot to match the desired style
    ax.legend(bbox_to_anchor=(0.5, 0), loc='upper center', ncol=5, frameon=False, prop={'size': 14})
    ax.set_xlim(0, 100)
    ax.xaxis.set_visible(False)
    ax.yaxis.set_visible(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)

    plt.tight_layout()
    plt.subplots_adjust(left=0.3)  # Adjust the left margin to provide more space for y-axis labels

    # Save the plot to a temporary file
    temp_img_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    plt.savefig(temp_img_file.name, format='png', dpi=300, transparent=True)
    temp_img_file.close()

    # Additional PDF content generation logic
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Heading1'],
        alignment=1,
        fontSize=20,
        textColor=colors.black
    )

    # Create the table data
    title_table_data = [[Paragraph("Summary", title_style)]]
    title_table = Table(title_table_data, colWidths=[550])
    title_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]))

    elements.append(title_table)
    # Add the plot image to the PDF elements
    elements.append(Image(temp_img_file.name, width=letter[0] * 0.93, height=letter[1] * 0.4))
    elements.append(Spacer(1, 12))
    elements.append(PageBreak())

    # Build the PDF with headers
    pdf.build(elements, onFirstPage=lambda c, d: culture_header(c, d, company_name, logo_path),
              onLaterPages=lambda c, d: culture_header(c, d, company_name, logo_path))

    pdf_buf.seek(0)

    # Clean up the temporary file
    os.unlink(temp_img_file.name)

    return pdf_buf


def subgroup_table(raw_df, org_name, subgroup, output_df, demo, company_name, logo_path, excel_file):
    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle(
        name='Normal',
        parent=styles['Normal'],
        fontName="Lexend",
        fontSize=9
    )

    font_path = "fonts/Lexend-Bold.ttf"
    pdfmetrics.registerFont(TTFont('Lexend-Bold', font_path))
    font_path = "fonts/Lexend-Regular.ttf"
    pdfmetrics.registerFont(TTFont('Lexend', font_path))
    font_path = "fonts/Lexend-Thin.ttf"
    pdfmetrics.registerFont(TTFont('Lexend-Light', font_path))

    raw_df['reportedAt'] = pd.to_datetime(raw_df['reportedAt'])
    latest_time = raw_df['reportedAt'].max()
    month = latest_time.strftime('%b %Y')
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_file:
        temp_file_path = temp_file.name

    # Define the PDF path
    pdf_path = temp_file_path
    pdf = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        rightMargin=0.35 * inch,
        leftMargin=0.35 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch
    )
    elements = []

    first_sheet_name = f"Avg. Scores for {demo}"
    scores_df = output_df[first_sheet_name]
    demographic_column = scores_df.columns[scores_df.columns.str.contains("Demographic", case=False)].tolist()[0]
    org_total_row = scores_df[scores_df[demographic_column].str.contains(subgroup, case=False, na=False)]

    demographic_column = scores_df.columns[scores_df.columns.str.contains("Demographic", case=False)].tolist()[0]

    # Extract the relevant rows for the subgroup and organization total
    org_total_row = scores_df[scores_df[demographic_column].str.contains(subgroup, case=False, na=False)]
    group_total = scores_df[scores_df[demographic_column].str.contains("Org Total", case=False, na=False)]

    # Elation norm values
    elation_norm = [62, 69, 65, 67]

    # Find the relevant rows and extract their values
    metrics = ["Wellbeing/Performance Potential", "Job Satisfaction", "Job Engagement", "Intent to Stay"]

    score_data = {}
    for metric in metrics:
        score_data[metric] = float(org_total_row[metric].values[0])

    # Calculate deltas for organization total and store in a new dictionary
    delta_org_data = {}
    for metric in metrics:
        org_value = float(group_total[metric].values[0])
        subgroup_value = float(org_total_row[metric].values[0])
        delta = subgroup_value - org_value
        delta_org_data[metric] = delta

    # Calculate deltas for elation norm and store in a new dictionary
    delta_elation_data = {}
    for metric, elation_value in zip(metrics, elation_norm):
        subgroup_value = float(org_total_row[metric].values[0])
        delta = subgroup_value - elation_value
        delta_elation_data[metric] = delta

    # Sort the score_data by values in descending order
    sorted_score_data = sorted(score_data.items(), key=lambda item: item[1], reverse=True)

    # Create the structured data for PDF
    outcomes = [
        ["Influencers", f"{month}", "vs Organization", "vs Elation Norm"],
    ]

    for metric, value in sorted_score_data:
        delta_org = delta_org_data[metric]
        delta_elation = delta_elation_data[metric]
        delta_org_str = f"{'+' if delta_org > 0 else ''}{int(delta_org)}"
        delta_elation_str = f"{'+' if delta_elation > 0 else ''}{int(delta_elation)}"
        outcomes.append([metric, str(int(value)), delta_org_str, delta_elation_str])


    sheet_name = 'Completion Rate'
    completion_rate_df = output_df[sheet_name]

    def get_demographic_data(df, subgroup):
        # Filter the data for the specified demographic
        filtered_data = df[df.iloc[:, 0] == subgroup]

        if not filtered_data.empty:
            # Extract values for total_members, completed_members, and completion_data
            total_members = filtered_data['total_members'].values[0]
            completed_members = filtered_data['completed_members'].values[0]
            completion_data = filtered_data['completion_rate'].values[0]

            return total_members, completed_members, completion_data
        else:
            return None, None, None

    total_members, completed_members, completion_data = get_demographic_data(completion_rate_df, subgroup)
    not_completed = total_members-completed_members
    datetime_obj = datetime.today()
    formatted_date = datetime_obj.strftime("%m/%d/%Y")
    date = formatted_date.replace("/0", "/").replace(" 0", " ")

    # Sample style
    styles = getSampleStyleSheet()

    # Data
    org_date = [
        ["Organization", f"{org_name}"],
        ["Date", f"{date}"]
    ]

    participation = [
        [f"Month", "Completed", "Not completed", "Total", "Participation Rate"],
        [f"{month}", f"{completed_members}", f"{not_completed}", f"{total_members}", f"{completion_data}"],
    ]

    demographic_column = scores_df.columns[scores_df.columns.str.contains("Demographic", case=False)].tolist()[0]

    # Extract the relevant rows for the subgroup and organization total
    org_total_row = scores_df[scores_df[demographic_column].str.contains(subgroup, case=False, na=False)]
    group_total = scores_df[scores_df[demographic_column].str.contains("Org Total", case=False, na=False)]

    # Elation norm values
    elation_norm = [65, 65, 65, 65, 65, 65, 65, 65, 65, 65, 65, 65, 65, 65, 65, 65, 65]

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
        score_data[metric] = float(org_total_row[metric].values[0])

    # Calculate deltas for organization total and store in a new dictionary
    delta_org_data = {}
    for metric in metrics:
        org_value = float(group_total[metric].values[0])
        subgroup_value = float(org_total_row[metric].values[0])
        delta = subgroup_value - org_value
        delta_org_data[metric] = delta

    # # Calculate deltas for elation norm and store in a new dictionary
    # delta_elation_data = {}
    # for metric, elation_value in zip(metrics, elation_norm):
    #     subgroup_value = float(org_total_row[metric].values[0])
    #     delta = subgroup_value - elation_value
    #     delta_elation_data[metric] = delta

    # Sort the score_data by values in descending order
    sorted_score_data = sorted(score_data.items(), key=lambda item: item[1], reverse=True)

    # Create the structured data for PDF
    workplace_indicators = [
        ["Influencers", f"{month}", "vs Organization"],
    ]

    for metric, value in sorted_score_data:
        delta_org = delta_org_data[metric]
        # delta_elation = delta_elation_data[metric]
        delta_org_str = f"{'+' if delta_org > 0 else ''}{int(delta_org)}"
        # delta_elation_str = f"{'+' if delta_elation > 0 else ''}{int(delta_elation)}"
        workplace_indicators.append([metric, str(int(value)), delta_org_str])

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
        'Workout Frequency and length'
    ]

    # Calculate score_data
    score_data = {}
    for metric in metrics:
        score_data[metric] = float(org_total_row[metric].values[0])

    # Calculate deltas for organization total
    delta_org_data = {}
    for metric in metrics:
        org_value = float(group_total[metric].values[0])
        subgroup_value = float(org_total_row[metric].values[0])
        delta = subgroup_value - org_value
        delta_org_data[metric] = delta

    # # Calculate deltas for elation norm
    # delta_elation_data = {}
    # for metric, elation_value in zip(metrics, elation_norm):
    #     subgroup_value = float(org_total_row[metric].values[0])
    #     delta = subgroup_value - elation_value
    #     delta_elation_data[metric] = delta

    # Sort the score_data by values in descending order
    sorted_score_data = sorted(score_data.items(), key=lambda item: item[1], reverse=True)

    # Create the structured data for PDF
    personal_indicators = [
        ["Influencers", f"{month}", "vs Organization"],
    ]

    for metric, value in sorted_score_data:
        delta_org = delta_org_data[metric]
        # delta_elation = delta_elation_data[metric]
        delta_org_str = f"{'+' if delta_org > 0 else ''}{int(delta_org)}"
        # delta_elation_str = f"{'+' if delta_elation > 0 else ''}{int(delta_elation)}"
        personal_indicators.append([metric, str(int(value)), delta_org_str])

    # Separate header from body
    header_data = participation[0]
    header_data2 = outcomes[0]
    header_data3 = workplace_indicators[0]
    header_data4 = personal_indicators[0]

    body_data = participation[1:]
    body_data2 = outcomes[1:]
    body_data3 = workplace_indicators[1:]
    body_data4 = personal_indicators[1:]

    # Wrap the body data in Paragraphs
    body_data_wrapped = [[Paragraph(cell, normal_style) for cell in row] for row in body_data]
    body_data_wrapped2 = [[Paragraph(cell, normal_style) for cell in row] for row in body_data2]
    body_data_wrapped3 = [[Paragraph(cell, normal_style) for cell in row] for row in body_data3]
    body_data_wrapped4 = [[Paragraph(cell, normal_style) for cell in row] for row in body_data4]

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

    title_table_data = [[Paragraph(f"{subgroup}", title_style)]]
    title_table = Table(title_table_data, colWidths=[550])
    title_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Lexend-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]))

    elements.append(title_table)
    elements.append(Spacer(1, 0.2 * inch))

    # Create and style the body table
    body_table = Table(org_date, colWidths=[130,420])
    body_table.setStyle(body_style)
    elements.append(body_table)
    elements.append(Spacer(1, 0.1 * inch))

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
    elements.append(Spacer(1, 0.1 * inch))

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
    header_table_data = [[Paragraph(header_data2[0], header_style), Paragraph(header_data2[1], header_style),
                          Paragraph(header_data2[2], header_style), Paragraph(header_data2[3], header_style)]]
    header_table = Table(header_table_data, colWidths=[235,105,105,105])
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
    table = Table(body_data_wrapped2, colWidths=[235,105,105,105])
    style = TableStyle([
        # ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        # ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    # Add colors for deltas
    for i in range(len(body_data2)):
        if int(body_data2[i][2]) >= 0:
            style.add('BACKGROUND', (2, i), (2, i), colors.HexColor("#A6CF5C"))
        else:
            style.add('BACKGROUND', (2, i), (2, i), colors.HexColor("#DE9C95"))

        if int(body_data2[i][3]) >= 0:
            style.add('BACKGROUND', (3, i), (3, i), colors.HexColor("#A6CF5C"))
        else:
            style.add('BACKGROUND', (3, i), (3, i), colors.HexColor("#DE9C95"))

    # Apply style to table
    table.setStyle(style)
    elements.append(table)
    elements.append(Spacer(1, 0.1 * inch))

    subtitle_table_data = [[Paragraph("Workplace Performance Influencers", subtitle_style)]]
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

    header_table_data = [[Paragraph(header_data3[0], header_style), Paragraph(header_data3[1], header_style),
                          Paragraph(header_data3[2], header_style)]]
    header_table = Table(header_table_data, colWidths=[235,105,210])
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
    table = Table(body_data_wrapped3, colWidths=[235,105,210])
    style = TableStyle([
        # ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        # ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    # Add colors for deltas
    for i in range(len(body_data3)):
        if int(body_data3[i][2]) >= 0:
            style.add('BACKGROUND', (2, i), (2, i), colors.HexColor("#A6CF5C"))
        else:
            style.add('BACKGROUND', (2, i), (2, i), colors.HexColor("#DE9C95"))

        # if int(body_data3[i][3]) >= 0:
        #     style.add('BACKGROUND', (3, i), (3, i), colors.HexColor("#A6CF5C"))
        # else:
        #     style.add('BACKGROUND', (3, i), (3, i), colors.HexColor("#DE9C95"))

    # Apply style to table
    table.setStyle(style)
    elements.append(table)
    # Add a page break
    elements.append(PageBreak())

    subtitle_table_data = [[Paragraph("Personal Performance Influencers", subtitle_style)]]
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

    header_table_data = [[Paragraph(header_data3[0], header_style), Paragraph(header_data3[1], header_style),
                          Paragraph(header_data3[2], header_style)]]
    header_table = Table(header_table_data, colWidths=[235,105,210])
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
    table = Table(body_data_wrapped4, colWidths=[235,105,210])
    style = TableStyle([
        # ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        # ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    # Add colors for deltas
    for i in range(len(body_data4)):
        if int(body_data4[i][2]) >= 0:
            style.add('BACKGROUND', (2, i), (2, i), colors.HexColor("#A6CF5C"))
        else:
            style.add('BACKGROUND', (2, i), (2, i), colors.HexColor("#DE9C95"))

    # Apply style to table
    table.setStyle(style)
    elements.append(table)

    # Define a style for the header
    subtitle_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Heading1'],
        alignment=0,
        fontSize=14,
        fontName="Lexend-Bold",
        textColor=colors.black
    )


    class HorizontalLine(Flowable):
        def __init__(self, width):
            super().__init__()
            self.width = width

        def draw(self):
            self.canv.line((self.canv._pagesize[0] - self.width) / 2, 0, (self.canv._pagesize[0] + self.width) / 2, 0)

    def add_horizontal_line(elements, width=1000):
        elements.append(Spacer(1, 24))
        elements.append(HorizontalLine(width))
        elements.append(Spacer(1, 24))

    add_horizontal_line(elements)

    elements.append(Paragraph("Recommendations", title_style))

    # Ensure you have a common column for merging, typically the demographic name
    common_column = demo

    raw_df['reportedAt'] = pd.to_datetime(raw_df['reportedAt'])
    sheet_name = 'Completion Rate'
    lowest_scores_sheet = 'Lowest Scores'

    # Read the data from the sheets
    data = pd.read_excel(excel_file, sheet_name=sheet_name)
    lowest_scores = pd.read_excel(excel_file, sheet_name=lowest_scores_sheet)

    body_style = normal_style

    # Merge the two DataFrames
    merged_df = pd.merge(data, lowest_scores, on=common_column, how='inner')

    # Filter the merged DataFrame for the specific subgroup
    subgroup_merged = merged_df[merged_df[common_column] == subgroup]

    # print(subgroup_merged)

    descriptions_df = pd.read_csv('CSVs/descriptions.csv')

    # Set to keep track of influencers that have already been added
    added_influencers = set()

    def calculate_dynamic_cutoff(row):
        lowest_inf = float(row['Lowest Influencer'].split(':')[1].strip())
        second_lowest_inf = float(row['Second Lowest Influencer'].split(':')[1].strip())
        return (lowest_inf + second_lowest_inf) / 2

    def add_influencer(influencer, style, cutoff):
        if float(influencer[1]) <= cutoff:
            influencer_text = f"{influencer[0]} - {round(float(influencer[1]))}"
            elements.append(Paragraph(influencer_text, style))

            # Check if the influencer has already been added
            if influencer[0] not in added_influencers:
                # Look up the definition and recommendation from the descriptions_df
                description_row = descriptions_df[descriptions_df['Influencer'] == influencer[0]]
                if not description_row.empty:
                    definition = description_row['Definition'].values[0].replace('\n', '<br/>')
                    recommendation = description_row['Recommendation'].values[0].replace('\n', '<br/>')

                    definition_paragraph = Paragraph(f"<b>Definition:</b><br/>{definition}", body_style)
                    recommendation_paragraph = Paragraph(f"<b>Recommendations:</b><br/>{recommendation}", body_style)

                    elements.append(definition_paragraph)
                    elements.append(recommendation_paragraph)

                    # Mark this influencer as added
                    added_influencers.add(influencer[0])
                else:
                    elements.append(Paragraph("No description available.", body_style))


    # Loop through the filtered DataFrame for the specific subgroup
    for _, row in subgroup_merged.iterrows():
        if row['completed_members'] >= 1:

            # Calculate the dynamic cutoff value
            dynamic_cutoff = calculate_dynamic_cutoff(row) + 50

            # Extract and sort influencers
            lowest_inf = row['Lowest Influencer'].split(':')
            second_lowest_inf = row['Second Lowest Influencer'].split(':')

            # Create a list of influencers
            influencers = [
                (lowest_inf[0], lowest_inf[1]),
                (second_lowest_inf[0], second_lowest_inf[1])
            ]

            # Retrieve additional lowest influencers if available
            additional_influencers = []
            if 'Third Lowest Influencer' in row:
                third_lowest_inf = row['Third Lowest Influencer'].split(':')
                additional_influencers.append((third_lowest_inf[0], third_lowest_inf[1]))

            # Merge all influencers and sort by their scores
            all_influencers = influencers + additional_influencers
            all_influencers.sort(key=lambda x: float(x[1]))

            # Limit to a maximum of 3 influencers but at least 1
            num_to_print = min(max(len(all_influencers), 2), 2)

            large_bold_style = ParagraphStyle(
                'LargeBold',
                parent=styles['BodyText'],
                fontSize=11,
                leading=20,
                spaceAfter=8,
                fontName='Lexend-Bold'
            )

            for i in range(num_to_print):
                add_influencer(all_influencers[i], large_bold_style, dynamic_cutoff)



    # Build the PDF with headers
    pdf.build(elements, onFirstPage=lambda c, d: culture_header(c, d, company_name, logo_path),
              onLaterPages=lambda c, d: culture_header(c, d, company_name, logo_path))

    with open(pdf_path, 'rb') as pdf_file:
        pdf_data = pdf_file.read()

    return pdf_data


# Function to merge glossary and culture report into one PDF
def merge_pdfs(glossary_pdf_bytes, culture_report_pdf_bytes, table_pdf_bytes, company_name, logo_path):
    merged_pdf = PdfWriter()

    # Read table PDF and add it first
    table_report_reader = PdfReader(BytesIO(table_pdf_bytes))
    for page in range(len(table_report_reader.pages)):
        merged_pdf.add_page(table_report_reader.pages[page])

    # Read glossary PDF
    glossary_reader = PdfReader(BytesIO(glossary_pdf_bytes))
    for page in range(len(glossary_reader.pages)):
        merged_pdf.add_page(glossary_reader.pages[page])

    # Read culture report PDF
    culture_report_reader = PdfReader(BytesIO(culture_report_pdf_bytes))
    for page in range(len(culture_report_reader.pages)):
        merged_pdf.add_page(culture_report_reader.pages[page])


    merged_pdf_bytes = BytesIO()
    merged_pdf.write(merged_pdf_bytes)
    merged_pdf_bytes.seek(0)

    return merged_pdf_bytes

