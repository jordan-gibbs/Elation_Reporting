import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import Flowable
from datetime import datetime
import pandas as pd
from overall_report_table import subgroup_table

def create_pdf_with_header_and_recommendations(excel_file, output_pdf, company_name, logo_path, raw_df, demo):
    sheet_name = 'Completion Rate'
    lowest_scores_sheet = 'Lowest Scores'

    # Read the data from the sheets
    data = pd.read_excel(excel_file, sheet_name=sheet_name)
    lowest_scores = pd.read_excel(excel_file, sheet_name=lowest_scores_sheet)

    doc = SimpleDocTemplate(output_pdf, pagesize=letter)
    elements = []

    # Add subgroup table elements
    output_df = pd.read_excel(excel_file, sheet_name=None)
    demo2 = "Org Total"
    subgroup_elements = subgroup_table(raw_df, company_name, demo2, output_df)
    elements.extend(subgroup_elements)

    # Define styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name='TitleStyle', parent=styles['Title'], alignment=0)
    normal_style = styles['BodyText']
    normal_style.fontName = 'Helvetica'
    normal_style.fontSize = 12
    bold_style = ParagraphStyle(name='BoldStyle', parent=styles['Normal'], spaceAfter=12, fontName='Helvetica-Bold')
    body_style = ParagraphStyle(name='BodyStyle', parent=styles['Normal'], spaceBefore=12)

    large_bold_style = ParagraphStyle(
        'LargeBold',
        parent=styles['BodyText'],
        fontSize=12,
        leading=20,
        spaceAfter=10,
        fontName='Helvetica-Bold'
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
    # Add recommendations
    elements.append(Paragraph("RECOMMENDATIONS", title_style))
    elements.append(Spacer(1, 12))

    # Ensure you have a common column for merging, typically the demographic name
    common_column = demo

    # Merge the two DataFrames
    merged_df = pd.merge(data, lowest_scores, on=common_column, how='inner')

    # Sort the merged DataFrame by 'Demographic Size' in descending order and select the top 5
    top_5_merged = merged_df.sort_values(by='Demographic Size', ascending=False).head(5)

    descriptions_df = pd.read_csv('CSVs/descriptions.csv')

    # Function to add influencer details to elements
    def add_influencer(influencer, style):
        if float(influencer[1]) <= 55:
            influencer_text = f"{influencer[0]} - {round(float(influencer[1]))}"
            elements.append(Paragraph(influencer_text, style))

            # Look up the definition and recommendation from the descriptions_df
            description_row = descriptions_df[descriptions_df['Influencer'] == influencer[0]]
            if not description_row.empty:
                definition = description_row['Definition'].values[0]
                recommendation = description_row['Recommendation'].values[0]

                definition_paragraph = Paragraph(f"<b>Definition:</b> {definition}", body_style)
                recommendation_paragraph = Paragraph(f"<b>Recommendation:</b> {recommendation}", body_style)

                elements.append(definition_paragraph)
                elements.append(recommendation_paragraph)
            else:
                elements.append(Paragraph("No description available.", body_style))

            elements.append(Spacer(1, 12))

    # Iterate through the top 5 largest demographics
    for _, row in top_5_merged.iterrows():
        group_name = row[common_column]
        completed = row['completed_members']
        total = row['total_members']
        completion_rate = row['completion_rate']

        section_title = f"{group_name} – {completed} of {total}, {completion_rate}"
        elements.append(Paragraph(section_title, styles['Heading2']))
        elements.append(Spacer(1, 12))

        lowest_inf = row['Lowest Influencer'].split(':')
        second_lowest_inf = row['Second Lowest Influencer'].split(':')

        add_influencer(lowest_inf, large_bold_style)
        add_influencer(second_lowest_inf, large_bold_style)

        elements.append(Spacer(1, 12))

    additional_recommendations = Paragraph(
        "Additional recommendations for each Influencer are listed on the <b>Read More</b> pages within the visualization platform.",
        body_style)
    elements.append(additional_recommendations)

    # Build the PDF
    doc.build(elements, onFirstPage=lambda canvas, doc: header(canvas, doc, company_name, logo_path),
              onLaterPages=lambda canvas, doc: header(canvas, doc, company_name, logo_path))

def header(canvas, doc, company_name, logo_path, scale_factor=0.14):
    canvas.saveState()
    canvas.setFont('Helvetica-Bold', 12)
    canvas.setFillColor(colors.grey)
    header_text = f"Wb^2 Report – {company_name}, ALL"
    canvas.drawString(inch, letter[1] - 0.60 * inch, header_text)

    # Load the logo
    logo = Image(logo_path)

    # Calculate scaled dimensions
    scaled_width = logo.imageWidth * scale_factor
    scaled_height = logo.imageHeight * scale_factor

    # Draw the logo at the top right corner with scaling
    logo.drawWidth = scaled_width
    logo.drawHeight = scaled_height
    logo.drawOn(canvas, letter[0] - scaled_width - 1.25 * inch, letter[1] - scaled_height - 0.4 * inch)

    canvas.restoreState()
