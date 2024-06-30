import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image

def create_pdf_with_header_and_recommendations(excel_file, output_pdf, company_name, logo_path):
    sheet_name = 'Completion Rate by Department'
    lowest_scores_sheet = 'Lowest Scores'

    # Read the data from the sheets
    data = pd.read_excel(excel_file, sheet_name=sheet_name)
    lowest_scores = pd.read_excel(excel_file, sheet_name=lowest_scores_sheet)

    doc = SimpleDocTemplate(output_pdf, pagesize=letter)
    elements = []

    # Define styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name='TitleStyle', parent=styles['Title'], alignment=0)
    normal_style = styles['BodyText']
    normal_style.fontName = 'Helvetica'
    normal_style.fontSize = 12
    bold_style = ParagraphStyle(name='BoldStyle', parent=styles['Normal'], spaceAfter=12, fontName='Helvetica-Bold')
    body_style = ParagraphStyle(name='BodyStyle', parent=styles['Normal'], spaceBefore=12)

    # Add recommendations
    elements.append(Paragraph("RECOMMENDATIONS", title_style))
    elements.append(Spacer(1, 12))

    for _, row in data.iterrows():
        group_name = row['groupName']
        completed = row['completed_members']
        total = row['total_members']
        completion_rate = row['completion_rate']

        section_title = f"{group_name} – {completed} of {total}, {completion_rate}%"
        elements.append(Paragraph(section_title, styles['Heading2']))
        elements.append(Spacer(1, 12))

        lowest_row = lowest_scores[lowest_scores['Department'] == group_name]
        if not lowest_row.empty:
            lowest_inf = lowest_row.iloc[0]['Lowest Influencer'].split(':')
            second_lowest_inf = lowest_row.iloc[0]['Second Lowest Influencer'].split(':')
            third_lowest_inf = lowest_row.iloc[0]['Third Lowest Influencer'].split(':')

            def add_influencer(influencer, style):
                if float(influencer[1]) <= 40:
                    influencer_text = f"{influencer[0]} - {round(float(influencer[1]))}"
                    elements.append(Paragraph(influencer_text, style))
                    ai = Paragraph("[Insert AI Generated Recommendation]", body_style)
                    elements.append(ai)
                    elements.append(Spacer(1, 12))

            add_influencer(lowest_inf, bold_style)
            add_influencer(second_lowest_inf, bold_style)
            add_influencer(third_lowest_inf, bold_style)

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
