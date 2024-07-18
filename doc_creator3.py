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
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image


def create_pdf_with_header_and_recommendations(excel_file, output_pdf, company_name, logo_path, raw_df, demo):

    font_path = "fonts/Lexend-Bold.ttf"
    pdfmetrics.registerFont(TTFont('Lexend-Bold', font_path))
    font_path = "fonts/Lexend-Regular.ttf"
    pdfmetrics.registerFont(TTFont('Lexend', font_path))
    font_path = "fonts/Lexend-Thin.ttf"
    pdfmetrics.registerFont(TTFont('Lexend-Light', font_path))

    # Define styles
    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle(
        name='Normal',
        parent=styles['BodyText'],
        fontName="Lexend",
        fontSize=9
    )

    bold_style = ParagraphStyle(
        name='Normal',
        parent=styles['Normal'],
        fontName="Lexend-Bold",
        fontSize=9
    )

    title_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Heading1'],
        alignment=1,
        fontSize=20,
        valign="TOP",
        fontName="Lexend-Bold",
        textColor=colors.black
    )

    subtitle_style = ParagraphStyle(
        name='LeftH2',
        parent=styles['Normal'],
        alignment=0,
        fontSize=14,
        valign="TOP",
        fontName="Lexend-Bold",
        textColor=colors.black
    )

    raw_df['reportedAt'] = pd.to_datetime(raw_df['reportedAt'])
    sheet_name = 'Completion Rate'
    lowest_scores_sheet = 'Lowest Scores'

    # Read the data from the sheets
    data = pd.read_excel(excel_file, sheet_name=sheet_name)
    lowest_scores = pd.read_excel(excel_file, sheet_name=lowest_scores_sheet)

    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=letter,
        rightMargin=0.35 * inch,
        leftMargin=0.35 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch
    )
    elements = []

    # Add subgroup table elements
    output_df = pd.read_excel(excel_file, sheet_name=None)
    demo2 = "Org Total"
    subgroup_elements = subgroup_table(raw_df, company_name, demo2, output_df)
    elements.extend(subgroup_elements)

    body_style = normal_style

    large_bold_style = ParagraphStyle(
        'LargeBold',
        parent=styles['BodyText'],
        fontSize=11,
        leading=20,
        spaceAfter=8,
        fontName='Lexend-Bold'
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
    elements.append(Paragraph("Recommendations", title_style))
    elements.append(Spacer(1, 12))

    # Ensure you have a common column for merging, typically the demographic name
    common_column = demo

    # Merge the two DataFrames
    merged_df = pd.merge(data, lowest_scores, on=common_column, how='inner')

    # Sort the merged DataFrame by 'Demographic Size' in descending order and select the top 5
    top_5_merged = merged_df.sort_values(by='Demographic Size', ascending=False).head(5)

    print(top_5_merged)

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

            elements.append(Spacer(1, 12))


    for _, row in top_5_merged.iterrows():
        group_name = row[common_column]
        completed = row['completed_members']
        total = row['total_members']
        completion_rate = row['completion_rate']

        section_title = f"{group_name} – {completed} of {total}, {completion_rate}"
        elements.append(Paragraph(section_title, subtitle_style))
        elements.append(Spacer(1, 12))

        # Calculate the dynamic cutoff value
        dynamic_cutoff = calculate_dynamic_cutoff(row) + 3

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
        num_to_print = min(max(len(all_influencers), 1), 3)

        for i in range(num_to_print):
            add_influencer(all_influencers[i], large_bold_style, dynamic_cutoff)

        elements.append(Spacer(1, 12))

    additional_recommendations = Paragraph(
        "Additional recommendations for each Influencer are listed on the <b>Read More</b> pages within the visualization platform.",
        body_style)
    elements.append(additional_recommendations)

    # Build the PDF
    doc.build(elements, onFirstPage=lambda canvas, doc: header(canvas, doc, company_name, logo_path),
              onLaterPages=lambda canvas, doc: header(canvas, doc, company_name, logo_path))

def header(canvas, doc, company_name, logo_path, scale_factor=0.14):
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

    # Load the logo
    logo = Image(logo_path)

    # Calculate scaled dimensions
    scaled_width = logo.imageWidth * scale_factor
    scaled_height = logo.imageHeight * scale_factor

    # Draw the logo at the top right corner with scaling
    logo.drawWidth = scaled_width
    logo.drawHeight = scaled_height
    logo.drawOn(canvas, letter[0] - scaled_width - .35 * inch, letter[1] - scaled_height - 0.4 * inch)

    canvas.restoreState()
