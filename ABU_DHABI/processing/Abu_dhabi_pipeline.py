"""
Generate professional PDF documentation for Abu Dhabi DB1 Pipeline.
"""

from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

OUTPUT_PDF = Path(__file__).parent / "Abu_Dhabi_Pipeline_Documentation.pdf"
DOC_DATE = date.today().strftime("%d %B %Y")
DOC_VERSION = "1.0"


def _header_footer(canvas, doc):
    canvas.saveState()
    width, height = A4

    if doc.page > 1:
        canvas.setStrokeColor(colors.HexColor("#1F4E79"))
        canvas.setLineWidth(0.5)
        canvas.line(2 * cm, height - 1.6 * cm, width - 2 * cm, height - 1.6 * cm)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.setFillColor(colors.HexColor("#1F4E79"))
        canvas.drawString(2 * cm, height - 1.25 * cm, "Abu Dhabi Real Estate Data Pipeline")
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawRightString(width - 2 * cm, height - 1.25 * cm, f"Version {DOC_VERSION}")

    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawString(2 * cm, 1.2 * cm, f"Confidential | Generated {DOC_DATE}")
    canvas.drawRightString(width - 2 * cm, 1.2 * cm, f"Page {doc.page}")
    canvas.restoreState()


def _cover(canvas, doc):
    canvas.saveState()
    width, height = A4

    canvas.setFillColor(colors.HexColor("#1F4E79"))
    canvas.rect(0, height - 8 * cm, width, 8 * cm, fill=1, stroke=0)

    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 26)
    canvas.drawCentredString(width / 2, height - 4.5 * cm, "Abu Dhabi Real Estate")
    canvas.drawCentredString(width / 2, height - 5.5 * cm, "Data Processing Pipeline")
    canvas.setFont("Helvetica", 14)
    canvas.drawCentredString(width / 2, height - 6.6 * cm, "DB1 Standardization & Enrichment")

    canvas.setFillColor(colors.HexColor("#2E75B6"))
    canvas.rect(0, 0, width, 0.4 * cm, fill=1, stroke=0)

    canvas.setFillColor(colors.HexColor("#333333"))
    canvas.setFont("Helvetica", 11)
    canvas.drawCentredString(width / 2, height - 11 * cm, "Technical Documentation")
    canvas.setFont("Helvetica", 10)
    canvas.drawCentredString(width / 2, height - 12 * cm, f"Document Version: {DOC_VERSION}")
    canvas.drawCentredString(width / 2, height - 12.7 * cm, f"Date: {DOC_DATE}")
    canvas.drawCentredString(width / 2, height - 13.4 * cm, "Source: ADREC (Abu Dhabi Real Estate Centre)")

    canvas.restoreState()


def build_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="DocTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=22,
            textColor=colors.HexColor("#1F4E79"),
            spaceAfter=14,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Section",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=14,
            textColor=colors.HexColor("#1F4E79"),
            spaceBefore=16,
            spaceAfter=10,
            borderPadding=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SubSection",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=11,
            textColor=colors.HexColor("#2E75B6"),
            spaceBefore=10,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Body",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            alignment=TA_JUSTIFY,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="DocBullet",
            parent=styles["Body"],
            leftIndent=14,
            bulletIndent=0,
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="TOC",
            parent=styles["Body"],
            fontSize=10,
            leftIndent=0,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="DocCode",
            parent=styles["Normal"],
            fontName="Courier",
            fontSize=8,
            leading=10,
            backColor=colors.HexColor("#F5F5F5"),
            borderPadding=4,
        )
    )
    return styles


def table(data, col_widths=None, header=True):
    tbl = Table(data, colWidths=col_widths, repeatRows=1 if header else 0)
    style = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    if header:
        style.extend(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ]
        )
        style.append(("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFD")]))
    tbl.setStyle(TableStyle(style))
    return tbl


def build_story(styles):
    story = []

    story.append(Spacer(1, 9 * cm))
    story.append(Paragraph("Prepared for internal data operations and DB1 ingestion workflows.", styles["Body"]))
    story.append(PageBreak())

    # Table of contents
    story.append(Paragraph("Table of Contents", styles["DocTitle"]))
    toc_items = [
        "1. Executive Summary",
        "2. Document Purpose & Scope",
        "3. Pipeline Overview",
        "4. Input Data Specification",
        "5. Processing Stages",
        "6. Field Mapping & Transformations",
        "7. Standardization Rules",
        "8. Geocoding Enrichment",
        "9. Index Assignment",
        "10. Output Deliverables",
        "11. Dependencies & Execution",
        "12. Data Quality & Limitations",
        "13. Appendix: DB1 Output Schema",
    ]
    for item in toc_items:
        story.append(Paragraph(item, styles["TOC"]))
    story.append(PageBreak())

    # 1. Executive Summary
    story.append(Paragraph("1. Executive Summary", styles["Section"]))
    story.append(
        Paragraph(
            "This document describes the end-to-end data processing pipeline implemented in "
            "<b>pipeline.ipynb</b> for Abu Dhabi property transaction records sourced from ADREC. "
            "The pipeline ingests raw transaction data, applies deduplication and business-rule "
            "transformations, standardizes records into the DB1 schema, enriches location and project "
            "coordinates, and assigns non-RERA project indexes for downstream analytics and ingestion.",
            styles["Body"],
        )
    )
    story.append(
        Paragraph(
            "<b>Key outcomes:</b> ~109,388 deduplicated transaction records transformed into 63 DB1 "
            "columns; 132 unique districts geocoded via ArcGIS; 367 unique projects geocoded via "
            "Google Maps (362 successful); non-RERA indexes assigned from <i>nr101</i> onward.",
            styles["Body"],
        )
    )

    # 2. Purpose & Scope
    story.append(Paragraph("2. Document Purpose & Scope", styles["Section"]))
    story.append(
        Paragraph(
            "<b>Purpose:</b> Provide a complete technical reference for developers, analysts, and "
            "data engineers responsible for maintaining, auditing, or extending the Abu Dhabi pipeline.",
            styles["Body"],
        )
    )
    story.append(Paragraph("<b>In scope:</b>", styles["SubSection"]))
    for item in [
        "Raw CSV ingestion and duplicate removal",
        "Property type and project type standardization",
        "DB1 column mapping and export",
        "District-level geocoding (ArcGIS)",
        "Project-level geocoding (Google Maps / Selenium)",
        "Non-RERA index generation",
    ]:
        story.append(Paragraph(f"• {item}", styles["DocBullet"]))
    story.append(Paragraph("<b>Out of scope:</b>", styles["SubSection"]))
    for item in [
        "RERA-indexed project matching",
        "LLM-based field extraction",
        "Buyer/seller identity enrichment",
    ]:
        story.append(Paragraph(f"• {item}", styles["DocBullet"]))

    # 3. Pipeline Overview
    story.append(Paragraph("3. Pipeline Overview", styles["Section"]))
    story.append(
        Paragraph(
            "The pipeline follows a sequential ETL architecture with optional enrichment stages. "
            "Each stage produces an intermediate CSV that can be validated independently.",
            styles["Body"],
        )
    )
    flow_data = [
        ["Stage", "Input", "Output", "Tool / Method"],
        ["1", "Abu_Dhabi.csv", "In-memory DataFrame", "pandas read_csv"],
        ["2", "DataFrame", "Deduped DataFrame", "drop_duplicates()"],
        ["3", "DataFrame", "Transformed DataFrame", "transform_abudhabi_transactions()"],
        ["4", "Transformed data", "Abu_Dhabi_Cleaned.csv", "DB1 column selection"],
        ["5", "Cleaned CSV", "geocoded_location_output.csv", "ArcGIS geocoder"],
        ["6", "Abu_Dhabi_DB1.csv", "Abu_Dhabi_DB1_Project_Geocoded.csv", "Google Maps + Selenium"],
        ["7", "Abu_Dhabi_DB1.csv", "Abu_Dhabi_DB1_Index_Filled.csv", "Index assignment logic"],
    ]
    story.append(table(flow_data, col_widths=[1.2 * cm, 4.2 * cm, 5.5 * cm, 5.5 * cm]))
    story.append(Spacer(1, 0.3 * cm))
    story.append(
        Paragraph(
            "<b>Workflow diagram (logical flow):</b><br/>"
            "Raw CSV → Dedupe → Map & Transform → DB1 Export → Location Geocode → "
            "Project Geocode → Index Fill → Final DB1 Dataset",
            styles["Body"],
        )
    )

    # 4. Input Data
    story.append(PageBreak())
    story.append(Paragraph("4. Input Data Specification", styles["Section"]))
    story.append(
        Paragraph(
            "<b>Source file:</b> Abu_Dhabi.csv<br/>"
            "<b>Data provider:</b> ADREC (Abu Dhabi Real Estate Centre)<br/>"
            "<b>Original record count:</b> 115,514 rows<br/>"
            "<b>After deduplication:</b> 109,388 rows<br/>"
            "<b>Source columns:</b> 14",
            styles["Body"],
        )
    )
    input_cols = [
        ["#", "Source Column", "Description"],
        ["1", "Asset Class", "High-level asset category (residential, commercial, etc.)"],
        ["2", "Property Type", "Detailed property classification (apartment, villa, plot, etc.)"],
        ["3", "Sale Application Date", "Transaction date in ISO format (YYYY-MM-DD)"],
        ["4", "Property Sold Area (SQM)", "Net sold area in square metres"],
        ["5", "Land Plot Ground Area (SQM)", "Land/plot ground area in square metres"],
        ["6", "Property Layout", "Unit layout (studio, 1 bed, 2 beds, unclassified, etc.)"],
        ["7", "District", "District / location name"],
        ["8", "Community", "Community / micro-market identifier"],
        ["9", "Project Name", "Development or project name"],
        ["10", "Property Sale Price (AED)", "Transaction price in AED"],
        ["11", "Property Sold Share", "Fractional share sold (typically 1.0)"],
        ["12", "Rate (AED per SQM)", "Derived price per square metre"],
        ["13", "Sale Application Type", "off-plan or ready"],
        ["14", "Sale Sequence", "primary or secondary sale"],
    ]
    story.append(table(input_cols, col_widths=[0.8 * cm, 4.5 * cm, 11 * cm]))

    # 5. Processing Stages
    story.append(Paragraph("5. Processing Stages", styles["Section"]))

    stages = [
        (
            "5.1 Read Raw Data",
            "Loads Abu_Dhabi.csv using pandas. Displays sample records for visual validation.",
        ),
        (
            "5.2 Drop Duplicates",
            "Removes exact duplicate rows using df.drop_duplicates(), reducing records from "
            "115,514 to 109,388 (6,126 duplicates removed).",
        ),
        (
            "5.3 Property Type Mapping",
            "Maps 43 raw property types to standardized DB1 categories: Flat, Villa, Office, Shop, "
            "Plot, and Others using PROPERTY_TYPE_MAP.",
        ),
        (
            "5.4 Project Type Mapping",
            "Maps Asset Class values to standardized project types (Residential, Commercial, "
            "Industrial, etc.) using PROJECT_TYPE_MAP.",
        ),
        (
            "5.5 Standard Transformation",
            "Core function transform_abudhabi_transactions() applies date parsing, field renaming, "
            "derived fields (year, quarter), fixed metadata, and null defaults for unavailable DB1 fields.",
        ),
        (
            "5.6 DB1 Column Selection",
            "Selects and orders 63 standard DB1 columns. Missing columns are created with null values. "
            "Output saved as Abu_Dhabi_Cleaned.csv.",
        ),
        (
            "5.7 Location Geocoding",
            "Geocodes 132 unique district names using ArcGIS API. Address format: "
            "'{District}, Abu Dhabi, United Arab Emirates'. Results written to location_latitude "
            "and location_longitude.",
        ),
        (
            "5.8 Project Geocoding",
            "Geocodes 367 unique projects using Google Maps via Selenium WebDriver. Search format: "
            "'{Project Name}, Abu Dhabi, UAE'. 362 projects successfully geocoded; 5 failed.",
        ),
        (
            "5.9 Non-RERA Index Assignment",
            "Assigns unique project indexes in format nr101, nr102, ... for records without existing "
            "indexes. Same location + project combination reuses the same index.",
        ),
    ]
    for title, text in stages:
        story.append(Paragraph(title, styles["SubSection"]))
        story.append(Paragraph(text, styles["Body"]))

    # 6. Field Mapping
    story.append(PageBreak())
    story.append(Paragraph("6. Field Mapping & Transformations", styles["Section"]))
    mapping_data = [
        ["Source Field", "DB1 Target Field", "Transformation Rule"],
        ["Project Name", "project_name", "Direct copy"],
        ["District", "location_name", "Direct copy"],
        ["District", "city_name", "Fixed to 'Abu Dhabi'"],
        ["Community", "sub_locality, micro_market", "Direct copy"],
        ["Property Sale Price (AED)", "agreement_price", "Numeric conversion"],
        ["Property Sold Area (SQM)", "net_carpet_area_sq_m", "Numeric conversion"],
        ["Sale Application Date", "transaction_date", "Parse YYYY-MM-DD → format DD-MM-YYYY"],
        ["Sale Application Date", "date_of_agreement_execution", "Same as transaction_date"],
        ["Sale Application Date", "year, quarter", "Derived (e.g. Q2-2026)"],
        ["Property Type", "property_type_raw", "Lowercase + strip"],
        ["Property Type", "property_type", "Mapped via PROPERTY_TYPE_MAP"],
        ["Property Layout", "property_layout", "Direct copy (original value retained)"],
        ["Property Layout", "unit_configuration", "Lowercase + strip (text retained)"],
        ["Asset Class", "project_type", "Mapped via PROJECT_TYPE_MAP"],
        ["Sale Application Type", "furnishing_status", "Direct copy (off-plan / ready)"],
        ["Sale Sequence", "sale_type", "Direct copy (primary / secondary)"],
        ["Property Type + Asset Class", "property_description", "Concatenated with ' | '"],
        ["—", "transaction_category", "Fixed: 'Sale'"],
        ["—", "country_name", "Fixed: 'UAE'"],
        ["—", "state_name", "Fixed: 'Abu Dhabi'"],
        ["—", "data_source", "Fixed: 'Adrec'"],
        ["—", "source_accessibility", "Fixed: 'Download'"],
        ["—", "source_accessibility_way", "Fixed: 'Easy'"],
    ]
    story.append(table(mapping_data, col_widths=[4.5 * cm, 4.5 * cm, 7.4 * cm]))

    # 7. Standardization Rules
    story.append(Paragraph("7. Standardization Rules", styles["Section"]))
    story.append(Paragraph("7.1 Property Type Mapping (PROPERTY_TYPE_MAP)", styles["SubSection"]))
    prop_map = [
        ["Raw Value (sample)", "Standardized Value"],
        ["apartment, duplex, residential complex", "Flat"],
        ["villa, townhouse / attached villa, palace", "Villa"],
        ["office, office complex", "Office"],
        ["retail, mall / market / retail center", "Shop"],
        ["plot for *, other * plot", "Plot"],
        ["clinic, factory, hotel, mosque, other, etc.", "Others"],
        ["Unmapped values", "Others (default)"],
    ]
    story.append(table(prop_map, col_widths=[8 * cm, 8.4 * cm]))

    story.append(Paragraph("7.2 Project Type Mapping (PROJECT_TYPE_MAP)", styles["SubSection"]))
    proj_map = [
        ["Raw Asset Class", "Standardized project_type"],
        ["residential", "Residential"],
        ["commercial", "Commercial"],
        ["agricultural", "Agricultural"],
        ["educational", "Educational"],
        ["healthcare", "Healthcare"],
        ["industrial & storage", "Industrial"],
        ["infrastructural", "Infrastructure"],
        ["recreational", "Recreational"],
        ["religious", "Religious"],
        ["other", "Others"],
    ]
    story.append(table(proj_map, col_widths=[8 * cm, 8.4 * cm]))

    story.append(Paragraph("7.3 Date Handling", styles["SubSection"]))
    story.append(
        Paragraph(
            "Source dates are in ISO 8601 format (YYYY-MM-DD). The pipeline parses using "
            "format='%Y-%m-%d' and exports in DD-MM-YYYY format. Quarter is derived as "
            "Q{1-4}-{YYYY} (e.g. Q2-2026).",
            styles["Body"],
        )
    )

    # 8. Geocoding
    story.append(PageBreak())
    story.append(Paragraph("8. Geocoding Enrichment", styles["Section"]))
    story.append(Paragraph("8.1 District Geocoding (ArcGIS)", styles["SubSection"]))
    geo1 = [
        ["Parameter", "Value"],
        ["Geocoder", "geopy.geocoders.ArcGIS"],
        ["Input column", "location_name (District)"],
        ["Search template", "{District}, Abu Dhabi, United Arab Emirates"],
        ["Output columns", "location_latitude, location_longitude"],
        ["Unique locations", "132"],
        ["Rate limiting", "1 second sleep between requests"],
        ["Output file", "geocoded_location_output.csv"],
    ]
    story.append(table(geo1, col_widths=[4.5 * cm, 12.3 * cm]))

    story.append(Paragraph("8.2 Project Geocoding (Google Maps)", styles["SubSection"]))
    geo2 = [
        ["Parameter", "Value"],
        ["Method", "Selenium WebDriver + Google Maps search URL"],
        ["Input column", "project_name"],
        ["Search template", "{Project Name}, Abu Dhabi, UAE"],
        ["Output columns", "project_latitude, project_longitude"],
        ["Unique projects", "367"],
        ["Success rate", "362 OK / 5 failed"],
        ["Output file", "Abu_Dhabi_DB1_Project_Geocoded.csv"],
    ]
    story.append(table(geo2, col_widths=[4.5 * cm, 12.3 * cm]))

    # 9. Index Assignment
    story.append(Paragraph("9. Index Assignment", styles["Section"]))
    story.append(
        Paragraph(
            "Records without an existing <b>index</b> value receive a non-RERA identifier starting "
            "from <b>nr101</b>. The assignment logic:",
            styles["Body"],
        )
    )
    for item in [
        "Groups records by (location_name, project_name) combination",
        "Reuses existing index if the same combination already has one",
        "Increments sequentially (nr101, nr102, nr103, ...) for new combinations",
        "Skips rows where project_name is null or blank",
        "Output saved as Abu_Dhabi_DB1_Index_Filled.csv",
    ]:
        story.append(Paragraph(f"• {item}", styles["DocBullet"]))

    # 10. Output Deliverables
    story.append(Paragraph("10. Output Deliverables", styles["Section"]))
    outputs = [
        ["File Name", "Description", "Records"],
        ["Abu_Dhabi_Cleaned.csv", "DB1-standardized transactions (63 columns)", "~109,388"],
        ["geocoded_location_output.csv", "Cleaned data + district lat/long", "~109,388"],
        ["Abu_Dhabi_DB1_Project_Geocoded.csv", "DB1 data + project coordinates", "~109,388"],
        ["Abu_Dhabi_DB1_Index_Filled.csv", "Final dataset with nr-indexes assigned", "~109,388"],
    ]
    story.append(table(outputs, col_widths=[5.5 * cm, 7.5 * cm, 3.4 * cm]))

    # 11. Dependencies
    story.append(Paragraph("11. Dependencies & Execution", styles["Section"]))
    story.append(Paragraph("<b>Python libraries required:</b>", styles["SubSection"]))
    deps = [
        ["Library", "Purpose"],
        ["pandas", "Data loading, transformation, export"],
        ["numpy", "Numeric operations and null handling"],
        ["geopy", "ArcGIS district geocoding"],
        ["selenium", "Browser automation for Google Maps"],
        ["webdriver_manager", "ChromeDriver management"],
        ["tqdm", "Progress tracking during geocoding"],
    ]
    story.append(table(deps, col_widths=[4.5 * cm, 12.3 * cm]))

    story.append(Paragraph("<b>Execution order in pipeline.ipynb:</b>", styles["SubSection"]))
    exec_steps = [
        "Cell 0–2: Import libraries and load Abu_Dhabi.csv",
        "Cell 4: Remove duplicate rows",
        "Cell 7, 9: Define PROPERTY_TYPE_MAP and PROJECT_TYPE_MAP",
        "Cell 11: Define transform_abudhabi_transactions()",
        "Cell 12–16: Transform, select DB1 columns, export Abu_Dhabi_Cleaned.csv",
        "Cell 19: District geocoding (ArcGIS)",
        "Cell 21: Project geocoding (Google Maps / Selenium)",
        "Cell 23: Non-RERA index assignment",
    ]
    for step in exec_steps:
        story.append(Paragraph(f"• {step}", styles["DocBullet"]))

    # 12. Limitations
    story.append(PageBreak())
    story.append(Paragraph("12. Data Quality & Limitations", styles["Section"]))
    limits = [
        ["Area", "Limitation", "Impact"],
        [
            "Furnishing status",
            "No furnishing field in source; Sale Application Type used as proxy",
            "Values are off-plan/ready, not furnished/unfurnished",
        ],
        [
            "Unit configuration",
            "Non-bedroom layouts (unclassified, commercial sizes) not mapped to numeric codes",
            "Some unit_configuration values remain as raw text",
        ],
        [
            "DB1 null fields",
            "35+ DB1 fields unavailable in source (buyer, seller, stamp duty, etc.)",
            "Populated as null in output",
        ],
        [
            "Geocoding",
            "API-based; subject to rate limits and address ambiguity",
            "5 projects failed geocoding; some coordinates may be approximate",
        ],
        [
            "transaction_type",
            "Listed in DB1_COLUMNS but not explicitly mapped in transform",
            "Column exported as null unless added manually",
        ],
    ]
    story.append(table(limits, col_widths=[3.5 * cm, 6.5 * cm, 6.8 * cm]))

    # 13. Appendix
    story.append(Paragraph("13. Appendix: DB1 Output Schema (63 Columns)", styles["Section"]))
    db1_cols = [
        "project_name", "location_name", "registered_document_village_name", "agreement_price",
        "transaction_date", "net_carpet_area_sq_m", "sub_registrar_office_code",
        "sub_registrar_office_name", "document_number", "transaction_type", "city_name",
        "guideline_value", "property_description", "floor_number", "unit_number",
        "property_type_raw", "property_type", "year", "balcony_sq_m", "terrace_sq_m",
        "seller_name", "buyer_name", "transaction_category", "internal_document_number",
        "micr_number", "bank_type", "party_code", "date_of_agreement_execution",
        "stamp_duty_paid", "registration_fee", "project_latitude", "project_longitude",
        "location_latitude", "location_longitude", "quarter", "unit_configuration",
        "property_layout", "buyer_pincode", "buyer_locality", "buyer_district", "buyer_state",
        "is_llm_processed", "is_manual_processed", "tower_name", "is_duplicate", "sale_type",
        "project_type", "country_name", "state_name", "micro_market", "sub_locality", "pincode",
        "parking_count", "facing_direction", "view_type", "furnishing_status", "condition_status",
        "source_accessibility", "source_accessibility_way", "sourcing_cost", "sourcing_time",
        "data_type", "data_source",
    ]
    col_rows = [["#", "Column Name"]]
    for i, col in enumerate(db1_cols, 1):
        col_rows.append([str(i), col])
    story.append(table(col_rows, col_widths=[1.2 * cm, 15.6 * cm]))

    story.append(Spacer(1, 1 * cm))
    story.append(
        Paragraph(
            "<i>End of Document — Abu Dhabi Real Estate Data Processing Pipeline</i>",
            ParagraphStyle(
                name="FooterNote",
                parent=styles["Body"],
                alignment=TA_CENTER,
                fontSize=9,
                textColor=colors.grey,
            ),
        )
    )

    return story


def main():
    styles = build_styles()
    doc = BaseDocTemplate(
        str(OUTPUT_PDF),
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2.2 * cm,
        bottomMargin=2 * cm,
        title="Abu Dhabi Pipeline Documentation",
        author="Data Engineering Team",
    )

    frame = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        doc.width,
        doc.height - 0.4 * cm,
        id="normal",
    )

    doc.addPageTemplates(
        [
            PageTemplate(id="cover", frames=[frame], onPage=_cover),
            PageTemplate(id="content", frames=[frame], onPage=_header_footer),
        ]
    )

    story = build_story(styles)
    story.insert(0, NextPageTemplate("content"))
    doc.build(story)
    print(f"Documentation saved to: {OUTPUT_PDF}")


if __name__ == "__main__":
    main()
