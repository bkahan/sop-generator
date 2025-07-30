#!/usr/bin/env python3
"""
Create a sample Word template matching the LS SOP format
This creates a template with all the required placeholders
"""

from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pathlib import Path


def create_sop_template():
    """Create a Word template with the LS SOP structure"""

    # Create a new document
    doc = Document()

    # Set default font for the document
    style = doc.styles['Normal']
    style.font.name = 'Verdana'
    style.font.size = Pt(10)

    # Add special notes section (will be removed in final)
    special_notes = doc.add_paragraph()
    special_notes.add_run('Special Notes:').bold = True
    doc.add_paragraph('N/A')
    doc.add_paragraph('-' * 70)

    # Add instruction to remove highlighted info
    instruction = doc.add_paragraph()
    run1 = instruction.add_run('Remove all highlighted information prior to submission')
    run1.font.highlight_color = 7  # Yellow highlight

    # Add formatting instructions
    format_inst1 = doc.add_paragraph()
    run2 = format_inst1.add_run('Heading1, font Verdana, 12 bold')
    run2.font.highlight_color = 7

    format_inst2 = doc.add_paragraph()
    run3 = format_inst2.add_run('Body of doc, font Verdana 10')
    run3.font.highlight_color = 7

    # 1. OBJECTIVE
    heading1 = doc.add_paragraph('1.  OBJECTIVE')
    heading1.runs[0].font.bold = True
    heading1.runs[0].font.size = Pt(12)
    doc.add_paragraph('{{objective}}')
    doc.add_paragraph()

    # 2. SCOPE
    heading2 = doc.add_paragraph('2.  SCOPE')
    heading2.runs[0].font.bold = True
    heading2.runs[0].font.size = Pt(12)
    doc.add_paragraph('{{scope}}')
    doc.add_paragraph()

    # 3. RESPONSIBILITIES
    heading3 = doc.add_paragraph('3.  RESPONSIBILITIES')
    heading3.runs[0].font.bold = True
    heading3.runs[0].font.size = Pt(12)
    doc.add_paragraph('{{responsibilities}}')
    doc.add_paragraph()

    # 4. DEFINITIONS
    heading4 = doc.add_paragraph('4.  DEFINITIONS')
    heading4.runs[0].font.bold = True
    heading4.runs[0].font.size = Pt(12)
    doc.add_paragraph('{{definitions}}')
    doc.add_paragraph()

    # 5. PROCEDURE
    heading5 = doc.add_paragraph('5.  PROCEDURE')
    heading5.runs[0].font.bold = True
    heading5.runs[0].font.size = Pt(12)
    doc.add_paragraph('See Mango File {{mango_pptx_id}} -- {{pptx_title}}.')

    # Add page break before revision sheet
    doc.add_page_break()

    # REVISION SHEET
    rev_heading = doc.add_paragraph()
    rev_heading.add_run('REVISION SHEET').bold = True
    rev_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Create revision table
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'

    # Set header row
    header_cells = table.rows[0].cells
    header_cells[0].text = 'Date'
    header_cells[1].text = 'Version'
    header_cells[2].text = 'Nature of Changes'
    header_cells[3].text = 'Changed By'

    # Make header bold
    for cell in header_cells:
        cell.paragraphs[0].runs[0].font.bold = True

    # Add template row for dynamic content
    # This will be replaced by docxtpl with actual revision data
    row_cells = table.add_row().cells
    row_cells[0].text = '{{date}}'
    row_cells[1].text = '1.0'
    row_cells[2].text = 'Initial Release'
    row_cells[3].text = '{{user_name}}'

    # Add empty rows for future revisions
    for _ in range(2):
        table.add_row()

    # Save as template
    template_path = Path('templates/LS_SOP_Template.docx')
    template_path.parent.mkdir(exist_ok=True)
    doc.save(template_path)

    print(f"✅ Created template: {template_path}")

    # Now convert to docxtpl format for dynamic content
    tpl = DocxTemplate(template_path)

    # Test render with sample data
    test_context = {
        'objective': 'This is a test objective.',
        'scope': 'This SOP applies to all Cassette Manufacturing operations at the US Stoneham facility',
        'responsibilities': 'Technician: Performs the procedure; Engineer: Reviews results',
        'definitions': 'N/A',
        'mango_pptx_id': 'MANGO-2024-001',
        'pptx_title': 'Test Procedure',
        'date': '12/15/2024',
        'user_name': 'John Doe'
    }

    # Save template with markers
    tpl.save(template_path)

    # Create a test output to verify
    test_output = Path('templates/LS_SOP_Template_TEST.docx')
    tpl.render(test_context)
    tpl.save(test_output)

    print(f"✅ Created test output: {test_output}")
    print("\nTemplate placeholders:")
    print("- {{objective}}")
    print("- {{scope}}")
    print("- {{responsibilities}}")
    print("- {{definitions}}")
    print("- {{mango_pptx_id}}")
    print("- {{pptx_title}}")
    print("- {{date}}")
    print("- {{user_name}}")


def create_template_with_dynamic_table():
    """Create a more advanced template with dynamic revision table"""

    doc = Document()

    # Set default font
    style = doc.styles['Normal']
    style.font.name = 'Verdana'
    style.font.size = Pt(10)

    # Main content sections
    sections = [
        ('1.  OBJECTIVE', '{{objective}}'),
        ('2.  SCOPE', '{{scope}}'),
        ('3.  RESPONSIBILITIES', '{{responsibilities}}'),
        ('4.  DEFINITIONS', '{{definitions}}'),
        ('5.  PROCEDURE', 'See Mango File {{mango_pptx_id}} -- {{pptx_title}}.')
    ]

    for heading_text, content in sections:
        heading = doc.add_paragraph(heading_text)
        heading.runs[0].font.bold = True
        heading.runs[0].font.size = Pt(12)
        doc.add_paragraph(content)
        doc.add_paragraph()  # Empty line

    # Page break
    doc.add_page_break()

    # Revision sheet with dynamic table
    rev_heading = doc.add_paragraph()
    rev_heading.add_run('REVISION SHEET').bold = True
    rev_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Add docxtpl tags for dynamic table
    p = doc.add_paragraph()
    p.add_run('{%tr for revision in revisions %}')

    # Create table structure
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'

    # Header
    header = table.rows[0].cells
    header[0].text = 'Date'
    header[1].text = 'Version'
    header[2].text = 'Nature of Changes'
    header[3].text = 'Changed By'

    for cell in header:
        cell.paragraphs[0].runs[0].font.bold = True

    # Dynamic row
    row = table.add_row().cells
    row[0].text = '{{revision.date}}'
    row[1].text = '{{revision.version}}'
    row[2].text = '{{revision.nature_of_changes}}'
    row[3].text = '{{revision.changed_by}}'

    # End tag
    p2 = doc.add_paragraph()
    p2.add_run('{%tr endfor %}')

    # Save
    advanced_path = Path('templates/LS_SOP_Template_Advanced.docx')
    doc.save(advanced_path)
    print(f"✅ Created advanced template: {advanced_path}")


if __name__ == "__main__":
    print("Creating LS SOP Word Template...")
    print("=" * 50)

    try:
        # Create basic template
        create_sop_template()

        # Create advanced template with dynamic table
        create_template_with_dynamic_table()

        print("\n✅ Templates created successfully!")
        print("\nNext steps:")
        print("1. Review the generated templates in the templates/ folder")
        print("2. Modify formatting as needed in Microsoft Word")
        print("3. Use 'LS_SOP_Template.docx' with the main application")

    except Exception as e:
        print(f"❌ Error creating template: {e}")
        print("\nMake sure you have docxtpl installed:")
        print("pip install python-docx docxtpl")