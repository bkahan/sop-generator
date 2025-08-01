#!/usr/bin/env python3
"""
Analyze how docxtpl reads and parses Word document templates
This script examines a Word document using ONLY docxtpl and shows all template variables
"""

from docxtpl import DocxTemplate
from pathlib import Path
import re
from typing import Dict, List, Any
from io import BytesIO


def analyze_docxtpl_template(template_path: str):
    """
    Analyze a Word template using only docxtpl

    Args:
        template_path: Path to the Word document template
    """

    print("=" * 80)
    print(f"DOCXTPL TEMPLATE ANALYSIS: {template_path}")
    print("=" * 80)

    # Check if file exists
    if not Path(template_path).exists():
        print(f"❌ File not found: {template_path}")
        return

    try:
        # Load with DocxTemplate
        template = DocxTemplate(template_path)
        print("✅ Successfully loaded template with DocxTemplate")

    except Exception as e:
        print(f"❌ Error loading template: {e}")
        return

    print("\n" + "=" * 60)
    print("1. EXTRACT TEXT FROM DOCXTPL")
    print("=" * 60)

    # Extract all text using docxtpl's internal document
    all_text = extract_text_from_docxtpl(template)
    print(f"📄 Total characters extracted: {len(all_text)}")

    # Show preview of text
    if all_text:
        preview = all_text.replace('\n', ' ').strip()[:200]
        print(f"📝 Text preview: {preview}{'...' if len(all_text) > 200 else ''}")

    print("\n" + "=" * 60)
    print("2. TEMPLATE VARIABLES FOUND")
    print("=" * 60)

    # Extract all Jinja2 template variables using regex
    variables = find_jinja_variables(all_text)

    if variables:
        total_constructs = sum(len(var_list) for var_list in variables.values())
        print(f"🎯 Found {total_constructs} template constructs:")

        for var_type, var_list in variables.items():
            if var_list:
                unique_vars = sorted(set(var_list))
                print(f"\n  {var_type.upper()} ({len(unique_vars)}):")
                for var in unique_vars:
                    print(f"    • {var}")
    else:
        print("❌ No Jinja2 template variables found")

    print("\n" + "=" * 60)
    print("3. CREATE TEST CONTEXT")
    print("=" * 60)

    # Create test context based on found variables
    test_context = create_test_context(variables)

    if test_context:
        print(f"🧪 Created test context with {len(test_context)} variables:")
        for key, value in test_context.items():
            print(f"    • {key}: '{value}'")
    else:
        print("⚠️  No variables found to create test context")

    print("\n" + "=" * 60)
    print("4. TEST TEMPLATE RENDERING")
    print("=" * 60)

    if test_context:
        try:
            # Test rendering with dummy data
            test_template = DocxTemplate(template_path)
            test_template.render(test_context)

            # Save to memory to test if it works
            output_stream = BytesIO()
            test_template.save(output_stream)

            output_size = len(output_stream.getvalue())
            print("✅ Template rendering successful!")
            print(f"   Generated document size: {output_size:,} bytes")

            # Optionally save test output
            test_output_path = template_path.replace('.docx', '_TEST_OUTPUT.docx')
            test_template.save(test_output_path)
            print(f"💾 Test output saved: {test_output_path}")

        except Exception as e:
            print(f"❌ Template rendering failed: {e}")
            print("   This might indicate:")
            print("   - Missing variables in test context")
            print("   - Syntax errors in template")
            print("   - Invalid Jinja2 constructs")

            # Try to identify the problematic variable
            error_str = str(e)
            if "undefined" in error_str.lower():
                print(f"   - Undefined variable error: {error_str}")
    else:
        print("⚠️  Skipping render test - no variables found")

    print("\n" + "=" * 60)
    print("5. VARIABLE ANALYSIS")
    print("=" * 60)

    # Analyze variable patterns
    simple_vars = variables.get('variables', [])
    if simple_vars:
        print("🔍 Variable pattern analysis:")

        # Group by common patterns
        patterns = {
            'dates': [v for v in simple_vars if any(x in v.lower() for x in ['date', 'time'])],
            'names': [v for v in simple_vars if any(x in v.lower() for x in ['name', 'user', 'author'])],
            'ids': [v for v in simple_vars if any(x in v.lower() for x in ['id', 'number', 'version'])],
            'content': [v for v in simple_vars if
                        any(x in v.lower() for x in ['objective', 'scope', 'title', 'description'])]
        }

        for pattern_type, pattern_vars in patterns.items():
            if pattern_vars:
                print(f"   • {pattern_type.title()}: {pattern_vars}")

        # Show variables that don't match common patterns
        all_pattern_vars = set()
        for pattern_vars in patterns.values():
            all_pattern_vars.update(pattern_vars)

        other_vars = [v for v in simple_vars if v not in all_pattern_vars]
        if other_vars:
            print(f"   • Other: {other_vars}")

    return {
        'variables': variables,
        'test_context': test_context,
        'text_length': len(all_text)
    }


def extract_text_from_docxtpl(template: DocxTemplate) -> str:
    """
    Extract all text from a DocxTemplate using its internal document
    """
    try:
        # Access the internal python-docx document object
        doc = template.docx
        all_text = []

        # Extract from paragraphs
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                all_text.append(paragraph.text)

        # Extract from tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        all_text.append(cell.text)

        return '\n'.join(all_text)

    except Exception as e:
        print(f"⚠️  Error extracting text: {e}")
        return ""


def find_jinja_variables(text: str) -> Dict[str, List[str]]:
    """
    Find all Jinja2 template variables in text
    Returns categorized variables by type
    """

    # Regular expressions for different Jinja2 constructs
    patterns = {
        'variables': r'\{\{\s*([^}]+)\s*\}\}',  # {{ variable }}
        'blocks': r'\{\%\s*([^%]+)\s*\%\}',  # {% block %}
        'comments': r'\{\#\s*([^#]+)\s*\#\}',  # {# comment #}
    }

    results = {}

    for pattern_type, pattern in patterns.items():
        matches = re.findall(pattern, text)

        # Clean up the matches
        cleaned_matches = []
        for match in matches:
            # Remove filters and get just the variable name
            if pattern_type == 'variables':
                # Split on | to separate variable from filters
                var_name = match.split('|')[0].strip()
                # Split on . to get base variable name
                base_var = var_name.split('.')[0].strip()
                cleaned_matches.append(base_var)
            else:
                cleaned_matches.append(match.strip())

        results[pattern_type] = cleaned_matches

    return results


def create_test_context(variables: Dict[str, List[str]]) -> Dict[str, Any]:
    """Create dummy context data for testing template rendering"""

    context = {}

    # Extract simple variables (not blocks or comments)
    simple_vars = variables.get('variables', [])

    for var in simple_vars:
        # Clean up variable name
        clean_var = var.strip()

        # Skip if already added
        if clean_var in context:
            continue

        # Generate appropriate dummy data based on variable name
        if any(keyword in clean_var.lower() for keyword in ['date', 'time']):
            context[clean_var] = '2024-12-15'
        elif any(keyword in clean_var.lower() for keyword in ['name', 'user', 'author']):
            context[clean_var] = 'Test User'
        elif any(keyword in clean_var.lower() for keyword in ['id', 'number', 'version']):
            context[clean_var] = 'TEST-001'
        elif any(keyword in clean_var.lower() for keyword in ['title', 'subject']):
            context[clean_var] = 'Test Document Title'
        elif any(keyword in clean_var.lower() for keyword in ['objective']):
            context[clean_var] = 'This is a test objective for the procedure.'
        elif any(keyword in clean_var.lower() for keyword in ['scope']):
            context[clean_var] = 'This SOP applies to all test operations.'
        elif any(keyword in clean_var.lower() for keyword in ['responsibilities']):
            context[clean_var] = 'Technician: Execute procedure\nEngineer: Review results'
        elif any(keyword in clean_var.lower() for keyword in ['definitions']):
            context[clean_var] = 'N/A'
        else:
            context[clean_var] = f'Test value for {clean_var}'

    return context


def analyze_template_file(file_path: str):
    """Main analysis function"""

    if not Path(file_path).exists():
        print(f"❌ File not found: {file_path}")
        print("Please provide a valid Word document path.")
        return None

    # Run the analysis
    results = analyze_docxtpl_template(file_path)

    # Summary
    print("\n" + "=" * 80)
    print("ANALYSIS SUMMARY")
    print("=" * 80)

    if results:
        total_vars = sum(len(vars_list) for vars_list in results['variables'].values())
        simple_vars = len(results['variables'].get('variables', []))

        print(f"📊 Total template constructs: {total_vars}")
        print(f"🎯 Simple variables: {simple_vars}")
        print(f"📄 Text extracted: {results['text_length']:,} characters")

        if results['test_context']:
            print(f"🧪 Test context: {len(results['test_context'])} variables")
            print("\n🔧 Ready for docxtpl rendering with:")
            print("   template = DocxTemplate('your_file.docx')")
            print("   template.render(context)")
            print("   template.save('output.docx')")
        else:
            print("⚠️  No variables found - this may not be a template")

    return results


if __name__ == "__main__":
    print("DocxTemplate Variable Analyzer (Pure docxtpl)")
    print("=" * 50)

    # Get file path from user or command line
    import sys

    if len(sys.argv) > 1:
        template_file = sys.argv[1]
    else:
        template_file = input("Enter path to Word template: ").strip()
        if not template_file:
            print("❌ No file path provided")
            sys.exit(1)

    # Run the analysis
    try:
        results = analyze_template_file(template_file)
        if results:
            print(f"\n✅ Analysis complete for: {template_file}")

    except Exception as e:
        print(f"❌ Analysis failed: {e}")
        import traceback

        traceback.print_exc()