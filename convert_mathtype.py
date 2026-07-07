"""
Convert MathType to OMML - 将Word文档中MathType公式批量转换为可编辑格式

Pipeline: MTEF binary -> MathML (via MathType COM) -> OMML (via MML2OMML.XSL)

Usage:
    python convert_mathtype.py "C:\\path\\to\\file.docx"

Requirements:
    - Windows OS
    - MathType installed (for COM interface Equation.DSMT4)
    - Microsoft Office installed (for MML2OMML.XSL)
    - Python packages: pywin32, olefile, lxml
      pip install pywin32 olefile lxml
"""
import win32com.client
import struct
import olefile
import zipfile
import tempfile
import os
import sys
import shutil
from lxml import etree
from io import BytesIO

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# === Configuration ===
XSL_PATH = r'C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL'

# Auto-detect XSL path if default doesn't exist
if not os.path.exists(XSL_PATH):
    possible_paths = [
        r'C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL',
        r'C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL',
        r'C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL',
        r'C:\Program Files\Microsoft Office\Office15\MML2OMML.XSL',
        r'C:\Program Files (x86)\Microsoft Office\Office15\MML2OMML.XSL',
    ]
    for p in possible_paths:
        if os.path.exists(p):
            XSL_PATH = p
            break

# Get input file from command line or prompt
if len(sys.argv) > 1:
    INPUT_FILE = sys.argv[1]
else:
    INPUT_FILE = input('请输入docx文件路径: ').strip().strip('"')

OUTPUT_FILE = INPUT_FILE

# === Namespaces ===
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
O = '{urn:schemas-microsoft-com:office:office}'
V = '{urn:schemas-microsoft-com:vml}'

def extract_mtef_from_ole(ole_binary_data):
    """Extract MTEF data from an OLE binary stream."""
    tmp_path = tempfile.mktemp(suffix='.bin')
    try:
        with open(tmp_path, 'wb') as f:
            f.write(ole_binary_data)
        ole = olefile.OleFileIO(tmp_path)
        if ole.exists('Equation Native'):
            eq_native = ole.openstream('Equation Native').read()
            ole.close()
            # Parse header: first 2 bytes = header size
            hdr_size = struct.unpack_from('<H', eq_native, 0)[0]
            mtef_data = eq_native[hdr_size:]
            return mtef_data
        ole.close()
        return None
    except Exception as e:
        print(f'  [WARN] OLE parse error: {e}')
        return None
    finally:
        try:
            os.unlink(tmp_path)
        except:
            pass


def mtef_to_mathml(mt_com, mtef_data):
    """Convert MTEF binary to MathML using MathType COM."""
    try:
        mt_com.SetMTEF(bytes(mtef_data))
        mathml = mt_com.GetMathML
        return mathml
    except Exception as e:
        print(f'  [WARN] MathType COM error: {e}')
        return None


def mathml_to_omml(transform, mathml_str):
    """Convert MathML string to OMML Element using XSLT."""
    # Add MathML namespace if not present
    if 'xmlns' not in mathml_str:
        mathml_str = mathml_str.replace('<math', '<math xmlns="http://www.w3.org/1998/Math/MathML"', 1)

    try:
        mathml_tree = etree.fromstring(mathml_str.encode('utf-8'))
        omml_result = transform(mathml_tree)
        return omml_result.getroot() if omml_result.getroot() is not None else None
    except Exception as e:
        print(f'  [WARN] XSLT transform error: {e}')
        return None


def add_font_properties(omml_elem):
    """Add Cambria Math font properties to all m:r elements in OMML."""
    for mr in omml_elem.iter(f'{M}r'):
        # Check if w:rPr already exists
        existing_rpr = mr.find(f'{W}rPr')
        if existing_rpr is None:
            rpr = etree.SubElement(mr, f'{W}rPr')
            mr.insert(0, rpr)  # rPr should be first child
        else:
            rpr = existing_rpr
        # Add font if not present
        rfonts = rpr.find(f'{W}rFonts')
        if rfonts is None:
            rfonts = etree.SubElement(rpr, f'{W}rFonts')
            rfonts.set(f'{W}ascii', 'Cambria Math')
            rfonts.set(f'{W}hAnsi', 'Cambria Math')


def is_display_equation(w_r_elem, parent_p):
    """Determine if an equation should be display (block) or inline.

    If the w:r containing the OLE object is essentially alone in the paragraph
    (no significant text siblings), treat it as display equation.
    """
    if parent_p is None:
        return False

    # Check all w:r siblings for text content or other equations
    has_text_siblings = False
    for child in parent_p:
        if child == w_r_elem:
            continue
        tag = child.tag
        if tag == f'{W}r':
            # Check if this run has actual text content (not just properties)
            t_elem = child.find(f'{W}t')
            obj_elem = child.find(f'{W}object')
            if (t_elem is not None and t_elem.text and t_elem.text.strip()) or obj_elem is not None:
                has_text_siblings = True
                break
        elif tag == f'{M}oMath' or tag == f'{M}oMathPara':
            has_text_siblings = True
            break

    return not has_text_siblings


def build_ole_to_file_map(zip_file, rels_xml_bytes):
    """Build mapping from relationship IDs to OLE file paths."""
    rels_root = etree.fromstring(rels_xml_bytes)
    mapping = {}
    for rel in rels_root:
        rid = rel.get('Id')
        target = rel.get('Target', '')
        if 'oleObject' in target or 'embeddings' in target:
            mapping[rid] = 'word/' + target if not target.startswith('word/') else target
    return mapping


def main():
    print(f'Input: {INPUT_FILE}')
    print(f'Output: {OUTPUT_FILE}')
    print()

    # Initialize MathType COM and XSLT transform
    print('Initializing MathType COM object...')
    mt_com = win32com.client.Dispatch('Equation.DSMT4')
    print('Loading MML2OMML.XSL transform...')
    xsl_tree = etree.parse(XSL_PATH)
    transform = etree.XSLT(xsl_tree)
    print('Ready.\n')

    # Read the docx as ZIP
    with open(INPUT_FILE, 'rb') as f:
        docx_bytes = f.read()

    zip_in = zipfile.ZipFile(BytesIO(docx_bytes), 'r')

    # Read document.xml and relationships
    doc_xml_bytes = zip_in.read('word/document.xml')
    rels_xml_bytes = zip_in.read('word/_rels/document.xml.rels')

    # Build OLE relationship map
    ole_map = build_ole_to_file_map(zip_in, rels_xml_bytes)

    # Parse document.xml preserving all namespaces
    parser = etree.XMLParser(remove_blank_text=False)
    doc_tree = etree.fromstring(doc_xml_bytes, parser)

    # Find all w:object elements with Equation.DSMT4
    # Use XPath to find o:OLEObject with ProgID containing Equation.DSMT4
    # then get their parent w:object elements
    eq_objects = []
    for ole_elem in doc_tree.iter(f'{O}OLEObject'):
        if 'Equation.DSMT4' in ole_elem.get('ProgID', ''):
            w_obj = ole_elem.getparent()
            if w_obj is not None and w_obj.tag == f'{W}object':
                eq_objects.append(w_obj)

    print(f'Found {len(eq_objects)} MathType equations to convert.\n')

    # Track which OLE files and image files we've replaced
    removed_ole_files = set()
    removed_image_rids = set()
    converted = 0
    failed = 0

    # Process each equation
    for idx, obj in enumerate(eq_objects):
        # Get relationship ID for OLE file
        ole_elem = obj.find(f'.//{O}OLEObject')
        r_id = ole_elem.get(f'{R}id')

        # Get relationship ID for image
        shape = obj.find(f'.//{V}shape')
        img_rid = None
        if shape is not None:
            imagedata = shape.find(f'.//{V}imagedata')
            if imagedata is not None:
                img_rid = imagedata.get(f'{R}id')

        # Find the OLE file in the ZIP
        ole_file_path = ole_map.get(r_id)
        if not ole_file_path:
            print(f'  [{idx+1}] SKIP: No OLE file for {r_id}')
            failed += 1
            continue

        # Read OLE data and extract MTEF
        try:
            ole_binary = zip_in.read(ole_file_path)
        except KeyError:
            # Try without word/ prefix
            alt_path = ole_file_path.replace('word/', '')
            try:
                ole_binary = zip_in.read(alt_path)
            except:
                print(f'  [{idx+1}] SKIP: Cannot read {ole_file_path}')
                failed += 1
                continue

        mtef_data = extract_mtef_from_ole(ole_binary)
        if mtef_data is None:
            print(f'  [{idx+1}] SKIP: No MTEF in {ole_file_path}')
            failed += 1
            continue

        # Convert MTEF -> MathML
        mathml = mtef_to_mathml(mt_com, mtef_data)
        if mathml is None or '<mrow></mrow>' in mathml.replace(' ', ''):
            # Empty equation - try with full eq_native data
            print(f'  [{idx+1}] SKIP: Empty MathML from {ole_file_path}')
            failed += 1
            continue

        # Convert MathML -> OMML
        omml_elem = mathml_to_omml(transform, mathml)
        if omml_elem is None:
            print(f'  [{idx+1}] SKIP: OMML conversion failed')
            failed += 1
            continue

        # Determine context: get parent w:r and parent w:p using lxml getparent()
        w_r = obj.getparent()
        if w_r is None:
            print(f'  [{idx+1}] SKIP: No parent element')
            failed += 1
            continue

        # Verify parent is a w:r element; if not, skip
        if w_r.tag != f'{W}r':
            print(f'  [{idx+1}] SKIP: Unexpected parent tag: {w_r.tag.split("}")[-1]}')
            failed += 1
            continue

        w_p = w_r.getparent()
        if w_p is None:
            print(f'  [{idx+1}] SKIP: No grandparent element')
            failed += 1
            continue

        # Decide inline vs display
        is_display = is_display_equation(w_r, w_p)

        # Prepare OMML element
        # The transform output is either m:oMathPara or m:oMath
        omml_tag = omml_elem.tag
        if is_display:
            # Use m:oMathPara > m:oMath
            if omml_tag == f'{M}oMathPara':
                final_elem = omml_elem
            else:
                # Wrap in oMathPara
                para = etree.Element(f'{M}oMathPara')
                para.append(omml_elem)
                final_elem = para
        else:
            # Use just m:oMath (inline)
            if omml_tag == f'{M}oMathPara':
                # Extract the m:oMath inside
                inner = omml_elem.find(f'{M}oMath')
                if inner is not None:
                    final_elem = inner
                else:
                    final_elem = omml_elem
            else:
                final_elem = omml_elem

        # Add font properties
        add_font_properties(final_elem)

        # Replace w:r with OMML in the parent paragraph
        if w_p is not None:
            try:
                pos = list(w_p).index(w_r)
            except ValueError:
                print(f'  [{idx+1}] SKIP: w:r not found in parent')
                failed += 1
                continue
            w_p.remove(w_r)
            w_p.insert(pos, final_elem)
        else:
            print(f'  [{idx+1}] WARN: No parent paragraph')
            failed += 1
            continue

        # Track removed files
        removed_ole_files.add(ole_file_path)
        if img_rid:
            removed_image_rids.add(img_rid)

        converted += 1
        if (converted % 20) == 0:
            print(f'  Converted {converted}/{len(eq_objects)}...')

    print(f'\nConversion complete: {converted} converted, {failed} failed.')

    # Write modified docx
    print('\nSaving modified document...')

    # Serialize modified document.xml
    new_doc_xml = etree.tostring(doc_tree, xml_declaration=True, encoding='UTF-8', standalone=True)

    # Create output ZIP
    output_buffer = BytesIO()
    with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_out:
        for item in zip_in.infolist():
            if item.filename == 'word/document.xml':
                zip_out.writestr(item, new_doc_xml)
            elif item.filename in removed_ole_files:
                # Skip removed OLE files
                continue
            else:
                zip_out.writestr(item, zip_in.read(item.filename))

    zip_in.close()

    # Write output file
    # Create backup first
    backup_path = INPUT_FILE + '.bak'
    if not os.path.exists(backup_path):
        shutil.copy2(INPUT_FILE, backup_path)
        print(f'Backup saved: {backup_path}')

    with open(OUTPUT_FILE, 'wb') as f:
        f.write(output_buffer.getvalue())

    print(f'Output saved: {OUTPUT_FILE}')
    print('Done!')


if __name__ == '__main__':
    main()

