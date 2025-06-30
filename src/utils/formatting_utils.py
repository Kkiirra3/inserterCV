from copy import deepcopy as python_deepcopy
from lxml import etree
from docx.shared import Pt

class FormattingUtils:
    def __init__(self):
        self.nsmap = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
        }

    def deepcopy(self, element):
        """
        Creates a deep copy of an XML element while preserving namespaces
        """
        return python_deepcopy(element)

    def convert_to_twips(self, value):
        """
        Converts value to twips.
        Handles various measurement formats.
        """
        if value is None:
            return None
        
        try:
            # If value is already in twips (integer)
            if isinstance(value, int):
                return value
            
            # If value is a Pt, Inches, Twips etc. object
            if hasattr(value, 'twips'):
                return value.twips
            
            # If value is a string with number
            if isinstance(value, str):
                # Remove all non-numeric characters except dot and minus
                clean_value = ''.join(c for c in value if c.isdigit() or c in '.-')
                # Convert to points (1 point = 20 twips)
                points = float(clean_value)
                return int(points * 20)
            
            # If value is float
            if isinstance(value, float):
                return int(value * 20)  # Assume value is in points
            
            return None
        except (ValueError, TypeError):
            return None

    def safe_get_attribute(self, obj, attr_name, default=None):
        """
        Safely gets object's attribute value.
        """
        try:
            value = getattr(obj, attr_name, default)
            if value is None:
                return default
            return value
        except Exception:
            return default

    def copy_run_formatting(self, source_run, target_run):
        """
        Copies all formatting from one run to another
        """
        if source_run._element.rPr is not None:
            if target_run._element.rPr is None:
                target_run._element.get_or_add_rPr()
            # Copy all formatting attributes
            for attr in source_run._element.rPr.attrib:
                target_run._element.rPr.set(attr, source_run._element.rPr.get(attr))
            # Copy all child elements (color, size, font etc.)
            for child in source_run._element.rPr:
                target_run._element.rPr.append(python_deepcopy(child))

    def copy_paragraph_formatting(self, source_para, target_para):
        """
        Copies all formatting from one paragraph to another
        """
        if source_para._element.pPr is not None:
            if target_para._element.pPr is None:
                target_para._element.get_or_add_pPr()
            # Copy all paragraph formatting attributes
            for attr in source_para._element.pPr.attrib:
                target_para._element.pPr.set(attr, source_para._element.pPr.get(attr))
            # Copy all child elements (indents, alignment, numbering etc.)
            for child in source_para._element.pPr:
                target_para._element.pPr.append(python_deepcopy(child))

    def copy_table_style(self, source_table, target_table):
        """
        Copies table style from source to target
        with attribute checking and value conversion
        """
        # List of attributes to copy
        style_attrs = {
            'alignment': None,
            'style': None,
            'autofit': None
        }
        
        # Copy attributes with checking
        for attr, default in style_attrs.items():
            try:
                value = self.safe_get_attribute(source_table, attr, default)
                if value != default:
                    setattr(target_table, attr, value)
            except Exception as e:
                print(f"Warning: Could not copy table style attribute {attr}: {str(e)}")
                continue
        
        # Copy table width if set
        try:
            source_width = self.safe_get_attribute(source_table, 'width')
            if source_width is not None:
                width_twips = self.convert_to_twips(source_width)
                if width_twips is not None:
                    target_table._element.tblPr.tblW.type = 'dxa'
                    target_table._element.tblPr.tblW.w = str(width_twips)
        except Exception as e:
            print(f"Warning: Could not copy table width: {str(e)}")
        
        # Copy border style
        try:
            if hasattr(source_table, '_element') and hasattr(target_table, '_element'):
                source_borders = source_table._element.find('.//w:tblBorders')
                if source_borders is not None:
                    target_borders = target_table._element.find('.//w:tblBorders')
                    if target_borders is not None:
                        for border in source_borders:
                            target_border = target_borders.find(border.tag)
                            if target_border is not None:
                                for key, value in border.attrib.items():
                                    target_border.set(key, value)
        except Exception as e:
            print(f"Warning: Could not copy table borders: {str(e)}")

    def copy_xml_element_with_props(self, source_element, target_element, prop_tag):
        """
        Copies XML element properties from source to target
        """
        # Get or create properties element
        source_props = source_element.find(prop_tag)
        if source_props is not None:
            # Remove existing properties in target element
            target_props = target_element.find(prop_tag)
            if target_props is not None:
                target_element.remove(target_props)
            # Copy properties from source
            target_element.append(python_deepcopy(source_props))

    def has_list_properties(self, para):
        """
        Checks if paragraph is a list element
        """
        try:
            # Check for numbering properties
            ppr = para.find('.//w:pPr', namespaces=self.nsmap)
            if ppr is not None:
                num_pr = ppr.find('.//w:numPr', namespaces=self.nsmap)
                if num_pr is not None:
                    return True
            return False
        except Exception:
            return False

    def copy_list_properties(self, source_para, target_para):
        """
        Copies list properties (bullet points) from source paragraph to target
        with full copying of definitions from numbering.xml
        """
        try:
            # Get pPr element of source paragraph
            source_ppr = source_para.find('.//w:pPr', namespaces=self.nsmap)
            if source_ppr is not None:
                # Get numPr (numbering properties)
                num_pr = source_ppr.find('.//w:numPr', namespaces=self.nsmap)
                if num_pr is not None:
                    # Get or create pPr in target paragraph
                    target_ppr = target_para.find('.//w:pPr', namespaces=self.nsmap)
                    if target_ppr is None:
                        target_ppr = etree.SubElement(target_para, '{%s}pPr' % self.nsmap['w'])
                    
                    # Copy numPr with all child elements
                    new_num_pr = etree.SubElement(target_ppr, '{%s}numPr' % self.nsmap['w'])
                    
                    # Copy ilvl (list level)
                    source_ilvl = num_pr.find('.//w:ilvl', namespaces=self.nsmap)
                    if source_ilvl is not None:
                        new_ilvl = etree.SubElement(new_num_pr, '{%s}ilvl' % self.nsmap['w'])
                        new_ilvl.set('{%s}val' % self.nsmap['w'], source_ilvl.get('{%s}val' % self.nsmap['w']))
                    
                    # Copy numId (numbering ID)
                    source_numId = num_pr.find('.//w:numId', namespaces=self.nsmap)
                    if source_numId is not None:
                        new_numId = etree.SubElement(new_num_pr, '{%s}numId' % self.nsmap['w'])
                        new_numId.set('{%s}val' % self.nsmap['w'], source_numId.get('{%s}val' % self.nsmap['w']))
                    
                    # Copy indent properties
                    source_ind = source_ppr.find('.//w:ind', namespaces=self.nsmap)
                    if source_ind is not None:
                        new_ind = etree.SubElement(target_ppr, '{%s}ind' % self.nsmap['w'])
                        for key, value in source_ind.attrib.items():
                            if '}' in key:  # If attribute has namespace
                                ns, local_name = key.split('}')
                                new_ind.set('{%s}%s' % (self.nsmap['w'], local_name), value)
                            else:
                                new_ind.set(key, value)
                    
                    # Copy numbering definitions from document
                    source_doc = source_para.getroottree()
                    target_doc = target_para.getroottree()
                    
                    # Find or create numbering.xml in target document
                    target_numbering = target_doc.find('.//w:numbering', namespaces=self.nsmap)
                    if target_numbering is None:
                        target_numbering = etree.Element('{%s}numbering' % self.nsmap['w'])
                        target_doc.getroot().append(target_numbering)
                    
                    # Copy numbering definitions
                    source_numbering = source_doc.find('.//w:numbering', namespaces=self.nsmap)
                    if source_numbering is not None and source_numId is not None:
                        num_id_val = source_numId.get('{%s}val' % self.nsmap['w'])
                        if num_id_val:
                            # Copy numbering definition
                            source_num = source_numbering.find('.//w:num[@w:numId="%s"]' % num_id_val, namespaces=self.nsmap)
                            if source_num is not None:
                                # Check existence in target document
                                target_num = target_numbering.find('.//w:num[@w:numId="%s"]' % num_id_val, namespaces=self.nsmap)
                                if target_num is None:
                                    # Copy definition
                                    new_num = etree.SubElement(target_numbering, '{%s}num' % self.nsmap['w'])
                                    new_num.set('{%s}numId' % self.nsmap['w'], num_id_val)
                                    
                                    # Copy abstractNumId
                                    source_abstract_num_id = source_num.find('.//w:abstractNumId', namespaces=self.nsmap)
                                    if source_abstract_num_id is not None:
                                        abstract_num_id_val = source_abstract_num_id.get('{%s}val' % self.nsmap['w'])
                                        if abstract_num_id_val:
                                            new_abstract_num_id = etree.SubElement(new_num, '{%s}abstractNumId' % self.nsmap['w'])
                                            new_abstract_num_id.set('{%s}val' % self.nsmap['w'], abstract_num_id_val)
                                            
                                            # Copy abstract numbering definition
                                            source_abstract_num = source_numbering.find(
                                                './/w:abstractNum[@w:abstractNumId="%s"]' % abstract_num_id_val,
                                                namespaces=self.nsmap
                                            )
                                            if source_abstract_num is not None:
                                                new_abstract_num = etree.SubElement(target_numbering, '{%s}abstractNum' % self.nsmap['w'])
                                                new_abstract_num.set('{%s}abstractNumId' % self.nsmap['w'], abstract_num_id_val)
                                                
                                                # Copy all levels and their properties
                                                for lvl in source_abstract_num.findall('.//w:lvl', namespaces=self.nsmap):
                                                    new_lvl = etree.SubElement(new_abstract_num, '{%s}lvl' % self.nsmap['w'])
                                                    for key, value in lvl.attrib.items():
                                                        if '}' in key:
                                                            ns, local_name = key.split('}')
                                                            new_lvl.set('{%s}%s' % (self.nsmap['w'], local_name), value)
                                                        else:
                                                            new_lvl.set(key, value)
                                                    
                                                    # Copy all level child elements
                                                    for child in lvl:
                                                        new_child = etree.SubElement(new_lvl, child.tag)
                                                        for key, value in child.attrib.items():
                                                            if '}' in key:
                                                                ns, local_name = key.split('}')
                                                                new_child.set('{%s}%s' % (self.nsmap['w'], local_name), value)
                                                            else:
                                                                new_child.set(key, value)
                                                        if child.text:
                                                            new_child.text = child.text
                
        except Exception as e:
            print(f"Warning: Could not copy list properties: {str(e)}")

    def copy_paragraph_format_with_ns(self, source_para, target_para):
        """
        Copies paragraph formatting with Word namespace consideration
        """
        try:
            # Register namespaces
            for prefix, uri in self.nsmap.items():
                etree.register_namespace(prefix, uri)
            
            # Copy paragraph properties (pPr)
            source_ppr = source_para.find('.//w:pPr', namespaces=self.nsmap)
            if source_ppr is not None:
                # Remove existing properties
                target_ppr = target_para.find('.//w:pPr', namespaces=self.nsmap)
                if target_ppr is not None:
                    target_para.remove(target_ppr)
                # Copy new properties
                target_para.append(python_deepcopy(source_ppr))
                
            # Copy list properties separately
            self.copy_list_properties(source_para, target_para)
            
        except Exception as e:
            print(f"Warning: Could not copy paragraph format: {str(e)}")

    def copy_run_format_with_ns(self, source_run, target_run):
        """
        Copies run formatting with Word namespace consideration
        """
        try:
            # Register namespaces
            for prefix, uri in self.nsmap.items():
                etree.register_namespace(prefix, uri)
            
            # Copy run properties (rPr)
            source_rpr = source_run.find('.//w:rPr', namespaces=self.nsmap)
            if source_rpr is not None:
                # Remove existing properties
                target_rpr = target_run.find('.//w:rPr', namespaces=self.nsmap)
                if target_rpr is not None:
                    target_run.remove(target_rpr)
                # Copy new properties
                target_run.append(python_deepcopy(source_rpr))
        except Exception as e:
            print(f"Warning: Could not copy run format: {str(e)}")

    def apply_format_to_run(self, run, format_info):
        """
        Applies formatting to run
        """
        run.bold = format_info['bold']
        run.italic = format_info['italic']
        run.underline = format_info['underline']
        if format_info['font']:
            run.font.name = format_info['font']
        if format_info['size']:
            run.font.size = format_info['size']
        if format_info['color']:
            run.font.color.rgb = format_info['color'] 