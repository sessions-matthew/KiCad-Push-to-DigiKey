import os
import sys
import pathlib
import xml.etree.ElementTree as ET

def get_sch_file_name(p: str):
    # `p`: pcb_path = board.GetFileName()
    _pcb_file_name = pathlib.Path(p).resolve().name
    if _pcb_file_name.lower().endswith('.kicad_pcb')\
            or _pcb_file_name.lower().endswith('.kicad_sch'):
        _sch_file_name = _pcb_file_name[:-10]
    else:
        _sch_file_name = 'From KiCad plugin'
    if len(_sch_file_name.strip()) == 0:
        return 'From KiCad plugin'
    return _sch_file_name


def pcb_2_sch_path(p: str):
    _parent_path = pathlib.Path(p).resolve().parent
    _pcb_file_name = pathlib.Path(p).resolve().name
    if _pcb_file_name.lower().endswith('.kicad_pcb'):
        _sch_file_name = _pcb_file_name[:-10] + '.kicad_sch'
    else:
        _sch_file_name = _pcb_file_name + '.kicad_sch'
    return _parent_path.joinpath(_sch_file_name)

def json_from_bom__with_pn_as_key(bom):
    # comply with: /mylists/api/thirdparty
    json_object = []
    for _pn in bom:
        _item = bom[_pn]
        json_object.append({
            "requestedPartNumber": _pn,
            "quantities": [
                {
                    "quantity": _item.get('qty', 0)
                }
            ],
            "customerReference": _item.get('cusRef', ''),
            "notes": _item.get('note', ''),
        })
    return json_object

def parse_bom(file_path):
    """
    Parse the bom.xml file and extract component information.

    Args:
        file_path (str): Path to the bom.xml file.

    Returns:
        list[dict]: A list of components with their details.
    """
    tree = ET.parse(file_path)
    root = tree.getroot()

    parts = {}
    for comp in root.findall(".//comp"):
        ref = comp.get("ref", "Unknown")
        value = comp.findtext("value", "Unknown")
        footprint = comp.findtext("footprint", "Unknown")
        datasheet = comp.findtext("datasheet", "Unknown")
        fields = comp.findall(".//field")
        partno = None
        for field in fields:
            if field.get("name") == "Partno":
                partno = field.text
                break
        
        if partno is None:
            print(f"Warning: No Partno found for component {ref}. Skipping.")
            continue
        if partno not in parts:
            parts[partno] = {
                "References": ref,
                "Value": value,
                "Footprint": footprint,
                "Datasheet": datasheet,
                "Quantity": 1,
            }
        else:
            parts[partno]["Quantity"] += 1
            parts[partno]["References"] += f", {ref}"

    return parts

def get_symbol_dict(kicad_sch_path):
    destxml = pathlib.Path(kicad_sch_path).resolve().parent.joinpath('bom.xml')
    if sys.platform == 'darwin':
        os.system("/Applications/KiCad/KiCad.app/Contents/MacOS/kicad-cli sch export python-bom -o " + str(destxml))
    else:
        raise Exception("Unsupported platform")
    return parse_bom(destxml)
