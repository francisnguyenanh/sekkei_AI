from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from schemas.models import StyleDef

def _parse_composite_border(border_str: str) -> Border:
    thin = Side(border_style="thin", color="000000")
    medium = Side(border_style="medium", color="000000")
    none_side = Side()
    style_map = {"thin": thin, "medium": medium, "none": none_side}
    sides = {"L": none_side, "R": none_side, "T": none_side, "B": none_side}
    for part in border_str.split(","):
        part = part.strip()
        if ":" in part:
            key, val = part.split(":", 1)
            sides[key.strip().upper()] = style_map.get(val.strip(), none_side)
    return Border(left=sides["L"], right=sides["R"], top=sides["T"], bottom=sides["B"])

def get_border(border_type: str) -> Border:
    if not border_type or border_type == "none":
        return Border()
        
    if ":" in border_type:
        return _parse_composite_border(border_type)
        
    thin = Side(border_style="thin", color="000000")
    medium = Side(border_style="medium", color="000000")
    
    borders = {
        "all_thin": Border(left=thin, right=thin, top=thin, bottom=thin),
        "outer_thin": Border(left=thin, right=thin, top=thin, bottom=thin),
        "all_medium": Border(left=medium, right=medium, top=medium, bottom=medium),
        "top_thin": Border(top=thin),
        "bottom_thin": Border(bottom=thin)
    }
    
    return borders.get(border_type, Border())

def create_named_style(name: str, config: StyleDef) -> dict:
    kwargs = {}
    if config.fill:
        kwargs['fill'] = PatternFill(start_color=config.fill, end_color=config.fill, fill_type="solid")
        
    if config.font_color or config.font_size or config.bold or config.font_name:
        font_name = config.font_name if config.font_name else "Meiryo UI"
        kwargs['font'] = Font(
            color=config.font_color, 
            size=config.font_size, 
            bold=config.bold,
            name=font_name
        )
        
    if config.align_h or config.align_v or config.wrap_text is not None:
        kwargs['alignment'] = Alignment(
            horizontal=config.align_h, 
            vertical=config.align_v,
            wrap_text=config.wrap_text
        )
        
    if config.border:
        kwargs['border'] = get_border(config.border)
        
    return kwargs
