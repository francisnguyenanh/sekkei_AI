from pydantic import BaseModel, Field, field_validator, ConfigDict, model_validator
from typing import Dict, List, Optional, Any
import re

class FontConfig(BaseModel):
    model_config = ConfigDict(extra='forbid')
    name: str = "Meiryo UI"
    size: float = 9.0

class GlobalConfig(BaseModel):
    model_config = ConfigDict(extra='forbid')
    default_font: FontConfig = Field(default_factory=FontConfig)
    column_widths: Dict[str, float] = Field(default_factory=dict)
    row_heights: Dict[int, float] = Field(default_factory=dict)

class StyleDef(BaseModel):
    model_config = ConfigDict(extra='forbid')
    fill: Optional[str] = None
    font_color: Optional[str] = None
    font_size: Optional[float] = None
    font_name: Optional[str] = None
    bold: bool = False
    align_h: Optional[str] = None
    align_v: Optional[str] = None
    wrap_text: Optional[bool] = None
    border: Optional[str] = None

    @field_validator('fill', 'font_color')
    @classmethod
    def validate_hex_color(cls, v: Optional[str]):
        if v is not None:
            if not re.match(r'^[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$', v):
                raise ValueError("Color must be a 6-character hex RGB or 8-character ARGB string")
        return v

class LayoutBlock(BaseModel):
    model_config = ConfigDict(extra='forbid')
    range: str
    style: Optional[str] = None
    merge: bool = False
    static_text: Optional[str] = None

    @field_validator('range')
    @classmethod
    def validate_range(cls, v: str):
        if not re.match(r'^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$', v):
            raise ValueError(f"Invalid Excel range: {v}")
        return v

class TemplateConfig(BaseModel):
    model_config = ConfigDict(extra='forbid')
    sheet_name: str = "Sheet1"
    global_config: GlobalConfig = Field(default_factory=GlobalConfig)
    styles: Dict[str, StyleDef] = Field(default_factory=dict)
    layout_blocks: List[LayoutBlock] = Field(default_factory=list)
    mapping_anchors: Dict[str, str] = Field(default_factory=dict)

class LogicData(BaseModel):
    model_config = ConfigDict(extra='forbid')
    single_values: Dict[str, Any] = Field(default_factory=dict)
    table_data: Dict[str, List[List[Any]]] = Field(default_factory=dict)

class GenerateRequest(BaseModel):
    model_config = ConfigDict(extra='forbid')
    template: TemplateConfig
    logic: LogicData

    @model_validator(mode='after')
    def check_size_limits(self):
        if len(self.template.layout_blocks) > 100_000:
            raise ValueError("layout_blocks exceeds maximum of 100,000 entries")
        return self

class MultiSheetGenerateRequest(BaseModel):
    model_config = ConfigDict(extra='forbid')
    sheets: List[GenerateRequest]

    @model_validator(mode='after')
    def check_sheet_count(self):
        if len(self.sheets) == 0:
            raise ValueError("sheets list must not be empty")
        if len(self.sheets) > 50:
            raise ValueError("sheets list exceeds maximum of 50 sheets per request")
        return self

def parse_request(d: dict) -> GenerateRequest:
    return GenerateRequest.model_validate(d)

def parse_multi_request(d: dict) -> MultiSheetGenerateRequest:
    return MultiSheetGenerateRequest.model_validate(d)
