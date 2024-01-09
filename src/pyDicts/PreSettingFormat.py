from enum import Enum

from openpyxl.styles import Font, PatternFill, Alignment


class PreSettingFormat(Enum):
    DEFAULT = {
        "font": Font(name='Arial Nova Cond',size=11, color="CCCCCC"),
        "fill": PatternFill(fill_type='solid', start_color='0000CCFF', end_color="0000CCFF"),
        "alignment": Alignment(horizontal='center', vertical='center')
    }



