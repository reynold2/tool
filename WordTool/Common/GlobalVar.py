from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


class TableRules:
    rows = ""
    column = ""


class Location_Table:
    row = ""
    column = ""
    value = ""


class EnumLocation:
    "高级结果设置需要的枚举"
    One_Column = 0
    Multiple_Column = 1
    Combination_Column = 2


class EnumRepeat:
    "高级结果设置需要的枚举"
    No_Processing = 0
    One_Processing = 1
    All_Processing = 2


class EnumConversion:
    "高级结果设置需要的枚举"
    NO_processing = 0
    PINYIN_Processing = 1
    ENGLISH_Processing = 2
    PINYIN_1_Processing = 3
    ENGLISH_1_Processing = 4


class EnumWidget:
    "高级结果设置需要的枚举"
    Label = 0
    SpinBox = 1
    LineEdit = 2


class ContentStyle:
    '''
    字号‘八号’对应磅值5
    字号‘七号’对应磅值5.5
    字号‘小六’对应磅值6.5
    字号‘六号’对应磅值7.5
    字号‘小五’对应磅值9
    字号‘五号’对应磅值10.5
    字号‘小四’对应磅值12
    字号‘四号’对应磅值14
    字号‘小三’对应磅值15
    字号‘三号’对应磅值16
    字号‘小二’对应磅值18
    字号‘二号’对应磅值22
    字号‘小一’对应磅值24
    字号‘一号’对应磅值26
    字号‘小初’对应磅值36
    字号‘初号’对应磅值42
    '''
    font = u'宋体'
    size = 12
    is_bold = True
    is_italic = True
    alignment = WD_PARAGRAPH_ALIGNMENT


class EnumFont:
    "字体枚举支持宋、黑体"
    Default = 0
    Song = 1
    Black = 2


class EnumFontSize:
    "字体大小3.5,5.0,4.5,4.0"
    SizeDefault = 0
    Size3_5 = 1
    Size5_0 = 2
    Size4_5 = 3
    Size4_0 = 4


class EnumFontAlignment:
    "对齐方式：默认，居中，左边，右边，两端"
    AlignmentDefault = 0
    Alignment_center = 1
    Alignment_Left = 2
    Alignment_right = 3
    Alignment_both_end = 4


class EnumFontBold:
    "默认，加粗，去加粗"
    BoldDefault = 0
    Bold_On = 1
    Bold_off = 2
