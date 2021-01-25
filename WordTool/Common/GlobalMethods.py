import xpinyin
from WordTool.Common.GlobalVar import *
from translate import Translator


def conversion_decorator(func):
    def call_fun(data, conversion_type=EnumConversion.NO_processing):
        p = xpinyin.Pinyin()
        if conversion_type == EnumConversion.PINYIN_Processing:
            data = p.get_pinyin(data, u'')
        elif conversion_type == EnumConversion.ENGLISH_Processing:
            data = Translator(from_lang='chinese', to_lang='english').translate(data)
        elif conversion_type == EnumConversion.PINYIN_1_Processing:
            data = p.get_initials(data, u'')
        elif conversion_type == EnumConversion.ENGLISH_1_Processing:
            data = Translator(from_lang='chinese', to_lang='english').translate(data).upper()
        return func(data, conversion_type)

    return call_fun
