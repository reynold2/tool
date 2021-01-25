from WordTool.DataManger.RulesHandler import *
from WordTool.Common.GlobalMethods import *


class AdvancedManger(object):
    index = 0

    @staticmethod
    def location_val_deal(location_val):
        """生成需要处理的行，依据配置信息，转换为统一的需求列"""
        location_range = []
        if len(location_val) == 3 and location_val:
            location1 = location_val[0]
            location2 = location_val[1]
            location3 = location_val[2]
            if location1 > 0:
                location_range.append(location1)
            elif "," in location3:
                location_range.clear()
                for data in location3.split(","):
                    location_range.append(int(data))
            elif location2:
                location_range = [x for x in range(1, 20)]

        return location_range

    @staticmethod
    def advanced_data_deal(data_list):
        """处理Excel表格数据，对展示或存储数据进行二次结果处理"""
        AdvancedManger.spinbox_index_clear()

        adv_config = RULES_HANDLER_INSTANCE.get_advanced_settings()
        location_val = adv_config.get("location")
        repeat_val = adv_config.get("repeat")
        slice_val = adv_config.get("slice")
        combination_val = adv_config.get("combination")
        eval_val = adv_config.get("eval")
        conversion_val = adv_config.get("conversion")
        """依据配置将位置参数转换为list集合，用于数据处理位置匹配"""
        location_val = AdvancedManger.location_val_deal(location_val)

        all_data_val = []
        repeat_data_val = []
        data = []
        for row, clo, val in data_list:
            if clo + 1 in location_val:
                """判断处理模式支持手动处理和函数处理两种模式"""
                if eval_val[0]:
                    val = AdvancedManger.eval_val_cal(val, eval_val[1])
                else:
                    val = AdvancedManger.configuration_val_cal(val, combination_val, slice_val, False, conversion_val)

                """将全部数据进行重复判断并数据位置记录"""
                if repeat_val != EnumRepeat.No_Processing:
                    if val in all_data_val:
                        repeat_data_val.append([row, clo, val])

            data.append([row, clo, val])
            all_data_val.append(val)
        """依据重复模式条件将多余的重复记录数据清除"""
        if repeat_val == EnumRepeat.One_Processing:
            for repeat_data in repeat_data_val:
                if repeat_data[1] in location_val is False:
                    repeat_data_val.remove(repeat_data)
        if repeat_data_val:
            """在源数据中依据重复数据将源数据进行整理"""
            data = AdvancedManger.repeat_data_deal(data, repeat_data_val)
        return data

    @staticmethod
    def repeat_data_deal(data_list, repeat_data):
        """在源数据中依据重复数据将源数据进行整理并将新数据返回"""
        new_data_list = []
        remove_row = []
        for repeat_data_row in repeat_data:
            remove_row.append(repeat_data_row[0])
        row = 0
        for data in data_list:

            if data[0] in remove_row:
                data_list.remove(data)
                continue

            if data[0] != row:
                row = +1
            new_data_list.append([row, data[1], data[2]])

        return new_data_list

    @staticmethod
    def eval_val_cal(val, eval_text):
        """结果函数模式处理"""
        try:
            if "val" in eval_text and eval_text is not None:
                return eval(eval_text)
        except:
            return val

    @staticmethod
    def configuration_val_cal(data, combination, slice_val, is_demo=False,
                              conversion_type=EnumConversion.NO_processing):
        """结果配置模式处理"""
        per_data = ""
        data = AdvancedManger.configuration_val_cal_conversion(data, conversion_type)
        context = AdvancedManger.configuration_val_cal_slice(data, slice_val)
        if combination:
            for type_val, val in combination:
                if type_val == EnumWidget.Label:
                    per_data = per_data + "".join(context)
                elif type_val == EnumWidget.SpinBox:
                    if is_demo is False:
                        val = AdvancedManger.index + val
                    per_data = per_data + "".join(str(val))
                elif type_val == EnumWidget.LineEdit:
                    per_data = per_data + "".join(val)
        if is_demo is False:
            AdvancedManger.index = AdvancedManger.index + 1
        return per_data

    @staticmethod
    def configuration_val_cal_slice(data, slice_val):
        """进行数据截取"""
        context = data[slice_val[0]:slice_val[1]]
        return context

    @staticmethod
    def configuration_val_cal_conversion(data, conversion_type=EnumConversion.NO_processing):
        p = xpinyin.Pinyin()
        if conversion_type == EnumConversion.PINYIN_Processing:
            data = p.get_pinyin(data, u'')
        elif conversion_type == EnumConversion.ENGLISH_Processing:

            data = Translator(from_lang='chinese', to_lang='english').translate(data)

            data = data.replace("\'m", " am")

        elif conversion_type == EnumConversion.PINYIN_1_Processing:

            data = p.get_initials(data, u'')
        elif conversion_type == EnumConversion.ENGLISH_1_Processing:
            data = Translator(from_lang='chinese', to_lang='english').translate(data).upper()
        return data

    @staticmethod
    def spinbox_index_clear():
        """重置combination中SpinBox的序号"""
        AdvancedManger.index = 0


if __name__ == '__main__':
    pass
