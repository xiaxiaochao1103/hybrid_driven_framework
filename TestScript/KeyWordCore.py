from Util.Excel import *
from Util.Log import info
from Util.TakePic import take_pic
from Conf.ProjVar import *
from Action.Action import *
from Util.DateAndTime import *
import re


def get_test_info(excel_file_path, sheet_name):  # 把指定的sheet所有内容从excel中读到了一个列表里
    wb = ExcelUtil(excel_file_path)
    wb.set_sheet_by_name(sheet_name)
    return wb.get_sheet_all_data()


# 把联系人数据的sheet内容，存在了一个列表里，列表里面的每个元素对应测试数据的一行，且以字典方式保存
def get_test_data_from_sheet(excel_file_path, sheet_name):  # 把list的数据转换成子元素都是字典的列表
    test_data_list = get_test_info(excel_file_path, sheet_name)
    result = []
    for data in test_data_list[1:]:
        temp = {}  # 生成了一个新的空字典
        for index in range(len(test_data_list[0])):
            key = test_data_list[0][index]
            temp[key] = data[index]
        # print(temp)
        result.append(temp)
    return result  ##


# print(get_test_data_from_sheet(test_data_file_path,"联系人数据"))
# 变量test_datas_dict---》对应存储字典格式数据的列表


# 从用例sheet中，把状态为y的用例筛选出来，放到一个列表中，以字典的方式去保存
def get_test_cases(excel_file_path, sheet_name):
    test_cases = get_test_data_from_sheet(excel_file_path, sheet_name)
    result = []
    for i in test_cases:
        if i["是否执行"] != "y":
            continue
        else:
            result.append(i)
    return result


# 变量test_datas_dict会传给execute_test_case:所有的测试数据就传入到函数内容了。
# 通过遍历test_datas_dict，就可以每一次取到一行联系人的测试数据？
def execute_test_case(test_step_sheet_name, test_datas_dict):
    wb = ExcelUtil(test_data_file_path)
    print("要执行的sheet name:", test_step_sheet_name)
    if not test_step_sheet_name in wb.get_sheet_names():  # 判断了一下测试步骤sheet名字在不在
        print("设定的测试用例sheet名称不存在！", test_step_sheet_name)
        return None
    try:
        test_result = []
        test_steps_info = get_test_info(test_data_file_path, test_step_sheet_name)
        # print("所有的测试步骤：",test_steps_info)
        wb.set_sheet_by_name(test_result_sheet)
        wb.write_a_line_in_sheet(test_steps_info[0], fgcolor="CD9B9B")
        for test_data_dict in test_datas_dict:  # 把列表中的每个子字典做遍历：
            if test_data_dict["是否执行"] != "y":
                continue
            test_steps_info = get_test_info(test_data_file_path, test_step_sheet_name)
            flag = True
            for test_step in test_steps_info[1:]:
                test_step_description = test_step[test_case_description_col_no]
                # 以下三行：将测试步骤的三部分取出来，赋值给三个变量
                # 第一个是关键字、第二个是定位表达式，第三个表达式是操作值
                keyword = test_step[keyword_col_no]
                locator = test_step[locator_col_no]
                value = test_step[value_col_no]
                print("***********:", test_data_dict, type(value), value)
                if value is not None and re.search(r"\$\{.*?\}", str(value)):
                    key = re.search(r"\$\{(.*?)\}", value).group(1)
                    value = test_data_dict[key]
                    test_step[value_col_no] = value
                    print("@@@@@@@@@@@@替换${%s}后的value是%s" % (key, value))

                if locator is None and value is None:
                    command = keyword + "()"
                elif locator is not None and value is None:
                    command = "%s('%s')" % (keyword, locator)
                elif locator is None and value is not None:
                    command = "%s('%s')" % (keyword, value)
                else:
                    command = "%s('%s','%s')" % (keyword, locator, value)
                print("------------:", command)
                test_step[test_step_time_col_no] = TimeUtil().get_chinesedatetime()
                try:
                    temp = eval(command)
                    if "open_browser" in command:
                        driver = temp
                    test_step[test_step_result_col_no] = "成功"
                except AssertionError as e:
                    print("断言失败")
                    flag = False
                    test_step[test_step_result_col_no] = "断言失败"
                    take_pic(driver)
                    wb.write_a_line_in_sheet(test_step, font_color="red")
                    break
                except Exception as eak:
                    print("突发异常")
                    flag = False
                    test_step[test_step_result_col_no] = "异常失败"
                    take_pic(driver)  # Action文件中的drive变量是""
                    wb.write_a_line_in_sheet(test_step, font_color="red")
                    break
                wb.write_a_line_in_sheet(test_step, font_color="green")
                wb.set_sheet_by_name(test_result_sheet)
            wb.write_a_line_in_sheet(["", "", ""])
            wb.save()
            test_time = TimeUtil().get_chinesedatetime()
            if flag:
                result = "成功"
            else:
                result = "失败"
            test_result.append({"测试时间": test_time, "测试结果": result})
    except Exception  as e:
        traceback.print_exc()
    return test_result


def dict_to_list(l):
    result = []
    result.append(list(l[0].keys()))
    for d in l[1:]:
        temp = []
        for value in d.values():
            temp.append(value)
        result.append(temp)
    return result


if __name__ == "__main__":
    # execute_test_case("联系人1",test_datas_dict)
    # print(get_test_cases(test_data_file_path, test_case_info_sheet))
    test_cases = get_test_cases(test_data_file_path, test_case_info_sheet)
    for test_case in test_cases:
        print(test_case)
        test_step_sheet_name = test_case["测试步骤sheet名称"]
        test_data_sheet_name = test_case["测试数据sheet名称"]
        test_datas_dict = get_test_data_from_sheet(test_data_file_path, test_data_sheet_name)
        a_group_test_data_result = execute_test_case(test_step_sheet_name, test_datas_dict)
        print(a_group_test_data_result)
        for i in range(len(test_datas_dict)):
            if test_datas_dict[i]["是否执行"] != "y": continue
            test_datas_dict[i]["执行时间"] = a_group_test_data_result[i]["测试时间"]
            test_datas_dict[i]["测试结果"] = a_group_test_data_result[i]["测试结果"]
        print(test_datas_dict)
        wb = ExcelUtil(test_data_file_path)
        wb.set_sheet_by_name(test_result_sheet)
        test_data_list = dict_to_list(test_datas_dict)
        print(test_data_list)
        for i in test_data_list:
            wb.write_a_line_in_sheet(i)
        wb.save()
