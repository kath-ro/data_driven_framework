from Action.Action import login,add_person_info
from Util.Excel import *
from Conf.ProjVar import *
from Action.Action import login,add_person_info
import traceback
from Util.DateAndTime import *
from Util.Log import *
from Util.TakePic import take_pic

wb = ExcelUtil(test_data_file_path)
def get_test_data(sheet_name):
    global wb
    wb.set_sheet_by_name(sheet_name)
    user_info = wb.get_sheet_all_data()
    return user_info

def run_test_case():
    user_info = get_test_data(test_user_info_sheet)
    info("________________"*5+"\n"+"测试开始了！\n")

    for line in user_info[1:]:
        if line[execute_flag_col_no] is None:
            continue
        if "y" in line[execute_flag_col_no]:
            info("此用户数据要被执行：%s,%s,%s" %(line[username_col_no],line[password_col_no],line[test_data_sheet_name_col_no]))
            #print("此数据要被执行",line[username_col_no],line[password_col_no],line[test_data_sheet_name_col_no])
            username = line[username_col_no]
            password = line[password_col_no]
            test_data_sheet = line[test_data_sheet_name_col_no]
            add_contacts_data = get_test_data(test_data_sheet)
            print("---:",add_contacts_data)
            driver = login(username,password) #调用一下登录逻辑
            wb.set_sheet_by_name(test_result_sheet)
            wb.write_a_line_in_sheet(add_contacts_data[0],fgcolor="CD9B9B") #在测试结果页写了一个表头
            flag = True
            for i in add_contacts_data[1:]:  #遍历所有新建联系人的数据，在登录状态下进行新建操作
                if i[test_data_execute_flag_col_no] is None:continue
                i[test_data_time_col_no]=TimeUtil().get_chinesedatetime()
                if "y" in i[test_data_execute_flag_col_no]:
                    contact_name = i[contact_name_col_no]
                    contact_email =i[contact_email_col_no]
                    contact_star = i[contact_star_col_no]
                    contact_mobile =i[contact_mobile_col_no]
                    contact_comment = i[contact_comment_col_no]
                    assert_word = i[assert_word_col_no]
                    info("当前联系人数据行：%s,%s,%s,%s,%s," %(contact_name,contact_email,contact_star,contact_mobile,contact_comment))
                    try:
                        add_person_info(driver,contact_name,contact_email,contact_star, contact_mobile,contact_comment)
                        assert assert_word in driver.page_source
                        i[test_data_result_no]="成功"#设定一下测试结果列中的状态
                        info("当前测试数据执行成功了")
                    except AssertionError:
                        info(traceback.format_exc())
                        i[test_data_result_no] = "断言失败"  # 设定一下测试结果列中的状态
                        flag = False
                        info("当前测试数据执行断言失败了")
                        take_pic(driver)

                    except Exception as e:
                        info(traceback.format_exc())
                        i[test_data_result_no] = "失败"  #设定一下测试结果列中的状态
                        flag = False
                        info("当前测试数据执行失败了")
                        take_pic(driver)

                if i[test_data_result_no]=="成功":
                    wb.write_a_line_in_sheet(i,font_color="green")

                else:
                    wb.write_a_line_in_sheet(i, font_color="red")
                    #wb.write_a_line_in_sheet(i)  #把每一行测试数据的测试结果写入到文件中
                if flag == True:
                    line[test_user_info_result_col_no] = '成功'
                else:
                    line[test_user_info_result_col_no] = '失败'
                line[test_time_col_no] = TimeUtil().get_chinesedatetime()
                wb.save()
            wb.write_a_line_in_sheet(user_info[0],"CD9B9B")
            if line[test_user_info_result_col_no] == '成功':
                info("当前测试用户 %s 的所有测试数据均执行成功了" % line[username_col_no])
                wb.write_a_line_in_sheet(line,font_color="green")
            else:
                wb.write_a_line_in_sheet(line, font_color="red")
                info("当前测试用户 %s 的所有测试数据中有失败的测试用例！" % line[username_col_no])
            wb.write_a_line_in_sheet([])
            wb.write_a_line_in_sheet([])
            wb.save()
            driver.quit()

    info("测试结束了！\n")