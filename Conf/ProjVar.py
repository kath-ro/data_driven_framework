import os

#proj_path = os.path.abspath(__file__)
#proj_path = os.path.dirname(os.path.abspath(__file__))
proj_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
print(proj_path)

PageElementLocator_file_path = os.path.join(proj_path,r"Conf\PageElementLocator.ini")
print(PageElementLocator_file_path)

test_data_file_path = os.path.join(proj_path,r"TestData\126邮箱联系人.xlsx")
print(test_data_file_path)

test_user_info_sheet = "126账号"
test_result_sheet = "测试结果"

username_col_no = 1
password_col_no = 2
test_data_sheet_name_col_no =3
execute_flag_col_no = 4
test_time_col_no = 5
test_user_info_result_col_no =6

contact_name_col_no = 1
contact_email_col_no =2
contact_star_col_no = 3
contact_mobile_col_no = 4
contact_comment_col_no = 5
assert_word_col_no = 6
test_data_execute_flag_col_no = 7
test_data_time_col_no = 8
test_data_result_no = 9






