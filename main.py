'''
The main executable module.
'''


from site_services import CourtSearch
from excel import ExcelWB


if __name__ == '__main__':

    excel = ExcelWB()
    name_list = excel.read_list()
    driver = CourtSearch()
    act_dict = {}

    for name in name_list:
        found_info = driver.search_by_name(*name)
        record_title = found_info.pop(0)
        if found_info:
            act_dict[record_title] = found_info
        else:
            act_dict[record_title] = []

    driver.driver.quit()
    excel.write_acts(act_dict)
    print('The script is finished')
