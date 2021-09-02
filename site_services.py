'''
Module for searching fo the trial acts in Moscow
'''


from selenium import webdriver
import selenium

class CourtSearch:
    '''
    Utils to search for acts of courts of Moscow
    '''

    def __init__(self, headless=True):

        driver_path = './chromedriver/chromedriver'
        chrome_options = webdriver.ChromeOptions()
        chrome_options.headless = headless

        self.driver = webdriver.Chrome(
                executable_path=driver_path,
                options=chrome_options,
                )

    def search_by_name(self, last_name, first_name, patronymic):
        '''
        Searchs by name for the trial acts in Moscow, returns result in a form
        list where the first index is the concatenated name and other are tuples
        with a found information.
        Arguments are selfexplained
        '''
        url = 'https://sudrf.ru/index.php'

        params = [
            '?id=300&page=0&act=go_sp_search&searchtype=sp&court_subj=77&suds_subj=',
            '&num_d=&f_name=',
            '&date_num_in=&date_num_out=&suds_vid=&spkatg=&suds_pip=&st_cat=&sud_pip=',
            ]

        courts_subj_set = {'77OV0000','77GV0005','77KJ0002','77GV0001','77RS0035',
            '77OS0000','77AJ0001','77RS0001','77RS0002','77RS0003','77RS0004',
            '77RS0005','77RS0006','77RS0007','77RS0008','77RS0009','77RS0010',
            '77RS0011','77RS0012','77RS0013','77RS0014','77RS0015','77RS0016',
            '77RS0017','77RS0018','77RS0019','77RS0020','77RS0021','77RS0022',
            '77RS0023','77RS0024','77RS0025','77RS0026','77RS0027','77RS0028',
            '77RS0029','77RS0030','77RS0031','77RS0032','77RS0033','77RS0034'}

        name = '{}+{}+{}'.format(last_name, first_name, patronymic)

        result_list = ['{} {} {}'.format(last_name, first_name, patronymic)]

        for court in courts_subj_set:
            self.driver.get(url+params[0]+court+params[1]+name+params[2])
            self.driver.implicitly_wait(1.3)

            try:
                table = self.driver.find_element_by_tag_name('table')
                rows = [row for row in table.find_elements_by_tag_name('tr')
                           if row.get_attribute('id') != 'head_num']

            except selenium.common.exceptions.NoSuchElementException:
                continue

            else:
                for row in rows:
                    cells = row.find_elements_by_tag_name('td')
                    court = cells[0].text
                    number = cells[1].find_element_by_tag_name('a')
                    number = (number.text, number.get_attribute('href'))
                    number = number[0] + '\n' + number[1]
                    date = cells[2].text
                    info = cells[3].text
                    judge = cells[4].text
                    resolution = cells[5].text
                    act = cells[6].find_element_by_xpath('//a')
                    act = (act.text, act.get_attribute('href'))
                    act = act[0] + '\n' + act[1]
                    data = (court, number, date, info, judge, resolution, act)
                    result_list.append(data)

        return result_list
