import os
import re
from flask import current_app
from bs4 import BeautifulSoup as BS
from .get_awr_file import FileOperation


class AnalyzeBase(object):

    def __init__(self):
        self.text = None

    # always return data like (status=True/False, data=[[], [], []])
    def parse(self, soup, id: str, flag=None):
        result = self.get_result_from_awrrpt(soup, id)
        if not result or (flag and result and flag > len(result)-1):
            return False, None, None

        flag = len(result) - 1 if flag is None else flag
        title = [i.string for i in result[flag].find_all('th')]
        tmp = [i.string.strip() for i in result[flag].find_all('td')]
        if not len(title):
            return True, tmp, []

        rows = [tmp[i: len(title) + i] for i in range(0, len(tmp), len(title))]
        return True, rows, title

    def get_result_from_awrrpt(self, soup, id):
        summary = DataGetMapping.data.get(id)
        return soup.find_all("table", summary=re.compile(summary))

    @staticmethod
    def get_soup(history_id):
        file = FileOperation(history_id).get_html_location()
        if not os.path.exists(file):
            return None

        return AnalyzeBase.get_soup_from_file(file)

    @staticmethod
    def get_soup_from_file(file):
        try:
            soup = BS(open(file, 'r', encoding='utf-8'), 'html.parser')
        except Exception:
            soup = BS(open(file, 'r', encoding='gbk'), 'html.parser')
        return soup

    def check_release(self, soup):
        status, data, title = self.parse(soup, '1', 0)
        if not status:
            return False

        res = dict(zip(title, data[0]))
        release = res.get('Release', '0')
        return int(release.replace('.', '')) > 112020


class DataGetMapping:
    ''' mapping the table summary with your customize keys.
        when change the ways to get awr content, you can add a new variable and a
        new function like get_soup_result_from_sql '''

    data = {'1': 'database instance information',
            '2': 'host information',
            '3': 'snapshot information',
            '4': 'This table displays operating systems statistics',
            '5': 'wait class statistics ordered by total wait time',
            '6': 'memory statistics',
            '7': 'global cache load',
            '8': 'instance efficiency percentages',
            '9': 'PGA aggregate target histograms',
            '10': 'different time model statistics',
            '11': 'load profile',
            '12': 'IO profile',
            '13': 'IO Statistics for different physical files',
            '14': 'name and value of init.ora parameters',
            '15': 'memory dynamic component statistics',
            '16': 'shared pool advisory',
            '17': 'This table displays MTTR advisory',
            '18': 'PGA memory advisory for different estimated PGA target sizes',
            '19': 'SGA target advisory for different SGA target sizes.',
            '20': 'background wait events statistics',
            '21': 'Foreground Wait Events and their wait statistics',
            '22': 'top SQL by number of parse calls',
            '23': 'top SQL by version counts',
            '24': 'top SQL by elapsed time',
            '25': 'top SQL by number of executions',
            '26': 'top SQL by CPU time',
            '27': 'top SQL by buffer gets',
            '28': 'top SQL by physical reads',
            '29': 'top segments by logical reads',
            '30': 'top segments by physical reads.',
            '31': 'top segments by direct physical reads',
            '32': 'Key Instance activity statistics',
            '33': 'top segments by row lock waits',
            '34': 'workload characteristics for global',
            '35': 'IC ping latency statistics',
            '36': 'Dynamic Remastering Stats',
            '37': 'top segments by global cache buffer busy waits.',
            '38': 'buffer pool statistics for different types of buffers'}
