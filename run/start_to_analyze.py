import re
from flask import render_template

from .get_parse_data import AnalyzeBase
from .get_awr_file import FileOperation
from .data_serialize import CommonMethod
from .recommend_params import RECOMMEND
from app.models import AwrHistory, Hosts
from app.commons import TimeFormat
from app.extensions import db


def do_sth(func, soup, id, keep_rows=5):
    status, data, _ = func(soup, id)
    if not status:
        return False, None
    return True, data[:keep_rows]


class DbInstInfo(AnalyzeBase):
    def parse(self, soup):
        status1, data1, title1 = super().parse(soup, '1', 0)
        status2, data2, title2 = super().parse(soup, '1', 1)
        if not status1:
            return False, None, None

        data = data1[0] + data2[0] if status2 else data1[0]
        title = title1 + title2 if status2 else title1

        self.text = '\n'.join(['|  %s  |' % (' | '.join(title)),
                               '| ' + '--- | '*len(data),
                               '|  %s  |' % (' | '.join(data))])
        return True, self.text, dict(zip(title, data))


class PdbInfo(AnalyzeBase):
    def parse(self, soup):
        *self.text, _ = super().parse(soup, '39')
        return self.text


class HostInfo(AnalyzeBase):
    def parse(self, soup):
        *self.text, _ = super().parse(soup, '2')
        return self.text


class SnapshotInfo(AnalyzeBase):
    def parse(self, soup):
        st, dt, ti = super().parse(soup, '3')
        if not st:
            return False, None, None
        title = "| " + " | ".join([i if i else " " for i in ti]) + " |\n"
        flag = "| " + "--- | "*len(ti) + "\n"
        data = "".join(["| " + "|".join(i) + " |\n" for i in dt])

        return True, dt, "".join([title, flag, data])


class SystemStatTime(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '4')
        self.text = CommonMethod.dict_two_colums(data, [0, 1])
        return self.text


class TotalWaitTime(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '5')
        self.text = CommonMethod.dict_two_colums(data, [0, 3], True)
        return self.text


class MemoryStat(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '6')
        self.text = CommonMethod.dict_two_colums(data, [0, 1], True)
        return self.text


class GlobalCacheLoadProfile(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '7')
        self.text = CommonMethod.dict_two_colums(data, [0, 1], True)
        return self.text


class InstEfficiencyPer(AnalyzeBase):
    def parse(self, soup):
        status, data, _ = super().parse(soup, '8')
        if status:
            keys = [re.sub(r'\s+', ' ', j)
                    for i, j in enumerate(data) if i % 2 == 0]
            values = [j.strip() for i, j in enumerate(data) if i % 2 == 1]
            data = dict(zip(keys, values))

        self.text = (status, data)
        return self.text


class PGATarget(AnalyzeBase):
    def parse(self, soup):
        *self.text, _ = super().parse(soup, '9')
        return self.text

    def suggest(self, soup):
        status, text = self.parse(soup)
        if not status:
            return False, None

        data = [max(int(j[-1]), int(j[-2])) for j in text if j[-1] and j[-2]]
        status = True if max(data) == 0 else False
        return True, status

    @staticmethod
    def get_pass_data(data):
        one_pass = all([True if not int(i[-2]) else False for i in data])
        m_pass = all([True if not int(i[-1]) else False for i in data])

        return all([one_pass, m_pass])


class TimeModelStat(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '10')
        self.text = CommonMethod.dict_two_colums(data, [0, 2], True)
        return self.text


class LoadProfile(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '11')
        self.text = CommonMethod.dict_two_colums(data, [0, 1], True)
        return self.text


class IoProfile(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '12')
        self.text = CommonMethod.dict_two_colums(data, [0, 1], True)
        return self.text


class FileIOStat(AnalyzeBase):
    def parse(self, soup):
        status, data, title = super().parse(soup, '13')
        if status:
            pattern = [r'Av\s+Rd\(ms\)', ]
            flag, *_ = CommonMethod.get_title_col_index(pattern, title)
            __tmp = [CommonMethod.int(i[flag]) for i in data if i[flag]]
            data = max(__tmp) if __tmp else 0

        return status, data


class InitOraParam(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '14')
        self.text = CommonMethod.dict_two_colums(data, [0, 1])
        return self.text

    def suggest(self, soup, wanted_dict):
        status, data = self.parse(soup)
        self.flag = 0
        if status:
            result = list()
            for key, value in wanted_dict.items():
                default_value = data.get(key, '/')
                self.flag += 1 if value != default_value else 0

                result.append([key, default_value, value])
            data = result
        return status, data


class MemDynamicStat(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '15')
        self.text = CommonMethod.dict_two_colums(data, [0, 1], True)
        return self.text

    def suggest(self):
        status, data = self.text
        if not status:
            return False, data
        result = [True if i[0] == i[1] else False for i in data.values()]
        return True, all(result)


class SharePoolAdv(AnalyzeBase):
    def suggest(self, soup):
        status, data, title = super().parse(soup, '16')
        self.text = status, data, title
        if not status:
            return False, None

        pattern = [r'SP\s+Size\s+Factr', r'Est\s+LC\s+Size\s+\(M\)']
        factor_col, size_col = CommonMethod.get_title_col_index(pattern, title)
        for i in data:
            if float(i[factor_col]) == 1.0:
                return True, i[size_col]
        return False, None

    def graph(self):
        status, data, title = self.text
        if not status:
            return False, data

        pattern = [r'Shared\s+Pool\s+Size.*', r'Est\s+LC\s+Load\sTime\s*\(s\)',
                   r'SP\s+Size\s+Fact.*']
        size_col, time_col, _ = CommonMethod.get_title_col_index(
            pattern, title)

        xdata = [CommonMethod.int(i[size_col]) for i in data]
        ydata = [CommonMethod.int(i[time_col]) for i in data]
        conf = [i[size_col] for i in data if CommonMethod.int(i[_]) == 1.0]

        return True, (xdata, ydata, conf)


class BufferPoolAdv(AnalyzeBase):
    def graph(self, soup):
        status, data, title = super().parse(soup, '17')
        if not status:
            return False, None

        pattern = [r'Size\s+for\s+Est\s+\S+', r'Estimated\s+Phys\s+Reads.*',
                   r'Size\s+Factor.*']
        est_col, phy_col, _ = CommonMethod.get_title_col_index(pattern, title)

        xdata = [CommonMethod.int(i[est_col]) for i in data if i[0] == 'D']
        ydata = [CommonMethod.int(i[phy_col]) for i in data if i[0] == 'D']
        conf = [i[est_col]
                for i in data if CommonMethod.int(i[_]) == 1.0 and i[0] == 'D']

        return True, (xdata, ydata, conf)


class PgaMemAdv(AnalyzeBase):
    def graph(self, soup):
        status, data, title = super().parse(soup, '18')
        self.text = status, data, title
        if not status:
            return False, None

        pattern = [r'PGA\s+Target\s+Est.*', r'Estd\s+PGA\s+Overalloc\s+Count.*',
                   r'Size\s+Fact.*']
        est_col, count_col, _ = CommonMethod.get_title_col_index(
            pattern, title)

        xdata = [CommonMethod.int(i[est_col]) for i in data]
        ydata = [CommonMethod.int(i[count_col]) for i in data]
        conf = [i[est_col] for i in data if CommonMethod.int(i[_]) == 1.0]

        return True, (xdata, ydata, conf)


class SgaTargetAdv(AnalyzeBase):
    def graph(self, soup):
        status, data, title = super().parse(soup, '19')
        self.text = status, data, title
        if not status:
            return False, None

        pattern = [r'SGA\s+Target\s+Size.*', r'Est\s+Physical\s+Reads.*',
                   r'Size\s+Fact.*']
        size_col, read_col, _ = CommonMethod.get_title_col_index(
            pattern, title)

        xdata = [CommonMethod.int(i[size_col]) for i in data]
        ydata = [CommonMethod.int(i[read_col]) for i in data]
        conf = [i[size_col] for i in data if CommonMethod.int(i[_]) == 1.0]

        return True, (xdata, ydata, conf)


class BackgroundWaitEvent(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '20')
        self.text = CommonMethod.dict_two_colums(data, [0, 1], True)
        return self.text


class ForegroundWaitEvent(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '21')
        self.text = CommonMethod.dict_two_colums(data, [0, 1], True)
        return self.text


class SqlParseCall(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '22', keep_rows)


class SqlVersionCount(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '23', keep_rows)


class SqlElapsedTime(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '24', keep_rows)


class SqlExecutions(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '25', keep_rows)


class SqlCpuTime(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '26', keep_rows)


class SqlGets(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '27', keep_rows)


class SqlPhyReads(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '28', keep_rows)


class SegmentsLogReads(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '29', keep_rows)


class SegmentsPhyReads(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '30', keep_rows)


class SegmentsDirectReads(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '31', keep_rows)


class KeyInstActivityStat(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '32')
        self.text = CommonMethod.dict_two_colums(data, [0, 1], True)
        return self.text


class InstActivityStat(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '40')
        self.text = CommonMethod.dict_two_colums(data, [0, 1], True)
        return self.text


class SegRowLockWaits(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '33', keep_rows)


class GCEnqueueService(AnalyzeBase):
    def parse(self, soup):
        status, data, _ = super().parse(soup, '34')
        if not status:
            return False, None

        keys = [j for i, j in enumerate(data) if i % 2 == 0]
        values = [j for i, j in enumerate(data) if i % 2 == 1]
        data = dict(zip(keys, values))
        self.text = (True, data)

        return self.text


class PingLatencyStats(AnalyzeBase):
    def parse(self, soup):
        *self.text, _ = super().parse(soup, '35')
        return self.text


class DynamicRemastering(AnalyzeBase):
    def parse(self, soup):
        *self.text, _ = super().parse(soup, '36')
        return self.text


class SegmentsBufferBusy(AnalyzeBase):
    def parse(self, soup, keep_rows):
        return do_sth(super().parse, soup, '37', keep_rows)


class BufferPoolStat(AnalyzeBase):
    def parse(self, soup):
        *data, _ = super().parse(soup, '38')
        return CommonMethod.dict_two_colums(data, [0, 1], True)


def choose_template(soup):
    st, _, dt = DbInstInfo().parse(soup)
    st2, _, tx = SnapshotInfo().parse(soup)
    if not st or not st2:
        raise Exception('DB Info get error')
    cdb = dt.get('CDB')
    rel = dt.get('Release')

    if (cdb and cdb.upper() == 'YES') or ("CDB" in tx and 'YES' in tx):
        f_key = 'cdb'
        template = 'awr_rt_cdb.md'
        if re.match(r'12\.1.', rel):
            s_key = '12.1'
        else:
            s_key = '12.2'
    else:
        template = 'awr_rt.md'
        f_key = 'normal'
        if re.match(r'12\.1.', rel):
            s_key = '12.1'
        elif re.match(r'11\.0.', rel):
            s_key = '11.4'
        else:
            s_key = '12.2'
    return template, RECOMMEND[f_key][s_key]


def GetMarkdownStr(history_id):
    db.session.rollback()

    awr_hy = AwrHistory.query.get(history_id)
    host_id = awr_hy.host_id
    soup = AnalyzeBase.get_soup(history_id)
    template, dt = choose_template(soup)

    string = render_template(template,
                             all=all, any=any, soup=soup,
                             host=Hosts.query.get(host_id),
                             wanted_dict=dt,
                             date=TimeFormat.timestp2date(awr_hy.finish_at),
                             CommonMethod=CommonMethod,
                             FileFunc=FileOperation(history_id),
                             kw=globals())

    return string
