# coding=utf-8
import os
import re


class DataSerialize:

    def snapshot_format(data):
        snap_id = [i[0] for i in data]
        time_str = [i[1]for i in data]

        return dict(zip(snap_id, time_str))

    def awr_data_format(data):
        null = None
        string = '\n'.join([i[0] for i in data if i[0]])
        return string

    def graph_link_sub(md_location):
        path, filename = os.path.split(md_location)
        file = os.path.join(path, 'tmp.md')
        folder = path.split('/static/')[1]
        input_str = ''
        with open(md_location, 'r+', encoding='utf-8') as f:
            input_str = re.sub('/static/%s' % folder, '.', f.read())

        with open(file, 'w+', encoding='utf-8') as f:
            f.write(input_str)


class CommonMethod:

    def int(data):
        try:
            data = float(str(data).replace(',', '').strip()) if data else 0
            return round(data, 2)
        except Exception:
            return 0

    def sort(data, key, keep_rows):
        return sorted(data, key=lambda x: x[key], reverse=True)[:keep_rows]

    def str_format(string):
        return string.strip().replace(' ', '').replace(':', '')

    def re_sub_unneed(string):
        data = string.replace('$', r"\$").replace('_', r'\_', 1)
        return data

    def get_title_col_index(patterns, str_list):
        result = list()
        for i in patterns:
            pattern = re.compile(i)
            data = [i for i, j in enumerate(str_list) if pattern.findall(j)]
            result += data
        return result

    @classmethod
    def dict_two_colums(cls, value, cols, all_=False):

        status, data = value
        if not status:
            return status, data

        keys = [re.sub(r'\s+', ' ', i[0]) for i in data]
        values = [i[1:] for i in data] if all_ else [i[cols[1]] for i in data]
        data = dict(zip(keys, values))

        return status, data

    @classmethod
    def merge_info_from_tables(cls, dict1, dict2, key_num, nums, wanted_list):

        dict1 = dict1 if isinstance(dict1, dict) else {}
        dict2 = dict2 if isinstance(dict2, dict) else {}
        result = list()

        for key in wanted_list:
            val1 = dict1.get(key)
            val2 = dict2.get(key)
            _ = [i for i in (val1, val2) if i]
            if not _:
                continue
            value = map(list, zip(*_))
            result.append([key] + [sum(cls.int(k) for k in j) for j in value])

        return cls.sort(result, key_num, nums)
