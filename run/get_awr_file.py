import os
from flask import current_app
from pyecharts import Line
from app.commons import DownloadZip
from app.models import AwrHistory
from app.servers.rules import DelExpiredFile


#  get the data from sql and save as raw awr report
class FileOperation(object):

    def __init__(self, history_id):
        self.raw_html_folder = current_app.config['RAW_AWR_FOLDER']
        self.history_id = str(history_id)
        self.base_loca = os.path.join(self.raw_html_folder, self.history_id)

    def mkdir(self):
        if not os.path.exists(self.base_loca):
            os.mkdir(self.base_loca)

    def data_save_as_file(self, data, file_name):
        file = os.path.join(self.base_loca, file_name)
        with open(file, 'w+', encoding='utf-8') as f:
            f.write(data)

    def dele_expire_zip(self):
        files = [os.path.join(root, name)
                 for root, _, files in os.walk(self.base_loca)
                 for name in files if name.split('.')[1] == 'zip']
        try:
            for file in files:
                os.remove(file)
        except Exception as e:
            current_app.logger.err('delete expired zip err:%s' % str(e))

    def check_file_exist(self, file_name):
        awr_html = os.path.join(self.base_loca, file_name)
        string = '/static/awr_html/%s/%s' % (self.history_id, file_name)

        if os.path.exists(awr_html):
            return True, string
        return False, 'file has been clean'

    def get_html_location(self):
        obj = AwrHistory.query.get(self.history_id)
        awr_html = os.path.join(self.base_loca, '%s.html' % obj.name)

        return awr_html

    def download(self, file):
        DelExpiredFile.do_sth(self.base_loca, ['zip'])
        return DownloadZip.zip(self.base_loca, file, kepp_source_file=True)

    def generate_graph(self, legend, xdata, ydata, file, axis_name):
        try:
            xaxis_name, yaxis_name = axis_name
            line = Line(background_color='white')
            xdata = [str(i) for i in xdata]
            line.add(legend, xdata, ydata, is_smooth=True,
                     xaxis_name=xaxis_name,
                     yaxis_name=yaxis_name,
                     xaxis_name_pos='middle',
                     yaxis_name_pos='end',
                     yaxis_interval=2,)

            path = os.path.join(self.base_loca, '%s.jpeg' % file)
            line.render(path)
            string = '/static/awr_html/%s/%s.jpeg' % (self.history_id, file)

            return string
        except Exception as e:
            current_app.logger.err('generate graph err:%s' % str(e))
