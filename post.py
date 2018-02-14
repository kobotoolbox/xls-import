import os
import netrc
import requests
import glob
from datetime import datetime
import csv
import re
import sys


# Set .netrc host, server URL, and log filename prefix.
netrc_host = 'kc.kobotoolbox.org' # This should match entry in .netrc file
server_url = 'https://kc.kobotoolbox.org/submission'
log_prefix = 'kc'

homedir = os.path.expanduser('~')
info = netrc.netrc(homedir + "/.netrc")
username, account, password = info.authenticators(netrc_host)


# https://stackoverflow.com/questions/20255124/too-many-open-files-using-requests-module?rq=1
class LazyFile(object):

    def __init__(self, filename, mode):
        self.filename = filename
        self.mode = mode

    def read(self):
        with open(self.filename, self.mode) as f:
            return f.read()


# http://stackoverflow.com/questions/3173320/text-progress-bar-in-the-console/27871113#comment50529068_27871113
def progress(count, total, status=''):
    bar_len = 60
    filled_len = int(round(bar_len * count / float(total)))

    percents = int(round(100 * count / float(total), 0))
    bar = '=' * filled_len + '-' * (bar_len - filled_len)

    sys.stdout.write('[%s] %s%s ...%s\r' % (bar, percents, '%', status))
    sys.stdout.flush()  # As suggested by Rom Ruben




filelist = glob.glob("tempfiles/*.xml")

filelist_tuples = [(f, LazyFile(f, 'rb')) for f in filelist]
filelist_tuples = tuple(filelist_tuples)

logdata = []
logfilename = datetime.now().strftime('postlog__%Y-%m-%d_%H-%M.csv')

i = 1
total = len(filelist_tuples)
for ft in filelist_tuples:
    if i == total:
        progress(i, total, status='...Posting...')
        print "\nWriting log file..."
        break
    else:
        files = {'xml_submission_file': ft}

        response = requests.post(server_url, files=files, auth=(username, password))

        d = {}
        d['res'] = response
        # get uuid from filepath in files tuple
        d['uuid'] = re.split(r'[/.]', ft[0])[1]
        logdata.append(d)
        progress(i, total, status='Posting...')
        i += 1


with open(log_prefix + logfilename, 'w') as csvfile:
    fieldnames = ['code', 'response_text', 'date', 'uuid']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, dialect='excel')
    writer.writeheader()

    for ld in logdata:
        writer.writerow({
            'code':          ld['res'].status_code,
            'response_text': ld['res'].text,
            'date':          ld['res'].headers['Date'],
            'uuid':          ld['uuid']
        })
    print (log_prefix + logfilename + " Complete")
