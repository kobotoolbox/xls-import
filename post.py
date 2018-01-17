import os
import netrc
import requests
import glob
from datetime import datetime
import csv
import re

host = 'kc.kobotoolbox.org'
homedir = os.path.expanduser('~')
info = netrc.netrc(homedir + "/.netrc")
username, account, password = info.authenticators(host)

filelist = glob.glob("tempfiles/*.xml")

filelist_tuples = [(f, open(f, 'rb')) for f in filelist]
filelist_tuples = tuple(filelist_tuples)

logdata = []
logfilename = datetime.now().strftime('postlog__%Y-%m-%d_%H-%M.csv')


for ft in filelist_tuples:
    files = {'xml_submission_file': ft}

    response = requests.post('https://kc.kobotoolbox.org/submission', files=files, auth=(username, password))

    d = {}
    d['res'] = response
    # get uuid from filepath in files tuple
    d['uuid'] = re.split(r'[\/.]', ft[0])[1]
    logdata.append(d)


with open("kc"+logfilename, 'w') as csvfile:
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
