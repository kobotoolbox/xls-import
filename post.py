import os
import netrc
import requests
import glob
from datetime import datetime
import csv

host = 'kc.kobotoolbox.org'
homedir = os.path.expanduser('~')
info = netrc.netrc(homedir + "/.netrc-kobo")
username, account, password = info.authenticators(host)

filelist = glob.glob("tempfiles/*.xml")

filelist_tuples = [(f, open(f, 'rb')) for f in filelist]
filelist_tuples = (filelist_tuples)

logdata = []
logfilename = datetime.now().strftime('postlog__%Y-%m-%d_%H-%M.csv')


for ft in filelist_tuples:
    files = {'xml_submission_file': ft}

    response = requests.post('https://kc.kobotoolbox.org/submission', files=files, auth=(username, password))
    logdata.append(response)


with open("kc"+logfilename, 'w') as csvfile:
    fieldnames = ['code', 'response_text', 'date']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

    writer.writeheader()
    for r in logdata:
        writer.writerow({'code': r.status_code, 'response_text': r.text, 'date': r.headers['Date']})
