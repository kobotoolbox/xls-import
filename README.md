# xls-import

xls2xml.py and post.py run under python 2.7.


## xls2xml.py

```
$ pip install lxml
$ pip install xlrd
$ python xls2xml.py SPREADSHEETNAME.xlsx
```

#### Underlying Spreadsheet Rules
1. The spreadsheet must include a worksheet (following all the content worksheets) called ```IDSheet```. This worksheet contains a minimum of two cells. Cell B1 contains the KPI ID. Cell B2 contains the KC ID. Refer to this video to determine how to find those two IDs: https://youtu.be/vMz_Q0yqpm8

    ![screen shot 2017-12-20 at 10 58 17 pm](https://user-images.githubusercontent.com/192568/34240033-6399f582-e5d9-11e7-9e0f-fd86c946e5a9.png)

1.  \_\_version\_\_ and other meta columns follow content columns. \_\_version\_\_ is the first column following the content.
1.  Format all cells as text and pass them through verbatim to the XML output. See #4 (issuecomment-352513439) and #5 (issuecomment-352816376)
1. Select_multiples columns may be formatted either like column G in the image below, or as boolean columns like columns H - K.
1. If using boolean select_multiple columns, the headers use the character "/" as an element name/content delimiter. All columns need to be formatted as text prior to entering '1' or '0'
1. Repeating groups appear in a new sheet
1. Non-repeating groups appear in the first sheet. columns contain double colon characters (`::`) as a hierarchy separator, such as `group_recent_haircuts::Last_Haircut`



## post.py

post.py will post tempfiles/*.xml to kc.kobotools.org:

Before you run post.py:
- Check (and double-check) to see if a ~/.netrc file already exists. If not, create it.
- Ensure that the permissions on the ~/.netrc file are read/write, restricting access to only the owner. The chmod command below will do that.
- Using the example-.netrc file as a guide, update  ~/.netrc with your login credentials. (If the file already existed, just append the three lines to what was already there.)
- If you are not using kc.kobotoolbox.org as your destination server, update the .netrc host, server URL, and log filename prefix in post.py (lines 10–12).

```
$ chmod 600 $HOME/.netrc
$ pip install requests
$ python post.py
```

Note: post.py outputs a date-stamped log with a name such as: ```kcpostlog__2018-01-16_18-04.csv```. After running the program, this will usually be the last file created. To view the file, it may be easier to run the following command than to try to type out the full filename.
```
$ more "$(ls -rt | tail -n1)"
```

# Frequently Asked Questions

- [I get an XLRDError](#i-get-an-xlrderror)
- [How do I determine the KPI and KC IDs?](#how-do-i-determine-the-kpi-and-kc-ids)


## I get an XLRDError

If you get
```
xlrd.biffh.XLRDError: No sheet named <'IDSheet'>
```
be sure to add a new sheet named
'IDSheet'. Place the KPI ID in cell B1, and the KC ID in cell B2.

## How do I determine the KPI and KC IDs?
Refer to this video to determine how to find those two IDs: https://youtu.be/vMz_Q0yqpm8


