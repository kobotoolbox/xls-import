# xls-import

xls2xml.py and post.py run under python 2.7.


## xls2xml.py

```
$ pip install lxml
$ pip install xlrd
$ python xls2xml.py SPREADSHEETNAME.xlsx
```

#### Underlying Spreadsheet Rules
1. The spreadsheet must include a worksheet (following all the content worksheets) called ```IDSheet```. This worksheet contains a minimum of two cells. Cell B1 contains the KPI ID. Cell B2 contains the KC ID. These two IDs will be provided to you.

    ![screen shot 2017-12-20 at 10 58 17 pm](https://user-images.githubusercontent.com/192568/34240033-6399f582-e5d9-11e7-9e0f-fd86c946e5a9.png)

1.  \_\_version\_\_ and other meta columns follow content columns. \_\_version\_\_ is the first column following the content.
1.  Format all cells as text and pass them through verbatim to the XML output. See #4 (issuecomment-352513439) and #5 (issuecomment-352816376)
1. Select_multiples columns may be formatted either like column G in the image below, or as boolean columns like columns H - K.
1. If using boolean select_multiple columns, the headers use the character "/" as an element name/content delimiter. All columns need to be formatted as text prior to entering '1' or '0'
1. Repeating groups appear in a new sheet
1. Non-repeating groups appear in the first sheet. columns contain the colon character ':' as a hierarchy separator, such as three_favorite_haircuts:



## post.py

post.py will post tempfiles/*.xml to kc.kobotools.or:

Before you run post.py:
- Copy the example-.netrc file to your home directory
- Update with your login credentials
- Rename to .netrc, or append to existing ~/.netrc (Check to see if ~/.netrc already exists before renaming)

```
$ pip install requests
$ python post.py
```

Note: post.py outputs a date-stamped log with a name such as: ```kcpostlog__2018-01-16_18-04.csv```. After running the program, this will usually be the last file created. To view the file, it may be easier to run the following command than to try to type out the full filename.
```
$ more "$(ls -rt | tail -n1)"
```