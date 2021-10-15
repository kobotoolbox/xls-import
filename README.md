# xls-import

xls2xml.py and post.py run under python 2.7.


## xls2xml.py

```
$ pip install lxml
$ pip install xlrd
$ python xls2xml.py SPREADSHEETNAME.xls
```

#### Restrictions
- It currently only supports excel 97-2003 workbooks, so with the .xls extention
- Nested groups are not supported


#### Underlying Spreadsheet Rules
1. The spreadsheet should be formatted as an excel 97-2003 workbook, so with the .xls extention
1. The spreadsheet must include a worksheet (following all the content worksheets, so it should be the last worksheet in your xls-file) called ```IDSheet```. This worksheet contains a minimum of two cells. Cell B1 contains the KPI ID. Cell B2 contains the KC ID.

    ![IDs](https://user-images.githubusercontent.com/192568/34240033-6399f582-e5d9-11e7-9e0f-fd86c946e5a9.png)

	* The KPI ID is a 22 character long string and can be found in the url of the form in kobotoolbox (https://<span></span>kobonew.ifrc.org/#/forms/"KPI ID"/summary, or whatever server you are using)
	* The KC ID can be easily found with the api:
		* Navigate to https://kcnew.ifrc.org/api/v1/forms (or the v1 api of your server)
		* Search for KPI ID by hitting ctrl+F (or cmd+F on osx) and pasting the KPI ID, you should have two matches
		* Look for the corresponding "uuid", a 32 character long string. This is the KC ID
		* If this option does not work, refer to this video to determine how to find those two IDs: https://youtu.be/vMz_Q0yqpm8

1.  \_\_version\_\_ and other meta columns follow content columns. \_\_version\_\_ should be the first column following the content.
1. Select_multiples columns may be formatted either like column G in the image below, or as boolean columns like columns H - K.

	![multi-select](https://raw.githubusercontent.com/rodekruis/xls-import/NLRC-updates/multi-select.png)

1. If using boolean select_multiple columns, the headers use the character "/" as an element name/content delimiter.
1. Repeating groups appear in a new sheet
1. Non-repeating groups appear in the first sheet. columns contain the colon character ':' as a hierarchy separator, such as three_favorite_haircuts:



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


## I get an XLRDError

If you get
```
xlrd.biffh.XLRDError: No sheet named <'IDSheet'>
```
be sure to add a new sheet named ```IDSheet``` and make sure it's the last worksheet of the workbook. Place the KPI ID in cell B1, and the KC ID in cell B2.


