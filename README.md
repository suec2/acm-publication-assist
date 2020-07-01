# publication-assist
This repository contains functions that were used to generate spreadsheets to assist with the ACM publication process.

## Usage
### Step 1: Install libraries
Install the Python packages:
```
pip install -r requirements.txt
```

### Step 2: Run the code
#### publication-xml-to-csv.py
This function converts the XML file available on the HotCRP website
(via 'Settings' -> 'ACM' -> 'Download TOC' button) into a .csv file containing:
- Paper Type
- Title
- Author names
- Author affiliation
- Lead author email
- Author email

The resulting .csv file has format consistent with ACM's requirement (as listed in https://www.acm.org/publications/gi-proceedings). Once the .csv file has been generated, save the .csv file with "UTF-8 encoding with BOM" using Sublime Text to ensure that the special characters are encoded correctly.

To run the code, use the following syntax:
```
python publication-xml-to-csv.py <path-to-xmlfile> <output-filename>
```

To test the code, use the following syntax:
```
python publication-xml-to-csv.py data/sample-main-acmcms-toc.xml exported-data/sample-main-acmcms-toc.csv
```

#### publication-checklist.py
This function allows us to combine the information available in the ACM SIG Conference Management System with the paper information available on the HotCRP submission website to create a .xlsx file. For more details on the contents on the resulting .xlsx file, refer to the code.

To run the code, use the following syntax:
```
python publication-checklist.py <path-to-.csv-file-produced-by-ACM> <path-to-.xml-file-from-HotCRP> <preamble-found-in-.xml-file> <path-to-output-.xlsx-file>
```

To test the code, use the following syntax:
```
python publication-checklist.py data/sample-acm-cms-export.csv data/sample-main-acmcms-toc.xml 'eenergy20-p' exported-data/sample-checklist.xlsx
```

#### publication-registration-status.py

This function allows us to combine the following information:
- .xml file containing the paper information that is available on the HotCRP submission website,
- .csv file created by Google Forms that the authors have to fill in upon conference registration,

to create a .xlsx file which provides an overview of the publication status, as well as list the registration status of the papers, sorted by tracks.

In particular, the .xlsx file contains the following worksheets:
- _FullList_: which lists all the papers that have been accepted at the conference and the paper information (e.g. type, title, authors, email addresses, affiliations, country, tags),
- _Summary_By_Types_: which provides a summary of the types of paper accepted at each track,
- _RegSummary_By_Tracks_: which provides a summary of the registration status for each track,
- _Unregistered_Papers_: which lists all the papers at the conference that have not yet registered to present at the conference,
- _*tag*_status_: which lists all the papers accepted at a particular *tag* track, for e.g. _main_status_ lists the papers accepted at the main conference.

This .xlsx file is useful to provide an overview of the current registration status, and to also facilitate the registration process in which both publication/registration chairs could reach out to authors that have not yet registered to attend at the conference.

To run the code, use the following syntax:
```
python publication-registration-status.py <path-to-.txt-file-listing-.xml-files-produced-by-HotCRP> <path-to-.csv-file-exported-from-Google-form> <path-to-output-.xlsx-file>
```

To test the code, use the following syntax:
```
python publication-registration-status.py data/sample-xml_list.txt data/sample-google-form.csv exported-data/sample-registration-status.xlsx
```
