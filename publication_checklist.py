import sys
import xml.etree.ElementTree as ET
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas as pd
pd.options.mode.chained_assignment = None

"""
Description:
This function allows us to combine the information available in the ACM SIG
Conference Management System with the paper information available on the
HotCRP submission website to create a .xlsx file.

The purpose of the .xlsx file is to facilitate the compliance checks of
camera-ready papers at ACM conferences. At the moment, the check criteria are
as follows:
- DOI and ISBN information on paper are consistent with what ACM has provided
- template used by authors is consistent with conference-determined ACM template
- CCS concepts and keywords are included in papers
- authors' emails and affiliations are included in the author block
- ACM reference format is included in the paper

Moreover, users might use this .xlsx to look up authors and their email
addresses.

Input required:
- .csv file produced using the 'Create CSV' button on the ACM SIG Conference
Management System
- .xml file produced using the 'Download TOC' button on the HotCRP submission
site.
Output:
- .xlsx file which contains:
    - paper ID (as listed in HotCRP website)
    - paper title
    - paper type
    - author list (separated by semicolons)
    - author emails (separated by semicolons)
    - DOI information (available from ACM SIG Conference Management System)
    - check criteria, such as:
        - DOI checklist
        - template check
        - CCS and keywords check
        - email and affiliation check
        - ACM reference format check
    - overall status (pre-filled with .xlsx formula that produces "FAIL" or "OK")-

Syntax: python publication-checklist.py argv[1] argv[2] argv[3] argv[4]
argv[1]: this points to the .csv file produced by ACM
argv[2]: this points to the .xml file produced by HotCRP
argv[3]: this is the preamble found in the .xml file for the field
        "event_tracking_number"
argv[4]: this points to the resulting .xlsx file produced by this script

Syntax: python publication-checklist.py <path-to-.csv-file-produced-by-ACM>
        <path-to-.xml-file-from-HotCRP>
        <preamble-found-in-.xml-file> <path-to-output-.xlsx-file>
Example syntax: python publication_checklist.py data/sample-acm-cms-export.csv
                data/sample-main-acmcms-toc.xml 'eenergy20-p'
                exported-data/sample-checklist.xlsx
"""

def create_acm_df(acm_in):
    acm_data = pd.read_csv(acm_in)
    acm_data['P_title'] = acm_data['Title'].apply(lambda x: x.split('\\')[0].split('full strip notes')[0].rstrip())
    acm_data['Type'] = acm_data['Rights Granted'].apply(lambda x: x.split('pdf')[0].rstrip())
    acm_sdata = acm_data[['P_title', 'DOI']]

    return acm_sdata

def create_hotcrp_df(xml_in, handle_in):
    """
    handle_in is the preamble used in <event_tracking_number> of the .xml file.
    Unfortunately this cannot be automated.
    """

    def _read_xml(xml_in):
        tree = ET.parse(xml_in)
        xml_out = tree.getroot()

        return xml_out

    def _flatten_list(list_in):
        flattened_string = ';'.join(list_in)

        return flattened_string

    def _create_email_list(list_in):
        l_email = ';'.join(list_in)

        return l_email

    paper_list = []
    xml_data = _read_xml(xml_in)
    for member in xml_data.findall('paper'):
        paper_type = member.find('paper_type').text
        paper_title = member.find('paper_title').text
        paper_number = member.find('event_tracking_number').text
        paper_number = paper_number.split(handle_in)[1].rstrip()
        for author_in in member.findall('authors'):
            author_names = []
            author_emails = []
            for a_author in author_in.findall('author'):
                a_count = 0
                f_name = a_author.find('first_name').text
                m_name = a_author.find('middle_name').text
                l_name = a_author.find('last_name').text
                if m_name is None:
                    a_name = f_name + " " + l_name
                else:
                    a_name = f_name + " " + m_name + " " + l_name
                author_names.append(a_name)
                a_email = a_author.find('email_address').text
                author_emails.append(a_email)
                a_count += 1
        paper_authors = _flatten_list(author_names)
        paper_emails = _create_email_list(author_emails)
        paper_list.append([paper_number,
                           paper_type,
                           paper_title,
                           paper_authors, paper_emails])
    hotcrp_data = pd.DataFrame(paper_list, \
                            columns=['Id', 'Type', 'P_title', 'Author', 'Email'])
    return hotcrp_data

def create_merged_df(acm_in, hotcrp_in, columns_in, how_join):

    def _add_columns(df_in, col_list):
        for col_in in col_list:
            df_in[col_in] = ""

        return df_in

    if how_join == 'outer':
        if len(acm_in) == len(hotcrp_in):
            df_temp = pd.merge(acm_in, hotcrp_in, on='P_title', how=how_join)
            print('Merge success')
            df_temp2 = df_temp[['Id', 'P_title', 'Type', 'Author', 'Email', 'DOI']]
            df_temp2.astype({'Id':'int32'}).dtypes
            df_temp2.sort_values(by='Id', ascending=True)
            df_out = _add_columns(df_temp2, columns_in)
            return df_out
        else:
            print('Check dataframe')
    else:
        df_temp = pd.merge(acm_in, hotcrp_in, on='P_title', how=how_join)
        print('Merge success')
        df_temp2 = df_temp[['Id', 'P_title', 'Type', 'Author', 'Email', 'DOI']]
        df_temp2.astype({'Id':'int32'}).dtypes
        df_temp2.sort_values(by='Id', ascending=True)
        df_out = _add_columns(df_temp2, columns_in)
        return df_out

def write_to_excel(df_in, filetosave, columns_in):
    writer = pd.ExcelWriter(filetosave,
                            engine='xlsxwriter',
                            options={'strings_to_urls': False,
                                     'strings_to_numbers':True})
    df_in.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    last_col_idx = len(df_in.columns)
    overallstat_idx = df_in.columns.get_loc('Overall_check')

    # Add a format. Light red fill with dark red text.
    format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                   'font_color': '#9C0006'})

    # Add a format. Green fill with dark green text.
    format2 = workbook.add_format({'bg_color': '#C6EFCE',
                                   'font_color': '#006100'})

    for row_now in range(1, len(df_in)+1):
        cell_compute = xl_rowcol_to_cell(row_now, overallstat_idx)
        cell_start = xl_rowcol_to_cell(row_now, last_col_idx - len(columns_in))
        cell_end = xl_rowcol_to_cell(row_now, last_col_idx-2)
        formula_str = '=IF(COUNTIF('+cell_start+':'+cell_end+', "FALSE") > 0, "FAIL", "OK")'
        worksheet.write_formula(cell_compute, formula_str)
        worksheet.data_validation(cell_start + ':' + cell_end, \
                                  {'validate': 'list', \
                                  'source':['FALSE', 'OK'], \
                                  'error_message':'The value to be entered must be either "FALSE" or "OK"'})

    cell_compute_start = xl_rowcol_to_cell(1, last_col_idx-1)
    cell_compute_end = xl_rowcol_to_cell(len(df_in), last_col_idx-1)

    worksheet.conditional_format(cell_compute_start + ':' + cell_compute_end, \
                                 {'type': 'cell', \
                                 'criteria': 'equal to', \
                                 'value':'"FAIL"', \
                                 'format': format1})
    worksheet.conditional_format(cell_compute_start + ':' + cell_compute_end, \
                                 {'type': 'cell', \
                                 'criteria': 'equal to', \
                                 'value':'"OK"', \
                                 'format': format2})

    writer.save()

if __name__ == "__main__":
    # Create a dataframe using the ACM-produced .csv file
    acm_df = create_acm_df(acm_in=sys.argv[1])

    # Create a dataframe using the HotCRP-produced .xml file
    hotcrp_df = create_hotcrp_df(xml_in=sys.argv[2], handle_in=sys.argv[3])

    # Merge the information from ACM and HotCRP
    COLUMNS_ADD = ['DOI_check',
                   'Template_check',
                   'CCS_Keyword_check',
                   'Email_affiliation_check',
                   'ACM_ref_check',
                   'Author_header_check',
                   'Track_title_check',
                   'Overall_check']
    df_all = create_merged_df(acm_in=acm_df,
                              hotcrp_in=hotcrp_df,
                              columns_in=COLUMNS_ADD,
                              how_join='inner')

    # Write the information to .xlsx file
    write_to_excel(df_all, sys.argv[4], COLUMNS_ADD)
