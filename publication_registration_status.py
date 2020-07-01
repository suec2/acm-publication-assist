import sys
import xml.etree.ElementTree as ET
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell
import numpy as np

"""
Description:
This function allows us to combine the following information:
- .xml file containing the paper information that is available on the HotCRP
submission website,
- .csv file created by Google Forms that the authors have to fill in upon
conference registration,
to create a .xlsx file which provides an overview of the publication status,
as well as list the registration status of the papers, sorted by tracks.

The purpose of the .xlsx file is to facilitate the sharing of publication and
registration information in a manner that is consumable for the General Chairs
and potentially TPC chairs.

Input required:
- .csv file produced by Google Forms
- .xml file produced using the 'Download TOC' button on the HotCRP submission
site.

Output:
- .xlsx file which contains the following worksheets:
    - "FullList": which lists all the papers that have been accepted at the
    conference, and the paper information, e.g.
        - Type: paper type ("Short Paper", "Full Paper", "Poster Paper"),
        - P_title: paper title
        - Author: author names
        - Email: author emails
        - Affiliation: author affiliations
        - Country: author countries
        - Tag: track (main, etc)
    - "Summary_By_Types" which provides a summary of the types of paper at each track
    - "RegSummary_by_Tracks" which provides a summary of the registration status
    for each track
    - "Unregistered_Papers" which lists the papers at the conference that have
    not yet registered to present at the conference
    - "<tag>_status" which lists all the papers accepted at a particular <tag>
    track, for e.g. "main_status" lists the papers accepted at the main conference
    and contains the following information:
        - Type_x: paper type
        - P_title: paper_title
        - Author: author names
        - Email: author emails
        - Affiliation: author affiliations
        - Country: author countries
        - Tag: track (main, etc)
        - Username: the email used to register on the Google form
        - P_id: the paper ID in the conference submission website
        - Author_r: the name of the author registered to attend the conference
        - R_status: the registration status
            - "registered": at least one of the authors have registered
            - "not-registered": none of the authors have registered, or if they had,
            they did not fill in the Google form
        - P_status: the paper status
            - "included": the paper has been accepted in the conference proceedings
            - "excluded": the paper has not been accepted/has been withdrawn,
            but someone has registered for the conference

Syntax: python publication-checklist.py argv[1] argv[2] argv[3]
argv[1]: this points to the .text file listing the .xml files produced by HotCRP
argv[2]: this points to the .csv file produced by Google forms
argv[3]: this points to the resulting .xlsx file produced by this script

Example syntax: python publication_registration_status.py
                data/sample-xml_list.txt
                data/sample-google-form.csv
                exported-data/sample-registration-status.xlsx"""

def read_xml_text(xml_in):
    xml_d = {}
    with open(xml_in) as f_name:
        for line in f_name:
            (tag, file_location) = line.split()
            xml_d[tag] = file_location
    return xml_d

def create_dictdf(dict_in):
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

    def _create_df(xml_in, tag_in):
        paper_list = []
        xml_data = _read_xml(xml_in)
        for member in xml_data.findall('paper'):
            paper_type = member.find('paper_type').text
            paper_title = member.find('paper_title').text
            for author_in in member.findall('authors'):
                author_names = []
                author_emails = []
                author_affiliations = []
                author_countries = []
                for a_author in author_in.findall('author'):
                    a_count = 0
                    f_name = a_author.find('first_name').text
                    m_name = a_author.find('middle_name').text
                    l_name = a_author.find('last_name').text
                    if m_name is None:
                        a_name = f_name + " " + l_name
                    else:
                        a_name = f_name + " " + m_name + " " + l_name
                    a_affiliation = a_author.find('affiliation').text
                    a_country = a_author.find('country').text
                    author_affiliations.append(a_affiliation)
                    author_names.append(a_name)
                    author_countries.append(a_country)
                    a_email = a_author.find('email_address').text
                    author_emails.append(a_email)
                    a_count += 1
            paper_authors = _flatten_list(author_names)
            paper_emails = _create_email_list(author_emails)
            paper_affiliations = _flatten_list(author_affiliations)
            author_countries = list(set([country_x for country_x in \
                                         author_countries \
                                         if str(country_x) != 'None']))
            paper_countries = _flatten_list(author_countries)
            paper_list.append([paper_type, \
                               paper_title, \
                               paper_authors, \
                               paper_emails, \
                               paper_affiliations, \
                               paper_countries, \
                               tag_in])
        paper_df = pd.DataFrame(paper_list, columns=['Type',
                                                     'P_title',
                                                     'Author',
                                                     'Email',
                                                     'Affiliation',
                                                     'Country',
                                                     'Tag'])
        return paper_df

    data_dict = {}
    for tag_now in dict_in:
        data_dict[tag_now] = _create_df(dict_in[tag_now], tag_now)

    return data_dict

def create_googledf(csv_in):
    registration_df = pd.read_csv(csv_in, names=['Username', \
                                                'P_id', \
                                                'P_title', \
                                                'Type', \
                                                'Author_r'], \
                                  index_col=0, header=0)
    return registration_df

def merge_df(dict_in):
    merged_df = pd.concat(dict_in.values(), ignore_index=True, sort=False)
    return merged_df

def merge_regdf(paperdata_in, regdata_in):
    merged_df = pd.merge(paperdata_in, regdata_in, on='P_title', how='outer')
    merged_df = merged_df.fillna('N/A')
    merged_df['R_status'] = np.where(merged_df['Author_r'] != 'N/A', 'registered', 'not-registered')
    merged_df['P_status'] = np.where(merged_df['Tag'] != 'N/A', 'included', 'excluded')

    return merged_df

def create_tagdf_bytag(xml_dictin, merged_df):
    def _create_taglist(xml_dictin):
        tag_l = [tag_now for tag_now in xml_dictin.keys()]
        tag_l.append('other')
        return tag_l

    tag_list = _create_taglist(xml_dictin)
    merged_df.drop(['Type_y'], axis=1, inplace=True)

    dict_df = {}
    for tag_now in tag_list[:-1]:
        cond_tag = merged_df['Tag'] == tag_now
        cond_df = merged_df[cond_tag]
        cond_df = cond_df.groupby(['Type_x', \
                                   'P_title', \
                                   'Author', \
                                   'Email', \
                                   'Affiliation', \
                                   'Country', \
                                   'Tag', \
                                   'R_status', \
                                   'P_status'])['Author_r'].apply(lambda x: ', '.join(list(set(x)))).reset_index()
        dict_df[tag_now] = cond_df
        print(tag_now, 'has', len(cond_df), 'entries')
    dict_df[tag_list[-1]] = merged_df[merged_df['Tag'] == 'N/A']
    print('other has', len(dict_df[tag_list[-1]]), 'entries')

    return dict_df

def create_reg_summary(dict_in):
    full_df = merge_df(dict_in)
    temp_piv = pd.pivot_table(full_df, \
                              values='P_title', \
                              columns=['Tag'], \
                              index=['R_status'], \
                              aggfunc=len, \
                              fill_value=0)

    return pd.DataFrame(temp_piv)

def write_to_excel(fulldf_in, typepivot_in, regpivot_in, dict_df, unreg_in, filetosave):

    writer = pd.ExcelWriter(filetosave, \
                            engine='xlsxwriter', \
                            options={'strings_to_urls': False, \
                                     'strings_to_numbers':True})

    fulldf_in.to_excel(writer, sheet_name='FullList', index=False)
    typepivot_in.to_excel(writer, sheet_name='Summary_by_Types')
    regpivot_in.to_excel(writer, sheet_name='RegSummary_by_Tracks')
    unreg_in.to_excel(writer, sheet_name='Unregistered_Papers')

    for tag_now in dict_df:

        sname = tag_now+'_status'
        df_now = dict_df[tag_now]
        df_now.to_excel(writer, sheet_name=sname, index=False)
        workbook = writer.book
        # Add a format. Light red fill with dark red text.
        format1 = workbook.add_format({'bg_color': '#FFC7CE', \
                                       'font_color': '#9C0006'})
        # Add a format. Green fill with dark green text.
        format2 = workbook.add_format({'bg_color': '#C6EFCE', \
                                       'font_color': '#006100'})
        worksheet = writer.sheets[sname]
        r_status_idx = df_now.columns.get_loc('R_status')
        p_status_idx = df_now.columns.get_loc('P_status')
        # calculate cell_compute_start
        rcell_compute_start = xl_rowcol_to_cell(1, r_status_idx)
        pcell_compute_start = xl_rowcol_to_cell(1, p_status_idx)
        #calculate cell_compute_end
        rcell_compute_end = xl_rowcol_to_cell(len(df_now), r_status_idx)
        pcell_compute_end = xl_rowcol_to_cell(len(df_now), p_status_idx)
        worksheet.conditional_format(rcell_compute_start + ':' + rcell_compute_end, \
                                    {'type': 'cell', \
                                    'criteria': 'equal to', \
                                    'value':'"not-registered"', \
                                    'format': format1})
        worksheet.conditional_format(rcell_compute_start + ':' + rcell_compute_end, \
                                    {'type': 'cell', \
                                    'criteria': 'equal to', \
                                    'value':'"registered"', \
                                    'format': format2})
        worksheet.conditional_format(pcell_compute_start + ':' + pcell_compute_end, \
                                    {'type': 'cell', \
                                    'criteria': 'equal to', \
                                    'value':'"excluded"', \
                                    'format': format1})
        worksheet.conditional_format(pcell_compute_start + ':' + pcell_compute_end, \
                                    {'type': 'cell', \
                                    'criteria': 'equal to', \
                                    'value':'"included"', \
                                    'format': format2})

    writer.save()

if __name__ == "__main__":

    # Define the location of paper information from HotCRP
    xml_dict = read_xml_text(sys.argv[1])

    # Create a dictionary where each key stores a Pandas dataframe produced by
    # reading in the respective .xml file
    paper_dict = create_dictdf(xml_dict)

    # Create a dataframe for the registration information
    reg_df = create_googledf(sys.argv[2])

    # Merge the dataframes produced by the HotCRP xml files
    papermerged_df = merge_df(paper_dict)

    # -----------------
    # Some analyses
    # -----------------
    # Here we do some basic analyses to understand the distribution of paper types
    # for the different tracks at the conference
    piv_tagtype = pd.pivot_table(papermerged_df, \
                                index=['Tag'], \
                                columns="Type", \
                                values="P_title", \
                                aggfunc=len, fill_value=0)
    piv_tagtypedf = pd.DataFrame(piv_tagtype)

    # Here we do some analyses to find out who has registered. We first merge the
    # dataframe containing the registration information with the dataframe containing
    # the paper information
    paperreg_df = merge_regdf(papermerged_df, reg_df)
    # Then we create new dataframes based on the tracks, e.g. main
    paperreg_df_by_type = create_tagdf_bytag(xml_dict, paperreg_df)
    # Once we have created these dataframes, we compute a summary of the registration
    # status of papers at the different tracks
    paperreg_summarydf = create_reg_summary(paperreg_df_by_type)
    # We also want to create a dataframe containing the papers that have not yet
    # registered, regardless of tracks
    cond_unreg = paperreg_df['R_status'] == 'not-registered'
    paperreg_unregdf = paperreg_df[cond_unreg]

    # Finally we write our results to an Excel file
    write_to_excel(papermerged_df, piv_tagtypedf, paperreg_summarydf, \
                   paperreg_df_by_type, paperreg_unregdf, sys.argv[3])
