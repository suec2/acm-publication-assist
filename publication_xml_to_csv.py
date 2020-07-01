import xml.etree.ElementTree as ET
import csv
import sys

"""
Description: this function converts the XML file available on the HotCRP website
(via 'Settings' -> 'ACM' -> 'Download TOC' button) into a .csv file containing:
- Paper Type
- Title
- Author names
- Author affiliation
- Lead author email
- Author email
The resulting .csv file has format consistent with ACM's requirement (as listed
in https://www.acm.org/publications/gi-proceedings), i.e.
"Paper Type","Title",
"Lead Author:Affiliation;Author2:Affiliation;Author3:Affiliation;etc.",
"Lead Author e-mail",”Author e-mail;Author e-mail"

Once the .csv file has been generated, save the .csv file with "UTF-8 encoding
with BOM" using Sublime Text to ensure that the special characters are encoded
correctly.

Syntax: python publication_xml_to_csv.py <path-to-xmlfile> <output-filename>
Example syntax: python publication_xml_to_csv.py data/sample-main-acmcms-toc.xml
                exported-data/sample-main-acmcms-toc.csv
"""

# The convert_xmldata function takes the xml data that has been parsed,
# processes the data and converts it into the format that is required by ACM
# (see above description for format).
def convert_xmldata(xml_in):

    def _flatten_list(list_in):
        flattened_string = ';'.join(list_in)

        return flattened_string

    def _separate_emails(list_in):
        l_email = list_in[0]
        o_email = ';'.join(list_in[1:])

        return l_email, o_email

    paper_list = []
    for member in xml_in.findall('paper'):
        paper_type = member.find('paper_type').text
        paper_title = member.find('paper_title').text
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
                a_affiliation = a_author.find('affiliation').text
                a_name_aff = a_name + ":" + a_affiliation
                author_names.append(a_name_aff)
                a_email = a_author.find('email_address').text
                author_emails.append(a_email)
                a_count += 1
        paper_authors = _flatten_list(author_names)
        l_email, paper_emails = _separate_emails(author_emails)
        paper_list.append([paper_type, paper_title, paper_authors, l_email, paper_emails])

    return paper_list

def write_to_csv(arraytowrite, filetosave):
    with open(filetosave, 'wt', encoding='utf-8') as csvfile:
        paperwriter = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
        for row_now in arraytowrite:
            paperwriter.writerow(row_now)

if __name__ == "__main__":
    # Define the xml data
    tree = ET.parse(sys.argv[1])

    # Create the root object
    xml_data = tree.getroot()

    # Process the XML xml_data
    conv_xmldata = convert_xmldata(xml_data)

    # Write the processed data to csv
    write_to_csv(arraytowrite=conv_xmldata, filetosave=sys.argv[2])
