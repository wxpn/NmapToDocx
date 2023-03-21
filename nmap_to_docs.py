import sys
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Inches

def main():
    # check for correct usage
    if len(sys.argv) != 2:
        print("Usage: python nmap_to_word.py <nmap_xml_file>")
        sys.exit(1)
    
    # parse Nmap XML file using ElementTree
    tree = ET.parse(sys.argv[1])
    root = tree.getroot()
    
    # create a new Word document
    document = Document()
    
    # add a table to the document
    table = document.add_table(rows=1, cols=5)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'IP Address'
    hdr_cells[1].text = 'Hostname'
    hdr_cells[2].text = 'Port'
    hdr_cells[3].text = 'Service Name'
    hdr_cells[4].text = 'Version'
    
    # iterate over each host in the Nmap scan
    for host in root.iter('host'):
        ip_addr = host.find('address').attrib['addr']
        hostname = ''
        for hostnames in host.iter('hostname'):
            hostname = hostnames.attrib['name']
            break # Use the first hostname found
        # iterate over each port in the host
        for port in host.iter('port'):
            port_num = port.attrib['portid']
            service_name = port.find('service').attrib['name']
            version = ''
            if 'version' in port.find('service').attrib:
                version = port.find('service').attrib['version']
            # add a new row to the table with the host and port information
            row_cells = table.add_row().cells
            row_cells[0].text = ip_addr
            row_cells[1].text = hostname
            row_cells[2].text = port_num
            row_cells[3].text = service_name
            row_cells[4].text = version
            
    # save the Word document
    document.save('nmap_results.docx')
    print('Results saved to nmap_results.docx')

if __name__ == '__main__':
    main()
