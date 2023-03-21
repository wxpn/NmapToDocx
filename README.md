# NmapToDocx

A Simple Python script to convert Nmap File to Microsoft Word Docx.

The followinb libraries need to be imported:

```
import argparse
   import xml.etree.ElementTree as ET
   from docx import Document
   from docx.shared import Inches
```

The script accepts a Nmap XML file was input and creates a word document which includes fields like ip address, hostname, port, service and version number.
