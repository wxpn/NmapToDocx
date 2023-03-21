# NmapToDocx

A Simple Python script to convert Nmap File to Microsoft Word Docx.

The following libraries need to be imported:

```
import sys
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Inches
```

The script accepts a Nmap XML file as input and creates a word document which includes table with columns ip address, hostname, port, service and version number.
