#! /opt/local/bin/python
# coding=utf-8
from docx import Document
import re

doc = Document('doc_20200904.docx')
i = 0
for x in doc.paragraphs:
    if( x.style.name == 'Heading 2'):
        print( i)
        print(x.style.name)
        print(x.text)
    if( x.style.name == 'Heading 3' and x.text[:13] == 'Rule_E_Score_'):
        print( i)
        print(x.style.name)
        print(x.text)
    if( x.style.name == 'Heading 4' and x.text[:13] == 'Rule_E_Score_'):
        print( i)
        print(x.style.name)
        print(x.text)
    if( x.style.name == 'Heading 1'):
        print( i)
        print(x.style.name)
        print(x.text)
    i += 1
