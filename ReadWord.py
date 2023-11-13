#!/usr/bin/env python
# -*- coding: utf-8 -*-

from docx import Document
from lxml import etree


# 打开Word文件
doc = Document('C:\\Users\\Administrator\\Desktop\\3842562013440940385fenmh(1).docx')

# 创建XML根元素
root = etree.Element("document")

# 遍历段落并将其添加为XML元素
for paragraph in doc.paragraphs:
    para_element = etree.SubElement(root, "paragraph")
    para_element.text = paragraph.text

# 将XML根元素转换为字符串
xml_str = etree.tostring(root, pretty_print=True, encoding='utf-8')

# 将XML字符串写入文件
with open('C:\\Users\\Administrator\\Desktop\\3842562013440940385fenmh(1).xml', 'wb') as xml_file:
    xml_file.write(xml_str)

# 关闭Word文档
# doc.close()

