# !/usr/bin/env python
# -*- coding:utf-8 -*-

"""
拷贝多个.docx文件指定内容到目标.docx文件中

"""

import docx
import os

#获取所有文件路径
def walk_dir(path):
	file_path = []
	for root, dirs, files in os.walk(path):
		for f in files:
			if f.lower().endswith('docx'):
				file_path.append(os.path.join(root, f))

	return file_path


#提取传入.docx文件有用信息，并追加到新.docx文件中。
def transition_context(file_path, destination_document):
	doc_text = []
	num_paragraph = 0


	source_document = docx.Document(file_path)

	#doc_text = [paragraph.text for paragraph in source_document.paragraphs]

	for paragraph in source_document.paragraphs:
		if paragraph.text[0:9] == 'Keywords:':
			break
		else:
			doc_text.append(paragraph.text)
			num_paragraph +=1


#	destination_document.add_heading('Document Test', 0)
	for i in range(num_paragraph):
		if (i >= 2) and (i%2==0):
			continue
		if (i == 0):
			destination_document.add_paragraph(doc_text[i], style='ListNumber')
		else:
			destination_document.add_paragraph(doc_text[i])

	



if __name__ == "__main__":
	i = 0
	destination_document = docx.Document()

	for file_path in walk_dir(os.getcwd()):
		transition_context(file_path, destination_document)
		i += 1
	print i
	destination_document.save('demo1.docx')