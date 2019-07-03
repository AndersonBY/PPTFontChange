# -*- coding: utf-8 -*-
# @Author: Anderson
# @Date:   2019-07-03 15:36:58
# @Last Modified by:   Anderson
# @Last Modified time: 2019-07-03 18:00:30
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import glob
import os


def set_text_frame_font(text_frame):
	for paragraph in text_frame.paragraphs:
		for run in paragraph.runs:
			if run.font.name in fonts_to_be_replaced:
				run.font.name = fonts_to_be_replaced[run.font.name]


def check_shape(shape):
	if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
		for shape_in_group in shape.shapes:
			check_shape(shape_in_group)
	elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
		for cell in shape.table.iter_cells():
			text_frame = cell.text_frame
			set_text_frame_font(text_frame)
	else:
		if shape.has_text_frame:
			text_frame = shape.text_frame
			set_text_frame_font(text_frame)


fonts_to_be_replaced = {
	'微软雅黑': '思源黑体',
	'Microsoft YaHei': '思源黑体',
	'等线': '思源黑体'
}

for file in glob.glob('input/*.pptx'):
	print(f'Processing file: {file}')
	prs = Presentation(file)
	for index, slide in enumerate(prs.slides):
		for shape in slide.shapes:
			check_shape(shape)

	prs.save(f'output/{os.path.basename(file)}')
