import os
from psd_tools import PSDImage
from openpyxl import load_workbook
import numpy as np
import json

path_split = os.getcwd().split(os.path.sep)
del path_split[0]
path_jsx = '/'+'/'.join(path_split)

path_dir = './'
file_list = os.listdir(path_dir)


webtoon_height = 0
psd_height_list = []
psd_end_point_psd = []

excel_imgs_height = 0
excel_height_list = []

path_jsx_json=[]
psd_name_list = []
for path in (file_list):
    if path.startswith('~$'):
        continue

    file_path = path_jsx + '/' + path

    if os.path.splitext(path)[1] == '.xlsx':
        excel = load_workbook(file_path, data_only=True)
        excel_sheet = excel.active
        web_imgs_excel = excel_sheet._images
        for index, img in enumerate(web_imgs_excel):
            excel_height_list.append(img.height)
            excel_imgs_height += img.height
    # PSD 및 PSB 파일
    if os.path.splitext(path)[1] == '.psd' or os.path.splitext(path)[1] =='.psb':
        psd_name_list.append(path)
        path_jsx_json.append(file_path)
        psd = PSDImage.open(path)
        psd_height_list.append(psd.height)
        webtoon_height += psd.height
        psd_end_point_psd.append(webtoon_height)


webtoon_width = psd.width
excel2psd = webtoon_height / excel_imgs_height
psd2excel = excel_imgs_height / webtoon_height

end_row = web_imgs_excel[-1].anchor._from.row
final = np.around((web_imgs_excel[1].anchor._from.row/web_imgs_excel[0].height) *web_imgs_excel[-1].height)
final_end_row = end_row + final

psd_row_list_b = np.array(psd_height_list) * (final_end_row / webtoon_height)
psd_row_list = np.around(psd_row_list_b)
every_ratio_row2psd = np.array(psd_height_list)/ np.array(psd_row_list_b)


psd_num = len(psd_name_list)
max_col = excel_sheet.max_column
min_col = excel_sheet.min_column
max_row = excel_sheet.max_row


cell_count_per_psd_list = psd_row_list
with open('jsonFile.json', 'w', encoding='UTF-8') as w:
    json_file = {'psdPath' : path_jsx_json, 'psdName' : psd_name_list}

    for psd_name in psd_name_list:
        json_file[psd_name] = []

    count = 0
    boundary = 0
    json_psd = []


    for r in range(1,max_row+1):
        psd_x = webtoon_width//4
        count += 1
        if count > cell_count_per_psd_list[boundary]:
            json_file[psd_name_list[boundary]] = json_psd
            count = 0
            boundary += 1
            json_file[psd_name_list[boundary]]
            json_psd = []


        for c in range(min_col, max_col + 1):
            context = excel_sheet.cell(row=r, column=c).value

            if context == None:
                continue

            if '\n' in context :
                context = context.replace('\n','\r')

            psd_y = count * every_ratio_row2psd[boundary]
            line = [['text', context], ['x', round(psd_x, 2)], ['y', round(psd_y, 2)]]
            json_psd.append(dict(line))
            psd_x += webtoon_width // 2
    print(json_psd[2]['text']) # 두줄이상인거 찾기
    json_file[psd_name_list[boundary]] = json_psd

    json.dump(json_file, w, indent=4, ensure_ascii=False)

    print("json파일 생성완료")

