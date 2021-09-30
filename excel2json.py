import os
from psd_tools import PSDImage
from openpyxl import load_workbook
import json

path_split = os.getcwd().split(os.path.sep)
del path_split[0]
path_jsx = '/'+'/'.join(path_split)

path_dir = './'
file_list = os.listdir(path_dir)


webtoon_height = 0
psd_height_list = []
webtoon_start_point_psd = []

excel_imgs_height = 0
excel_height_list = []
excel_psd_end_imgs = []
excel_psd_end_imgs_index = [0]

path_jsx_json=[]
psd_name_list = []
for path in (file_list):
    if path.startswith('~$'):
        continue

    file_path = path_jsx + '/' + path

    if os.path.splitext(path)[1] == '.psd':
        psd_name_list.append(path)
        path_jsx_json.append(file_path)
        psd = PSDImage.open(path)
        psd_height_list.append(psd.height)
        print(psd.height,"psd.height")
        webtoon_height += psd.height
        webtoon_start_point_psd.append(webtoon_height)

    if os.path.splitext(path)[1] == '.xlsx':
        excel = load_workbook(file_path, data_only=True)
        excel_sheet = excel.active
        web_imgs_excel = excel_sheet._images[:]

        standard_height = web_imgs_excel[0].height

        for index, img in enumerate(web_imgs_excel):
            print(img.height,"img.height")
            excel_height_list.append(img.height)
            excel_imgs_height += img.height


            if img.height != standard_height:
                excel_psd_end_imgs.append(img.height)
                excel_psd_end_imgs_index.append(index+1)
    else :
        continue


webtoon_width = psd.width

excel2psd = webtoon_height / excel_imgs_height
psd2excel = excel_imgs_height / webtoon_height
# print(webtoon_height,"webtoon_height")
# print(excel_imgs_height, "excel_imgs_height")
# print(webtoon_height,"webtoon_height")
# print(psd_height_list[0]+1,psd_height_list[0]+psd_height_list[1])
print(psd_height_list, "psd_height_list")
print(psd2excel,"psd2excel") #엥?? 이게1이네.. 00
a = [] #엑셀로 변환
for i in psd_height_list:
    a.append(i*psd2excel)

print(a, "excel height로 전환 ... point")


print(excel_psd_end_imgs_index,"excel_psd_end_imgs_index")
print(len(excel_height_list),"excel_height_list")

excel_psd_height_list = []
# for i in range(len(psd_name_list)):
#     excel_psd = excel_height_list[excel_psd_end_imgs_index[i]:excel_psd_end_imgs_index[i + 1]]
#     excel_psd_height_list.append(sum(excel_psd))

start_psd = []
excel_height_list_psd = []

for index in excel_psd_end_imgs_index[1:-1]:
    end_img_per_psd = web_imgs_excel[index]
    end_of_psd = end_img_per_psd.anchor._from.row
    start_psd.append(end_of_psd+1)

psd_num = len(start_psd)

max_col = excel_sheet.max_column
min_col = excel_sheet.min_column
max_row = excel_sheet.max_row

start_psd.append(max_row)

cell_count_per_psd_list = [start_psd[0]]
for i in range(1,len(start_psd)):
    cell_count = start_psd[i] - start_psd[i-1]
    cell_count_per_psd_list.append(cell_count)


print(cell_count_per_psd_list,"cell_count_per_psd_list")
every_ratio_excel2psd = []
for i in range(len(cell_count_per_psd_list)):
    ratio = psd_height_list[i] / cell_count_per_psd_list[i]
    every_ratio_excel2psd.append(ratio)


with open('jsonFile.json', 'w', encoding='UTF-8') as w:
    json_file = {'psdPath' : path_jsx_json, 'psdName' : psd_name_list}

    for psd_name in psd_name_list:
        json_file[psd_name] = []

    count = 0
    boundary = 0
    json_psd = []


    for r in range(1,max_row):
        psd_x = webtoon_width//4
        count += 1


        if count >= cell_count_per_psd_list[boundary]:
            json_file[psd_name_list[boundary]] = json_psd
            count = 0
            boundary += 1
            json_file[psd_name_list[boundary]]
            json_psd = []


        for c in range(min_col, max_col + 1):
            context = excel_sheet.cell(row=r, column=c).value

            if context == None:
                continue

            psd_y = count * every_ratio_excel2psd[boundary]

            line = [['text', context], ['x', round(psd_x, 2)], ['y', round(psd_y, 2)]]
            json_psd.append(dict(line))
            psd_x += webtoon_width // 2

    json_file[psd_name_list[boundary]] = json_psd

    json.dump(json_file, w, indent=4, ensure_ascii=False)

    print("json파일 생성완료")

