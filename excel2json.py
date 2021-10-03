import os
from psd_tools import PSDImage
import openpyxl
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
        # print(excel_sheet._images[-1].anchor._from)

        # 이게 문제. 맨 마지막 img를 읽지를 못함 맞다... 얘네 시작점밖에 못찾지...

        for index, img in enumerate(web_imgs_excel):
            excel_height_list.append(img.height)
            excel_imgs_height += img.height

    if os.path.splitext(path)[1] == '.psd':
        psd_name_list.append(path)
        path_jsx_json.append(file_path)
        psd = PSDImage.open(path)
        psd_height_list.append(psd.height)
        webtoon_height += psd.height
        psd_end_point_psd.append(webtoon_height)


# TODO : 전체 사이즈
# print(webtoon_height)
webtoon_width = psd.width
# print(excel_imgs_height)



# TODO : 각 psd당 height
# print(psd_height_list)
# [11007, 12000, 11397, 10891, 11115, 8482]
# print(excel_height_list)
# [1000, 1000, 1000, 1000, ...


# TODO : ratio
excel2psd = webtoon_height / excel_imgs_height
psd2excel = excel_imgs_height / webtoon_height


# TODO : start point
# print(psd_end_point_psd)
psd_end_point_excel = np.array(psd_end_point_psd)*psd2excel
# print(psd_end_point_excel)
# [11007, 23007, 34404, 45295, 56410, 64892]
print(psd_end_point_excel)



point = 0
i = 0
excel_point_img = []
# 작을 때는 어떻게 할건데....[-1]


# while이 나을까
for index, height in enumerate(excel_height_list) :
    point += height
    # print(point)
    if point >= psd_end_point_excel[i] :
        print(point)
        i += 1
        # print(index)
        excel_point_img.append(index)

# print(excel_point_img)

# excel height가 더 작을 경우
# if len(excel_point_img) != len(psd_end_point_excel):
#     excel_point_img.append(psd_end_point_excel[-1])


 # 애초에 마지막은 예외 처리해줘야되는거, anchor가 시작점밖에 못구하는데
 # 이걸로 끝점 구하려고 했으면 마지막 이미지는 range 넘어가니까 그냥 예외 처리했어야지..
 # 하 머리가 이렇게 안좋아서 어떻게하지...? 기억력이라는게 없나.. 내가짠걸내가 까먹으면 ... 노답이네

excel_psd_start_imgs_index = []
print(psd_end_point_excel)
# [11007, 23007, 34404, 45295, 56410, 64892]

for ind, point in enumerate(psd_end_point_excel[:-1]):
    excel_psd = excel_height_list[:excel_point_img[ind]+1]

    ans_1 = abs(point - sum(excel_psd))
    ans_2 = abs(point - sum(excel_psd[:-1]))


    ans = min(ans_1,ans_2)  # End img
    # ans 가 end img
    if ans == ans_1 :
        print(ind, "ind")
        print(point)
        print(sum(excel_psd))
        print(ans_1)
        print(ans_2)
        excel_psd_start_imgs_index.append(excel_point_img[ind]+1)
    else:
        excel_psd_start_imgs_index.append(excel_point_img[ind])



start_cell_psd = [1]   # 1 or 0

#anchor가 시작점 밖에 못 찾기 때문에 다음 이미지의 시작점 -1하면 끝점
for index in excel_psd_start_imgs_index:
    start_img_per_psd = web_imgs_excel[index]
    start_cell_psd.append(start_img_per_psd.anchor._from.row+1)
    # print(start_cell_psd.append(start_img_per_psd.ext))
    # OneCellAnchor(_from=marker, ext=size)


psd_num = len(psd_name_list)

max_col = excel_sheet.max_column
min_col = excel_sheet.min_column
max_row = excel_sheet.max_row #맞다 이거 마지막 못찾지.....

# end_cell_psd.append(max_row)
# print(end_cell_psd)

# start_cell_psd.append(int(psd_end_point_excel[-1])+1)  # 마지막 psd의 cellcount를 위해
print(start_cell_psd)

cell_count_per_psd_list = []

for i in range(psd_num-1):
    cell_count = start_cell_psd[i+1] - start_cell_psd[i]
    cell_count_per_psd_list.append(cell_count)

print()
print(psd_end_point_psd)
print(psd_end_point_psd[-1]-psd_end_point_psd[-2])
print(cell_count_per_psd_list[0]/psd_end_point_excel[0])
final_psd = (cell_count_per_psd_list[0]/psd_end_point_excel[0])*(psd_end_point_psd[-1]-psd_end_point_psd[-2])
cell_count_per_psd_list.append(int(final_psd))
print()

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
        #
        # print(cell_count_per_psd_list,"몇갠데이래")
        # print(count)
        if count > cell_count_per_psd_list[boundary]:
            # print(psd_name_list)
            # print(boundary)
            json_file[psd_name_list[boundary]] = json_psd
            count = 0
            boundary += 1
            # print(json_file)
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

    print(json_file)

    json.dump(json_file, w, indent=4, ensure_ascii=False)

    print("json파일 생성완료")

