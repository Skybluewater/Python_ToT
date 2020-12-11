# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
#
# file = open("")
#
# for i in range(1024):
#     if (i / 32) % 2 == 0:
#         print("0x00", end=",")
#     else:
#         print("0x11", end=",")

import os
import cv2
import numpy as np


def rgb2gray(rgb):
    return np.dot(rgb[..., :3], [0.2989, 0.5870, 0.1140])


def resize_img(DATADIR, img_size):
    w = img_size[0]
    h = img_size[1]
    '''设置目标像素大小，此处设为300'''
    path = os.path.join(DATADIR)
    # 返回path路径下所有文件的名字，以及文件夹的名字，
    img_list = os.listdir(path)

    for i in img_list:
        if i.endswith('.jpg'):
            # 调用cv2.imread读入图片，读入格式为IMREAD_COLOR
            img_array = cv2.imread((path + '/' + i), cv2.IMREAD_COLOR)
            # 调用cv2.resize函数resize图片
            new_array = cv2.resize(img_array, (w, h), interpolation=cv2.INTER_CUBIC)
            new_array = rgb2gray(new_array)
            '''生成图片存储的目标路径'''
            save_path = path + '/_new' + str(i)
            # print(new_array.shape)
            '''调用cv.2的imwrite函数保存图片'''
            count = 0
            sum = 0
            temp = 0
            total = 0
            string_out = ""
            for i in range(new_array.shape[0]):
                for j in range(new_array.shape[1]):
                    if new_array[i][j] >= 127:
                        sum <<= 1
                        sum += 1
                    count += 1
                    if count == 4:
                        temp = sum
                        sum = 0
                    if count == 8:
                        count = 0
                        sum2 = sum
                        sum += temp << 4
                        print("0x%X%X," % (temp, sum2), end="")
                        string_out += str(hex(sum)) + ","
                        sum = 0
                        temp = 0
                        total += 1
            cv2.imwrite(save_path, new_array)
            object = open("file.txt", "w")
            object.write(string_out)
            print("")
            print(total)


if __name__ == '__main__':
    # 设置图片路径
    DATADIR = "./"
    # 需要修改的新的尺寸
    img_size = [128, 64]
    resize_img(DATADIR, img_size)
