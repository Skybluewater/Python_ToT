import shutil
import os


def copydirs(fromFile, toFile):
    files = os.listdir(fromFile)  # 获取文件夹中文件和目录列表
    for f in files:
        if os.path.isdir(fromFile + '/' + f):  # 判断是否是文件夹
            copydirs(fromFile + '/' + f, toFile)  # 递归调用本函数
        else:
            path = f
            if f != '.DS_Store':
                while os.path.exists(toFile + '/' + path):
                    path = path.split('.')[0] + '-1' + '.' + path.split('.')[-1]
                shutil.copy(fromFile + '/' + f, toFile + '/' + path)  # 拷贝文件


if __name__ == '__main__':
    from_file = ""
    to_file = ""
    if not os.path.exists(to_file):  # 如不存在目标目录则创建
        os.makedirs(to_file)
    copydirs(from_file, to_file)
