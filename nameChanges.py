import shutil
import os


def copydirs(fromFile):
    files = os.listdir(fromFile)  # 获取文件夹中文件和目录列表
    for f in files:
        if os.path.isdir(fromFile + '/' + f):  # 判断是否是文件夹
            copydirs(fromFile + '/' + f)  # 递归调用本函数
            if f.
        else:
            path = f
            if f != '.DS_Store':
                while os.path.exists(fromFile + '/' + path):
                    path = fromFile + '/' + path
                    os.rename()
                shutil.copy(fromFile + '/' + f)  # 拷贝文件


if __name__ == '__main__':
    from_file = ""
    if not os.path.exists(from_file):  # 如不存在目标目录则创建
        os.makedirs(from_file)
    copydirs(from_file)
