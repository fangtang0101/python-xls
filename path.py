# -*- coding: utf-8 -*-
import os
import shutil
import xlrd
import string

### 创建多层目录
def mkdirs(path):
    # 去除首位空格
    path = path.strip()
    # 去除尾部 \ 符号
    path = path.rstrip("\\")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False
    isExists = os.path.exists(path)

    # 判断结果
    if not isExists:
        # 创建目录操作函数
        os.makedirs(path)
        # 如果不存在则创建目录
        print(path + ' 创建成功')
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        shutil.rmtree(path_target) # 如果存在，先删除
        print(path + ' 目录已存在')
        os.makedirs(path)
        # 如果不存在则创建目录
        print(path + ' 创建成功')
        return False

def search_file(path, newpath):
	mkdirs(newpath)
	if os.path.isfile(path):
		name = os.path.basename(path)  # 获取文件名
		dirname = os.path.dirname(path)  # 获取文件目录
		print("fime-name:",name,"file-path",dirname)
		full_path = os.path.join(dirname, name)  # 将文件名与文件目录连接起来，形成完整路径
		des_path = newpath  #目标路径，将该文件夹信息添加进最后的文件名中
		print("full_path:",full_path,"des_path",des_path)
		# shutil.move(full_path, des_path)#移动文件到目标路径（移动+重命名）
		shutil.copy(full_path, newpath) 
		# 移动到新的文件里面
	else :
		print("错误 path 路径不是文件",path)

path_source = "./source/模版.xls"
path_target = "./target"

search_file(path_source,path_target)

