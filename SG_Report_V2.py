import zipfile
from pathlib import Path
import os
import logging
from PIL import Image
# import pylightxl as xl
import xlwings as xw
# import matplotlib.pyplot as plt
# import cv2
# import time
# import PySimpleGUI as sg
from shutil import copyfile

def copyFile(sourcePath ,targetPath):#copy raw data file
    cpcmd="cp -r {0} {1}".format(sourcePath,targetPath)
    print("copy cmd:",cpcmd)
    os.system(cpcmd)
def checkFile(files):#check and move .DS_Store 
    allfilePaths=files
    for path in files:
        if path==".DS_Store":
            allfilePaths.remove('.DS_Store')
    return allfilePaths
# Paths=[{'rawFileDirectory':"","reportfilePath":""}]
filePaths = []
def checkAndCreateDirectory(paths):
    if(os.path.exists(paths)):
        print("check path pass path:",paths)
    else:
        os.makedirs(paths)
        checkAndCreateDirectory(paths)

def getFullpath(fixtureNames):
    filePaths=[]
    for i, fixtureId in enumerate(fixtureNames):
        buff = {}
        buff["fixtureid"] = fixtureId
        rawDataFileSandboxDirectory = os.path.join(rawDataSandboxPath, fixtureId)  # 完整路径
        tempReportFilePath = os.path.join(tempReportPath, fixtureId+".xlsx")
        buff["rawFileDirectory"] = rawDataFileSandboxDirectory
        buff["tempReportFilePath"] = tempReportFilePath
        buff["reportPath"]=os.path.join(reportDirectoryPath,fixtureId+".xlsx")
        print(buff)
        filePaths.append(buff)
    return filePaths



# for i, fixtureID in enumerate(fixtureNamePaths):
#     buff = {}
#     completePath = os.path.join(rawDataPath, fixtureID)  # 完整路径
#     reportPath = os.path.join(reportDirectoryPath, fixtureID+".xlsx")
#     buff["rawFileDirectory"] = completePath
#     buff["reportfilePath"] = reportPath
#     buff["FixtureID"] = fixtureID
#     filePaths.append(buff)


def read_img(zipfile_path):
    # if not isfile_exist(zipfile_path):
    #     return False
    try:
        dir_path = os.path.dirname(zipfile_path)  # 获取文件所在目录
        file_name = os.path.basename(zipfile_path)  # 获取文件名
        pic_dir = 'xl' + os.sep + 'media'  # excel变成压缩包后，再解压，图片在media目录
        pic_path = os.path.join(dir_path, str(
            file_name.split('.')[0]), pic_dir)
        imgefilepath = ""
        file_list = os.listdir(pic_path)
        for file in file_list:
            imgefilepath = os.path.join(pic_path, file)
            # print(filepath)
        print(imgefilepath)
        return imgefilepath
    except IOError as e:
       print("read_img")
       exit(1)
    except:
       print("read_img")
       exit(1)

def unzip_file(zipfile_path):
    if os.path.splitext(zipfile_path)[1] != '.zip':
        print("It's not a zip file! %s" % zipfile_path)
        return False

    file_zip = zipfile.ZipFile(zipfile_path, 'r')
    file_name = os.path.basename(zipfile_path)  # 获取文件名
    zipdir = os.path.join(os.path.dirname(zipfile_path),
                          str(file_name.split('.')[0]))  # 获取文件所在目录

    if os.path.exists(zipdir) == False:
        os.mkdir(zipdir)
    print(file_name)
    for files in file_zip.namelist():
        file_zip.extract(files, zipdir)  # 解压到指定文件目录
    file_zip.close()
    return zipdir


# 修改指定目录下的文件类型名，将excel后缀名修改为.zip
def change_file_name(sourcepath, new_type='.zip'):
    backupath = os.path.join(os.getcwd(), "backup")
    checkAndCreateDirectory(backupath)
    file_name = os.path.basename(sourcepath)  # 获取文件名
    backupath_path = os.path.join(backupath, file_name)
    print(backupath_path)
    copyfile(sourcepath, backupath_path)
    extend = os.path.splitext(backupath_path)[1]  # 获取文件拓展名
    if extend != '.xlsx' and extend != '.xls':
        print("It's not a excel file! %s" % backupath_path)
        return False
    new_name = str(file_name.split('.')[0]) + new_type  # 新的文件名，命名为：xxx.zip
    dir_path = os.path.dirname(backupath_path)  # 获取文件所在目录
    new_path = os.path.join(dir_path, new_name)  # 新的文件路径
    if os.path.exists(new_path):
        os.remove(new_path)
    print(backupath_path)
    os.rename(backupath_path, new_path)  # 保存新文件，旧文件会替换掉
    return new_path  # 返回新的文件路径，压缩包


def writebuff(filePaths):
    print("Paths_buff", filePaths)
    startRowIndex = 25
    count=range(len(filePaths))
    for fileindex, filePath in enumerate(filePaths) :
        # sg.one_line_progress_meter('实时进度条', fileindex + 1, len(count), '-文件转换进度-')
        # app=xw.App(visible=False,add_book=False)
        templatewb = xw.Book(templateFileSandboxPath)
        templatesheet = templatewb.sheets["Sheet1"]
        rawFileDirectory = filePath["rawFileDirectory"]
        tempReportFilePath = filePath["tempReportFilePath"]
        templatesheet.range("B4").value = filePath["fixtureid"]
        reportPath=filePath["reportPath"]
        picturesindex = 0
        for index1 in range(2):
            for index2 in range(6):
                rawdatafilepath = os.path.join(
                    rawFileDirectory, "{0}-{1}.xlsx".format(index1+1, index2+1))
                print(rawdatafilepath)
                buffWb = xw.Book(rawdatafilepath)
                buffsheet = buffWb.sheets["Sheet1"]
                picturesindex = index1 == 0 and index2 or index1*6+index2
                for index in range(8):
                    rowindex = 17+index*7
                    rawDataMaxIndexKey = "C"+str(rowindex)
                    rawDataMinIndexKey = "D"+str(rowindex)
                    reportMaxIndexKey = "C" + \
                        str(startRowIndex*(index2+1) + index1*8+index)
                    reportMinIndexKey = "D" + \
                        str(startRowIndex*(index2+1) + index1*8+index)
                    print("Key1:{0} Key2{1}".format(reportMaxIndexKey,reportMinIndexKey))
                    templatesheet.range(reportMaxIndexKey).value = buffsheet.range(
                        rawDataMaxIndexKey).value
                    templatesheet.range(reportMinIndexKey).value = buffsheet.range(
                        rawDataMinIndexKey).value
                    print("MAX:{0} MIn:{1}".format(buffsheet.range(rawDataMaxIndexKey).value,buffsheet.range(rawDataMinIndexKey).value))
                buffWb.close()
                imgpath = read_img(unzip_file(
                    change_file_name(rawdatafilepath)))
                # img=cv2.imread(imgpath)
                # img = cv2.resize(img, (800, 200))
                # cv2.imshow(imgpath,img)
                # cv2.waitKey(300)
                # cv2.destroyAllWindows()
                picturesindexKey = "F"+str(22+index2*startRowIndex + index1*10)
                templatesheet.pictures.add(imgpath, left=templatesheet.range(
                    picturesindexKey).left, top=templatesheet.range(picturesindexKey).top)
        print(templatesheet.pictures)
        templatewb.save(tempReportFilePath)
        templatewb.close()
       
        os.system("mv {0} {1}".format(tempReportFilePath,reportPath))

        


print("user path",Path.home())  # User Path
reportDirectoryPath = os.path.join(os.getcwd(), "report/")
checkAndCreateDirectory(reportDirectoryPath)
# Sandbox path
excelSandboxPath = "Library/Containers/com.microsoft.Excel/Data/" 
# Excel full sandbox path
excelSandboxPath = os.path.join(Path.home(), excelSandboxPath)  
print("excel SandBox directory:",excelSandboxPath)
# template file Sandbox Path
templateFileSandboxPath = os.path.join(excelSandboxPath, "template/")
print("template SandBox directory:",templateFileSandboxPath)
checkAndCreateDirectory(templateFileSandboxPath)
#Template file path
templateFileSandboxPath=os.path.join(templateFileSandboxPath,"usiTemplate.xlsx")
print("template SandBox path:",templateFileSandboxPath)
templateFilePath=os.path.join(os.getcwd(),"template/usiTemplate.xlsx")
print("template file path:",templateFilePath)
rawDataSandboxPath = os.path.join(excelSandboxPath, "RawData/")  # 
rawDataPath=os.path.join(os.getcwd(),"data/*")
fixtureNamePaths =checkFile(os.listdir(rawDataSandboxPath))   # 获取下面的文件路径
tempReportPath=os.path.join(excelSandboxPath,"tempReport/")
checkAndCreateDirectory(tempReportPath)#   报告临时存放路径
os.system("open {0}".format(tempReportPath))
print(tempReportPath)
print(fixtureNamePaths)
filePaths=getFullpath(fixtureNamePaths)
writebuff(filePaths)

# for i, fixtureID in enumerate(fixtureNamePaths):
#     buff = {}
#     completePath = os.path.join(rawDataPath, fixtureID)  # 完整路径
#     reportPath = os.path.join(reportDirectoryPath, fixtureID+".xlsx")
#     buff["rawFileDirectory"] = completePath
#     buff["reportfilePath"] = reportPath
#     buff["FixtureID"] = fixtureID
#     filePaths.append(buff)

# print(filePaths)



# templatePath="./template/usiTemplate.xlsx"
# savePath=os.path.join(os.getcwd(),"report/debug_1.xlsx")
# imgpath=os.path.join(os.getcwd(),"PointImage/Image1.png")
# templatewb=xw.Book(templatePath)
# templatesheet=templatewb.sheets["Sheet1"]
# BuffPaths1=["./data/CYG DC80006-S02/1-1.xlsx",
#            "./data/CYG DC80005-S02/1-2.xlsx",
#            "./data/CYG DC80005-S02/1-3.xlsx",
#            "./data/CYG DC80005-S02/1-4.xlsx",
#            "./data/CYG DC80005-S02/1-5.xlsx",
#            "./data/CYG DC80005-S02/1-6.xlsx"
#            ]
# BuffPaths2=[
#            "./data/CYG DC80005-S02/2-1.xlsx",
#            "./data/CYG DC80005-S02/2-2.xlsx",
#            "./data/CYG DC80005-S02/2-3.xlsx",
#            "./data/CYG DC80005-S02/2-4.xlsx",
#            "./data/CYG DC80005-S02/2-5.xlsx",
#            "./data/CYG DC80005-S02/2-6.xlsx"
#            ]
# reportStatRows=[25,50,75,100,125,150]
# for i,path in enumerate(BuffPaths1):
#     buffWb=xw.Book(path)
#     buffsheet=buffWb.sheets["Sheet1"]
#     sensor_list=[1,2]
#     # imge=buffsheet.pictures
#     # print(imge)

#     for index in range(8):
#         rowindex=17+index*7
#         rawDataMaxIndexKey="C"+str(rowindex)
#         rawDataMinIndexKey="D"+str(rowindex)
#         reportMaxIndexKey="C"+str(index+reportStatRows[i])
#         reportMinIndexKey="D"+str(index+reportStatRows[i])
#         templatesheet.range(reportMaxIndexKey).value=buffsheet.range(rawDataMaxIndexKey).value
#         templatesheet.range(reportMinIndexKey).value=buffsheet.range(rawDataMinIndexKey).value
#     templatesheet.range("A137").value=buffsheet.range("L100").value


#     buffWb.close()

# for i,path in enumerate(BuffPaths2):
#     buffWb=xw.Book(path)
#     buffsheet=buffWb.sheets["Sheet1"]
#     sensor_list=[1,2]

#     for index in range(8):
#         rowindex=17+index*7
#         rawDataMaxIndexKey="C"+str(rowindex)
#         rawDataMinIndexKey="D"+str(rowindex)
#         reportMaxIndexKey="C"+str(index+reportStatRows[i]+8)
#         reportMinIndexKey="D"+str(index+reportStatRows[i]+8)
#         templatesheet.range(reportMaxIndexKey).value=buffsheet.range(rawDataMaxIndexKey).value
#         templatesheet.range(reportMinIndexKey).value=buffsheet.range(rawDataMinIndexKey).value
#     buffWb.close()

# templatewb.save(savePath)
# templatewb.close()
