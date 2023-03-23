import zipfile
from pathlib import Path
import os
import logging
# from PIL import Image
# import pylightxl as xl
import xlwings as xw
# import matplotlib.pyplot as plt
# import cv2
# import time
# import PySimpleGUI as sg
from shutil import copyfile
print("main")
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
        rawDataFileSandboxDirectory = os.path.join(rawDataSandboxPath, fixtureId)  # 
        tempReportFilePath = os.path.join(tempReportPath, fixtureId+".xlsx")
        buff["rawFileDirectory"] = rawDataFileSandboxDirectory
        buff["tempReportFilePath"] = tempReportFilePath
        buff["reportPath"]=os.path.join(reportDirectoryPath,fixtureId+".xlsx")
        print(buff)
        filePaths.append(buff)
    return filePaths

def moveSamdboxRawFile(dirpath):
    print('moveSamdboxRawFile')
    cmd="rm -rf {0}".format(dirpath)
    print(cmd)
    os.system(cmd)

# for i, fixtureID in enumerate(fixtureNamePaths):
#     buff = {}
#     completePath = os.path.join(rawDataPath, fixtureID)  # 
#     reportPath = os.path.join(reportDirectoryPath, fixtureID+".xlsx")
#     buff["rawFileDirectory"] = completePath
#     buff["reportfilePath"] = reportPath
#     buff["FixtureID"] = fixtureID
#     filePaths.append(buff)


def read_img(zipfile_path):
    # if not isfile_exist(zipfile_path):
    #     return False
    try:
        dir_path = os.path.dirname(zipfile_path)  # 
        file_name = os.path.basename(zipfile_path)  # 
        pic_dir = 'xl' + os.sep + 'media'  # excel
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
    file_name = os.path.basename(zipfile_path)  # 
    zipdir = os.path.join(os.path.dirname(zipfile_path),
                          str(file_name.split('.')[0]))  # 

    if os.path.exists(zipdir) == False:
        os.mkdir(zipdir)
    print(file_name)
    for files in file_zip.namelist():
        file_zip.extract(files, zipdir)  # 解压到指定文件目录
    file_zip.close()
    return zipdir


# .zip
def change_file_name(sourcepath, new_type='.zip'):
    backupath = os.path.join(os.getcwd(), "backup")
    checkAndCreateDirectory(backupath)
    file_name = os.path.basename(sourcepath)  # 
    backupath_path = os.path.join(backupath, file_name)
    print(backupath_path)
    copyfile(sourcepath, backupath_path)
    extend = os.path.splitext(backupath_path)[1]  # 
    if extend != '.xlsx' and extend != '.xls':
        print("It's not a excel file! %s" % backupath_path)
        return False
    new_name = str(file_name.split('.')[0]) + new_type  # .zip
    dir_path = os.path.dirname(backupath_path)  # 
    new_path = os.path.join(dir_path, new_name)  # 
    if os.path.exists(new_path):
        os.remove(new_path)
    print(backupath_path)
    os.rename(backupath_path, new_path)  # 保存新文件，旧文件会替换掉
    return new_path  # 返回新的文件路径，压缩包

#Panel FCT
def writebuff(filePaths):
    print("Paths_buff", filePaths)
    startRowIndex = 25
    count=range(len(filePaths))
    for fileindex, filePath in enumerate(filePaths) :
        # sg.one_line_progress_meter('实时进度条', fileindex + 1, len(count), '-文件转换进度-')
        # app=xw.App(visible=False,add_book=False)
        FCT_1templatefilepath=os.path.join(templateFileSandboxPath,"FCT1.xlsx")
        templatewb = xw.Book(FCT_1templatefilepath)
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

#BI

def writebuff2(filePaths,reportType,plot):
    print("Paths_buff", filePaths)
    startRowIndex = 26
    count=range(len(filePaths))
    for fileindex, filePath in enumerate(filePaths) :
        # sg.one_line_progress_meter('实时进度条', fileindex + 1, len(count), '-文件转换进度-')
        # app=xw.App(visible=False,add_book=False)
        # templateFileSandboxPath
        reportTemplate=""
        fixtureidrange=""
        slots=0
        if reportType=="BI":
            slots=8
            reportTemplate=os.path.join(templateFileSandboxPath,"BI.xlsx")
            fixtureidrange="C5"
            plot=17
        if reportType=="PFCT":
            fixtureidrange="B4"
            slots=4
            if plot==16:
                reportTemplate=os.path.join(templateFileSandboxPath,"PFCT16.xlsx")
            elif plot==17:
                reportTemplate=os.path.join(templateFileSandboxPath,"PFCT17.xlsx")
       
        templatewb = xw.Book(reportTemplate)
        templatesheet = templatewb.sheets["Sheet1"]
        rawFileDirectory = filePath["rawFileDirectory"]
        tempReportFilePath = filePath["tempReportFilePath"]
        
        templatesheet.range(fixtureidrange).value = filePath["fixtureid"]
        reportPath=filePath["reportPath"]
        picturesindex = 0
       
       
        for index1 in range(slots):
            rawdatafilepath = os.path.join(rawFileDirectory, "{0}.xlsx".format(index1+1))
            print(rawdatafilepath)
            buffWb = xw.Book(rawdatafilepath)
            buffsheet = buffWb.sheets["Sheet1"]
            for index in range(plot):
                rawdatarowindex = 18+index*7
                rawDataMaxIndexKey = "C"+str(rawdatarowindex)
                rawDataMinIndexKey = "D"+str(rawdatarowindex)
                reportrowindex=startRowIndex+index+index1*26
                reportMaxIndexKey = "D" + str(reportrowindex)
                reportMinIndexKey = "E" + str(reportrowindex)
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
            picturesindexKey = "G"+str(startRowIndex+index1*26)
            templatesheet.pictures.add(imgpath, left=templatesheet.range(
                picturesindexKey).left, top=templatesheet.range(picturesindexKey).top)
        print(templatesheet.pictures)
        templatewb.save(tempReportFilePath)
        templatewb.close()
       
        os.system("mv {0} {1}".format(tempReportFilePath,reportPath))

        
type_index = input('请选择要生成的报告设备类型：\n1.BI\n2.PFCT-16\n3.PFCT-17\n')
Type_str = ""
plot=0
if int(type_index) == 1:
    Type_str = "BI"
    pass
elif int(type_index) == 2:
    Type_str = "PFCT"
    plot=16
    pass
elif int(type_index) == 3:
    Type_str = "PFCT"
    plot=17
    pass


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
copyFile(os.path.join(os.getcwd(),"template/"),templateFileSandboxPath)#拷贝模版
print("template SandBox directory:",templateFileSandboxPath)
checkAndCreateDirectory(templateFileSandboxPath)
#Template file path
print("template SandBox path:",templateFileSandboxPath)
# templateFilePath=os.path.join(os.getcwd(),"template/usiTemplate.xlsx")
# print("template file path:",templateFilePath)
rawDataSandboxPath = os.path.join(excelSandboxPath, "RawData/")  # 
moveSamdboxRawFile(rawDataSandboxPath)
checkAndCreateDirectory(rawDataSandboxPath)
rawDataPath=os.path.join(os.getcwd(),"data/*")
copyFile(rawDataPath,rawDataSandboxPath)

fixtureNamePaths =checkFile(os.listdir(rawDataSandboxPath))   # 获取下面的文件路径
tempReportPath=os.path.join(excelSandboxPath,"tempReport/")
checkAndCreateDirectory(tempReportPath)#   报告临时存放路径
# os.system("open {0}".format(tempReportPath))
print(tempReportPath)
print(fixtureNamePaths)
filePaths=getFullpath(fixtureNamePaths)
# writebuff(filePaths)
# templateFileSandboxPath=
print(">>>>",templateFileSandboxPath)
writebuff2(filePaths,Type_str,plot)

