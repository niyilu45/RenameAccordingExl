import os
import zipfile
import openpyxl
import shutil
from pprint import pprint
def DecompressZip(filePath, outFolder=None):
    supportSuffix = [".zip", ".rar"]
    suffix = os.path.splitext(filePath)[1].lower()
    # 0) check suffix
    if not suffix in supportSuffix:
        print(f"Do not support suffix as {suffix}")
        return -1
    # 1) decompress zip
    if suffix == ".zip":
        pprint(f"Decompressing {filePath}")
        with zipfile.ZipFile(filePath, 'r') as zipRef:
            for info in zipRef.infolist():
                fileName = info.filename.encode("cp437").decode("gbk")
                zipRef.extract(info, outFolder)
                # use try to solve chinese filename error
                tarPath = os.path.join(outFolder, fileName)
                try:
                    os.rename(os.path.join(outFolder, info.filename), tarPath)
                except FileExistsError:
                    os.remove(tarPath)
                    os.rename(os.path.join(outFolder, info.filename), tarPath)

    pass
if __name__ == '__main__':
    # 0) user configs
    inputPath = "input/" # zip folder
    outputPath = "output/" # zip decompress folder
    exlPath = "./Test.xlsx"
    keyConnect = ["Tdoc", "Source"]

    # 1) find all zips
    zipSuffix = [".zip", ".rar"]
    allZips = []
    for root, dirs, files in os.walk(inputPath):
        for f in files:
            suffix = os.path.splitext(f)[1].lower()
            if suffix in zipSuffix:
                allZips.append(f)
    print(f"find %d zips in {inputPath}:" % len(allZips))
    pprint(allZips)

    # 2) express zips
    usrInput = input("Do you need to decompress all zips?(y/n)?")
    if not usrInput or usrInput[0] == 'y':
        for item in allZips:
            filePath = os.path.join(inputPath, item)
            DecompressZip(filePath, outputPath)
        print("decompress all files")

    # 3) load excel
    exlBook = openpyxl.load_workbook(exlPath)
    exlSheets = exlBook.sheetnames

    exlData = []
    for sheetName in exlSheets:
        # Parse tarSheet
        sheet = exlBook[sheetName]
        for row in sheet.rows:
            line = []
            for cell in row:
                line.append(cell.value)
            exlData.append(line)
        break # only read first sheet

    # 4) parse key idx
    keyDict = {}
    for key in keyConnect:
        keyDict[key] = exlData[0].index(key)

    # 5) find name pairs accoding to excel
    sourceList = [x[0] for x in exlData[1:]]
    for root, dirs, files in os.walk(outputPath):
        for d in dirs:
            try:
                tarLineIdx = 1 + sourceList.index(d)
                tarLine = exlData[tarLineIdx]
                nameElement = [tarLine[keyDict[x]] for x in keyConnect]
                tarName = " ".join(nameElement)
                tarPath = os.path.join(outputPath, tarName)
                try:
                    os.rename(os.path.join(outputPath, d), tarPath)
                except FileExistsError:
                    shutil.rmtree(tarPath)
                    os.rename(os.path.join(outputPath, d), tarPath)

            except ValueError:
                pprint(f"Cannot find {d} in excel!!")
