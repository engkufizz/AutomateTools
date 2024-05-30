import os

def getFileList(path):
    fileList = os.listdir(path)
    num = 0
    while num <= len(fileList) - 1:
        cExt = fileList[num][len(fileList[num]) - 4:]
        if cExt == '.log' or cExt == 'dblg':
            num += 1
        else:
            del fileList[num]
    return fileList

def analysize(fName, outputMode=1):
    keywords = ['q0w:', 'q16w:', 'q32w:', 'q48w:']
    
    with open(fName, encoding='utf-8') as file:
        lines = file.readlines()
    
    cLN = 0
    while cLN <= len(lines) - 1:
        cLine = lines[cLN].strip()
        cLineLength = len(cLine)
        if cLineLength < 4:
            del lines[cLN]
            continue
        if 'q0w:' not in cLine:
            del lines[cLN]
        else:
            cLN += 5
            print("Preprocessing : ", cLN, '/', len(lines))
    if keywords[3] not in lines[len(lines) - 1]:
        del lines[len(lines) - 1]
        
    cLN = 0
    while cLN <= len(lines) - 1:
        cLine = lines[cLN].strip()
        cLineLength = len(cLine)

        if 'q0w:' not in cLine:
            cLN += 1
        else:
            if outputMode == 1:
                lines[cLN] = lines[cLN][0:6] + lines[cLN][15:cLineLength] + ' ' + lines[cLN + 1][7:15] + '\n'
                lines[cLN + 1] = lines[cLN + 1][0:6] + lines[cLN + 1][15:cLineLength] + ' ' + lines[cLN + 2][7:15] + '\n'
                lines[cLN + 2] = lines[cLN + 2][0:6] + lines[cLN + 2][15:cLineLength] + ' ' + lines[cLN + 3][7:15] + '\n'
                lines[cLN + 3] = lines[cLN + 3][0:6] + lines[cLN + 3][15:]
                cLN += 5
                print("Recombination : ", cLN, '/', len(lines))
            else:
                lines[cLN] = '000000  ' + lines[cLN][16:18] + ' ' + lines[cLN][18:20] + ' ' + lines[cLN][20:22] + ' ' + lines[cLN][22:24] + ' ' \
                           + lines[cLN][25:27] + ' ' + lines[cLN][27:29] + ' ' + lines[cLN][29:31] + ' ' + lines[cLN][31:33] + ' ' \
                           + lines[cLN][34:36] + ' ' + lines[cLN][36:38] + ' ' + lines[cLN][38:40] + ' ' + lines[cLN][40:42] + ' ' \
                           + lines[cLN + 1][7:9] + ' ' + lines[cLN + 1][9:11] + ' ' + lines[cLN + 1][11:13] + ' ' + lines[cLN + 1][13:15] + '\n'
                lines[cLN + 1] = '000010  ' + lines[cLN + 1][16:18] + ' ' + lines[cLN + 1][18:20] + ' ' + lines[cLN + 1][20:22] + ' ' + lines[cLN + 1][22:24] + ' ' \
                           + lines[cLN + 1][25:27] + ' ' + lines[cLN + 1][27:29] + ' ' + lines[cLN + 1][29:31] + ' ' + lines[cLN + 1][31:33] + ' ' \
                           + lines[cLN + 1][34:36] + ' ' + lines[cLN + 1][36:38] + ' ' + lines[cLN + 1][38:40] + ' ' + lines[cLN + 1][40:42] + ' ' \
                           + lines[cLN + 2][7:9] + ' ' + lines[cLN + 2][9:11] + ' ' + lines[cLN + 2][11:13] + ' ' + lines[cLN + 2][13:15] + '\n'
                lines[cLN + 2] = '000020  ' + lines[cLN + 2][16:18] + ' ' + lines[cLN + 2][18:20] + ' ' + lines[cLN + 2][20:22] + ' ' + lines[cLN + 2][22:24] + ' ' \
                           + lines[cLN + 2][25:27] + ' ' + lines[cLN + 2][27:29] + ' ' + lines[cLN + 2][29:31] + ' ' + lines[cLN + 2][31:33] + ' ' \
                           + lines[cLN + 2][34:36] + ' ' + lines[cLN + 2][36:38] + ' ' + lines[cLN + 2][38:40] + ' ' + lines[cLN + 2][40:42] + ' ' \
                           + lines[cLN + 3][7:9] + ' ' + lines[cLN + 3][9:11] + ' ' + lines[cLN + 3][11:13] + ' ' + lines[cLN + 3][13:15] + '\n'
                lines[cLN + 3] = '000030  ' + lines[cLN + 3][16:18] + ' ' + lines[cLN + 3][18:20] + ' ' + lines[cLN + 3][20:22] + ' ' + lines[cLN + 3][22:24] + ' ' \
                           + lines[cLN + 3][25:27] + ' ' + lines[cLN + 3][27:29] + ' ' + lines[cLN + 3][29:31] + ' ' + lines[cLN + 3][31:33] + ' ' \
                           + lines[cLN + 3][34:36] + ' ' + lines[cLN + 3][36:38] + ' ' + lines[cLN + 3][38:40] + ' ' + lines[cLN + 3][40:42] + '\n'    
                cLN += 5
                print("Recombination : ", cLN, '/', len(lines))                  
    
    with open(fName, mode='w+', encoding='utf-8') as file:
        file.writelines(lines)

if __name__ == '__main__':
    path = "./"
    fileList = getFileList(path)
    
    for ff in range(0, len(fileList)):
        fName = fileList[ff]
        fName = path + fName
        print(ff, len(fileList), fName)
        
        analysize(fName, 2)
