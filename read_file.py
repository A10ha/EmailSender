Lines = []


def File_Read(file_path):
    global Lines
    Lines = []
    with open(file_path, 'r') as f:
        while True:
            line = f.readline()  # 逐行读取
            if not line:  # 到 EOF，返回空字符串，则终止循环
                break
            File_Data(line, flag=1)
    File_Data(line, flag=0)


def File_Data(line, flag):
    global Lines
    if flag == 1:
        Lines.append(line)
    else:
        return Lines


if __name__ == '__main__':
    File_Read('./test_email.txt')
    print(Lines)
    File_Read('./email.txt')
    print(Lines)
    # print(line)
