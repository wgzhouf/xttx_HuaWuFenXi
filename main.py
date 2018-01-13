from tool import read_csv


def main():
    # read_csv(r'C:\Users\15305\Desktop\通话20180105_0111.csv')
    # print('==================')
    datagrid = read_csv(r'C:\Users\15305\Desktop\20180105-20180111日协同通信活跃情况.csv', 1)
    for i in datagrid:
        print(i[1])


if __name__ == '__main__':
    main()
