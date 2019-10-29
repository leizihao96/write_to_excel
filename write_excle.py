import pymysql
import xlsxwriter


class Database:
    def __init__(self, host='localhost',
                 port=3306,
                 charset='utf8',
                 user='root',
                 passwd='123456',
                 database='stu'):
        self.host = host
        self.port = port
        self.user = user
        self.passwd = passwd
        self.database = database
        self.charset = charset
        self.db = pymysql.connect(host=self.host, port=self.port, passwd=self.passwd, user=self.user,
                                  database=self.database, charset=self.charset)
        self.cur = self.db.cursor()

    def select_(self):
        """

        :return:读取到表中的所有数据
        """
        sql = 'select * from stu.class1;'
        try:
            self.cur.execute(sql)
            self.db.commit()
        except Exception as e:
            print(e)
            self.db.rollback()
        self.data = [ list(i) for i in self.cur.fetchall()]

        return self.data  # data --> (（），（）)

    def mkdir_excel_sheek(self):
        #创建工作簿
        self.workbook = xlsxwriter.Workbook('myworkbook1.xlsx')
        #创建工作表
        self.worksheet = self.workbook.add_worksheet('sheet_1')
        #写入表头
        self.fileds = ['学生id','学生姓名','学生年龄','学生性别','学生成绩']
        for item in range(len(self.fileds)):
            self.worksheet.write(0,item,self.fileds[item])#三个参数分别代表的行列和数据
        self.worksheet.write(0,len(self.fileds),'成绩总和')
    def write_into_excel(self):
        self.select_()
        self.mkdir_excel_sheek()
        data = 0
        for row in range(1,len(self.data)+1):
            data += self.data[row - 1][4]
            for col in range(len(self.fileds)):
                if self.data[row-1][col] == None:
                    self.data[row-1][col] = 'null'
                self.worksheet.write(row,col,self.data[row-1][col])
        self.worksheet.write(1,len(self.fileds),data)
        self.workbook.close()

if __name__ == "__main__":
    target = Database()
    target.write_into_excel()
