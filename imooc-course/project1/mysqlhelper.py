import pymysql

class dbhelper():
    def __init__(self,host,port,user,passwd,db,charset="utf8"):
        self.host=host
        self.port=port
        self.user=user
        self.passwd=passwd
        self.db=db
        self.charset=charset
    #创建一个链接
    def connection(self):
        #1. 创建连接
        self.conn = pymysql.connect(host=self.host, port=self.port, user=self.user, passwd=self.passwd, db=self.db,charset=self.charset)
        #2. 创建游标
        self.cur = self.conn.cursor()
    #关闭链接
    def closeconnection(self):
        self.cur.close()
        self.conn.close()
    #查询一条数据
    def getonedata(self,sql):
        try:
            self.connection()
            self.cur.execute(sql)
            result=self.cur.fetchone()
            self.closeconnection()
        except Exception:
            print(Exception)
        return result
    #查询多条数据
    def getalldata(self,sql):
        try:
            self.connection()
            self.cur.execute(sql)
            result=self.cur.fetchall()
            self.closeconnection()
        except Exception:
            print(Exception)
        return result
    #添加/修改/删除
    def executedata(self,sql):
        try:
            self.connection()
            self.cur.execute(sql)
            self.conn.commit()
            self.closeconnection()
        except Exception:
            print(Exception)
    #批量插入
    def executemanydata(self,sql,vals):
        try:
            self.connection()
            self.cur.executemany(sql,vals)
            self.conn.commit()
            self.closeconnection()
        except Exception as e:
            print(e)