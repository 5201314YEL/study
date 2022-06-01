import sqlite3
#创建数据库
cn = sqlite3.connect('student.db')
#创建表
cr = cn.cursor()
sql1 = '''create table if not exists student(
                    sid varchar(6) primary key,
                    name varchar(50) 
                    )
                    '''

sql2 = '''create table if not exists course(
        cid varchar(6) primary key,
        cname varchar(50)
        )
        '''
sql3 = '''create table if not exists score(
          sid varchar(6),
          gid varchar(6) primary key,
          cid varchar(6),
          grade int
          )
          '''
cr.execute(sql1)
cr.execute(sql2)
cr.execute(sql3)
cn.commit()
cr.close()
cn.close()