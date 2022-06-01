from re import I
from sqlite3.dbapi2 import Row
from tkinter import *
import tkinter
from tkinter.messagebox import askokcancel, showinfo
import sqlite3
import tkinter.ttk
#import xlsxwriter


# 创建顶级窗口
root = Tk()
systitle = '班级信息管理系统'
frame = Frame()
frame.pack()
dbfile = 'student.db'
cn = sqlite3.connect(dbfile)
cr = cn.cursor()

# 定义主函数
def main():
    root.geometry('600x400')
    root.title(systitle)
    # 创建顶级菜单
    menubar = Menu(root, tearoff=0)
    # 创建学生管理菜单并增加菜单项
    smenu = Menu(menubar)
    smenu.add_command(label='添加新学生', font=('宋体', 10), command=addstudent)
    smenu.add_command(label='显示学生全部信息', font=('宋体', 10), command=showallstudent)
    smenu.add_command(label='查找/修改/删除学生信息', font=('宋体', 10), command=checkupdatestudent)
    smenu.add_separator()
    smenu.add_command(label='退出', font=('宋体', 10), command=goexit)
    menubar.add_cascade(label='学生管理', font=('宋体', 10), menu=smenu)

    # 创建科目管理菜单并添加菜单项
    subject = Menu(menubar)
    subject.add_command(label='录入新科目', font=('宋体', 10), command=addcourse)
    subject.add_command(label='显示全部科目', font=('宋体', 10), command=showallcourse)
    subject.add_command(label='查找/修改/删除科目信息', font=('宋体', 10), command=checkupdetecourse)
    menubar.add_cascade(label='科目管理', font=('宋体', 10), menu=subject)

    # 创建成绩管理菜单并添加菜单项
    grade = Menu(menubar)
    grade.add_command(label='录入成绩', font=('宋体', 10), command=addgrade)
    grade.add_command(label='查询成绩', font=('宋体', 10), command=searchbycid)
    grade.add_command(label='成绩输出为Excel', font=('宋体', 10))
    menubar.add_cascade(label='成绩管理', font=('宋体', 10), menu=grade)

    # 创建其他菜单并添加菜单项
    other = Menu(menubar)
    other.add_command(label='查看操作日志', font=('宋体', 10))
    other.add_command(label='版本信息', font=('宋体', 10))

    menubar.add_cascade(label='其他', font=('宋体', 10), menu=other, command=edition)

    root.config(menu=menubar)
    root.mainloop()

# 定义版本信息函数
def edition():
    showinfo(systitle, '菜鸟程序员\n1.0版本')
def addstudent():
    for widget in frame.winfo_children():
        widget.destroy()

    f1 = Frame(frame)
    f1.pack()

    sidvar = StringVar()
    snamevar = StringVar()
    lsid = Label(f1, text='学号:')
    lsname = Label(f1, text='姓名:')
    esid = Entry(f1, textvariable=sidvar)
    ename = Entry(f1, textvariable=snamevar)

    f2 = Frame(frame)
    f2.pack(pady=20)

    breset = Button(f2, text='重置')
    bsave = Button(f2, text='保存')

    lsid.grid(row=1, column=1)
    lsname.grid(row=2, column=1)
    esid.grid(row=1, column=2)
    ename.grid(row=2, column=2)

    breset.grid(row=1, column=1)
    bsave.grid(row=1, column=2)

    def reset():
        sidvar.set('')
        snamevar.set('')

    def save():
        try:
            id = sidvar.get()
            if not id.isdigit():
                raise EXCEPTION('学号必须为数字')
            elif len(id) != 4:
                raise EXCEPTION('学号必须为四位')

            name = snamevar.get()

            cr.execute('insert into student values(?,?)', (id, name))

            cn.commit()
            showinfo(systitle, '成功添加新学生')

            reset()
            esid.focus()

            # 测试
            stu_list = cr.execute('select * from student').fetchall()
            print(stu_list)
        except EXCEPTION as ex:
            showinfo(systitle, ex)

    breset.config(command=reset)
    bsave.config(command=save)
def showallstudent():
    for widget in frame.winfo_children():
        widget.destroy()
    frame.columnconfigure(1, minsize=50)
    frame.columnconfigure(2, minsize=100)
    Label(frame, text='学号', font=('宋体', 10)).grid(row=0, column=1)
    Label(frame, text='姓名', font=('宋体', 10)).grid(row=0, column=2)
    stulist = cr.execute('select * from student').fetchall()
    rownum = 1
    for stu in stulist:
        colum = 1
        for info in stu:
            Label(frame, text=str(info), font=('宋体', 10)).grid(row=rownum, column=colum)
            colum += 1
        rownum += 1
def checkupdatestudent():
    for widget in frame.winfo_children():
        widget.destroy()
    f1 = LabelFrame(frame, text='查找学生')
    f1.pack()
    searchsidvar = StringVar()
    Label(f1, text='学生学号', ).grid(row=1, column=1)
    esearchsid = Entry(f1, textvariable=searchsidvar)
    esearchsid.grid(row=1, column=2)
    bserach = Button(f1, text='查询', font=('宋体', 10))
    bserach.grid(row=1, column=3)
    f2 = LabelFrame(frame, text='删除与修改')
    f2.pack(pady=30)
    fdel = Frame(f2)
    fdel.pack()
    bdel = Button(f2, text="删除", font=('宋体', 10))
    bdel.pack()
    fupdate = Frame(f2)
    fupdate.pack()
    sidvar = StringVar()
    snamevar = StringVar()
    Label(fupdate, text='学号', font=('宋体', 10), ).grid(row=1, column=1)
    Label(fupdate, text='姓名', font=('宋体', 10)).grid(row=2, column=1)
    Entry(fupdate, textvariable=sidvar, ).grid(row=1, column=2)
    Entry(fupdate, textvariable=snamevar, ).grid(row=2, column=2)
    bsave = Button(fupdate, text='保存', font=('宋体', 10), state=NORMAL)
    bsave.grid(row=3, column=3)

    def quit_():
        f2.destroy()
        f1.destroy()

    bexit = Button(fupdate, text='退出', command=quit_)
    bexit.grid(row=3, column=1)

    def search():
        sidvar.set('')
        snamevar.set('')
        searchid = searchsidvar.get()
        stu = cr.execute('select sid,name from student where sid=?', (searchid,)).fetchone()
        if stu == None:
            showinfo(systitle, '无该学生信息')
            bsave.config(state=DISABLED)
            bdel.config(state=DISABLED)
        else:
            sidvar.set(stu[0])
            snamevar.set(stu[1])
            bsave.config(state=NORMAL)
            bdel.config(state=NORMAL)
    bserach.config(command=search)
    esearchsid.bind('<Return>', search)
    def delete():
        cn.execute('delete from student where sid=?', (sidvar.get(),))
        cn.commit()
        showinfo(systitle, '删除成功!')
        sidvar.set('')
        snamevar.set('')
    bdel.config(command=delete)
    def save():
        cn.execute('update student set name=? where sid=?',(snamevar.get(), sidvar.get()))
        cn.commit()
        showinfo(systitle, '修改成功!')
    bsave.config(command=save)
# 创建添加科目函数
def addcourse():
    for widget in frame.winfo_children():
        widget.destroy()
    f1 = Frame(frame)
    f1.pack()
    cidvar = StringVar()
    cnamevar = StringVar()
    lcid = Label(f1, text='科目名称', font=('宋体', 10))
    lcname = Label(f1, text='科目编号', font=('宋体', 10))
    ecid = Entry(f1, textvariable=cidvar)
    ecname = Entry(f1, textvariable=cnamevar)
    f2 = Frame(frame)
    f2.pack()
    breset = Button(f2, text='重置', font=('宋体', 10))
    bsave = Button(f2, text='保存', font=('宋体', 10))
    lcname.grid(row=2, column=1)
    lcid.grid(row=1, column=1)
    ecid.grid(row=2, column=2)
    ecname.grid(row=1, column=2)
    breset.grid(row=1, column=1)
    bsave.grid(row=1, column=2)
    def reset():
        cidvar.set('')
        cnamevar.set('')
    def save():
        
        name = cnamevar.get()
        id = cidvar.get()
        cr.execute('insert into course values(?,?)', (id, name))
        cn.commit()
        showinfo(systitle, '成功添加科目')
        reset()
        ecid.focus()
        # 测试
        course_list = cr.execute('select * from course').fetchall()
        print(course_list)
    breset.config(command=reset)
    bsave.config(command=save)
# 创建显示科目函数
def showallcourse():
    for widget in frame.winfo_children():
        widget.destroy()
    frame.columnconfigure(1, minsize=50)
    frame.columnconfigure(2, minsize=100)
    Label(frame, text='学科', font=('宋体', 15))
    Label(frame, text='编号', font=('宋体', 10))
    courselist = cr.execute('select * from course').fetchall()
    rownum = 1
    for cou in courselist:
        colnum = 1
        for info in cou:
            Label(frame, text=str(info), font=('宋体', 10)).grid(row=rownum, column=colnum)
            colnum += 1
        rownum += 1

# 创建查找,修改,删除科目函数
def checkupdetecourse():
    for widget in frame.winfo_children():
        widget.destroy()
    f1 = LabelFrame(frame, text='查询课程')
    f1.pack()
    searchcnamevar = StringVar()
    Label(f1, text='课程名称', font=('宋体, 10')).grid(row=1, column=1)
    esearchcname = Entry(f1, textvariable=searchcnamevar)
    esearchcname.grid(row=1, column=2)
    bsearch = Button(f1, text='查询', font=('宋体 ', 10))
    bsearch.grid(row=1, column=3)
    f2 = LabelFrame(frame, text='删除与修改')
    f2.pack(pady=30)
    fdel = Frame(f2)
    fdel.pack()
    bdel = Button(f2, text='删除', font=('宋体', 10))
    bdel.pack()
    fupdate = Frame(f2)
    fupdate.pack()
    cidvar = StringVar()
    cnamevar = StringVar()
    Label(fupdate, text='课程编号', font=('宋体', 10)).grid(row=1, column=1)
    Label(fupdate, text='课程名称', font=('宋体', 10)).grid(row=2, column=1)
    Entry(fupdate, textvariable=cidvar).grid(row=1, column=2)
    Entry(fupdate, textvariable=cnamevar).grid(row=2, column=2)
    bsave = Button(fupdate, text='保存', font=('宋体', 10))
    bsave.grid(row=3, column=3)
    def quit_():
        f2.destroy()
        f1.destroy()

    bexit = Button(fupdate, text='退出', command=quit_)
    bexit.grid(row=3, column=1)

    # 创建查询函数
    def search():
        cidvar.set('')
        cnamevar.set('')
        searchname = searchcnamevar.get()
        cou = cr.execute('select cname,cid from course where c=?', searchname).fetchone()
        if cou is None:
            showinfo(systitle, '没有该科目信息')
            bsave.config(state=DISABLED)
            bdel.config(state=DISABLED)
        else:
            cidvar.set(cou[0])
            cnamevar.set(cou[1])
            bsave.config(state=NORMAL)
            bdel.config(state=NORMAL)
    bsearch.config(command=search)
    esearchcname.bind('<Return>', search)
    # 创建删除函数
    def delete():
        cn.execute('delete from course where cname=?', (cnamevar.get(),))
        cn.commit()
        showinfo(systitle, '删除成功!')
        cidvar.set('')
        cnamevar.set('')
    bdel.config(command=delete)
    # 定义修改保存函数
    def save():
        cn.execute('update course set cname=? where cid=?', (cidvar.get(), cnamevar.get()))
        cn.commit()
        showinfo(systitle, '修改成功!')
    bsave.config(command=save)

# 定义成绩录入函数
def addgrade():
    for widget in frame.winfo_children():
        widget.destroy()
    f1 = Frame(frame)
    f1.pack()
    sidvar = StringVar()
    scnamevar = StringVar()
    sgradevar = StringVar()
    lsid = Label(f1, text='学号', font=('宋体', 10))
    lscanme = Label(f1, text='科目', font=('宋体', 10))
    lsgrade = Label(f1, text='成绩', font=('宋体', 10))
    esid = Entry(f1, textvariable=sidvar)
    escname = Entry(f1, textvariable=scnamevar)
    esgragde = Entry(f1, textvariable=sgradevar)
    f2 = Frame(frame)
    f2.pack()
    breset = Button(f2, text='重置', font=('宋体', 10))
    bsave = Button(f2, text='保存', font=('宋体', 10))
    lsid.grid(row=1, column=1)
    lscanme.grid(row=2, column=1)
    lsgrade.grid(row=3, column=1)
    esid.grid(row=1, column=2)
    escname.grid(row=2, column=2)
    esgragde.grid(row=3, column=2)
    breset.grid(row=4, column=1)
    bsave.grid(row=4, column=3)
    # 创建重置函数
    def reset():
        sidvar.set('')
        scnamevar.set('')
        sgradevar.set('')
    
    # 定义保存函数
    
    def save():
        id = sidvar.get()
        name = scnamevar.get()
        grade = sgradevar.get()
        cr.execute('insert into score values(?,?,?)', (id, name, grade))
        cn.commit()
        showinfo(systitle, '成功添加新的学生成绩')
        reset()
        esid.focus()
        # 测试
        stu_list = cr.execute('select * from student').fetchall()
        print(stu_list)

    breset.config(command=reset)
    bsave.config(command=save)
    

# 创建查询成绩函数
def searchbycid():
    for widget in frame.winfo_children():
        widget.destroy()
    f1 = LabelFrame(frame)
    f1.pack()

    searchvar = StringVar()
    Label(f1, text='科目名称', font=('宋体, 10')).grid(row=1, column=1)
    combosearch = tkinter.ttk.Combobox(f1, textvariable=searchvar)
    combosearch.grid(row=1, column=2)
    courselist = cr.execute('select cid,cname from course').fetchall()
    cnamelist = []
    for c in courselist:
        cnamelist.append(c[1])
    combosearch['values'] = cnamelist

    
    f2 = Frame(frame)
    f2.pack()
    
    Label(f2, text='序号', font=('宋体', 10), width=10).grid(row=1, column=1)
    Label(f2, text='学号', font=('宋体', 10), width=10).grid(row=1, column=2)
    Label(f2, text='姓名', font=('宋体', 10), width=10).grid(row=1, column=3)
    Label(f2, text='科目名称', font=('宋体', 10), width=10).grid(row=1, column=4)
    Label(f2, text='成绩', font=('宋体', 10), width=10).grid(row=1, column=5)

    f3 = Frame(frame)
    f3.pack()
    def search(*args):
        for widget in f3.winfo_children():
            widget.destroy()
        searchname = searchvar.get()
        for sub in courselist:
            if sub in courselist:
                searchid = sub[0]
        scorelist = cn.execute('''select sc.gid,st.cid,st.name,c.cname,sc.grade formstudent,st,
        score sc,course c where sc.sid=st.sid and sc.cid=c.cid and sc.cid=? ''',(searchid,)).fetchall()
        r = 2
        for s in scorelist:
            c = 1
            for info in s:
                Label(f3, text=info, width=10).grid(row=r, column=c)
    combosearch.bind('<<ComboboxSelected>>', search)
# def toexcel():
#     workbook = xlsxwriter.Workbook('成绩表.xlsx')
#     worksheet = workbook.add_worksheet()
#     scorelist = cn.execute( scorelist = cn.execute('''select sc.gid,st.sid,st.name,c.cname,sc.grade form student st,score sc,course c where sc.sid=st.sid and sc.cid=c.cid''',).fetchall()
    # for i, j in [('A1', '序号'), ('B1', '学号'), ('C1','姓名'), ('D1', '科目'), ('E1', '成绩')]:
    #     worksheet.write(i, j)
    # r = 1
    # for s in scorelist:
    #     c = 0
    #     for info in s:
    #         worksheet.write(r, c, info)
    #         c += 1
    #     r += 1
    # workbook.close()

def goexit():
    if askokcancel('班级信息管理系统', '确定退出系统?'):
        exit(0)
if __name__ == '__main__':
    main()
