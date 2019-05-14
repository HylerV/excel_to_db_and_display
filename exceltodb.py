import wx       #WXPYTHON模块，界面模块
import wx.grid  #WXPYTHON表格组件模块
import MySQLdb  #数据库操作模块
import os       #系统模块
import time     #时间模块
import xlrd     #EXCEL文件读取模块

# 执行sql语句，并返回结果
def ex_sql(sql):
    try:
        # 执行sql语句
        cursor.execute(sql)
        # 提交到数据库执行
        db.commit()
        # 返回执行结果
        results = cursor.fetchall()
        print('数据库中')
        return results
    except:
        # Rollback in case there is any error
        db.rollback()
        print('数据库出错')
        # 关闭数据库连接
        db.close()

# 生成sql语句
def sql_table(part_sql, table):
    sql_list = [part_sql, table]
    return "".join(sql_list)

# 获取时间（年月，例201905）
def get_time():
    # 获取本地时间
    return time.strftime("%Y%m", time.localtime(time.time()))

# 判断数据库或者数据库内表格是否已存在，返回结果
def judge_bool(name, tag):
    # 判断数据库是否已存在
    if tag == 1:

        the_name = ''.join(["'", name, "'"])

        j = ex_sql(sql_table("show databases like ", the_name))

        print(j)
    else:
        j = ex_sql("show tables like ", name)
    
    j_bool = j is not ()
    return j_bool

# 根据本地时间，获取数据库名称
def get_dbname():
    # 得到今天的时间
    now_time = get_time()

    dbname = ''.join(['DB', now_time])

    return dbname

# 连接数据库，判断数据库存在，存在为真，判断是否已有表格，，为真，显示第一个表格，没有表格，显示示例表；
# 不存在数据库的话，建立数据库，并显示示例表
def dbconnect(tag):

    dbname = get_dbname()

    print(judge_bool(dbname, 1))

    if judge_bool(dbname, 1):

        ex_sql(sql_table("use ", dbname))

        tables_name = ex_sql("show tables")

        tables = []

        for i in tables_name:

            tables.append(i[0])

        if tag == 1:

            if tables != []:

                table_select = tables[0]

        else:

            table_select = frame.tables_names.GetStringSelection()

            print(table_select)

        if tables != []:

            print(sql_table('select * from ', table_select))

            values = ex_sql(sql_table('select * from ',table_select))

            co =  ex_sql(sql_table('show columns from ',table_select))

            titles = []

            for i in co:

                titles.append(i[0])

            return values,titles,tables

        else:
            return test_data()

    else:

        ex_sql(sql_table('create database ', dbname))

        # 建立新数据库成功
        print("建立新数据库成功")

        ex_sql(sql_table("use ", dbname))

        return test_data()


# 示例表数据
def test_data():

    values = (('1', '1', '2', '3'), ('2', '4', '5', '6'),  ('3', '7', '8', '9'))
    titles = ['ID','实例A', '实例B', '实例C']
    tables = []
    # tbname = get_tablename()

    return values,titles,tables

# 更新数据库内表格数据
def dbchange(data_part):
    print(sql_table("update ", data_part))
    ex_sql(sql_table("update ", data_part))

# 根据EXCEL表格名称，获取数据库表格名称
def get_tablename(ta_name):
    
    tbname = ''.join(['tb', os.path.splitext(os.path.split(ta_name)[-1])[0]])
    print(tbname)

    operate_excel(ta_name, tbname)


# 操作excel，打开表格，获得数据
def operate_excel(fname, ta_d_name):

    bk = xlrd.open_workbook(fname)
    shxrange = range(bk.nsheets)
    try:
        sh = bk.sheet_by_index(0)
    except:
        print("no sheet in %s named Sheet1" % fname)
    # 获取行数
    nrows = sh.nrows
    # 获取列数
    ncols = sh.ncols
    print("nrows %d, ncols %d" % (nrows, ncols))
    # 获取第一行第一列数据
    # cell_value = sh.cell_value(0, 2)
    # print(cell_value)

    part_name = ""
    part_value = ""

    part_title = sh.row_values(0)

    for i in part_title:
        n = "".join([str(i)," char(20),"]) 
        part_name += n

    part_value = ' values '
    for i in range(1, nrows):
        row_data = sh.row_values(i)
        part_value = "".join([part_value, str(tuple(row_data)), ','])

    create_table_sql = "".join(['create table ', ta_d_name, ' (', part_name[:-1], ')engine=InnoDB default charset=utf8'])
    insert_value_sql = "".join(['insert into ', ta_d_name, part_value[:-1]])

    print(insert_value_sql)

    ex_sql(create_table_sql)
    ex_sql(insert_value_sql)
    print("建表成功")

# 建立图形界面
class GridFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title = "数据表", pos = (200,200), size = (500,400))

        datas = dbconnect(1)

        values = datas[0]

        titles = datas[1]

        tables = datas[2]

        self.panel = wx.Panel(self)

        self.path_text = wx.TextCtrl(self.panel, pos = (5,5), size = (350,24))

        self.path_text.Bind(wx.EVT_LEFT_DOWN, self.openfile, self.path_text)

        self.tables_names = wx.ComboBox(self.panel, -1, value='图表名字', choices=tables, style=wx.CB_SORT, pos = (370,5), size = (80,24))

        self.create_table(values, titles)

        # 导入数据表
        self.import_button = wx.Button(self.panel, label = "导入") 

        self.import_button.Bind(wx.EVT_BUTTON, self.importtable, self.import_button)

        # 显示表格
        self.show_button = wx.Button(self.panel, label = "显示") 

        self.show_button.Bind(wx.EVT_BUTTON, self.showtable, self.show_button)

        # 允许修改表格
        self.change_button = wx.Button(self.panel, label = "修改") 

        self.change_button.Bind(wx.EVT_BUTTON, self.changetable, self.change_button)

        # 修改数据库
        self.renovate_button = wx.Button(self.panel, label = "刷新")

        self.renovate_button.Bind(wx.EVT_BUTTON, self.renovatedatabase, self.renovate_button)

        self.setsizer()

        self.Show()
       
    # 建立组件尺寸器
    def setsizer(self):
        self.box1 = wx.BoxSizer() # 不带参数表示默认实例化一个水平尺寸器
        self.box1.Add(self.path_text, proportion = 6, flag = wx.EXPAND|wx.ALL, border = 3)
        self.box1.Add(self.tables_names, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 3)

        self.box2 = wx.BoxSizer() # 不带参数表示默认实例化一个水平尺寸器
        self.box2.Add(self.grid, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 3)

        self.box3 = wx.BoxSizer() # 不带参数表示默认实例化一个水平尺寸器
        self.box3.Add(self.import_button, flag = wx.EXPAND|wx.ALL, border = 3)
        self.box3.Add(self.show_button, flag = wx.EXPAND|wx.ALL, border = 3)
        self.box3.Add(self.change_button, flag = wx.EXPAND|wx.ALL, border = 3)
        self.box3.Add(self.renovate_button, flag = wx.EXPAND|wx.ALL, border = 3)

        self.v_box = wx.BoxSizer(wx.VERTICAL) # wx.VERTICAL参数表示实例化一个垂直尺寸器
        self.v_box.Add(self.box1, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 3) # 添加组件
        self.v_box.Add(self.box2, proportion = 7, flag = wx.EXPAND|wx.ALL, border = 3) # 添加组件
        self.v_box.Add(self.box3, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 3) # 添加组件

        self.panel.SetSizer(self.v_box)

        self.panel.SetBackgroundColour('#B9D4DB') 

    # 建立WXGrid表格组件
    def create_table(self, values, titles):
         # Create a wxGrid object
        self.grid = wx.grid.Grid(self.panel, -1)

        self.grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.value_change, self.grid)

        self.m = len(values)

        self.n = len(values[0])

        self.grid.CreateGrid(self.m, self.n)
 
        # We can specify that some cells are read.only
        for a in range(self.m):

            self.grid.SetRowSize(a, 30)

            for b in range(self.n):

                self.grid.SetCellValue(a, b, values[a][b])
            
                self.grid.SetReadOnly(a, b, True)

                self.grid.SetCellAlignment(a, b, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE);

        for c in range(self.n):

            self.grid.SetColSize(c, 80)

            self.grid.SetColLabelValue(c, titles[c])


    # 对显示表格数据改动动作监视，并显示
    def value_change(self, event):
        print(self.grid.GetGridCursorCol())
        print(self.grid.GetGridCursorRow())

        print(self.grid.GetCellValue(self.grid.GetGridCursorRow(), self.grid.GetGridCursorCol()))


    # 允许修改表格权限
    def changetable(self, event):

        self.panel.SetBackgroundColour('#D0E6A5') 
        self.panel.Refresh()

        for a in range(self.m):

            for b in range(self.n):

                self.grid.SetReadOnly(a, b, False)

    # 对已改动的表格数据，反馈到数据库，进行修改
    def renovatedatabase(self, event):
        value_change = self.grid.GetCellValue(self.grid.GetGridCursorRow(), self.grid.GetGridCursorCol())
        value_title = self.grid.GetColLabelValue(self.grid.GetGridCursorCol())
        id_title = self.grid.GetColLabelValue(0)
        value_id = self.grid.GetGridCursorRow()

        current_table = self.tables_names.GetStringSelection()

        if current_table is not "":

            data_part = ''.join([current_table, " set", " ", value_title, "='", \
                value_change, "' where", " ", id_title, "='", self.grid.GetCellValue(value_id, 0), "'"])
            print(data_part)
            dbchange(data_part)
            print("修改成功")

            for a in range(self.m):

                for b in range(self.n):

                    self.grid.SetReadOnly(a, b, True)

    # 复选框选择数据库内已存在表格，并进行展示
    def showtable(self, event):

        print(self.tables_names.IsListEmpty())
        print(self.tables_names.IsTextEmpty())

        if self.tables_names.IsListEmpty():

            pass

        else:

            self.grid.Destroy()
            # self.panel.Hide()

            datas = dbconnect(2)

            values = datas[0]

            titles = datas[1]

            self.create_table(values, titles)

            self.setsizer()

            self.Refresh(True)

            event.Skip()

            return self.tables_names.GetStringSelection()


    # “导入按钮动作”，EXCEL获得的数据，导入数据库
    def importtable(self, event):

        ta_name1 = ''.join(['tb', os.path.splitext(os.path.split(self.path_text.GetLineText(0))[-1])[0]])
        ta_name = self.path_text.GetLineText(0)

        if ta_name == "":

            pass

        else:

            get_tablename(ta_name)

            new_table_names = ex_sql("show tables")

            n_t_names = []

            for i in new_table_names:

                n_t_names.append(i[0])

            self.tables_names.Destroy()

            self.tables_names = wx.ComboBox(self.panel, -1, value='图表名字', choices=n_t_names, \
                style=wx.CB_SORT, pos = (370,5), size = (80,24))

            self.setsizer()

            self.tables_names.Refresh(True)

    # 文本框点击操作，限制只能打开EXCEL文件，并在文本框显示路径
    def openfile(self, event):

        with wx.FileDialog(self, "Open XYZ file", wildcard="EXCEL files (*.xls;*.xlsx)|*.xls;*.xlsx",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return     # the user changed their mind

            # Proceed loading the file chosen by the user
            pathname = fileDialog.GetPath()

            self.path_text.SetValue(pathname)

# 主函数
if __name__ == '__main__':

    app = wx.App(0)
    
    # 连接数据库
    db = MySQLdb.connect("localhost", "root", "password", charset="utf8")

    cursor = db.cursor()

    print("数据库已打开")
    
    frame = GridFrame(None)

    frame.Show()

    app.MainLoop()
