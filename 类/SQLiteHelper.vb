'作者：刮骨剑
'日期：2021-07-07
'使用说明：
'1、在网站（http://system.data.sqlite.org/index.html/doc/trunk/www/downloads.wiki）下载对应工程 .NET 版本的预编译包，
'   建议使用 32 位版本，如：Precompiled Binaries For 32-bit Windows (.NET Framework 4.5.1) 列表下的
'   sqlite-netFx46-binary-Win32-2015-1.0.114.0.zip
'2、把压缩包中的 System.Data.SQLite.dll、SQLite.Interop.dll 文件放进工程的生成输出路径下（如 bin\Debug\）。
'3、添加引用 System.Data.SQLite.dll。
Imports System.Data.SQLite
''' <summary>
''' SQLite 数据库操作类
''' </summary>
Public Class SQLiteHelper
    ''' <summary>
    ''' 数据库连接配置
    ''' </summary>
    Private ReadOnly dbCfgs As New Dictionary(Of String, Object)

    ''' <summary>
    ''' 数据库连接对象
    ''' </summary>
    Private ReadOnly sqlConn As SQLiteConnection

    ''' <summary>
    ''' SQl语句对象
    ''' </summary>
    Private sqlCmd As SQLiteCommand

    ''' <summary>
    ''' 事务控制对象
    ''' </summary>
    Private sqlTran As SQLiteTransaction

    ''' <summary>
    ''' 数据库操作状态
    ''' </summary>
    Public Property SQLStatus As Boolean

    ''' <summary>
    ''' 影响行数
    ''' </summary>
    Public Property RowsAffected As Integer = -1

    ''' <summary>
    ''' 字段集数组
    ''' </summary>
    Public Property ColsArr() As String()

    ''' <summary>
    ''' 数据集数组
    ''' </summary>
    Public Property DatasArr() As String(,)

    ''' <summary>
    ''' 数据集（DataTable）
    ''' </summary>
    Public Property DatasDt As New DataTable

    Public Property DatasReader As New SQLiteDataAdapter

    ''' <summary>
    ''' 数据集（DataSet）
    ''' </summary>
    Public Property DatasDs As New DataSet

    ''' <summary>
    ''' 数据库连接配置初始化
    ''' </summary>
    Private Sub DBCfgInit()
        Dim dbCfgId As String    '数据库连接配置ID

        '配置示例
        dbCfgId = "test"
        dbCfgs(dbCfgId) = New Dictionary(Of String, String) From {
            {"dbAddress", IO.Directory.GetCurrentDirectory & "\sqlite.db"}
        }

        'dbCfgId = "UserInfo"
        'dbCfgs(dbCfgId) = New Dictionary(Of String, String) From {
        '    {"dbAddress", IO.Directory.GetCurrentDirectory & "\sqlite.db"}
        '}
    End Sub

    ''' <summary>
    ''' 创建 SQLiteHelper 实例
    ''' </summary>
    ''' <param name="dbCfgId">数据库连接配置ID</param>
    Public Sub New(ByVal dbCfgId As String)
        Dim sqlConnStr As String
        Dim dbAddress As String '数据库连接地址

        '数据库连接配置初始化
        'Call DBCfgInit()

        '获取数据库配置
        'If dbCfgs.ContainsKey(dbCfgId) Then
        '    dbAddress = dbCfgs(dbCfgId)("dbAddress")
        'Else
        '    SQLStatus = False
        '    MsgBox("数据库连接信息配置 " & dbCfgId & " 不存在！", 16, "错误")
        '    Exit Sub
        'End If
        dbAddress = dbCfgId
        sqlConnStr = "Data Source=" & dbAddress & ";"

        Try
            '连接并打开数据库
            sqlConn = New SQLiteConnection(sqlConnStr)
            sqlConn.Open()
            SQLStatus = True
        Catch ex As Exception
            SQLStatus = False
            MsgBox("数据库连接错误！" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, 16, "错误")
            Exit Sub
        End Try
    End Sub

    ''' <summary>
    ''' 执行SQL增删改语句
    ''' </summary>
    ''' <param name="sqlStrs">SQl语句文本列表</param>
    Public Sub IUD(ByVal sqlStrs As ArrayList)
        Dim sqlStr As String
        Dim rowsAff As Integer

        '变量初始化
        RowsAffected = -1

        Try
            '开启事务
            sqlTran = sqlConn.BeginTransaction()

            For Each sqlStr In sqlStrs
                sqlCmd = New SQLiteCommand(sqlStr, sqlConn, sqlTran)

                rowsAff = sqlCmd.ExecuteNonQuery()

                Console.WriteLine(rowsAff)

                RowsAffected += rowsAff
            Next

            RowsAffected += 1

            '提交事务
            sqlTran.Commit()

            SQLStatus = True
        Catch ex As Exception
            '回滚事务
            sqlTran.Rollback()

            SQLStatus = False

            MsgBox("数据库操作失败！" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, 16, "错误")

            Exit Sub
        End Try
    End Sub

    ''' <summary>
    ''' 执行SQL查询语句并将结果集存入数组（仅返回第一个结果集）
    ''' </summary>
    ''' <param name="sqlStr">SQl语句文本</param>
    Public Sub SelectToArray(ByVal sqlStr As String)
        Dim sqlReader As Object

        '变量初始化
        RowsAffected = -1
        ColsArr = Nothing
        DatasArr = Nothing

        Try
            sqlCmd = New SQLiteCommand(sqlStr, sqlConn, sqlTran)

            '执行SQL语句
            sqlReader = sqlCmd.ExecuteReader()

            '判断是否有结果
            If sqlReader.HasRows = False Then
                RowsAffected = 0
            End If

            '注：因为 SqlDataReader 是一条一条语句的读取，只能获取列数，不能获取行数，因此需要将总的记录除以列数才能获取行数。

            ' 定义函数返回数组的列数和行数
            Dim lstReader As New List(Of String)
            Dim intColumnCount As Integer = sqlReader.FieldCount
            Dim intRowsCount As Integer

            '将字段集存入数组
            ReDim ColsArr(intColumnCount - 1)
            For i = 0 To intColumnCount - 1
                ColsArr(i) = sqlReader.GetName(i)
            Next

            '将数据集存入 lstReader 列表
            While sqlReader.Read()
                For i = 0 To intColumnCount - 1
                    lstReader.Add(sqlReader.GetValue(i).ToString)   '如果字段值是 NULL，则返回空字符串
                Next
            End While

            '获取数据集的行数
            intRowsCount = lstReader.Count / intColumnCount

            '将数据集存入二维数组
            ReDim DatasArr(intRowsCount - 1, intColumnCount - 1)
            Dim lstIdx As Integer = 0
            For j = 0 To UBound(DatasArr, 1)
                For i = 0 To UBound(DatasArr, 2)
                    DatasArr(j, i) = lstReader.Item(lstIdx)
                    lstIdx += 1
                Next
            Next

            RowsAffected = intRowsCount

            sqlReader.Close()

            SQLStatus = True
        Catch ex As Exception
            SQLStatus = False

            MsgBox("数据库操作失败！" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, 16, "错误")

            Exit Sub
        End Try
    End Sub

    Public Sub SelectToReader(ByVal sqlStr As String)
        Dim sqlReader As Object

        '变量初始化
        RowsAffected = -1

        Try
            'sqlCmd = New SQLiteCommand(sqlStr, sqlConn, sqlTran)
            sqlReader = New SQLiteDataAdapter(sqlStr, sqlConn)

            '判断是否有结果
            'If sqlReader.HasRows = False Then
            '    RowsAffected = 0
            'End If

            DatasReader = sqlReader

            sqlReader.Close()

            SQLStatus = True
        Catch ex As Exception
            SQLStatus = False

            MsgBox("数据库操作失败！" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, 16, "错误")

            Exit Sub
        End Try
    End Sub

    ''' <summary>
    ''' 执行SQL查询语句并将结果集存入DataTable对象（仅返回第一个结果集）
    ''' </summary>
    ''' <param name="sqlStr">SQl语句文本</param>
    Public Sub SelectToDt(ByVal sqlStr As String)
        Dim sqlReader As Object

        '变量初始化
        RowsAffected = -1
        DatasDt = New DataTable

        Try
            sqlCmd = New SQLiteCommand(sqlStr, sqlConn, sqlTran)

            '执行SQL语句
            sqlReader = sqlCmd.ExecuteReader()

            '判断是否有结果
            If sqlReader.HasRows = False Then
                RowsAffected = 0
            End If

            DatasDt.Load(sqlReader)

            RowsAffected = DatasDt.Rows.Count

            sqlReader.Close()

            SQLStatus = True
        Catch ex As Exception
            SQLStatus = False

            MsgBox("数据库操作失败！" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, 16, "错误")

            Exit Sub
        End Try
    End Sub

    ''' <summary>
    ''' 执行SQL查询语句并将结果集存入DataSet对象（返回所有结果集）
    ''' </summary>
    ''' <param name="sqlStr">SQl语句文本</param>
    Public Sub SelectToDs(ByVal sqlStr As String)
        '变量初始化
        RowsAffected = -1
        DatasDs = New DataSet

        Try
            sqlCmd = New SQLiteCommand(sqlStr, sqlConn, sqlTran)

            Dim sda As New SQLiteDataAdapter(sqlCmd)

            sda.Fill(DatasDs)

            SQLStatus = True
        Catch ex As Exception
            SQLStatus = False

            MsgBox("数据库操作失败！" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, 16, "错误")

            Exit Sub
        End Try
    End Sub

    Public Sub SaveToSql(dt As DataTable, sql As String)
        Dim sqlreader As SQLiteDataAdapter

        sqlreader = New SQLiteDataAdapter(sql, sqlConn)
        Dim ebuilder As New SQLiteCommandBuilder(sqlreader) With {
            .ConflictOption = ConflictOption.OverwriteChanges
        }

        sqlreader.Update(dt)

    End Sub

    ''' <summary>
    ''' 释放资源，关闭数据库连接
    ''' </summary>
    Public Sub Close()
        If IsNothing(sqlCmd) = False Then
            sqlCmd = Nothing
        End If

        If IsNothing(DatasDt) = False Then
            DatasDt = Nothing
        End If

        If IsNothing(DatasDs) = False Then
            DatasDs = Nothing
        End If

        If IsNothing(sqlConn) = False Then
            sqlConn.Close()
        End If
    End Sub
End Class
