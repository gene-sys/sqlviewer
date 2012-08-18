VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "1"
      Top             =   9060
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Страница"
      Height          =   435
      Left            =   2820
      TabIndex        =   2
      Top             =   9030
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Открыть"
      Height          =   465
      Left            =   180
      TabIndex        =   1
      Top             =   9000
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   8475
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   270
      Width           =   9375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   435
      Left            =   4380
      TabIndex        =   4
      Top             =   9030
      Width           =   5235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Con As ADODB.Connection
Attribute Con.VB_VarHelpID = -1
Private WithEvents Rst As ADODB.Recordset
Attribute Rst.VB_VarHelpID = -1
Private ttt, first As Boolean, j As Long
Private Const sPage = 40
'
'
Private Sub Command1_Click()
'
Set Con = New ADODB.Connection
Con.ConnectionTimeout = 10
Con.Open "Provider=SQLOLEDB.1;SERVER=SVR3;User ID=sa;Password=sa;Database=Alarms"
'Con.CursorLocation = adUseServer
Con.CursorLocation = adUseClient
Con.CommandTimeout = 200
Set Rst = New ADODB.Recordset
 j = 0
ttt = GetTickCount()
first = True
'Rst.Open "SELECT COUNT(ID_KKS) [countid] FROM LittleDynamic", Con, adOpenForwardOnly, adLockReadOnly
'Rst.Open "SELECT TOP 1000 COUNT(*) [countid] FROM LittleDynamic", Con, adOpenForwardOnly, adLockReadOnly
'Rst.Open "SELECT COUNT(*) [countid] FROM LittleDynamic", Con, adOpenDynamic, adLockReadOnly
Rst.Open "SELECT COUNT([MSGID]) [countid] FROM " & _
    "LittleDynamic where [TimeIn] between '2011-21-08 10:03:00.000' and " & _
    "'2011-01-09 00:00:01.000' and [textmessage] like '%X79%'", Con, _
                adOpenForwardOnly, adLockReadOnly, adAsyncExecute
'Rst.Open "SELECT  TOP 1000 [ID_KKS], [TimeIn], [TimeOut] FROM LittleDynamic " & _
 '   "where [textmessage] like '%X79%'", Con, adOpenDynamic, adLockReadOnly
'Rst.Open "SELECT [ID_KKS], [TimeIn], [TimeOut] FROM LittleDynamic " & _
'    "where [textmessage] like '%X79%'", Con, adOpenDynamic, adLockReadOnly
'Rst.Open "SELECT [ROWID], [ID_KKS], [TimeIn], [TimeOut], [MSGID] FROM (SELECT [ID_KKS], [TimeIn], [TimeOut], [MSGID], " & _
'        "ROW_NUMBER() OVER(ORDER BY MSGID DESC) AS ROWID FROM LittleDynamic " & _
'    "where [textmessage] like '%X79%') as LD WHERE [ROWID] BETWEEN 100000 AND 101000", Con, adOpenDynamic, adLockReadOnly
Do While first = True
    DoEvents
Loop
Rst.Open "SELECT [KKS], [TimeIn], [TimeOut],[MSGID] FROM " & _
        "(SELECT TOP " & sPage & " [KKS], [TimeIn], [TimeOut],[MSGID] FROM " & _
        "(SELECT TOP (" & Str(0) & " * " & sPage & "  + " & sPage & " ) " & _
        "ts.[ALM_TAGNAME] KKS, [TimeIn], [TimeOut],[MSGID] FROM LittleDynamic LD " & _
        "left outer join TagsSQL ts on ts.ID_KKS = LD.ID_KKS " & _
        "where [TimeIn] between '2011-21-08 10:03:00.000' and '2011-01-09 00:00:01.000' " & _
        "and [textmessage] like '%X79%' ORDER BY [TimeIn] ASC) SO1 " & _
        "ORDER BY [TimeIn] DESC) SO2 ORDER BY [TimeIn]", Con, adOpenDynamic, adLockReadOnly, adAsyncExecute
''Rst.MoveLast
''i = Rst.RecordCount
''Rst.MoveFirst
''ttt = (GetTickCount() - ttt) / 1000
''
''Me.Text1.Text = Me.Text1.Text & Rst.Fields(1) & " - " & ttt & "   " & Str(i) & vbCrLf
'''Do While Not Rst.EOF
'''    Me.Text1.Text = Me.Text1.Text & Rst.Fields(1) & vbCrLf
'''    Rst.MoveNext
'''Loop
'''Rst.Close
'''Rst.Open "SELECT  TOP 20 [ID_KKS], [TimeIn], [TimeOut] FROM LittleDynamic " & _
'''    "where [textmessage] like '%X79%' ORDER BY [MSGID] ASC", Con, adOpenDynamic, adLockReadOnly
'''Me.Text1.Text = Me.Text1.Text & vbCrLf
'''Do While Not Rst.EOF
'''    Me.Text1.Text = Me.Text1.Text & Rst.Fields(1) & vbCrLf
'''    Rst.MoveNext
'''Loop
''Rst.Close
''Set Rst = Nothing
''Con.Close
''Set Con = Nothing
End Sub

Private Sub Command2_Click()
Me.Text1.Text = ""
ttt = GetTickCount()
j = CLng(Me.Text2.Text)
Me.Text2.Text = Str(j + 1)
'Rst.Open "SELECT [ID_KKS], [TimeIn], [TimeOut],[MSGID] FROM " & _
'        "(SELECT TOP 40 [ID_KKS], [TimeIn], [TimeOut],[MSGID] FROM " & _
'        "(SELECT TOP (" & Str(j) & " * 40 + 40) * FROM LittleDynamic " & _
'        "where [textmessage] like '%X79%' ORDER BY [MSGID] ASC) SO1 " & _
'        "ORDER BY [MSGID] DESC) SO2 ORDER BY [MSGID]", Con, adOpenDynamic, adLockReadOnly, adAsyncExecute
Rst.Open "SELECT [KKS], [TimeIn], [TimeOut],[MSGID] FROM " & _
        "(SELECT TOP " & sPage & " [KKS], [TimeIn], [TimeOut],[MSGID] FROM " & _
        "(SELECT TOP (" & Str(j) & " * " & sPage & " + " & sPage & ") " & _
        "ts.[ALM_TAGNAME] KKS, [TimeIn], [TimeOut],[MSGID] FROM LittleDynamic LD " & _
        "left outer join TagsSQL ts on ts.ID_KKS = LD.ID_KKS " & _
        "where [TimeIn] between '2011-21-08 10:03:00.000' and '2011-01-09 00:00:01.000' " & _
        "and [textmessage] like '%X79%' ORDER BY [TimeIn] ASC) SO1 " & _
        "ORDER BY [TimeIn] DESC) SO2 ORDER BY [TimeIn]", Con, adOpenDynamic, adLockReadOnly, adAsyncExecute
End Sub

Private Sub Con_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
Dim i As Long
If first = False Then
    pRecordset.MoveLast
    i = pRecordset.RecordCount
    pRecordset.MoveFirst
    ttt = (GetTickCount() - ttt) / 1000
     Me.Text1.Text = Me.Text1.Text & ttt & " сек." & vbCrLf
    Do While Not Rst.EOF
        Me.Text1.Text = Me.Text1.Text & Rst.Fields(0) & "   " & Rst.Fields(1) & _
                                       "   " & Rst.Fields(2) & "   " & Rst.Fields(3) & vbCrLf
        Rst.MoveNext
    Loop
Else
    i = Int(pRecordset.Fields(0) / sPage)
    Me.Label1.Caption = "ВСЕГО строк = " & Str(pRecordset.Fields(0)) & "  Страниц = " & Str(i + 1) & _
                " по " & Str(sPage) & " строк " & vbCrLf
    first = False
End If
pRecordset.Close
Set pRecordset = Nothing
End Sub


''''                        SELECT * FROM
''''                        (SELECT TOP (@PageSize) * FROM
''''                            (SELECT TOP (@Page * @PageSize + @PageSize) * FROM sys.objects
''''                             ORDER BY name ASC) SO1
''''                          ORDER BY name DESC) SO2
''''                        ORDER BY name

''''            неработает как надо
''''            SELECT PriceRank, ProductName, UnitPrice FROM (SELECT ProductName, UnitPrice,
''''            ROW_NUMBER() OVER(ORDER BY UnitPrice DESC) AS PriceRank FROM Products ) AS ProductsWithRowNumber
''''            WHERE PriceRank BETWEEN 11 AND 20

''''                Микрософт советует извращаться следующим способом.
''''                Пусть например надо считать строки с 101 по 150 из вьюхи View1.
''''                Тогда записываем первые 150 строк во временную таблицу
''''
''''                SELECT TOP 150 into #temptable from View1 order by ...
''''
''''                Грохаем первые 100 записей
''''
''''                SET ROWCOUNT 100
''''                DELETE #temptable
''''
''''                (тут в принципе можно грохнуть и нужные записи, так как у DELETE нет ORDER BY :-)
''''
''''                Возвращаем на место ROWCOUNT и считываем нужные данные
''''
''''                SET ROWCOUNT 0
''''                SELECT * from #temptable order by ...



''''        Dim cn As ADODB.Connection, rs As ADODB.Recordset
''''        Set cn = New ADODB.Connection
''''        Set rs = New ADODB.Recordset
''''        cn.Open "Provider=SQLOLEDB;Server=xxx;uid=sa;pwd=;database=Northwind"
''''        rs.CursorLocation = adUseClient
''''        rs.Open "spJoinTest", cn, adOpenStatic, adLockOptimistic, adCmdStoredProc
''''        '  rs.Properties("Unique Table") = "Orders"
''''        '  rs.Properties("Resync Command") = "EXEC spJoinResync ?"
''''        rs!CustomerID = "ALFKI"
''''        rs.Update
''''        rs.Resync adAffectCurrent
''''        Debug.Print rs!CompanyName
''''        rs!CustomerID = "WOLZA"
''''        rs.Update
''''        rs.Resync adAffectCurrent
''''        Debug.Print rs!CompanyName
''''        rs.AddNew
''''        rs!EmployeeID = 1
''''        rs!CustomerID = "BLAUS"
''''        rs.UpdateS
''''        rs.Resync
''''        Debug.Print rs!CompanyName
''''        rs.Close
''''        cn.Close
            
''''    Public Sub DoStuff()
''''       On Error GoTo ErrHandler
''''        Dim oConn As ADODB.Connection
''''        Dim oCmd As ADODB.Command
''''        Dim oRS As ADODB.Recordset
''''
''''        Set oConn = New ADODB.Connection
''''        oConn.Open sConnectionString
''''        'Here we specify cursor locations to adUseClient (3)
''''        'because we will be reusing the connection afterwards
''''
''''        Set oCmd = New ADODB.Command
''''        Set oCmd.ActiveConnection = oConn
''''        oCmd.CommandText = "SELECT * FROM Authors"
''''        oCmd.CommandType = adCmdText
''''
''''        Set oRS = oCmd.Execute
''''        '...Operate on Recordset...
''''        '...and we won't be needing it anymore so let's clean up
''''        oRS.Close
''''        Set oRS = Nothing
''''        Set oCmd = Nothing
''''
''''        oConn.Execute "INSERT INTO Jobs (job_desc, min_lvl, max_lvl) values (    'edjez',25,100 )"
''''
''''        oConn.Close
''''        Set oConn = Nothing
''''
''''        GetObjectContext.SetComplete
''''       Exit Sub
''''ErrHandler:
''''        GetObjectContext.SetAbort
''''        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
''''       End Sub

