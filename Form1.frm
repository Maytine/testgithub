VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "โปรแกรมฐานข้อมูลลูกค้า"
   ClientHeight    =   6330
   ClientLeft      =   1920
   ClientTop       =   1530
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8340
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   1560
      Width           =   3615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2055
      Left            =   -120
      TabIndex        =   20
      Top             =   4200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3625
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   3840
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ออก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Picture         =   "Form1.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ลบข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Picture         =   "Form1.frx":2997
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "แก้ไข/บันทึก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Picture         =   "Form1.frx":5319
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "เพิ่ม/บันทึก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Picture         =   "Form1.frx":8013
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "เคลียร์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Picture         =   "Form1.frx":8CDD
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      ToolTipText     =   "ไปยังข้อมูลสุดท้าย"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      ToolTipText     =   "ไปยังข้อมูลต่อไป"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      ToolTipText     =   "ไปยังข้อมูลก่อนหน้านี้"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<"
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "ไปยังข้อมูลแรก"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "ที่เก็บรูปภาพ"
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "หมายเหตุ"
      Height          =   240
      Left            =   3240
      TabIndex        =   9
      Top             =   480
      Width           =   660
   End
   Begin VB.Label Label4 
      Caption         =   "จำนวนเงิน"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "ชื่อลูกค้า"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "รหัส"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "id"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As New ADODB.Connection
Dim RC As New ADODB.Recordset
Dim SQL As String
Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"

Private Sub Command1_Click()
With RC
            .MoveFirst
            
            Text1.Text = .Fields("id")
            Text2.Text = .Fields("รหัส")
            Text3.Text = .Fields("ชื่อ")
            Text4.Text = .Fields("จำนวนเงินที่ซื้อ")
            Text5.Text = .Fields("หมายเหตุ")
            Text6.Text = .Fields("ชื่อรูปภาพ")
            Image1.Picture = LoadPicture(App.Path & "\picture\" & Text6.Text)
End With

End Sub

Private Sub Command2_Click()
With RC
            .MovePrevious
            If .BOF = True Then .MoveLast
            Text1.Text = .Fields("id")
            Text2.Text = .Fields("รหัส")
            Text3.Text = .Fields("ชื่อ")
            Text4.Text = .Fields("จำนวนเงินที่ซื้อ")
            Text5.Text = .Fields("หมายเหตุ")
            Text6.Text = .Fields("ชื่อรูปภาพ")
            Image1.Picture = LoadPicture(App.Path & "\picture\" & Text6.Text)
End With
End Sub

Private Sub Command3_Click()
With RC
            .MoveNext
            If .EOF = True Then .MoveFirst
            Text1.Text = .Fields("id")
            Text2.Text = .Fields("รหัส")
            Text3.Text = .Fields("ชื่อ")
            Text4.Text = .Fields("จำนวนเงินที่ซื้อ")
            Text5.Text = .Fields("หมายเหตุ")
            Text6.Text = .Fields("ชื่อรูปภาพ")
            Image1.Picture = LoadPicture(App.Path & "\picture\" & Text6.Text)
End With
End Sub

Private Sub Command4_Click()
With RC
            .MoveLast
            Text1.Text = .Fields("id")
            Text2.Text = .Fields("รหัส")
            Text3.Text = .Fields("ชื่อ")
            Text4.Text = .Fields("จำนวนเงินที่ซื้อ")
            Text5.Text = .Fields("หมายเหตุ")
            Text6.Text = .Fields("ชื่อรูปภาพ")
            Image1.Picture = LoadPicture(App.Path & "\picture\" & Text6.Text)
End With

End Sub

Private Sub Command5_Click()
With RC
        .AddNew
            .Fields("รหัส") = Text2.Text
            .Fields("ชื่อ") = Text3.Text
            .Fields("จำนวนเงินที่ซื้อ") = Text4.Text
            .Fields("หมายเหตุ") = Text5.Text
            .Fields("ชื่อรูปภาพ") = Text6.Text
        .Update
        MsgBox "เพิ่มข้อมูลเรียบร้อยแล้วครับ", vbInformation, "เพิ่มข้อมูล"
End With

Call Form_Load
End Sub

Private Sub Command6_Click()
With RC
        Dim sname As String
        sname = .Fields("ชื่อ")
        .Delete
        .Requery
        MsgBox "ลบข้อมูลของ " & sname & " เรียบร้อยแล้วครับ", vbInformation, "ลบข้อมูล"
End With

Call Form_Load

End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Command8_Click()
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
End Sub

Private Sub Command9_Click()
With RC
            .Fields("รหัส") = Text2.Text
            .Fields("ชื่อ") = Text3.Text
            .Fields("จำนวนเงินที่ซื้อ") = Text4.Text
            .Fields("หมายเหตุ") = Text5.Text
            .Fields("ชื่อรูปภาพ") = Text6.Text
            MsgBox "แก้ไขข้อมูลของ " & .Fields("ชื่อ") & " เรียบร้อยแล้วครับ", vbInformation, "แก้ไขข้อมูล"
End With

Call Form_Load

End Sub

Private Sub Form_Load()
'เปิดฐานข้อมูล
With Conn
        If .State = 1 Then .Close
        .ConnectionString = strConn & ";Data Source=" & App.Path & "\db\customer.mdb"
        .Open
End With

'เปิดตาราง
With RC
        SQL = "SELECT * FROM myCustomer  ORDER BY id ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, Conn, 2, 3
        
        Label6.Caption = "มีลูกค้าทั้งหมด " & .RecordCount & " คน"
        If .RecordCount <> 0 Then
            Text1.Text = .Fields("id")
            Text2.Text = .Fields("รหัส")
            Text3.Text = .Fields("ชื่อ")
            Text4.Text = .Fields("จำนวนเงินที่ซื้อ")
            Text5.Text = .Fields("หมายเหตุ")
            Text6.Text = .Fields("ชื่อรูปภาพ")
            Image1.Picture = LoadPicture(App.Path & "\picture\" & Text6.Text)
        End If
        
With Adodc1
        .ConnectionString = strConn & ";Data Source=" & App.Path & "\db\customer.mdb"
        .RecordSource = SQL
        .Refresh
End With

End With


End Sub


