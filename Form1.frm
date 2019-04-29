VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame dataPane 
      BackColor       =   &H00FFFFC0&
      Height          =   5295
      Left            =   0
      TabIndex        =   17
      Top             =   3000
      Width           =   12255
      Begin VB.CommandButton btnRefresh 
         Caption         =   "Refresh"
         DisabledPicture =   "Form1.frx":0000
         Height          =   495
         Left            =   9720
         Picture         =   "Form1.frx":2726
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   7200
         Top             =   240
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=E:\MYBACKUP\VB6 PROJECTS\CRUD_VB6_Acess\CRUDVB6.mdb"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=E:\MYBACKUP\VB6 PROJECTS\CRUD_VB6_Acess\CRUDVB6.mdb"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "employee"
         Caption         =   "Previous Next"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dgEmployee 
         Bindings        =   "Form1.frx":4E4C
         Height          =   4455
         Left            =   -1560
         TabIndex        =   20
         Top             =   840
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483628
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
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
               LCID            =   1033
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
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   -1  'True
            AllowSizing     =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton btnFind 
         Caption         =   "Search"
         Height          =   495
         Left            =   6120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox btnSearch 
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Employee Details"
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      Begin VB.ComboBox cbCounty 
         DataField       =   "county"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":4E61
         Left            =   7560
         List            =   "Form1.frx":4E74
         TabIndex        =   15
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox cbDesignation 
         DataField       =   "designation"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":4EA2
         Left            =   7440
         List            =   "Form1.frx":4EB5
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "DELETE"
         Height          =   495
         Index           =   3
         Left            =   7800
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "UPDATE"
         Height          =   495
         Index           =   2
         Left            =   6120
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton btnadd 
         Caption         =   "ADD"
         Height          =   495
         Index           =   1
         Left            =   4200
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton btnNew 
         Caption         =   "NEW"
         Height          =   495
         Index           =   0
         Left            =   2160
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtMobileNo 
         DataField       =   "tel"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtIdno 
         DataField       =   "idno"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   6
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtOthername 
         DataField       =   "othernames"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtSurname 
         DataField       =   "surname"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "Home County"
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Designation"
         Height          =   255
         Left            =   6360
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Mobile No"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "IDNO/PassPort No"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Other Names"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Surname"
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub btnadd_Click(Index As Integer)
Adodc1.Recordset.Update
End Sub

Private Sub btnFind_Click()

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CRUDVB6.mdb;Persist Security Info=False"
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from employee where idno like '" & btnSearch.Text & "'"
    
    
If Adodc1.Recordset.RecordCount > 0 Then
   'txtSurname.Text = Adodc1.Recordset.Fields(0)
   MsgBox ("Value Exist")
   Adodc1.Refresh
End If


End Sub

Private Sub btnNew_Click(Index As Integer)
Adodc1.Recordset.AddNew
End Sub

Private Sub btnRefresh_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CRUDVB6.mdb;Persist Security Info=False"
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from employee order by idno asc"
    
    
If Adodc1.Recordset.RecordCount > 0 Then
   'txtSurname.Text = Adodc1.Recordset.Fields(0)
   MsgBox ("You are loading the list again")
   Adodc1.Refresh
End If
End Sub

Private Sub btnUpdate_Click(Index As Integer)
Adodc1.Recordset.Update

End Sub

Private Sub Command1_Click(Index As Integer)
Adodc1.Recordset.Delete
End Sub
