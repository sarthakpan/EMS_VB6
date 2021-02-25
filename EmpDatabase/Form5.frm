VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10410
   LinkTopic       =   "Form5"
   ScaleHeight     =   6360
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Report by Employee"
      Height          =   735
      Left            =   6720
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Report By Month"
      Height          =   735
      Left            =   6600
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Month of Pay"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Salary"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "User ID"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Index"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    Unload Me
    con.Close
End Sub

Private Sub Command1_Click()
    Dim sql As String
    sql = "insert into Salary (ID,User_ID,Sal,Mon)values("
    sql = sql & "'" & Text1 & "',"
    sql = sql & "'" & Text2 & "',"
    sql = sql & "'" & Text3 & "',"
    sql = sql & "'" & Text4 & "')"
    con.Execute sql
    MsgBox ("Salary paid")
End Sub

Private Sub Command3_Click()
    'con.Close
    DataReport1.Show
End Sub

Private Sub Command4_Click()
    'con.Close
    DataReport2.Show
End Sub

Private Sub Form_Load()
    Call loadcon
End Sub


