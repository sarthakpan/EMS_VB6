VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form4"
   ScaleHeight     =   6660
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Payment"
      Height          =   495
      Left            =   7200
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   8040
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCLR 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   6360
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtPhno 
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox txtDOB 
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtCity 
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtEmpID 
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Phone number"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Date of Birth"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "City"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Emp No"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    Dim sql As String
    sql = "insert into emp(e_no,e_name,city,phone_no, dob) values ("
    sql = sql & "'" & txtEmpID & "',"
    sql = sql & "'" & txtName & "',"
    sql = sql & "'" & txtCity & "',"
    sql = sql & "'" & txtPhno & "',"
    sql = sql & "'" & txtDOB & "')"
    con.Execute sql
    MsgBox ("Record Inserted")
End Sub

Private Sub cmdCLR_Click()
    txtEmpID = ""
    txtName = ""
    txtCity = ""
    txtDOB = ""
    txtPhno = ""

End Sub

Private Sub cmdDel_Click()
    Dim sql As String
    sql = "DELETE * FROM emp where e_no="
    sql = sql & "'" & txtEmpID & "'"
    con.Execute sql
    MsgBox ("Record Deleted")
    txtEmpID = ""

End Sub

Private Sub cmdExit_Click()
    Unload Me
    con.Close
End Sub

Private Sub cmdUpdate_Click()
    Dim sql As String
    sql = "UPDATE emp SET e_name = '" & txtName & "', city = ' " & txtCity & " ' , phone_no = '" & txtPhno & "' , dob = '" & txtDOB & "' WHERE e_no = '" & txtEmpID & "'"
    con.Execute sql
    MsgBox ("Record Updateded")
End Sub

Private Sub Command1_Click()
    con.Close
    Form5.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Call loadcon
    Dim str As String
    str = "select * from emp"
    rs.Open "select * from emp", con, adOpenDynamic, adLockPessimistic
    rs.MoveFirst
    txtEmpID.Text = rs(0)
    txtName.Text = rs(1)
    txtCity.Text = rs(2)
    txtDOB.Text = rs(4)
    txtPhno.Text = rs(3)
End Sub

