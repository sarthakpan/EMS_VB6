VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   LinkTopic       =   "Form2"
   ScaleHeight     =   6360
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User ID"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub Command2_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub Command1_Click()
    With rs
        .Open "SELECT * FROM Login", con, adOpenDynamic, adLockOptimistic

        .MoveFirst
        While Not .EOF

            If Text1 = !User_name And Text2 = !Passwd Then
                
                LoginSucceeded = True
                con.Close
                MsgBox ("Connection Closed")
                
                Form4.Show
                Unload Me
                Exit Sub
                
            ElseIf Text1 <> !User_name Then
                .MoveNext
            Else
                MsgBox "Invalid Password, try again!", , "Login"
                Exit Sub
            End If
        Wend
        .Close
    End With

End Sub

Private Sub Form_Load()
    Call loadcon
    Unload Form3
End Sub
