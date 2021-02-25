VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu opt 
      Caption         =   "Options"
      Index           =   1
      Begin VB.Menu Form4 
         Caption         =   "Employee Details"
         Index           =   1
      End
      Begin VB.Menu Form5 
         Caption         =   "Salary Pay"
         Index           =   2
      End
   End
   Begin VB.Menu Rptl 
      Caption         =   "Report"
      Index           =   2
      Begin VB.Menu RptMon 
         Caption         =   "By Month"
         Index           =   1
      End
      Begin VB.Menu RptEmp 
         Caption         =   "By Employee"
         Index           =   2
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form4_Click(Index As Integer)
    Form4.Show
End Sub

Private Sub Form5_Click(Index As Integer)
    Form5.Show
End Sub

