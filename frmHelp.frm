VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close ""Help"" Window"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   $"frmHelp.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "2.  If The Date Of A Event Is Repeated It Will Simply Be Over-Written Without Any Promting"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "1.  There Can Be Maximum Of One Change Every Minute"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmHelp.Visible = False

End Sub
