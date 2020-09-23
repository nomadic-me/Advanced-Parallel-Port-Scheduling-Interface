VERSION 5.00
Begin VB.Form frmView 
   Caption         =   "View / Edit Events"
   ClientHeight    =   3525
   ClientLeft      =   5250
   ClientTop       =   1140
   ClientWidth     =   5370
   LinkTopic       =   "Form2"
   ScaleHeight     =   3525
   ScaleWidth      =   5370
   Begin VB.CommandButton Command3 
      Caption         =   "&Publish"
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "1"
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   3000
      Width           =   495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "2"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox Check4 
      Caption         =   "3"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox Check5 
      Caption         =   "4"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox Check6 
      Caption         =   "5"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox Check7 
      Caption         =   "6"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   3000
      Width           =   495
   End
   Begin VB.CheckBox Check8 
      Caption         =   "7"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3480
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Hide"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   120
      Pattern         =   "*.evnt"
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Status Of Pins For Above Selected Event"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   2925
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "month(m) day(dd) hour(hh) minute (mm) .evnt"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "( Format Of File Name ::)"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List Of Events"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1425
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bin As Variant
Dim filenum As Variant

Private Sub cmdrefresh_Click()
File1.Refresh
End Sub

Private Sub Command1_Click()
frmView.Visible = False
End Sub


Private Sub Check1_Click()
bin = ((2 ^ (Val(Check1.Caption))) * Val(Check1.Value)) + ((2 ^ (Val(Check2.Caption))) * Val(Check2.Value)) + ((2 ^ (Val(Check3.Caption))) * Val(Check3.Value)) + ((2 ^ (Val(Check4.Caption))) * Val(Check4.Value)) + ((2 ^ (Val(Check5.Caption))) * Val(Check5.Value)) + ((2 ^ (Val(Check6.Caption))) * Val(Check6.Value)) + ((2 ^ (Val(Check7.Caption))) * Val(Check7.Value)) + ((2 ^ (Val(Check8.Caption))) * Val(Check8.Value))


End Sub

Private Sub Check2_Click()
bin = ((2 ^ (Val(Check1.Caption))) * Val(Check1.Value)) + ((2 ^ (Val(Check2.Caption))) * Val(Check2.Value)) + ((2 ^ (Val(Check3.Caption))) * Val(Check3.Value)) + ((2 ^ (Val(Check4.Caption))) * Val(Check4.Value)) + ((2 ^ (Val(Check5.Caption))) * Val(Check5.Value)) + ((2 ^ (Val(Check6.Caption))) * Val(Check6.Value)) + ((2 ^ (Val(Check7.Caption))) * Val(Check7.Value)) + ((2 ^ (Val(Check8.Caption))) * Val(Check8.Value))

End Sub

Private Sub Check3_Click()
bin = ((2 ^ (Val(Check1.Caption))) * Val(Check1.Value)) + ((2 ^ (Val(Check2.Caption))) * Val(Check2.Value)) + ((2 ^ (Val(Check3.Caption))) * Val(Check3.Value)) + ((2 ^ (Val(Check4.Caption))) * Val(Check4.Value)) + ((2 ^ (Val(Check5.Caption))) * Val(Check5.Value)) + ((2 ^ (Val(Check6.Caption))) * Val(Check6.Value)) + ((2 ^ (Val(Check7.Caption))) * Val(Check7.Value)) + ((2 ^ (Val(Check8.Caption))) * Val(Check8.Value))

End Sub

Private Sub Check4_Click()
bin = ((2 ^ (Val(Check1.Caption))) * Val(Check1.Value)) + ((2 ^ (Val(Check2.Caption))) * Val(Check2.Value)) + ((2 ^ (Val(Check3.Caption))) * Val(Check3.Value)) + ((2 ^ (Val(Check4.Caption))) * Val(Check4.Value)) + ((2 ^ (Val(Check5.Caption))) * Val(Check5.Value)) + ((2 ^ (Val(Check6.Caption))) * Val(Check6.Value)) + ((2 ^ (Val(Check7.Caption))) * Val(Check7.Value)) + ((2 ^ (Val(Check8.Caption))) * Val(Check8.Value))

End Sub

Private Sub Check5_Click()
bin = ((2 ^ (Val(Check1.Caption))) * Val(Check1.Value)) + ((2 ^ (Val(Check2.Caption))) * Val(Check2.Value)) + ((2 ^ (Val(Check3.Caption))) * Val(Check3.Value)) + ((2 ^ (Val(Check4.Caption))) * Val(Check4.Value)) + ((2 ^ (Val(Check5.Caption))) * Val(Check5.Value)) + ((2 ^ (Val(Check6.Caption))) * Val(Check6.Value)) + ((2 ^ (Val(Check7.Caption))) * Val(Check7.Value)) + ((2 ^ (Val(Check8.Caption))) * Val(Check8.Value))

End Sub

Private Sub Check6_Click()
bin = ((2 ^ (Val(Check1.Caption))) * Val(Check1.Value)) + ((2 ^ (Val(Check2.Caption))) * Val(Check2.Value)) + ((2 ^ (Val(Check3.Caption))) * Val(Check3.Value)) + ((2 ^ (Val(Check4.Caption))) * Val(Check4.Value)) + ((2 ^ (Val(Check5.Caption))) * Val(Check5.Value)) + ((2 ^ (Val(Check6.Caption))) * Val(Check6.Value)) + ((2 ^ (Val(Check7.Caption))) * Val(Check7.Value)) + ((2 ^ (Val(Check8.Caption))) * Val(Check8.Value))

End Sub

Private Sub Check7_Click()
bin = ((2 ^ (Val(Check1.Caption))) * Val(Check1.Value)) + ((2 ^ (Val(Check2.Caption))) * Val(Check2.Value)) + ((2 ^ (Val(Check3.Caption))) * Val(Check3.Value)) + ((2 ^ (Val(Check4.Caption))) * Val(Check4.Value)) + ((2 ^ (Val(Check5.Caption))) * Val(Check5.Value)) + ((2 ^ (Val(Check6.Caption))) * Val(Check6.Value)) + ((2 ^ (Val(Check7.Caption))) * Val(Check7.Value)) + ((2 ^ (Val(Check8.Caption))) * Val(Check8.Value))

End Sub

Private Sub Check8_Click()
bin = ((2 ^ (Val(Check1.Caption))) * Val(Check1.Value)) + ((2 ^ (Val(Check2.Caption))) * Val(Check2.Value)) + ((2 ^ (Val(Check3.Caption))) * Val(Check3.Value)) + ((2 ^ (Val(Check4.Caption))) * Val(Check4.Value)) + ((2 ^ (Val(Check5.Caption))) * Val(Check5.Value)) + ((2 ^ (Val(Check6.Caption))) * Val(Check6.Value)) + ((2 ^ (Val(Check7.Caption))) * Val(Check7.Value)) + ((2 ^ (Val(Check8.Caption))) * Val(Check8.Value))

End Sub

Private Sub Command2_Click()
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Check5.Enabled = True
Check6.Enabled = True
Check7.Enabled = True
Check8.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True

If Check1.Enabled = True Then

Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Check8.Enabled = False



Open File1.FileName For Output As #2
Print #2, Str(bin)
Close #2

Else

MsgBox ("You Need To Edit The Status Of Pins For A Event To Overwrite It!!")
End If
End Sub

Private Sub File1_Click()
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Check8.Enabled = False

End Sub

Private Sub File1_DblClick()
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Check8.Enabled = False

End Sub

Private Sub Timer1_Timer()

Dim Temp, Temp1 As Long
 Dim i As Integer
 Dim res As String
 Dim Done As Boolean
 Dim m As Variant
Dim a As Variant
Dim b As Variant




On Error GoTo trap
filenum = FreeFile
Open File1.FileName For Input As filenum
m = Val(Input(LOF(filenum), filenum))
Close filenum




 
 
 Temp = m
 Temp1 = m
 Do Until Temp \ 2 = 1
  Temp = Temp \ 2
  i = i + 1
 Loop
 res = ""
 For j = i + 1 To 0 Step -1
  If Temp1 - 2 ^ j > 0 And Done = False Then
   Temp1 = Temp1 - 2 ^ j
   res = res & "1"
  ElseIf Temp1 - 2 ^ j <> 0 And Done = False Then
   res = res & "0"
  ElseIf Temp1 - 2 ^ j = 0 And Done = False Then
   res = res & "1"
   Done = True
  ElseIf Done = True Then
   res = res & "0"
  End If
 Next
 



    a = Val(res)
    
    b = Int(a) - (Int(a / 10) * 10)
    Check1.Value = b
    a = Int(a / 10)
    
    b = Int(a) - (Int(a / 10) * 10)
    Check2.Value = b
    a = Int(a / 10)

    b = Int(a) - (Int(a / 10) * 10)
    Check3.Value = b
    a = Int(a / 10)

    b = Int(a) - (Int(a / 10) * 10)
    Check4.Value = b
    a = Int(a / 10)

    b = Int(a) - (Int(a / 10) * 10)
    Check5.Value = b
    a = Int(a / 10)

    b = Int(a) - (Int(a / 10) * 10)
    Check6.Value = b
    a = Int(a / 10)
    
    b = Int(a) - (Int(a / 10) * 10)
    Check7.Value = b
    a = Int(a / 10)

    b = Int(a) - (Int(a / 10) * 10)
    Check8.Value = b
    a = Int(a / 10)

trap:


Exit Sub
End Sub
