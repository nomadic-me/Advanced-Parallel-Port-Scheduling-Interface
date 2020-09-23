VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Advanced Parallel Port - Time Based Programmable Interface"
   ClientHeight    =   4275
   ClientLeft      =   3645
   ClientTop       =   5505
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   7245
   Begin VB.CommandButton Command6 
      Caption         =   "View / Edit Events"
      Height          =   495
      Left            =   2640
      TabIndex        =   33
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6120
      Top             =   2280
   End
   Begin VB.CommandButton Command4 
      Caption         =   "S&top"
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmbPub 
      Caption         =   "&Publish"
      Height          =   495
      Left            =   5760
      TabIndex        =   25
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdhelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   4560
      TabIndex        =   24
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4200
      TabIndex        =   14
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CheckBox Check8 
      Caption         =   "7"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox Check7 
      Caption         =   "6"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox Check6 
      Caption         =   "5"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.CheckBox Check5 
      Caption         =   "4"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.CheckBox Check4 
      Caption         =   "3"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Caption         =   "2"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "1"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "0"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "All On"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "All Off"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   2760
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buttons"
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   6855
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Current File name"
      Height          =   195
      Left            =   2280
      TabIndex        =   32
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3720
      TabIndex        =   31
      Top             =   120
      Width           =   15
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   495
      Left            =   3000
      TabIndex        =   30
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Calendar Controls"
      Height          =   195
      Left            =   240
      TabIndex        =   29
      Top             =   360
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderStyle     =   6  'Inside Solid
      Height          =   2655
      Left            =   120
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "( For The Above Mentioned Dat And Time Time )"
      Height          =   195
      Left            =   2160
      TabIndex        =   23
      Top             =   1920
      Width           =   3450
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(Hint :: For all pins to be ""Off""/""On"" just click on ""All Off""/""All On"" Button Below)"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   5685
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Status Of Pins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   21
      Top             =   1800
      Width           =   1755
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Minutes (mm)"
      Height          =   195
      Left            =   3840
      TabIndex        =   20
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hours (hh)"
      Height          =   195
      Left            =   2880
      TabIndex        =   19
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Day (dd)"
      Height          =   195
      Left            =   1560
      TabIndex        =   18
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Month (mm)"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter The Following Data For Setting Up your calender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   6600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   15
      Top             =   1320
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bin As Variant
Dim filestr As String
Dim filenum As Variant
Dim m As Variant
Dim strfilename As String





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

Private Sub cmbPub_Click()
Dim master_str As String




If ((Text1 = "") Or (Val(Text1) > 12) Or (Val(Text1) < 1)) Then
MsgBox ("Enter Valid Number For Month (ie. 1 - 12) ")
Else


If ((Text2 = "") Or (Text2 > 31) Or (Text2 < 1)) Then
MsgBox ("Enter Valid Number For Day (ie. 1 - 31) ")
Else


If ((Text3 = "") Or (Text3 > 23) Or (Text3 < 1)) Then
MsgBox ("Enter Valid Number For Hour (ie. 0 - 23) ")
Else

If ((Text4 = "") Or (Text4 > 59) Or (Text4 < 1)) Then
MsgBox ("Enter Valid Number For Minute (ie. 0 - 59) ")
Else

master_str = "month " + Text1 + "day " + Text2 + "hour " + Text3 + "minute " + Text4 + ".evnt"



Open master_str For Output As #1
Print #1, Str(bin)
Close #1

End If
End If
End If
End If

End Sub

Private Sub cmdhelp_Click()
frmHelp.Show
frmHelp.Visible = True
End Sub

Private Sub cmdSetup_Click()

End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
bin = 0
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0

End Sub

Private Sub Command3_Click()
bin = 255

Check1.Value = 1
Check2.Value = 1
Check3.Value = 1
Check4.Value = 1
Check5.Value = 1
Check6.Value = 1
Check7.Value = 1
Check8.Value = 1
End Sub


Private Sub Command4_Click()
Timer1.Enabled = False
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
frmView.Show
frmView.Visible = True
End Sub

Private Sub Form_Load()
Call PortOut(888, 255)
bin = 0
End Sub

Private Sub Timer1_Timer()

strfilename = "month" + Str(Month(Date)) + "day" + Str(Day(Date)) + "hour" + Str(Hour(Time)) + "minute" + Str(Minute(Time)) + ".evnt"
Label12.Caption = strfilename
On Error GoTo trap
filenum = FreeFile
Open strfilename For Input As filenum
m = Val(Input(LOF(filenum), filenum))
Call PortOut(888, m)
Close filenum
Kill (strfilename)
Exit Sub
Label12.Caption = strfilename
trap:
Exit Sub
End Sub
