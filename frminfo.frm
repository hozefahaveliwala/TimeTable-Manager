VERSION 5.00
Begin VB.Form frminfo 
   Caption         =   "S.G.S.I.T.S"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdexam 
      Caption         =   "EXAMINATION TIMETABLE"
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECT OPTION"
      Height          =   2055
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   7935
      Begin VB.CommandButton cmdclass 
         Caption         =   "CLASS TIMETABLE"
         Height          =   855
         Left            =   4680
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   7080
      ScaleHeight     =   1455
      ScaleWidth      =   1485
      TabIndex        =   2
      Top             =   120
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   600
      ScaleHeight     =   1455
      ScaleWidth      =   1485
      TabIndex        =   1
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label Label2 
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   720
      TabIndex        =   7
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "S.G.S.I.T.S TIME TABLE MANAGMENT SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frminfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclass_Click()
Unload Me
frmmain.Show
End Sub

Private Sub cmdexam_Click()
Unload Me
frmexaminfo.Show
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub Form_Load()
Picture1.Picture = LoadPicture(App.Path & "\gslogo.JPG")
Picture2.Picture = LoadPicture(App.Path & "\gslogo.JPG")
Label2.Caption = "Developed BY:" & vbCrLf & "Hozefa Haveliwala B.E.(CSE)" & vbCrLf & "Vivek Rawat B.E.(CSE)" & vbCrLf & "with the guidance of Ms Deepika."
End Sub
