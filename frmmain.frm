VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmmain 
   Caption         =   "S.G.S.I.T.S"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhome 
      Caption         =   "HOME"
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   6600
      Width           =   1785
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   7920
      ScaleHeight     =   1455
      ScaleWidth      =   1485
      TabIndex        =   24
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
      Left            =   1080
      ScaleHeight     =   1455
      ScaleWidth      =   1485
      TabIndex        =   23
      Top             =   120
      Width           =   1515
   End
   Begin VSFlex8Ctl.VSFlexGrid vstt 
      Height          =   735
      Left            =   360
      TabIndex        =   22
      Top             =   5760
      Visible         =   0   'False
      Width           =   6135
      _cx             =   10821
      _cy             =   1296
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   8880
      TabIndex        =   21
      Top             =   6600
      Width           =   1785
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8520
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   6840
      TabIndex        =   17
      Top             =   6600
      Width           =   1785
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   6600
      Width           =   1785
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "SHOW DETAILS"
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   6600
      Width           =   1785
   End
   Begin VB.Frame Frame2 
      Caption         =   "DAY / TIME / SUBJECT / FACULTY"
      Height          =   1935
      Left            =   360
      TabIndex        =   8
      Top             =   3840
      Width           =   10215
      Begin VB.TextBox txtsub 
         Height          =   285
         Left            =   5040
         TabIndex        =   20
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   7200
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox combday 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox combtime 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "SUBJECT"
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "FACULTY NAME"
         Height          =   255
         Left            =   7680
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "TIMING"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "DAY"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "YEAR / BRANCH / SECTION"
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   10215
      Begin VB.ComboBox combsection 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox combbranch 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   960
         Width           =   2895
      End
      Begin VB.ComboBox combyear 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "SECTION"
         Height          =   255
         Left            =   6480
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "BRANCH"
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "YEAR"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   600
      Width           =   15
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
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i, j As Integer
Dim d, t As String


Private Sub cmdadd_Click()
If (combyear.ListIndex = -1) Or (combday.ListIndex = -1) Or (combtime.ListIndex = -1) Or (txtname.Text = "") Or (txtsub.Text = "") Then
 j = -1
ElseIf (combbranch.Enabled = True) And (combbranch.ListIndex = -1) Then
    j = -1
ElseIf (combsection.Enabled = True) And (combsection.ListIndex = -1) Then
    j = -1
End If
 
 
If j = -1 Then
  MsgBox "Entry Missing Try Again", vbCritical, "invalid entry"
 cmdclear_Click
 
Else
 
 Adodc1.RecordSource = "select days.day_no, timing.time_no from days, timing where days.day ='" & combday.Text & "'" & "and timing.timing='" & combtime.Text & "'"
Adodc1.Refresh
 d = Adodc1.Recordset.Fields("day_no")
 t = Adodc1.Recordset.Fields("time_no")
 
 i = combyear.ListIndex
 If i = 0 Then
 
 Adodc1.RecordSource = "select year1.class_name, days.day,timing.timing,year1.subject,year1.faculty from year1, days,timing where year1.class_name = '" & combsection.Text & "'" & " and days.day = '" & combday.Text & "' and timing.timing = '" & combtime.Text & "'" & "and year1.day=days.day_no and year1.timing=timing.time_no"
 ElseIf i = 1 Then
 Adodc1.RecordSource = "select year2.branch, days.day,timing.timing,year2.subject,year2.faculty from year2, days,timing where year2.branch = '" & combbranch.Text & "'" & " and days.day = '" & combday.Text & "' and timing.timing = '" & combtime.Text & "'" & "and year2.day=days.day_no and year2.timing=timing.time_no"
 ElseIf i = 2 Then
 Adodc1.RecordSource = "select year3.branch, days.day,timing.timing,year3.subject,year3.faculty from year3, days,timing where year3.branch = '" & combbranch.Text & "'" & " and days.day = '" & combday.Text & "' and timing.timing = '" & combtime.Text & "'" & "and year3.day=days.day_no and year3.timing=timing.time_no"
 ElseIf i = 3 Then
 Adodc1.RecordSource = "select year4.branch, days.day,timing.timing,year4.subject,year4.faculty from year4, days,timing where year4.branch = '" & combbranch.Text & "'" & " and days.day = '" & combday.Text & "' and timing.timing = '" & combtime.Text & "'" & "and year4.day=days.day_no and year4.timing=timing.time_no"
 End If
 
 Adodc1.Refresh
 
 If Adodc1.Recordset.RecordCount = 0 Then
 
 If i = 0 Then
 Adodc1.RecordSource = "select * from year1"
 ElseIf i = 1 Then
 Adodc1.RecordSource = "select * from year2"
 ElseIf i = 2 Then
 Adodc1.RecordSource = "select * from year3"
 ElseIf i = 3 Then
 Adodc1.RecordSource = "select * from year4"
 End If
 
 Adodc1.Refresh
 
 Adodc1.Recordset.AddNew
  If (i = 0) Then
  Adodc1.Recordset.Fields("class_name") = combsection.Text
  Else
  Adodc1.Recordset.Fields("branch") = combbranch.Text
  End If
 
 Adodc1.Recordset.Fields("day") = d
 Adodc1.Recordset.Fields("timing") = t
 Adodc1.Recordset.Fields("subject") = txtsub.Text
 Adodc1.Recordset.Fields("faculty") = txtname.Text
 
 Adodc1.Recordset.Update

 MsgBox "Timetable Updated", vbInformation, "valid entry"
 cmdclear_Click

Else
Set vstt.DataSource = Adodc1
vstt.Visible = True
MsgBox " Period not available", vbCritical
cmdclear_Click
End If
End If

End Sub

Private Sub cmdclear_Click()
j = 0
combyear.ListIndex = -1
combbranch.ListIndex = -1
combsection.ListIndex = -1
combday.ListIndex = -1
combtime.ListIndex = -1
txtname.Text = ""
txtsub.Text = ""
vstt.Clear
vstt.Visible = False
combbranch.Enabled = False
combsection.Enabled = False


End Sub


Private Sub cmdexit_Click()
End
End Sub



Private Sub cmdhome_Click()
Unload Me
frminfo.Show
End Sub

Private Sub cmdshow_Click()
frmprint.Show
Unload Me
End Sub

Private Sub combyear_Click()
If combyear.ListIndex = 0 Then
combsection.Enabled = True
If combbranch.Enabled = True Then
combbranch.ListIndex = -1
combbranch.Enabled = False
End If

Else
combbranch.Enabled = True
If combsection.Enabled = True Then
combsection.ListIndex = -1
combsection.Enabled = False
End If
End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "timetable.mdb;Persist Security Info=False"

Picture1.Picture = LoadPicture(App.Path & "\gslogo.JPG")
Picture2.Picture = LoadPicture(App.Path & "\gslogo.JPG")

combyear.AddItem ("First Year")
combyear.AddItem ("Second Year")
combyear.AddItem ("Third Year")
combyear.AddItem ("Fourth Year")

combsection.AddItem ("Section A")
combsection.AddItem ("Section B")
combsection.AddItem ("Section C")
combsection.AddItem ("Section D")
combsection.AddItem ("Section E")
combsection.AddItem ("Section F")
combsection.AddItem ("Section G")
combsection.AddItem ("Section H")
combsection.AddItem ("Section I")
combsection.AddItem ("Section J")

combbranch.AddItem ("Biomedical")
combbranch.AddItem ("Civil")
combbranch.AddItem ("CS")
combbranch.AddItem ("Electrical")
combbranch.AddItem ("EI")
combbranch.AddItem ("E&TC")
combbranch.AddItem ("IP")
combbranch.AddItem ("IT")

combday.AddItem ("Monday")
combday.AddItem ("Tuesday")
combday.AddItem ("Wednesday")
combday.AddItem ("Thursday")
combday.AddItem ("Friday")

combtime.AddItem ("09am-10am")
combtime.AddItem ("10am-11am")
combtime.AddItem ("11am-12pm")
combtime.AddItem ("12pm-01pm")
combtime.AddItem ("02pm-03pm")
combtime.AddItem ("03pm-04pm")
combtime.AddItem ("04pm-05pm")
combtime.AddItem ("05pm-06pm")

End Sub
