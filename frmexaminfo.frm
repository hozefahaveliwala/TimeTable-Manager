VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmexaminfo 
   Caption         =   "S.G.S.I.T.S"
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin MSACAL.Calendar Calendar1 
      Height          =   2055
      Left            =   360
      TabIndex        =   41
      Top             =   2280
      Width           =   2415
      _Version        =   524288
      _ExtentX        =   4260
      _ExtentY        =   3625
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2012
      Month           =   3
      Day             =   22
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsexam 
      Height          =   855
      Left            =   360
      TabIndex        =   40
      Top             =   6720
      Visible         =   0   'False
      Width           =   12375
      _cx             =   21828
      _cy             =   1508
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
      HighLight       =   1
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
      ScrollBars      =   3
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
   Begin VB.CommandButton cmdhome 
      Caption         =   "HOME"
      Height          =   495
      Left            =   600
      TabIndex        =   39
      Top             =   7920
      Width           =   1900
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10560
      Top             =   7560
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
   Begin VB.TextBox txtintno 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   2640
      MaxLength       =   11
      TabIndex        =   32
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   9360
      ScaleHeight     =   1455
      ScaleWidth      =   1485
      TabIndex        =   12
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
      Left            =   1440
      ScaleHeight     =   1455
      ScaleWidth      =   1485
      TabIndex        =   11
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   9960
      TabIndex        =   10
      Top             =   7920
      Width           =   1900
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   7920
      Width           =   1900
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   7920
      Width           =   1900
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "SHOW DETAILS"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   7920
      Width           =   1900
   End
   Begin VB.Frame Frame2 
      Caption         =   " FACULTY DETAILS"
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   12255
      Begin VB.TextBox txtextno 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   1
         EndProperty
         Height          =   495
         Left            =   10200
         MaxLength       =   11
         TabIndex        =   37
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtextadd 
         Height          =   495
         Left            =   6960
         TabIndex        =   36
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtext 
         Height          =   495
         Left            =   4560
         TabIndex        =   35
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtint 
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "EXTERNAL PHONE NO"
         Height          =   255
         Left            =   9960
         TabIndex        =   38
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "EXTERNAL ADDRESS"
         Height          =   255
         Left            =   7560
         TabIndex        =   34
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "EXTERNAL NAME"
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "INTERNAL PHONE NO"
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "INTERNAL NAME"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CLASS DETAILS"
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   12375
      Begin VB.TextBox txtlabat 
         Height          =   495
         Left            =   10080
         TabIndex        =   25
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtlabtech 
         Height          =   495
         Left            =   7560
         TabIndex        =   24
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtlab 
         Height          =   495
         Left            =   5040
         TabIndex        =   23
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox combsection 
         Height          =   315
         Left            =   8880
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox combbranch 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox combyear 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox combtime 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtsub 
         Height          =   495
         Left            =   2640
         TabIndex        =   15
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtsubcode 
         Height          =   495
         Left            =   10680
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "LAB ATTENDENT"
         Height          =   255
         Left            =   10200
         TabIndex        =   28
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "LAB TECHNICIAN"
         Height          =   255
         Left            =   7680
         TabIndex        =   27
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "LAB NAME"
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "SECTION"
         Height          =   255
         Left            =   8760
         TabIndex        =   22
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "BRANCH"
         Height          =   255
         Left            =   6720
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "YEAR"
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "SUBJECT NAME"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "SUBJECT CODE"
         Height          =   255
         Left            =   10440
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "TIME"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "DATE"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   840
      TabIndex        =   9
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
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmexaminfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i, j As Integer
Dim d As String



Private Sub cmdadd_Click()

d = Calendar1.Value
If (combyear.ListIndex = -1) Or (combtime.ListIndex = -1) Or (txtsub.Text = "") Or (txtsubcode.Text = "") Or (txtlab.Text = "") Or (txtlabat.Text = "") Or (txtlabtech.Text = "") Then
 j = -1
ElseIf (txtint.Text = "") Or (txtintno.Text = "") Or (txtext.Text = "") Or (txtextadd.Text = "") Or (txtextno.Text = "") Then
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
 
 Adodc1.RecordSource = " select * from exam where val(exam_date) =" & Val(d) & " and lab_name='" & txtlab.Text & "' and exam_time ='" & combtime.Text & "'"
Adodc1.Refresh
 
 
 If Adodc1.Recordset.RecordCount = 0 Then
 
 Adodc1.Recordset.AddNew
 Adodc1.Recordset.Fields("exam_date") = d
 Adodc1.Recordset.Fields("exam_time") = combtime.Text
 Adodc1.Recordset.Fields("class") = combyear.Text
 Adodc1.Recordset.Fields("subject_code") = txtsubcode.Text
 Adodc1.Recordset.Fields("subject") = txtsub.Text
 Adodc1.Recordset.Fields("internal_name") = txtint.Text
 Adodc1.Recordset.Fields("internal_no") = txtintno.Text
 Adodc1.Recordset.Fields("external_name") = txtext.Text
 Adodc1.Recordset.Fields("external_add") = txtextadd.Text
 Adodc1.Recordset.Fields("external_no") = txtextno.Text
 Adodc1.Recordset.Fields("lab_name") = txtlab.Text
 Adodc1.Recordset.Fields("lab_technician") = txtlabtech.Text
 Adodc1.Recordset.Fields("lab_attendent") = txtlabat.Text
 If combsection.Enabled = True Then
 Adodc1.Recordset.Fields("branch") = combsection.Text
 ElseIf combbranch.Enabled = True Then
 Adodc1.Recordset.Fields("branch") = combbranch.Text
 End If
 Adodc1.Recordset.Update

 MsgBox "Timetable Updated", vbInformation, "valid entry"
 cmdclear_Click

Else
Set vsexam.DataSource = Adodc1
vsexam.Visible = True
MsgBox " not available", vbCritical

End If
End If
cmdclear_Click
End Sub

Private Sub cmdclear_Click()
j = 0
combyear.ListIndex = -1
combbranch.ListIndex = -1
combsection.ListIndex = -1
combtime.ListIndex = -1
txtsubcode.Text = ""
txtsub.Text = ""
txtlab.Text = ""
txtlabtech.Text = ""
txtlabat.Text = ""
txtint.Text = ""
txtintno.Text = ""
txtext.Text = ""
txtextadd.Text = ""
txtextno.Text = ""

vsexam.Clear
vsexam.Visible = False
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
Unload Me
frmexamprint.Show
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
combbranch.Enabled = False
combsection.Enabled = False
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



combtime.AddItem ("09am")
combtime.AddItem ("10am")
combtime.AddItem ("11am")
combtime.AddItem ("12pm")
combtime.AddItem ("02pm")
combtime.AddItem ("03pm")
combtime.AddItem ("04pm")
combtime.AddItem ("05pm")

End Sub

