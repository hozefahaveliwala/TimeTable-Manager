VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmexamprint 
   Caption         =   "S.G.S.I.T.S"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VSPrinter8LibCtl.VSPrinter VSP 
      Height          =   6495
      Left            =   240
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   9255
      _cx             =   16325
      _cy             =   11456
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1700
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   33.7789661319073
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.Frame Frame3 
      Caption         =   "SEARCH"
      Height          =   2415
      Left            =   3240
      TabIndex        =   15
      Top             =   1800
      Width           =   6255
      Begin VB.ComboBox combtime 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtname 
         Height          =   375
         Left            =   4200
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "DELETE"
         Height          =   495
         Left            =   4200
         TabIndex        =   18
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         Height          =   495
         Left            =   2400
         TabIndex        =   17
         Top             =   1680
         Width           =   1455
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2055
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
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
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   -1  'True
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "EXTERNAL NAME"
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "TIME"
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SORT BY"
      Height          =   1215
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   2775
      Begin VB.ComboBox combsort 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   480
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin VSFlex8Ctl.VSFlexGrid vsexam 
      Height          =   2415
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   9135
      _cx             =   16113
      _cy             =   4260
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
   Begin VB.CommandButton cmdprint 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   7080
      Width           =   1300
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   7080
      Width           =   1300
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "BACK"
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   7080
      Width           =   1300
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   8160
      TabIndex        =   8
      Top             =   7080
      Width           =   1300
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "SHOW"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   7080
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Caption         =   "YEAR / BRANCH / SECTION"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      Begin VB.ComboBox combyear 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox combbranch 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox combsection 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "YEAR"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "BRANCH"
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "SECTION"
         Height          =   255
         Left            =   6480
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmexamprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim str1, header As String
Dim d As String
Dim result As VbMsgBoxResult
Private Sub cmdback_Click()
frmexaminfo.Show
Unload Me
End Sub

Private Sub cmdclear_Click()
combyear.ListIndex = -1
combtime.ListIndex = -1
txtname.Text = ""
combsort.ListIndex = -1
combsection.ListIndex = -1
combsection.Enabled = False
combbranch.ListIndex = -1
combbranch.Enabled = False
Calendar1.Value = Null
vsexam.Clear

If VSP.Visible = True Then
VSP.Clear
VSP.Visible = False
End If

End Sub



Private Sub cmddel_Click()
Adodc2.RecordSource = "exam"
Adodc2.Refresh
 If (txtname.Text = "") Or (combtime.ListIndex = -1) Or (Calendar1.ValueIsNull = True) Then
 MsgBox "Entry Missing", vbCritical, "invalid entry"
 cmdclear_Click
 Else
  result = MsgBox("Are you sure you want to delete?", vbYesNo + vbQuestion, "")
   If result = vbYes Then
    d = Calendar1.Value
     Do While (Adodc2.Recordset.EOF = False)
       If ((Adodc2.Recordset.Fields("exam_time") = combtime.Text) And (Adodc2.Recordset.Fields("external_name") = txtname.Text) And (Adodc2.Recordset.Fields("exam_date").Value = d)) Then
       Adodc2.Recordset.Delete (adAffectCurrent)
       vsexam.Clear
       MsgBox "Records Deleted!!", , ""
       cmdclear_Click
        Exit Sub
       End If
   Adodc2.Recordset.MoveNext
   Loop
Else
Exit Sub
End If
MsgBox " No Entries Found!!!", , ""
cmdclear_Click
End If

End Sub

Private Sub cmdexit_Click()
End
End Sub



Private Sub cmdprint_Click()

VSP.Visible = True
VSP.PhysicalPage = True

VSP.Orientation = orLandscape
VSP.HdrFontSize = 20
VSP.header = "|Shri G S Institute Of Technology & Science" & vbCrLf & "Department Of Computer Engineering"
VSP.MarginLeft = 500
VSP.MarginRight = 0
VSP.StartDoc
VSP.RenderControl = vsexam.hWnd
VSP.EndDoc
End Sub

Private Sub cmdsearch_Click()
Dim s1 As Integer
s1 = 0
If (txtname.Text = "") And (combtime.ListIndex = -1) And (Calendar1.Value = "") Then
s1 = 1
MsgBox "Entry Missing", vbCritical, "invalid entry"
cmdclear_Click

ElseIf (combtime.ListIndex = -1) And (Calendar1.Value = "") And (IsNull(txtname.Text) = False) Then
 
Adodc1.RecordSource = " select * from exam where external_name='" & txtname.Text & "'"
Adodc1.Refresh


ElseIf (txtname.Text = "") And (combtime.ListIndex = -1) And IsNull(Calendar1.Value) = False Then
d = Calendar1.Value
Adodc1.RecordSource = " select * from exam where val(exam_date)=" & Val(d)
Adodc1.Refresh


ElseIf (txtname.Text = "") And (Calendar1.Value = "") And combtime.ListIndex <> -1 Then
Adodc1.RecordSource = " select * from exam where exam_time='" & combtime.Text & "'"
Adodc1.Refresh


ElseIf (IsNull(txtname.Text) = False) And (IsNull(Calendar1.Value) = False) And combtime.ListIndex = -1 Then
d = Calendar1.Value
Adodc1.RecordSource = " select * from exam where val(exam_date) =" & Val(d) & " and external_name='" & txtname.Text & "'"
Adodc1.Refresh

ElseIf (combtime.ListIndex <> -1) And (IsNull(txtname.Text) = False) And (IsNull(Calendar1.Value) = True) Then
Adodc1.RecordSource = " select * from exam where external_name='" & txtname.Text & "' and exam_time ='" & combtime.Text & "'"
Adodc1.Refresh

ElseIf (txtname.Text = "") And (IsNull(Calendar1.Value) = False) And (combtime.ListIndex <> -1) Then
d = Calendar1.Value
Adodc1.RecordSource = " select * from exam where val(exam_date) =" & Val(d) & "and exam_time ='" & combtime.Text & "'"
Adodc1.Refresh

Else
d = Calendar1.Value
Adodc1.RecordSource = " select * from exam where val(exam_date) =" & Val(d) & " and external_name='" & txtname.Text & "' and exam_time ='" & combtime.Text & "'"
Adodc1.Refresh
End If

If s1 = 0 Then
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "No Entries Found!!!", vbInformation, ""
cmdclear_Click
Else
Set vsexam.DataSource = Adodc1
vsexam.Refresh
End If
End If

End Sub

Private Sub cmdshow_Click()
If combyear.ListIndex = -1 Then
MsgBox "Entry Missing", vbCritical, "invalid entry"
ElseIf (combsection.Enabled = True) And combsection.ListIndex = -1 Then
MsgBox "Entry Missing", vbCritical, "invalid entry"
ElseIf (combbranch.Enabled = True) And combbranch.ListIndex = -1 Then
MsgBox "Entry Missing", vbCritical, "invalid entry"
Else
If combbranch.Enabled = True Then
str1 = combbranch.Text
ElseIf combsection.Enabled = True Then
str1 = combsection.Text
End If
Adodc1.RecordSource = " select * from exam where class = '" & combyear.Text & "' and branch= '" & str1 & "' order by " & combsort.Text
Adodc1.Refresh
Set vsexam.DataSource = Adodc1
vsexam.Refresh
header = combyear.Text & " " & str1
End If
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
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "timetable.mdb;Persist Security Info=False"
Adodc2.CommandType = adCmdTable
Calendar1.Value = Null

combsort.AddItem ("exam_date")
combsort.AddItem ("internal_name")
combsort.AddItem ("lab_technician")
combsort.AddItem ("lab_attendent")

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
