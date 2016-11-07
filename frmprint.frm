VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form frmprint 
   Caption         =   "S.G.S.I.T.S"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddeleteall 
      Caption         =   "DELETE TABLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   19
      Top             =   6120
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   8760
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   8640
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.CommandButton cmdprint 
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   17
      Top             =   6120
      Width           =   1455
   End
   Begin VSPrinter8LibCtl.VSPrinter VSP 
      Height          =   5775
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   10335
      _cx             =   18230
      _cy             =   10186
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
      MarginTop       =   2000
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
      Zoom            =   29.5008912655971
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
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   6120
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid vstt 
      Height          =   2895
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   10335
      _cx             =   18230
      _cy             =   5106
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   2
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   100
      RowHeightMax    =   0
      ColWidthMin     =   2700
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmprint.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
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
   Begin VB.Frame Frame2 
      Caption         =   "FACULTY TIME TABLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   9255
      Begin VB.TextBox txtreplace 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   20
         Top             =   460
         Width           =   2055
      End
      Begin VB.CommandButton cmdreplace 
         Caption         =   "REPLACE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   460
         Width           =   2295
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "FACULTY NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   8
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "SHOW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   6120
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8760
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Frame Frame1 
      Caption         =   "YEAR / BRANCH / SECTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      Begin VB.ComboBox combyear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "BRANCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "SECTION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n, i, j, m As Integer
Dim str1, header, t As String
Dim result As VbMsgBoxResult
Sub del()

Do While (Adodc2.Recordset.EOF = False)

Adodc2.Recordset.Fields("faculty") = Replace(Adodc2.Recordset.Fields("faculty"), txtname.Text, txtreplace.Text)
Adodc2.Recordset.Update
Adodc2.Recordset.MoveNext
Loop
End Sub
Function delall()
If combyear.Text = "First Year" Then
Do While (Adodc3.Recordset.EOF = False)
If (Adodc3.Recordset.Fields("class_name") = combsection.Text) Then
Adodc3.Recordset.Delete adAffectCurrent
Adodc3.Recordset.Update
End If
Adodc3.Recordset.MoveNext
Loop

Else
Do While (Adodc3.Recordset.EOF = False)
If (Adodc3.Recordset.Fields("branch") = combbranch.Text) Then
Adodc3.Recordset.Delete adAffectCurrent
Adodc3.Recordset.Update
End If
Adodc3.Recordset.MoveNext
Loop

End If

End Function
Function gridformat()
vstt.Rows = 10
vstt.Cols = 6
vstt.FixedCols = 1
vstt.FixedRows = 1
vstt.TextMatrix(0, 1) = "Monday"
vstt.TextMatrix(0, 2) = "Tuesday"
vstt.TextMatrix(0, 3) = "Wednesday"
vstt.TextMatrix(0, 4) = "Thursday"
vstt.TextMatrix(0, 5) = "Friday"

vstt.TextMatrix(1, 0) = "09am-10am"
vstt.TextMatrix(2, 0) = "10am-11am"
vstt.TextMatrix(3, 0) = "11am-12pm"
vstt.TextMatrix(4, 0) = "12pm-01pm"
vstt.TextMatrix(6, 0) = "02pm-03pm"
vstt.TextMatrix(7, 0) = "03pm-04pm"
vstt.TextMatrix(8, 0) = "04pm-05pm"
vstt.TextMatrix(9, 0) = "05pm-06pm"
End Function


Private Sub cmdback_Click()
frmmain.Show
Unload Me
End Sub

Private Sub cmdclear_Click()
combyear.ListIndex = -1
combsection.ListIndex = -1
combsection.Enabled = False
combbranch.ListIndex = -1
combbranch.Enabled = False
txtname.Text = ""
txtreplace.Text = ""
vstt.Clear
If VSP.Visible = True Then
VSP.Clear
VSP.Visible = False
End If

End Sub

Private Sub cmddelete_Click()

End Sub

Private Sub cmddeleteall_Click()
If combyear.ListIndex = -1 Then
j = -1
End If
If combsection.Enabled = True And combsection.ListIndex = -1 Then
j = -1
End If
If combbranch.Enabled = True And combbranch.ListIndex = -1 Then
j = -1
End If
If j = -1 Then
MsgBox "Entry Missing", vbCritical, "invalid entry"
Else

result = MsgBox("Are you sure you want to delete?", vbYesNo + vbQuestion, "")
If (result = vbYes) Then
Adodc3.CommandType = adCmdTable
If (combyear.Text = "First Year") Then
Adodc3.RecordSource = "year1"
ElseIf (combyear.Text = "Second Year") Then
Adodc3.RecordSource = "year2"
ElseIf (combyear.Text = "Third Year") Then
Adodc3.RecordSource = "year3"
ElseIf (combyear.Text = "Fourth Year") Then
Adodc3.RecordSource = "year4"
End If
Adodc3.Refresh
delall
vstt.Clear
MsgBox "Records Deleted", vbInformation, ""
Else
Exit Sub
End If
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

VSP.header = "|Shri G S Institute Of Technology & Science" & vbCrLf & "Department Of Computer Engineering" & vbCrLf & header
VSP.MarginLeft = 500
VSP.MarginRight = 0
VSP.StartDoc
VSP.RenderControl = vstt.hWnd
VSP.EndDoc
End Sub



Private Sub cmdreplace_Click()
If txtname.Text = "" Or txtreplace.Text = "" Then
MsgBox "Entry Missing", vbCritical, "invalid entry"
Else
result = MsgBox("Are you sure you want to replace?", vbYesNo + vbQuestion, "")
If (result = vbYes) Then
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select faculty  from year1 where faculty like '%" & txtname.Text & "%' "
Adodc2.Refresh
del
Adodc2.RecordSource = " select faculty from year2 where faculty like '%" & txtname.Text & "%'"
Adodc2.Refresh
del
Adodc2.RecordSource = " select faculty from year3 where faculty like '%" & txtname.Text & "%' "
Adodc2.Refresh
del
Adodc2.RecordSource = " select faculty from year4 where faculty like '%" & txtname.Text & "%'"
Adodc2.Refresh
del

txtname.Text = ""
txtreplace.Text = ""
MsgBox "Records Replaced", vbInformation, ""

Else
Exit Sub
End If
End If

End Sub

Private Sub cmdsearch_Click()
If txtname.Text = "" Then
MsgBox "Entry Missing", vbCritical, "invalid entry"
Else
str1 = "select class_name as mbr,subject,day as dayno ,timing as timeno,faculty  from year1 where faculty like '%" & txtname.Text & "%' UNION ALL"
str1 = str1 + " select branch as mbr,subject,day as dayno,timing as timeno,faculty from year2 where faculty like '%" & txtname.Text & "%' UNION ALL"
str1 = str1 + " select branch as mbr,subject,day as dayno,timing as timeno,faculty from year3 where faculty like '%" & txtname.Text & "%' UNION ALL"
str1 = str1 + " select branch as mbr,subject,day as dayno,timing as timeno,faculty from year4 where faculty like '%" & txtname.Text & "%'"

Adodc1.RecordSource = str1
Adodc1.Refresh

If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "No Entries Found", vbCritical, ""

Else
header = "TIME TABLE" & vbCrLf & txtname.Text
gridformat
Do While (Not Adodc1.Recordset.EOF = True)

i = Adodc1.Recordset.Fields("timeno")
j = Adodc1.Recordset.Fields("dayno")
vstt.TextMatrix(i, j) = Adodc1.Recordset.Fields("mbr") & "  " & Adodc1.Recordset.Fields("subject") & " " & Adodc1.Recordset.Fields("faculty")

Adodc1.Recordset.MoveNext
Loop
For m = 1 To 5
vstt.TextMatrix(5, m) = "LUNCH"
Next
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
n = combyear.ListIndex

If n = 0 Then
Adodc1.RecordSource = "select subject,day ,timing,faculty  from year1 where class_name='" & combsection.Text & "'"
header = "TIME TABLE" & vbCrLf & "FIRST YEAR " & combsection.Text

ElseIf n = 1 Then
Adodc1.RecordSource = "select subject,day ,timing,faculty  from year2 where branch='" & combbranch.Text & "'"
header = "TIME TABLE" & vbCrLf & "SECOND YEAR " & combbranch.Text

ElseIf n = 2 Then
Adodc1.RecordSource = "select subject,day ,timing,faculty from year3 where branch='" & combbranch.Text & "'"
header = "TIME TABLE" & vbCrLf & "THIRD YEAR " & combbranch.Text

ElseIf n = 3 Then
Adodc1.RecordSource = "select subject,day ,timing,faulty  from year4 where branch='" & combbranch.Text & "'"
header = "TIME TABLE" & vbCrLf & "FOURTH YEAR " & combbranch.Text

End If

Adodc1.Refresh

gridformat
Do While (Not Adodc1.Recordset.EOF = True)

i = Adodc1.Recordset.Fields("timing")
j = Adodc1.Recordset.Fields("day")
vstt.TextMatrix(i, j) = Adodc1.Recordset.Fields("subject") & "  " & Adodc1.Recordset.Fields("faculty")

Adodc1.Recordset.MoveNext
Loop

For m = 1 To 5
vstt.TextMatrix(5, m) = "LUNCH"
Next
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
Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "timetable.mdb;Persist Security Info=False"


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

End Sub

