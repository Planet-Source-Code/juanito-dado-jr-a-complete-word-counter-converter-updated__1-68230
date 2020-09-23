VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmWordCounter 
   Caption         =   "Word Counter"
   ClientHeight    =   11115
   ClientLeft      =   2535
   ClientTop       =   2670
   ClientWidth     =   14325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   14325
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   16536
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Setup Page"
      TabPicture(0)   =   "Form1.frx":1708A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ImageList1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPath"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdLocal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdNetwork"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "File1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Dir1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Report Page"
      TabPicture(1)   =   "Form1.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCopy"
      Tab(1).Control(1)=   "Option1(1)"
      Tab(1).Control(2)=   "Option1(2)"
      Tab(1).Control(3)=   "Option1(3)"
      Tab(1).Control(4)=   "Option1(4)"
      Tab(1).Control(5)=   "MSFlexGrid1"
      Tab(1).Control(6)=   "Label2"
      Tab(1).Control(7)=   "lblLineCount"
      Tab(1).ControlCount=   8
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   3840
         TabIndex        =   12
         Top             =   4440
         Width           =   9615
         Begin VB.CommandButton cmdRmvFile 
            Caption         =   "-- Item"
            Height          =   615
            Left            =   1920
            Picture         =   "Form1.frx":170C2
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "remove file from the list"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdWordCount 
            Caption         =   "Start Counting"
            Default         =   -1  'True
            Height          =   615
            Left            =   8040
            Picture         =   "Form1.frx":171E3
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Start Word Count!!!"
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "+ All"
            Height          =   615
            Left            =   960
            Picture         =   "Form1.frx":172BC
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Add all items on the list"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdRmvAll 
            Caption         =   "-- All"
            Height          =   615
            Left            =   2880
            Picture         =   "Form1.frx":17530
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Remove all files from the list"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdAddFile 
            Caption         =   "+ Item"
            Height          =   615
            Left            =   0
            Picture         =   "Form1.frx":1764E
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Add single item on the list"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdRemoveShared 
            Caption         =   "-- Shared"
            Height          =   615
            Left            =   3840
            Picture         =   "Form1.frx":177D4
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Remove Shared Workbook"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdFileDT 
            Caption         =   "Date/Time"
            Height          =   615
            Left            =   4800
            Picture         =   "Form1.frx":1790A
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Get Time Stamp"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdRtfToTxt 
            Caption         =   "Rtf-Txt"
            Height          =   615
            Left            =   5760
            Picture         =   "Form1.frx":179E5
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Convert RTF to TEXT"
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.DirListBox Dir1 
         Height          =   7965
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   3420
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   11760
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy to Clipboard"
         Height          =   615
         Left            =   -74760
         Picture         =   "Form1.frx":17B0C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Char With Spaces"
         Height          =   255
         Index           =   1
         Left            =   -72840
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Linecount"
         Height          =   255
         Index           =   2
         Left            =   -71160
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Char W/Out Spaces"
         Height          =   255
         Index           =   3
         Left            =   -69960
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Words"
         Height          =   255
         Index           =   4
         Left            =   -68040
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdNetwork 
         Caption         =   "Network Files"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdLocal 
         Caption         =   "Local Files"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         ToolTipText     =   "Enter Network Path"
         Top             =   600
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7935
         Left            =   -74760
         TabIndex        =   9
         Top             =   1200
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   13996
         _Version        =   393216
         Cols            =   6
         BackColorBkg    =   -2147483624
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         AllowUserResizing=   3
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   11760
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":17C2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":17D5E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":17E98
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":17FB4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3735
         Left            =   3840
         TabIndex        =   21
         Top             =   5280
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   14111
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3720
         Left            =   3840
         TabIndex        =   22
         ToolTipText     =   "Double click OR Drag the File"
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   6562
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Total Line Count:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66600
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblLineCount 
         Height          =   375
         Left            =   -64800
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   10080
      TabIndex        =   25
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   10860
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This also include getting date/time stamp, removing shared documents, and converting RTF-TXT format."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   720
      TabIndex        =   28
      Top             =   600
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This is a Complete Word Counter Application suited for counting different text formats such as excel, word, etc."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   720
      TabIndex        =   27
      Top             =   360
      Width           =   9975
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   12240
      Picture         =   "Form1.frx":180E2
      Top             =   240
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000018&
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      FillColor       =   &H80000001&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   240
      Top             =   120
      Width           =   13695
   End
   Begin VB.Menu mnuFil 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmWordCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
'oooooooooooooooo Word Counter by Juanito Dado Jr oooooooooooooooooooooo
'ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo


Option Explicit
'progress bar in status bar
Private Declare Function SetParent Lib "user32" _
        (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetInputState Lib "user32" () As Long
Dim fs As FileSystemObject
Dim cnt As Long, i As Long, X As Long, z As Long
Dim dirLocalPath As String, dirNetworkPath As String
Dim sLocal As Boolean, sNetwork As Boolean
Dim timeEnd As String, timeStart As String

Private Sub showfileinfo(fileSpec As String, X As Integer)
    Dim f
    Dim listitem1 As ListItem
    
    Set fs = New FileSystemObject
    Set f = fs.GetFile(fileSpec)
   
   'include path to the listview
    'replace \\ with \ if the  path is on c:\
    dirLocalPath = Replace(Dir1.Path & "\" & fs.GetFileName(fileSpec), "\\", "\")
    dirNetworkPath = Dir1.Path & "\" & fs.GetFileName(fileSpec)
    
    
    Set listitem1 = ListView1.ListItems.Add()
    
    ' FileName
    If sLocal Then
        listitem1.Text = dirLocalPath
    ElseIf sNetwork Then
        listitem1.Text = dirNetworkPath
    End If
    
    ' Date Created
    ListView1.ListItems(X + 1).ListSubItems.Add , , Format(f.DateCreated, "mm/dd/yyyy")
    ' Date last Modified
    ListView1.ListItems(X + 1).ListSubItems.Add , , Format(f.DateLastModified, "mm/dd/yyyy")
    ' File Type
    ListView1.ListItems(X + 1).ListSubItems.Add , , f.Type
End Sub

Private Sub cmdCopy_Click()
    'copy to clipboard
    Call CopySelected
End Sub

Private Sub cmdFileDT_Click()
Dim i%, j%, strText$

For i = 1 To ListView1.ListItems.Count
    If InStr(1, ListView1.ListItems(i).Text, "WAV") > 0 Or InStr(1, ListView1.ListItems(i).Text, "dct") > 0 _
        Or InStr(1, ListView1.ListItems(i).Text, "dss") > 0 Or InStr(1, ListView1.ListItems(i).Text, "wav") > 0 Then
        ListView2.ListItems.Add , , FileDateTime(ListView1.ListItems(i).Text)
    End If
Next


    Clipboard.Clear
    With ListView2
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                strText = strText & .ListItems(i).Text
                For j = 1 To .ColumnHeaders.Count - 1
                    strText = strText
                Next j
                strText = strText & vbNewLine
            End If
        Next i
    End With
    Clipboard.SetText strText

End Sub

Private Sub cmdLocal_Click()
    sLocal = True
    sNetwork = False
    Dir1.Path = "c:\"
End Sub

Private Sub cmdNetwork_Click()
    sLocal = False
    sNetwork = True
    Dir1.Path = "\\YourServer\NetworkFolder\" 'Put your network path here
End Sub

Private Sub cmdPath_Click()
On Error Resume Next
Dim InputStr As String
    sLocal = False
    sNetwork = True

    InputStr = InputBox("Enter Network Path", "WordCounter")
    Dir1.Path = InputStr

End Sub

Private Sub cmdRmvFile_Click()
    'if list is empty then exit
    If ListView2.ListItems.Count = 0 Then
        MsgBox "List Empty", vbInformation, "Word Counter"
        Exit Sub
    Else
        ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
    End If
End Sub

'RTF to TEXT conversion
Private Sub cmdRtfToTxt_Click()
Dim oWord As Word.Application
'Dim oDoc As Word.Document
Dim myFile As String
Dim iDel As Integer

If ListView2.ListItems.Count = 0 Then
    MsgBox "No files to convert", vbExclamation, App.EXEName
    Exit Sub
End If
Set oWord = New Word.Application

ProgressBar1.Min = 0
ProgressBar1.Max = ListView2.ListItems.Count

For X = 1 To ListView2.ListItems.Count
    'Set oDoc = New Word.Document
    ProgressBar1.Value = X
    myFile = ListView2.ListItems.Item(X).Text
    oWord.Documents.Open myFile
    oWord.ActiveDocument.SaveAs Mid$(myFile, 1, InStrRev(myFile, ".", -1)) & "txt", wdFormatText, False, _
    "", True, "", False, False, False, False, False, 1252, False, False, wdCRLF
    oWord.ActiveDocument.Close
    If GetInputState Then DoEvents
Next

ProgressBar1.Value = 0
oWord.Application.Quit
Set oWord = Nothing
MsgBox "Done Converting.", vbInformation, "Word Converter"

'delete the word files of the converted files
For iDel = 1 To ListView2.ListItems.Count
    Kill ListView2.ListItems(iDel).Text
Next
End Sub

Private Sub cmdWordCount_Click()
Dim wordObject As Word.Application
Dim charWithSpace As Long, charNoSpace As Long, Words As Long
Dim lineCount As Double
Dim listy As ListItem
Dim strVar() As String
Dim strFile As String
Dim totLineCount As Double

    'set flexgrid to number of items on listview
    MSFlexGrid1.Rows = ListView2.ListItems.Count + 1
    'for selected column color
   
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No files available to Count", vbInformation, "Word Counter"
        Exit Sub
    Else
        ProgressBar1.Min = 0
        ProgressBar1.Max = ListView2.ListItems.Count
        timeStart = Format(Time, "hh:nn:ss")
        
        'set wordObject
        Set wordObject = New Word.Application
        
        For X = 1 To ListView2.ListItems.Count
            
            DoEvents
            
            'set invalid line count to exe files
            If InStr(1, ListView2.ListItems.Item(X).Text, "exe") > 0 Then
                SSTab1.Tab = 1
                With MSFlexGrid1
                    'flexgrid alignment
                    .ColAlignment(1) = flexAlignLeftCenter
                    .TextMatrix(X, 1) = UseInStrRev(ListView2.ListItems.Item(X).Text)
                    .TextMatrix(X, 2) = "Invalid"
                    .TextMatrix(X, 3) = "Invalid"
                    .TextMatrix(X, 4) = "Invalid"
                    .TextMatrix(X, 5) = "Invalid"
                End With
            Else
                SSTab1.Tab = 1
                ProgressBar1.Value = X
                wordObject.Documents.Open ListView2.ListItems.Item(X).Text
                'compute for char with space
                charWithSpace = wordObject.ActiveDocument.Content.ComputeStatistics(wdStatisticCharactersWithSpaces)
                'compute for char without space
                charNoSpace = wordObject.ActiveDocument.Content.ComputeStatistics(wdStatisticCharacters)
                'compute for words
                Words = wordObject.ActiveDocument.Content.ComputeStatistics(wdStatisticWords)
                'linecount divided by 65
                lineCount = wordObject.ActiveDocument.Content.ComputeStatistics(wdStatisticLines)
                'flexgrid alignment
                With MSFlexGrid1
                    .ColAlignment(1) = flexAlignLeftCenter
                    'populate msflexgrid with data
                    .TextMatrix(X, 0) = X
                    .TextMatrix(X, 1) = UseInStrRev(ListView2.ListItems.Item(X).Text)
                    .TextMatrix(X, 2) = charWithSpace
                    .TextMatrix(X, 3) = lineCount
                    .TextMatrix(X, 4) = charNoSpace
                    .TextMatrix(X, 5) = Words
                    totLineCount = totLineCount + lineCount
                End With
                'wordObject.Documents.Close False
                wordObject.ActiveDocument.Close False
            End If
        Next X

        wordObject.Quit
        Set wordObject = Nothing
        
        timeEnd = Format(Time, "hh:nn:ss")
        MsgBox "Duration: " & (DateDiff("s", timeStart, timeEnd)) & " second(s)", , "Word Counter"
        lblLineCount.Caption = totLineCount
        ProgressBar1.Value = 0
        Call colBackColor(MSFlexGrid1, 2, RGB(255, 255, 163))
        Option1(1) = True
    End If
End Sub

'for extracting filenames only
Private Function UseInStrRev(ByVal strIn As String) As String
Dim intPos As Integer

    intPos = InStrRev(strIn, "\") + 1
    UseInStrRev = Mid(strIn, intPos)
End Function


Private Sub cmdAddFile_Click()
    'if nothing selected
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Select items on the list above.", vbInformation, "Word Counter"
    Else
        Call AddToListview
    End If
End Sub

Private Sub cmdAddAll_Click()
Dim Y As Long
    
    For Y = 1 To ListView1.ListItems.Count
        'listview icon purposes
        'this can be done by shell and API but i dunno how to do it. ^ ^
        If InStr(1, ListView1.ListItems.Item(Y).Text, "doc") > 0 Or InStr(1, ListView1.ListItems.Item(Y).Text, "DOC") > 0 _
        Or InStr(1, ListView1.ListItems.Item(Y).Text, "rtf") > 0 Or InStr(1, ListView1.ListItems.Item(Y).Text, "RTF") > 0 Then
            ListView2.ListItems.Add , , ListView1.ListItems.Item(Y).Text, , 1
        ElseIf InStr(1, ListView1.ListItems.Item(Y).Text, "xls") > 0 Or InStr(1, ListView1.ListItems.Item(Y).Text, "XLS") > 0 Then
            ListView2.ListItems.Add , , ListView1.ListItems.Item(Y).Text, , 2
        ElseIf InStr(1, ListView1.ListItems.Item(Y).Text, "txt") > 0 Or InStr(1, ListView1.ListItems.Item(Y).Text, "TXT") > 0 Then
            ListView2.ListItems.Add , , ListView1.ListItems.Item(Y).Text, , 3
        Else
            ListView2.ListItems.Add , , ListView1.ListItems.Item(Y).Text, , 4
        End If
    Next Y
End Sub

Private Sub cmdRmvAll_Click()
    ListView2.ListItems.Clear
    MSFlexGrid1.Clear
    'for columheaders title
    MSFlexGrid1.FormatString = "|Filename|Char With Spaces|Linecount|Char W/Out Spaces|Words"
End Sub


Private Sub Dir1_Change()
    Dim X As Integer
    
    ListView1.ListItems().Clear
    File1.Path = Dir1.Path
    Screen.MousePointer = vbHourglass
    For X = 0 To File1.ListCount - 1
        showfileinfo File1.Path & "/" & File1.List(X), X
    Next X
    Screen.MousePointer = vbDefault
    'remove default selection on listview
    Call removeSelection
End Sub

Private Sub removeSelection()
'remove default selection on the listview
    With ListView1
        For cnt = 1 To .ListItems.Count
        .ListItems(cnt).Selected = False
        Next
        Set .SelectedItem = Nothing
    End With
End Sub
    

Private Sub Form_Load()
    'attaching the progress bar to statbar
    'setparent
    SetParent ProgressBar1.hWnd, StatusBar1.hWnd
    ProgressBar1.Top = 55
    'position
    ProgressBar1.Left = StatusBar1.Panels(1).Width + 60
    'size
    ProgressBar1.Width = StatusBar1.Panels(2).Width - 60
    ProgressBar1.Height = StatusBar1.Height - 90
    
    
    Dir1.Path = "c:\"
    sLocal = True
    Dim X As Integer
    
    ListView1.ListItems().Clear
    
    Screen.MousePointer = vbHourglass
    For X = 0 To File1.ListCount - 1
        showfileinfo File1.Path & "/" & File1.List(X), X
    Next X
    Screen.MousePointer = vbDefault
    
    ListView1.View = lvwReport
    ' Set up headers for listView
    Dim colHeader As ColumnHeader
    Set colHeader = ListView1.ColumnHeaders.Add()
    colHeader.Text = "Name"
    colHeader.Width = 5000
    Set colHeader = ListView1.ColumnHeaders.Add()
    colHeader.Text = "Date Created"
    colHeader.Width = 1500
    Set colHeader = ListView1.ColumnHeaders.Add()
    colHeader.Text = "Date Modified"
    colHeader.Width = 2000
    Set colHeader = ListView1.ColumnHeaders.Add()
    colHeader.Text = "File Type"
    colHeader.Width = 2000
    
    With MSFlexGrid1
    'msflexgrid formats
        .FormatString = "|Filename|Char With Spaces|Linecount|Char W/Out Spaces|Words"
        .ColAlignment(1) = flexAlignLeftCenter
    
    'msflexgrid size
        .ColWidth(0) = 400
        .ColWidth(1) = 6500
        .ColWidth(2) = 1500
        .ColWidth(4) = 1600
        .ColWidth(5) = 1000
    End With
    'remove default selection on listview
    Call removeSelection
    SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fs = Nothing
End Sub

Private Sub AddToListview()
Dim z As Long
    For z = 1 To ListView2.ListItems.Count
            If ListView2.ListItems.Item(z).Text = ListView1.SelectedItem.Text Then
                MsgBox "Already added to the list.", vbInformation, "Word Counter"
                Exit Sub
            End If
    Next z
    
    If InStr(1, ListView1.SelectedItem.Text, "doc") > 0 Or InStr(1, ListView1.SelectedItem.Text, "DOC") > 0 _
    Or InStr(1, ListView1.SelectedItem.Text, "rtf") > 0 Or InStr(1, ListView1.SelectedItem.Text, "RTF") > 0 Then
        ListView2.ListItems.Add , , ListView1.SelectedItem.Text, , 1
    ElseIf InStr(1, ListView1.SelectedItem.Text, "xls") > 0 Or InStr(1, ListView1.SelectedItem.Text, "XLS") > 0 Then
        ListView2.ListItems.Add , , ListView1.SelectedItem.Text, , 2
    ElseIf InStr(1, ListView1.SelectedItem.Text, "txt") > 0 Or InStr(1, ListView1.SelectedItem.Text, "TXT") > 0 Then
        ListView2.ListItems.Add , , ListView1.SelectedItem.Text, , 3
    Else
        ListView2.ListItems.Add , , ListView1.SelectedItem.Text, , 4
    End If
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
    Call AddToListview
End Sub

'drag drop
Private Sub ListView2_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
            For X = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(X).Text = ListView1.SelectedItem.Text Then
                    MsgBox "Already added to the list.", vbInformation, "Word Counter"
                    Exit For
                Else
            
                    If InStr(1, ListView1.ListItems.Item(i).Text, "doc") > 0 Or InStr(1, ListView1.ListItems.Item(i).Text, "DOC") > 0 _
                    Or InStr(1, ListView1.ListItems.Item(i).Text, "rtf") > 0 Or InStr(1, ListView1.ListItems.Item(i).Text, "RTF") > 0 Then
                        ListView2.ListItems.Add , , ListView1.ListItems(i).Text, , 1
                    ElseIf InStr(1, ListView1.ListItems.Item(i).Text, "xls") > 0 Or InStr(1, ListView1.ListItems.Item(i).Text, "XLS") > 0 Then
                        ListView2.ListItems.Add , , ListView1.ListItems(i).Text, , 2
                    ElseIf InStr(1, ListView1.ListItems.Item(i).Text, "txt") > 0 Or InStr(1, ListView1.ListItems.Item(i).Text, "TXT") > 0 Then
                        ListView2.ListItems.Add , , ListView1.ListItems(i).Text, , 3
                    Else
                       ListView2.ListItems.Add , , ListView1.ListItems(i).Text, , 4
                    End If
                End If
            Next X
        End If
    Next i
End Sub

Private Sub ListView2_DblClick()
    'remove items from listview2
    ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
End Sub


Private Sub mnuAbout_Click()
    MsgBox "Word Counter by Juanito Dado Jr." & vbCrLf & "Please Vote for ME on PSC", , "About"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub CopySelected()
    'Copy the selection and put it on the Clipboard
    Clipboard.Clear
    Clipboard.SetText MSFlexGrid1.Clip
End Sub

'this will fill the background of the selected column with yellow color
'credits goes to gavio of vbforums
Private Sub colBackColor(mfg As MSFlexGrid, col As Long, color As Long)
    With mfg
        .Redraw = False
        .FillStyle = flexFillRepeat
        .col = col
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .CellBackColor = color
        .FillStyle = flexFillSingle
        .col = 1
        .Redraw = True
    End With
End Sub

'background color change
Private Sub Option1_Click(Index As Integer)
         Static lastCol As Long
        If lastCol <> 0 Then
            colBackColor MSFlexGrid1, lastCol, vbWhite
        End If
            lastCol = (Index + 1)
                colBackColor MSFlexGrid1, lastCol, RGB(255, 255, 163)
End Sub


Private Sub cmdRemoveShared_Click()
On Error Resume Next
Dim excelObject As Excel.Application
Dim X As Long
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No files available!", vbInformation, "Excel Remove Shared"
        Exit Sub
    Else
        ProgressBar1.Min = 0
        ProgressBar1.Max = ListView2.ListItems.Count
        timeStart = Format(Time, "hh:nn:ss")
        
        Set excelObject = New Excel.Application
        
        For X = 1 To ListView2.ListItems.Count
        
                ProgressBar1.Value = X
                With excelObject
                    .Workbooks.Open ListView2.ListItems.Item(X).Text
                    .ActiveWorkbook.Application.DisplayAlerts = False
                    .Application.AskToUpdateLinks = False
                    .DisplayAlerts = False
                    .ActiveWorkbook.ExclusiveAccess
                    .ActiveWorkbook.Save
                End With
        Next X

        excelObject.Quit
        Set excelObject = Nothing
        
        timeEnd = Format(Time, "hh:nn:ss")
        MsgBox "Duration: " & (DateDiff("s", timeStart, timeEnd)) & " second(s)", , "Word Counter"
        ProgressBar1.Value = 0
    End If
End Sub



Private Sub Form_Resize()
'autoscale objects in side the form
On Error Resume Next   ' this is needed because when the user resize it to minimum it'll send an error
     ProgressBar1.Left = StatusBar1.Panels(1).Width + 60
     Dir1.Height = Me.Height - 3670
     ListView1.Height = Frame1.Top - 750
     ListView1.Width = Me.Width - 4600
     ListView2.Width = Me.Width - 4600
     ListView2.Height = Me.Height - 8000
     Shape1.Width = Me.Width - 600
     Frame1.Top = ListView2.Top - 750
     SSTab1.Height = Me.Height - 2500
     SSTab1.Width = Me.Width - 600
     MSFlexGrid1.Width = Me.Width - 550
     MSFlexGrid1.Height = Me.Height - 2200
End Sub




