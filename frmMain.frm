VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Strings"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox chkPursuit 
      Caption         =   "Pursuit List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   3855
      Value           =   1  'Aktiviert
      Width           =   1170
   End
   Begin VB.CommandButton cmdSaveCharSet 
      Caption         =   "Sa&ve charset"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6750
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   525
      Width           =   585
   End
   Begin VB.PictureBox pProgressBar 
      AutoRedraw      =   -1  'True
      Height          =   165
      Left            =   3360
      ScaleHeight     =   105
      ScaleWidth      =   1890
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh List"
      Height          =   465
      Left            =   5430
      TabIndex        =   4
      Top             =   525
      Width           =   1320
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search!"
      Height          =   360
      Left            =   6585
      TabIndex        =   3
      Top             =   1470
      Width           =   750
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   5430
      TabIndex        =   2
      Top             =   1170
      Width           =   1905
   End
   Begin VB.TextBox ValidCharSet 
      Height          =   285
      Left            =   5430
      TabIndex        =   1
      Text            =   "abcdefghijklmnopqrstuvwxyzüöäÜÖÄß_-' 1234567890"
      Top             =   225
      Width           =   1905
   End
   Begin VB.ListBox lStrings 
      Height          =   3765
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   5265
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Ready!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5535
      TabIndex        =   8
      Top             =   3660
      Width           =   525
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5430
      TabIndex        =   6
      Top             =   990
      Width           =   555
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Valid Char Set:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5445
      TabIndex        =   5
      Top             =   45
      Width           =   1065
   End
   Begin VB.Menu menuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuCountThis 
         Caption         =   "Count This"
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "Remove All"
      End
      Begin VB.Menu mnuRemoveDuplicates 
         Caption         =   "Remove Duplicates"
      End
      Begin VB.Menu mnuRemoveThis 
         Caption         =   "Remove This"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear List"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileName As String

Private Sub chkPursuit_Click()
  SaveSetting "CodeXP", "StringExtractor", "PursuitList", chkPursuit.Value
End Sub

Private Sub cmdRefresh_Click()
  lStrings.Clear
  lStrings.SetFocus
  ReadFile
End Sub

Private Sub cmdSaveCharSet_Click()
  If Len(ValidCharSet) Then
    SaveSetting "CodeXP", "StringExtractor", "ValidCharSet", ValidCharSet
  End If
  cmdSaveCharSet.Enabled = False
End Sub

Private Sub cmdSearch_Click()
  Dim i As Long, j As Long
  If Len(txtSearch) = 0 Then
    If txtSearch.Enabled = False Then txtSearch.Enabled = True
    txtSearch.SetFocus
    Exit Sub
  End If
  cmdSearch.Enabled = False
  txtSearch.Enabled = False
  j = lStrings.ListIndex + 1
  If j < 0 Or j = lStrings.ListCount Then j = 0
  For i = j To lStrings.ListCount - 1
    If InStr(UCase(lStrings.List(i)), UCase(txtSearch)) Then
      lStrings.ListIndex = i
      Exit For
    End If
    DoEvents
  Next i
  cmdSearch.Enabled = True
  txtSearch.Enabled = True
End Sub

Private Sub Form_Load()
  FileName = Replace(Command, """", "")
  chkPursuit.Value = Val(GetSetting("CodeXP", "StringExtractor", "PursuitList", chkPursuit.Value))
  ValidCharSet = GetSetting("CodeXP", "StringExtractor", "ValidCharSet", ValidCharSet)
  cmdSaveCharSet.Enabled = False
  
  If Len(FileName) Then Me.Caption = "Strings [" & FileName & "]"
  Me.Show
  DoEvents
  
  ReadFile
End Sub

Private Sub ReadFile()
  Dim Buffer As String, BufferLen As Long, FileNr As Integer
  Dim Tmp As String, i As Long
  FileNr = FreeFile
  On Local Error Resume Next
  Open FileName For Input As #FileNr
  If Err Then
    Err.Clear
    Exit Sub
  End If
  Close #FileNr
  cmdRefresh.Enabled = False
  ValidCharSet.Enabled = False
  pProgressBar.Visible = True
  pProgressBar.Cls
  Open FileName For Binary Access Read As #FileNr
  BufferLen = 256
  Do Until EOF(FileNr)
    If BufferLen > LOF(FileNr) - Loc(FileNr) Then
      BufferLen = LOF(FileNr) - Loc(FileNr)
      If BufferLen < 1 Or Err Then Exit Do
    End If
    pProgressBar.Line (0, 0)-(pProgressBar.Width / LOF(FileNr) * Loc(FileNr), _
                              pProgressBar.Height), vbBlue, BF
    pProgressBar.Refresh
    Buffer = Space(BufferLen)
    Get #FileNr, , Buffer
    For i = 1 To Len(Buffer)
      If InStr(ValidCharSet, LCase(Mid(Buffer, i, 1))) Then
        Tmp = Tmp & Mid(Buffer, i, 1)
      Else
        If Len(Tmp) > 3 Then
          lStrings.AddItem Trim(Tmp)
          If chkPursuit.Value Then
            lStrings.ListIndex = lStrings.ListCount - 1
          End If
        End If
        Tmp = ""
      End If
    Next i
    DoEvents
  Loop
  Close FileNr
  If Err Then Err.Clear
  pProgressBar.Visible = False
  cmdRefresh.Enabled = True
  ValidCharSet.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub lStrings_DblClick()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    MsgBox lStrings, vbOKOnly, "Long String View"
  End If
End Sub

Private Sub lStrings_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  Select Case KeyCode
  Case 46 '[DEL]'
    If Shift = 0 Then
      mnuRemoveAll_Click
    Else
      mnuRemoveDuplicates_Click
      TrySelectItem lIndex
    End If
  Case 10, 13 '[ENTER]'
    lStrings_DblClick
  End Select
End Sub

Private Sub lStrings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    PopupMenu menuEdit
  End If
End Sub

Private Sub mnuClear_Click()
  lStrings.Clear
End Sub

Private Sub mnuCopy_Click()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    Clipboard.Clear
    Clipboard.SetText lStrings
  End If
End Sub

Private Sub mnuCountThis_Click()
  Dim ItemText As String
  Dim i As Long
  Dim c As Long
  
  i = lStrings.ListIndex + 1
  If i Then
    ItemText = UCase(Trim(lStrings))
    For i = 0 To lStrings.ListCount - 1
      If UCase(Trim(lStrings.List(i))) = ItemText Then c = c + 1
    Next i
    MsgBox "Count of """ & ItemText & """ is " & c
  End If
End Sub

Private Sub mnuRemoveAll_Click()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    RemoveItem lStrings
    TrySelectItem lIndex - 1
  End If
End Sub

Private Sub mnuRemoveDuplicates_Click()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    RemoveItem lStrings, lIndex
    TrySelectItem lIndex - 1
  End If
End Sub

Private Sub mnuRemoveThis_Click()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    lStrings.RemoveItem lIndex - 1
    TrySelectItem lIndex - 1
  End If
End Sub

Private Sub ValidCharSet_Change()
  cmdSaveCharSet.Enabled = True
End Sub


Private Sub RemoveItem(ByVal ItemText As String, Optional ByVal ExceptIndex As Long)
  Dim i As Long
  ItemText = UCase(Trim(ItemText))
  While i < lStrings.ListCount
    If i + 1 <> ExceptIndex Then
      If UCase(Trim(lStrings.List(i))) = ItemText Then
        lStrings.RemoveItem i
        i = i - 1
      End If
    End If
    i = i + 1
  Wend
End Sub


Private Sub TrySelectItem(ByVal lIndex As Long)
  If lStrings.ListCount < 1 Then Exit Sub
  If lStrings.ListCount < (lIndex + 1) Then lIndex = lStrings.ListCount - 1
  If lIndex < 0 Then lIndex = 0
  lStrings.ListIndex = lIndex
End Sub
