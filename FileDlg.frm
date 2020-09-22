VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "File Dialog Test"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Hide ReadOnly"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.CheckBox chkMultiSelect 
      Caption         =   "MultiSelect"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   180
      Width           =   2055
   End
   Begin VB.ListBox lstFiles 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   1620
      Width           =   4035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1635
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "File(s) Chosen:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1380
      Width           =   1035
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fD As cFileDialog

Private Sub cmdOpen_Click()
    On Error GoTo errorhandler
    Dim cD As New cFileDialog
    Dim sFiles() As String
    Dim filecount As Long
    Dim sDir As String
    Dim i As Long
    
    With fD
        .flags = OFN_FILEMUSTEXIST
        If chkReadOnly.Value = 1 Then .flags = .flags Or OFN_HIDEREADONLY
        If chkMultiSelect.Value = 1 Then .flags = .flags Or OFN_ALLOWMULTISELECT
        .hwnd = Me.hwnd
        
        .CancelError = True
        .ShowOpen
        lstFiles.Clear
        If chkMultiSelect.Value = 1 Then
            .ParseMultiFileName sDir, sFiles(), filecount
            For i = 0 To filecount - 1
                lstFiles.AddItem sFiles(i)
            Next
        Else
            lstFiles.AddItem .Filename
        End If
    End With
    
Exit Sub
errorhandler:
    MsgBox Err.Description
End Sub

Private Sub cmdSave_Click()
    
    With fD
        .flags = OFN_OVERWRITEPROMPT
        If chkReadOnly.Value = 1 Then .flags = .flags Or OFN_HIDEREADONLY
        .hwnd = Me.hwnd
        .DefaultExt = "vbp"
        .CancelError = False
        .ShowSave
        lstFiles.Clear
        lstFiles.AddItem .Filename
    End With
    
End Sub

Private Sub Form_Load()
    Set fD = New cFileDialog
    fD.Filter = "All Files (*.*)|*.*|Visual Basic Projects (*.vbp)|*.vbp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fD = Nothing
End Sub
