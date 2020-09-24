VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMergeBooklists 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merge Booklists"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdBrowse2 
      Caption         =   "Browse..."
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse1 
      Caption         =   "Browse..."
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtFilename2 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtFilename1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblFilename2 
      AutoSize        =   -1  'True
      Caption         =   "Into booklist filename"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label lblFilename1 
      AutoSize        =   -1  'True
      Caption         =   "Merge booklist filename"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmMergeBooklists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse1_Click()
    On Error GoTo ErrHandler
    cmd1.CancelError = True
    cmd1.Filter = "All Files (*.*)|*.*|Amazon Booklist Files (*.abl)|*.abl"
    cmd1.FilterIndex = 2
    
    cmd1.ShowOpen
    
    If (Dir(cmd1.FileName) = "") Then
        ' File not found.
        MsgBox "The file '" + cmd1.FileName + "' wasn't found!", vbExclamation, "Amazon Book Searcher"
        Exit Sub
    End If
    
    txtFilename1.Text = cmd1.FileName
    
    Exit Sub
    
ErrHandler:
    ' The user pressed "Cacnel".
    Exit Sub
End Sub


Private Sub cmdBrowse2_Click()
    On Error GoTo ErrHandler
    cmd1.CancelError = True
    cmd1.Filter = "All Files (*.*)|*.*|Amazon Booklist Files (*.abl)|*.abl"
    cmd1.FilterIndex = 2
    
    cmd1.ShowOpen
    
    If (Dir(cmd1.FileName) = "") Then
        ' File not found.
        MsgBox "The file '" + cmd1.FileName + "' wasn't found!", vbExclamation, "Amazon Book Searcher"
        Exit Sub
    End If
    
    txtFilename2.Text = cmd1.FileName
    
    Exit Sub
    
ErrHandler:
    ' The user pressed "Cacnel".
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If (Dir(txtFilename1.Text) = "") Then
        ' File not found.
        MsgBox "The file '" + txtFilename1.Text + "' wasn't found!", vbExclamation, "Amazon Book Searcher"
        Exit Sub
        
    ElseIf (Dir(txtFilename2.Text) = "") Then
        ' File not found.
        MsgBox "The file '" + txtFilename2.Text + "' wasn't found!", vbExclamation, "Amazon Book Searcher"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    ' Merge the two booklists.
    MergeBooklists txtFilename1.Text, txtFilename2.Text
    
    Me.MousePointer = vbDefault
    
    Unload Me
End Sub
