VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Amazon Book Searcher"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7260
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
   ScaleHeight     =   6180
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   6000
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picMatches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   2040
      ScaleHeight     =   2775
      ScaleWidth      =   4815
      TabIndex        =   37
      Top             =   4560
      Width           =   4815
      Begin ComctlLib.ListView lvwMatches 
         Height          =   1575
         Left            =   240
         TabIndex        =   38
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2778
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "No Matches Loaded"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   210
         Left            =   240
         TabIndex        =   39
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.PictureBox picAdvancedOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   3720
      ScaleHeight     =   3495
      ScaleWidth      =   4935
      TabIndex        =   26
      Top             =   3480
      Width           =   4935
      Begin VB.CheckBox chkAdditionalInfo 
         Caption         =   "Retrieve book's additional information (takes longer)"
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   3000
         Width           =   4095
      End
      Begin VB.TextBox txtPublicationDate 
         Height          =   285
         Left            =   2280
         TabIndex        =   35
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbPublicationDate 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2520
         Width           =   2055
      End
      Begin VB.ComboBox cmbLanguage 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox cmbReaderAge 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cmbFormat 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         ForeColor       =   &H8000000C&
         Height          =   210
         Left            =   2280
         TabIndex        =   36
         Top             =   2280
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblPublicationDate 
         AutoSize        =   -1  'True
         Caption         =   "Publication Date"
         ForeColor       =   &H8000000C&
         Height          =   210
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label lblLanguage 
         AutoSize        =   -1  'True
         Caption         =   "Language"
         ForeColor       =   &H8000000C&
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label lblReaderAge 
         AutoSize        =   -1  'True
         Caption         =   "Reader Age"
         ForeColor       =   &H8000000C&
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   870
      End
      Begin VB.Label lblFormat 
         AutoSize        =   -1  'True
         Caption         =   "Format"
         ForeColor       =   &H8000000C&
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      ScaleHeight     =   615
      ScaleWidth      =   4575
      TabIndex        =   22
      Top             =   120
      Width           =   4575
      Begin VB.Line Line2 
         X1              =   2760
         X2              =   2760
         Y1              =   120
         Y2              =   480
      End
      Begin VB.Line Line1 
         X1              =   1200
         X2              =   1200
         Y1              =   120
         Y2              =   480
      End
      Begin VB.Label lblAdvancedOptions 
         AutoSize        =   -1  'True
         Caption         =   "Advanced Options"
         Height          =   210
         Left            =   3000
         TabIndex        =   25
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label lblMatches 
         AutoSize        =   -1  'True
         Caption         =   "Matches"
         Height          =   210
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         Caption         =   "Search"
         Height          =   210
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   -2400
      ScaleHeight     =   5655
      ScaleWidth      =   6855
      TabIndex        =   0
      Top             =   3960
      Width           =   6855
      Begin VB.ComboBox cmbSubject 
         Height          =   330
         Left            =   120
         TabIndex        =   44
         Top             =   2640
         Width           =   5175
      End
      Begin VB.ComboBox cmbTitle 
         Height          =   330
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Width           =   5175
      End
      Begin VB.ComboBox cmbAuthor 
         Height          =   330
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   5175
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   495
         Left            =   5640
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
      End
      Begin VB.PictureBox picAuthorOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   5775
         TabIndex        =   12
         Top             =   840
         Width           =   5775
         Begin VB.OptionButton optAuthor 
            Caption         =   "Start of last name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   15
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton optAuthor 
            Caption         =   "Exact name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   4440
            TabIndex        =   14
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton optAuthor 
            Caption         =   "First name/initials and last name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin VB.PictureBox picTitleOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   4935
         TabIndex        =   8
         Top             =   1920
         Width           =   4935
         Begin VB.OptionButton optTitle 
            Caption         =   "Exact start of title"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   11
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton optTitle 
            Caption         =   "Start(s) of title word(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   10
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton optTitle 
            Caption         =   "Title word(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.PictureBox picSubjectOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   5775
         TabIndex        =   4
         Top             =   3000
         Width           =   5775
         Begin VB.OptionButton optSubject 
            Caption         =   "Subject word(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optSubject 
            Caption         =   "Start(s) of Subject"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3960
            TabIndex        =   6
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton optSubject 
            Caption         =   "Start(s) of Subject word(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   5
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.TextBox txtISBN 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   3840
         Width           =   5295
      End
      Begin VB.TextBox txtPublisher 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   4560
         Width           =   5295
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   1
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         Caption         =   "Author"
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Title"
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   285
      End
      Begin VB.Label lblSubject 
         AutoSize        =   -1  'True
         Caption         =   "Subject"
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label lblISBN 
         AutoSize        =   -1  'True
         Caption         =   "ISBN"
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   345
      End
      Begin VB.Label lblPublisher 
         AutoSize        =   -1  'True
         Caption         =   "Publisher"
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   4320
         Width           =   660
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open Booklist"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save Booklist"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Merge Booklists"
         Index           =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   4
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Help"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "About"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAuthor_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, c As Integer
Dim s As String

    ' Auto-complete the text.
    
    For i = 0 To cmbAuthor.ListCount - 1
        If ((Left$(cmbAuthor.List(i), Len(cmbAuthor.Text)) = cmbAuthor.Text) And (cmbAuthor.Text <> "") And _
            (KeyCode <> vbKeyBack) And (KeyCode <> vbKeyDelete) And (KeyCode <> vbKeyLeft) And (KeyCode <> vbKeyRight) And _
            (KeyCode <> vbKeyUp) And (KeyCode <> vbKeyDown)) Then
            ' Auto-complete the text.
            cmbAuthor.SelStart = Len(cmbAuthor.Text)
            c = Len(cmbAuthor.Text)
            cmbAuthor.SelText = Mid$(cmbAuthor.List(i), Len(cmbAuthor.Text) + 1)
            cmbAuthor.SelStart = c
            cmbAuthor.SelLength = Len(cmbAuthor.List(i))
            
            Exit For
        End If
    Next i
End Sub

Private Sub cmbPublicationDate_Change()
    If (cmbPublicationDate.ListIndex > 0) Then
        lblYear.Visible = True
        txtPublicationDate.Visible = True
    Else
        lblYear.Visible = False
        txtPublicationDate.Visible = False
    End If
End Sub

Private Sub cmbPublicationDate_Click()
    Call cmbPublicationDate_Change
End Sub

Private Sub cmbPublicationDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call cmbPublicationDate_Change
End Sub

Private Sub cmbPublicationDate_KeyPress(KeyAscii As Integer)
    Call cmbPublicationDate_Change
End Sub

Private Sub cmbPublicationDate_Scroll()
    Call cmbPublicationDate_Change
End Sub

Private Sub cmbSubject_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, c As Integer
Dim s As String

    ' Auto-complete the text.
    
    For i = 0 To cmbAuthor.ListCount - 1
        If ((Left$(cmbAuthor.List(i), Len(cmbAuthor.Text)) = cmbAuthor.Text) And (cmbAuthor.Text <> "") And _
            (KeyCode <> vbKeyBack) And (KeyCode <> vbKeyDelete) And (KeyCode <> vbKeyLeft) And (KeyCode <> vbKeyRight) And _
            (KeyCode <> vbKeyUp) And (KeyCode <> vbKeyDown)) Then
            ' Auto-complete the text.
            cmbAuthor.SelStart = Len(cmbAuthor.Text)
            c = Len(cmbAuthor.Text)
            cmbAuthor.SelText = Mid$(cmbAuthor.List(i), Len(cmbAuthor.Text) + 1)
            cmbAuthor.SelStart = c
            cmbAuthor.SelLength = Len(cmbAuthor.List(i))
            
            Exit For
        End If
    Next i
End Sub

Private Sub cmbTitle_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, c As Integer
Dim s As String

    ' Auto-complete the text.
    
    For i = 0 To cmbAuthor.ListCount - 1
        If ((Left$(cmbAuthor.List(i), Len(cmbAuthor.Text)) = cmbAuthor.Text) And (cmbAuthor.Text <> "") And _
            (KeyCode <> vbKeyBack) And (KeyCode <> vbKeyDelete) And (KeyCode <> vbKeyLeft) And (KeyCode <> vbKeyRight) And _
            (KeyCode <> vbKeyUp) And (KeyCode <> vbKeyDown)) Then
            ' Auto-complete the text.
            cmbAuthor.SelStart = Len(cmbAuthor.Text)
            c = Len(cmbAuthor.Text)
            cmbAuthor.SelText = Mid$(cmbAuthor.List(i), Len(cmbAuthor.Text) + 1)
            cmbAuthor.SelStart = c
            cmbAuthor.SelLength = Len(cmbAuthor.List(i))
            
            Exit For
        End If
    Next i
End Sub

Private Sub cmdReset_Click()
    ' Reset All Fields.
    cmbAuthor.Text = ""
    cmbTitle.Text = ""
    cmbSubject.Text = ""
    txtISBN.Text = ""
    txtPublisher.Text = ""
    txtPublicationDate.Text = ""
End Sub

Private Sub cmdSearch_Click()
' The URL to be loaded.
Dim strURL As String
' The Data to be downloaded.
Dim vtData As String
Dim strData As String: strData = ""
' Counter\Temp Variables.
Dim i As Integer, s() As String
' The Book class to be used when retrieving additional book info.
Dim clsB As clsBook
' Used for adding books to the Listview.
Dim itmBook As ListItem
' A flag used to for adding\not adding a search string
' to the history.
Dim blnAddSearchString As Boolean
' Used for connecting to the internet and waiting
' for the connection to be completed.
Dim process_id As Long
Dim process_handle As Long

    On Error GoTo ErrHandler

    If (IsConnected = False) Then
        ' The user isn't connected to the internet.
        If (MsgBox("You're not seem to be connected to the internet, would you like to connect to it?", vbYesNo Or vbQuestion, "Amazon Book Searcher") = vbNo) Then Exit Sub
        
        ' Connect to the internet and wait for
        ' the connection to be completed.
        process_id = Shell("Rundll32.exe rnaui.dll,RnaDial", vbNormalFocus)
        On Error GoTo 0
        
        Me.MousePointer = vbHourglass
        
        DoEvents
    
        ' Wait for the program to finish.
        ' Get the process handle.
        process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
        If (process_handle <> 0) Then
            WaitForSingleObject process_handle, INFINITE
            CloseHandle process_handle
        End If
        
        Me.MousePointer = vbDefault
    End If

    If (cmdSearch.Caption = "Cancel") Then
        ' The User Canceled the search.
        gblnStopSearch = True
        lblStat.Caption = "Search Canceled. " & glngResultsCount & " matches found."
        cmdSearch.Caption = "Search"
        Inet1.Cancel
        Exit Sub
    End If
    
    gblnStopSearch = False
    
    ' Enable the "Save Booklist" Menu item.
    mnuFileItem(1).Enabled = True
    
    
    ' Save the Search terms in the history.
    If (Trim$(cmbAuthor.Text) <> "") Then
        ' Add the search string only if it's not in
        ' the combobox already.
        blnAddSearchString = True
        For i = 0 To cmbAuthor.ListCount - 1
            If (cmbAuthor.List(i) = cmbAuthor.Text) Then
                blnAddSearchString = False
            End If
        Next i
        
        If (blnAddSearchString = True) Then
            cmbAuthor.AddItem cmbAuthor.Text, 0
            
            If (cmbAuthor.ListCount > HISTORY_COUNT) Then
                cmbAuthor.RemoveItem cmbAuthor.ListCount - 1
            End If
        End If
    End If
    If (Trim$(cmbTitle.Text) <> "") Then
        ' Add the search string only if it's not in
        ' the combobox already.
        blnAddSearchString = True
        For i = 0 To cmbTitle.ListCount - 1
            If (cmbTitle.List(i) = cmbTitle.Text) Then
                blnAddSearchString = False
            End If
        Next i
        
        If (blnAddSearchString = True) Then
            cmbTitle.AddItem cmbTitle.Text, 0
            
            If (cmbTitle.ListCount > HISTORY_COUNT) Then
                cmbTitle.RemoveItem cmbTitle.ListCount - 1
            End If
        End If
    End If
    If (Trim$(cmbSubject.Text) <> "") Then
        ' Add the search string only if it's not in
        ' the combobox already.
        blnAddSearchString = True
        For i = 0 To cmbSubject.ListCount - 1
            If (cmbSubject.List(i) = cmbSubject.Text) Then
                blnAddSearchString = False
            End If
        Next i
        
        If (blnAddSearchString = True) Then
            cmbSubject.AddItem cmbSubject.Text, 0
            
            If (cmbSubject.ListCount > HISTORY_COUNT) Then
                cmbSubject.RemoveItem cmbSubject.ListCount - 1
            End If
        End If
    End If
    
    If (Trim$(txtISBN.Text) <> "") Then
        ' Search the book by its ISBN.
    
        ' Initialize the Global Search Variables.
        glngResultsCount = 0
        gintCurrentPage = 0
        glngCurrentBook = 0
        ReDim gclsBooks(0 To 0)
        
        lvwMatches.ListItems.Clear
        
        Call lblMatches_Click
        
        cmdSearch.Caption = "Cancel"
        
        lblStat.Caption = "Searching... Downloading book page"
    
        ' Open the First Results Page.
        vtData = Inet1.OpenURL(Replace(BOOK_PAGE_URL, "<ISBN>", Trim$(txtISBN.Text)))
        
        ' Get the Rest of the file.
        
        strData = strData & vtData
        
        ' Get the first chunk.
        vtData = Inet1.GetChunk(1024, icString)
        DoEvents
        Do
            ' Append the Chunk to the end of the Data String.
            strData = strData & vtData
            
            DoEvents
            
            ' Get the next chunk.
            vtData = Inet1.GetChunk(1024, icString)
            
            ' Exit the loop if there are no more chunks to retrieve.
            If (Len(vtData) = 0) Then Exit Do
        Loop
        
        ' Parse the Book Page.
        lblStat.Caption = "Searching... Parsing book page"
        
        Set clsB = ParseBookPage(strData)
    
        If (clsB.Title = "") Then
            ' No matching book was found.
            lblStat.Caption = "No matches found!"
            cmdSearch.Caption = "Search"
        Else
            ' A Match was found - Add the book to the list.
            ReDim gclsBooks(1 To 1)
            Set gclsBooks(1) = clsB
            ' Add the Book's Title.
            Set itmBook = lvwMatches.ListItems.Add(, gclsBooks(UBound(gclsBooks)).Title, gclsBooks(UBound(gclsBooks)).Title)
            
            If (gclsBooks(UBound(gclsBooks)).AuthorCount > 0) Then
                ' Build the Book Author Array.
                ReDim s(1 To gclsBooks(UBound(gclsBooks)).AuthorCount)
                For i = 1 To gclsBooks(UBound(gclsBooks)).AuthorCount
                    s(i) = gclsBooks(UBound(gclsBooks)).Authors(i)
                Next i
            End If
            
            ' Add the Book's Authors.
            itmBook.SubItems(1) = Join(s, ", ")
            ' Add the Book's Average Rating.
            If (gclsBooks(UBound(gclsBooks)).AverageRating > 0) Then
                itmBook.SubItems(2) = gclsBooks(UBound(gclsBooks)).AverageRating & " out of 5"
            Else
                itmBook.SubItems(2) = "N\A"
            End If
            If (gclsBooks(UBound(gclsBooks)).Cover <> "") Then
                ' Add the Book's Cover Type.
                itmBook.SubItems(3) = gclsBooks(UBound(gclsBooks)).Cover
            Else
                itmBook.SubItems(3) = "N\A"
            End If
            ' Add the Book's Price.
            itmBook.SubItems(4) = "$" & gclsBooks(UBound(gclsBooks)).Price
            
            lvwMatches.Refresh
            
            lblStat.Caption = "Search complete. 1 match found."
            cmdSearch.Caption = "Search"
            
            Exit Sub
        End If
    End If
    

    ' See if the search should be done.
    cmbAuthor.Text = Trim$(cmbAuthor.Text)
    cmbTitle.Text = Trim$(cmbTitle.Text)
    cmbSubject.Text = Trim$(cmbSubject.Text)
    txtISBN.Text = Trim$(txtISBN.Text)
    txtPublisher.Text = Trim$(txtPublisher.Text)
    txtPublicationDate.Text = Trim$(txtPublicationDate.Text)
    
    If ((cmbAuthor.Text = "") And (cmbTitle.Text = "") And _
        (cmbSubject.Text = "") And (txtISBN.Text = "") And (txtPublisher.Text = "")) Then
        ' Missing Search String.
        MsgBox "You must enter the Book's Author, Title, Subject, ISBN or Publisher in order to search!", vbExclamation, "Amazon Book Searcher"
        Exit Sub
    ElseIf ((cmbPublicationDate.ListIndex > 0) And (txtPublicationDate.Text = "")) Then
        ' An invalid publication date was found.
        MsgBox "You must enter a valid year in the Publication Date field!", vbExclamation, "Amazon Book Searcher"
        Exit Sub
    End If

    ' Construct the Search URL.
    
    strURL = SEARCH_URL_PREFIX
    
    If (Trim$(cmbAuthor.Text) <> "") Then
        If (optAuthor(0).Value = True) Then
            strURL = strURL & "author-like"
        ElseIf (optAuthor(1).Value = True) Then
            strURL = strURL & "author-begins"
        ElseIf (optAuthor(2).Value = True) Then
            strURL = strURL & "author-exact"
        End If
        
        strURL = strURL & "%01" & EncodeURL(Trim$(cmbAuthor.Text))
    End If

    If (Trim$(cmbTitle.Text) <> "") Then
        If (Trim$(cmbAuthor.Text) <> "") Then
            strURL = strURL & "%02"
        End If
        
        If (optTitle(0).Value = True) Then
            strURL = strURL & "title"
        ElseIf (optTitle(1).Value = True) Then
            strURL = strURL & "title-words-begin"
        ElseIf (optTitle(2).Value = True) Then
            strURL = strURL & "title-begins"
        End If
        
        strURL = strURL & "%01" & EncodeURL(Trim$(cmbTitle.Text))
    End If

    If (Trim$(cmbSubject.Text) <> "") Then
        If ((Trim$(cmbAuthor.Text) <> "") Or (Trim$(cmbTitle.Text) <> "")) Then
            strURL = strURL & "%02"
        End If
        
        If (optSubject(0).Value = True) Then
            strURL = strURL & "subject"
        ElseIf (optSubject(1).Value = True) Then
            strURL = strURL & "subject-words-begin"
        ElseIf (optSubject(2).Value = True) Then
            strURL = strURL & "subject-begins"
        End If
        
        strURL = strURL & "%01" & EncodeURL(Trim$(cmbSubject.Text))
    End If

    If (Trim$(txtPublisher.Text) <> "") Then
        If ((Trim$(cmbAuthor.Text) <> "") Or (Trim$(cmbTitle.Text) <> "") Or (Trim$(cmbSubject.Text) <> "")) Then
            strURL = strURL & "%02"
        End If
        
        strURL = strURL & "publisher" & "%01" & EncodeURL(Trim$(txtPublisher.Text))
    End If
    
    If (cmbFormat.ListIndex > 0) Then
        If ((Trim$(cmbAuthor.Text) <> "") Or (Trim$(cmbTitle.Text) <> "") Or (Trim$(cmbSubject.Text) <> "") Or (Trim$(txtPublisher.Text) <> "")) Then
            strURL = strURL & "%02"
        End If
        
        strURL = strURL & "binding" & "%01"
        
        Select Case cmbFormat.ListIndex
            Case 1
                strURL = strURL & "hardcover"
            Case 2
                strURL = strURL & "paperback"
            Case 3
                strURL = strURL & "digital"
            Case 4
                strURL = strURL & "audio%20cassette%7Caudio%20cd"
            Case 5
                strURL = strURL & "large%20print"
        End Select
    End If
    
    If (cmbReaderAge.ListIndex > 0) Then
        If ((Trim$(cmbAuthor.Text) <> "") Or (Trim$(cmbTitle.Text) <> "") Or (Trim$(cmbSubject.Text) <> "") Or (Trim$(txtPublisher.Text) <> "") Or (cmbFormat.ListIndex > 0)) Then
            strURL = strURL & "%02"
        End If
        
        strURL = strURL & "age" & "%01"
        
        Select Case cmbReaderAge.ListIndex
            Case 1
                strURL = strURL & "baby/preschool"
            Case 2
                strURL = strURL & "4-8"
            Case 3
                strURL = strURL & "9-12"
            Case 4
                strURL = strURL & "books%20-%20young%20adult"
        End Select
    End If
    
    If (cmbLanguage.ListIndex > 0) Then
        If ((Trim$(cmbAuthor.Text) <> "") Or (Trim$(cmbTitle.Text) <> "") Or (Trim$(cmbSubject.Text) <> "") Or (Trim$(txtPublisher.Text) <> "") Or (cmbFormat.ListIndex > 0) Or (cmbReaderAge.ListIndex > 0)) Then
            strURL = strURL & "%02"
        End If
        
        strURL = strURL & "language" & "%01" & "spanish"
    End If

    If (cmbPublicationDate.ListIndex > 0) Then
        If ((Trim$(cmbAuthor.Text) <> "") Or (Trim$(cmbTitle.Text) <> "") Or (Trim$(cmbSubject.Text) <> "") Or (Trim$(txtPublisher.Text) <> "") Or (cmbFormat.ListIndex > 0) Or (cmbReaderAge.ListIndex > 0) Or (cmbLanguage.ListIndex > 0)) Then
            strURL = strURL & "%02"
        End If
        
        strURL = strURL & "dateop" & "%01"
        
        Select Case cmbPublicationDate.ListIndex
            Case 1
                strURL = strURL & "before"
            Case 2
                strURL = strURL & "during"
            Case 3
                strURL = strURL & "after"
        End Select
        
        strURL = strURL & "%02datemod%010%02dateyear%01" & txtPublicationDate.Text
    End If

    strURL = AMAZON_URL & strURL & SEARCH_URL_SUFFIX
    
    ' Initialize the Global Search Variables.
    glngResultsCount = 0
    gintCurrentPage = 0
    glngCurrentBook = 0
    ReDim gclsBooks(0 To 0)
    
    lvwMatches.ListItems.Clear
    
    Call lblMatches_Click
    
    cmdSearch.Caption = "Cancel"
    
    lblStat.Caption = "Searching... Downloading results page"

    ' Open the First Results Page.
    vtData = Inet1.OpenURL(Replace(strURL, "<Page Number>", "1"))
    
    ' Get the Rest of the file.
    
    strData = strData & vtData
    
    ' Get the first chunk.
    vtData = Inet1.GetChunk(1024, icString)
    DoEvents
    Do
        ' Append the Chunk to the end of the Data String.
        strData = strData & vtData
        
        DoEvents
        
        ' Get the next chunk.
        vtData = Inet1.GetChunk(1024, icString)
        
        ' Exit the loop if there are no more chunks to retrieve.
        If (Len(vtData) = 0) Then Exit Do
    Loop
    
    ' Parse the Results Page.
    lblStat.Caption = "Searching... Parsing results page"
    
    ParseResultPage (strData)

    ' Calculate the Result Page Count.
    gintResultPageCount = glngResultsCount / RESULTS_PER_PAGE
    If (glngResultsCount Mod RESULTS_PER_PAGE > 0) Then gintResultPageCount = gintResultPageCount + 1
    
    ' Download the rest of the result pages.
    For i = 2 To gintResultPageCount
        strData = ""
        
        ' Download the Results Page.
        lblStat.Caption = "Searching... " & glngResultsCount & " matches. Downloading results page (" & i & " out of " & gintResultPageCount & " pages)"
        vtData = Inet1.OpenURL(Replace(strURL, "<Page Number>", i))
         
        ' Get the Rest of the file.
        
        strData = strData & vtData
        
        ' Get the first chunk.
        vtData = Inet1.GetChunk(1024, icString)
        DoEvents
        Do
            ' Append the Chunk to the end of the Data String.
            strData = strData & vtData
            
            DoEvents
            
            ' Get the next chunk.
            vtData = Inet1.GetChunk(1024, icString)
            
            DoEvents
            
            If (gblnStopSearch = True) Then
                ' Stop the Search.
                gblnStopSearch = False
                lblStat.Caption = "Search Canceled. " & glngResultsCount & " matches found."
                cmdSearch.Caption = "Search"
                Inet1.Cancel
                Exit Sub
            End If
            
            ' Exit the loop if there are no more chunks to retrieve.
            If (Len(vtData) = 0) Then Exit Do
        Loop
        
        If (gblnStopSearch = True) Then
            ' Stop the Search.
            gblnStopSearch = False
            lblStat.Caption = "Search Canceled. " & glngResultsCount & " matches found."
            cmdSearch.Caption = "Search"
            Inet1.Cancel
            Exit Sub
        End If
        
        ' Parse the Results Page.
        lblStat.Caption = "Searching... " & glngResultsCount & " matches. Parsing results page (" & i & " out of " & gintResultPageCount & " pages)"
        ParseResultPage (strData)
    Next i
    
    If (chkAdditionalInfo.Value = vbChecked) And (UBound(gclsBooks) > 0) Then
        ' Retrieve the Additional Info of every match.
        
        For i = 1 To UBound(gclsBooks)
            strData = ""
            
            ' Download the Results Page.
            lblStat.Caption = glngResultsCount & " matches. Downloading additional book information (" & i & " out of " & glngResultsCount & " books)"
            vtData = Inet1.OpenURL(gclsBooks(i).URL)
            
            ' Get the Rest of the file.
            
            strData = strData & vtData
            
            ' Get the first chunk.
            vtData = Inet1.GetChunk(1024, icString)
            DoEvents
            Do
                ' Append the Chunk to the end of the Data String.
                strData = strData & vtData
                
                DoEvents
                
                ' Get the next chunk.
                vtData = Inet1.GetChunk(1024, icString)
                
                DoEvents
                
                If (gblnStopSearch = True) Then
                    ' Stop the Search.
                    gblnStopSearch = False
                    lblStat.Caption = "Search Canceled. " & glngResultsCount & " matches found."
                    cmdSearch.Caption = "Search"
                    Inet1.Cancel
                    Exit Sub
                End If
                
                ' Exit the loop if there are no more chunks to retrieve.
                If (Len(vtData) = 0) Then Exit Do
            Loop
            
            If (gblnStopSearch = True) Then
                ' Stop the Search.
                gblnStopSearch = False
                lblStat.Caption = "Search Canceled. " & glngResultsCount & " matches found."
                cmdSearch.Caption = "Search"
                Inet1.Cancel
                Exit Sub
            End If
            
            ' Parse the Book Page.
            lblStat.Caption = glngResultsCount & " matches. Parsing additional book information (" & i & " out of " & glngResultsCount & " books)"
            Set clsB = ParseBookPage(strData)
            
            ' Replace the Old book info with the new book info.
            clsB.URL = gclsBooks(i).URL
            Set gclsBooks(i) = clsB

        Next i
    End If
    

    If (glngResultsCount = 0) Then
        ' No matches found.
        lblStat.Caption = "No matches found!"
        cmdSearch.Caption = "Search"
    Else
        ' Search complete.
        lblStat.Caption = "Search complete. " & glngResultsCount & " matches found."
        cmdSearch.Caption = "Search"
    End If

    Exit Sub

ErrHandler:
    If (Err.Number = 35764) Then
        ' Request Timed Out.
        MsgBox "The Amazon server isn't responding. Please try searching later.", vbInformation, "Amazon Book Searcher"
        lblStat.Caption = "Search canceled. " & glngResultsCount & " matches found."
        cmdSearch.Caption = "Search"
        Exit Sub
    Else
    
    End If
End Sub


Private Sub Form_Load()
    ' Load the settings.
    OpenSettings (App.Path & "\" & SETTINGS_FILENAME)
    
    Call lblSearch_Click
    
    ' Disable the "Save Booklist" Menu item.
    mnuFileItem(1).Enabled = False
    
    ReDim gclsBooks(0 To 0)
    
    ' Initialize the comboboxes in the Advanced Options Tab.
    
    cmbFormat.AddItem "All Formats"
    cmbFormat.AddItem "Hardcover"
    cmbFormat.AddItem "Paperback"
    cmbFormat.AddItem "e-Book"
    cmbFormat.AddItem "Audiobook (Cassete or CD)"
    cmbFormat.AddItem "Large Print"
    
    cmbReaderAge.AddItem "All ages"
    cmbReaderAge.AddItem "Baby-3 Years"
    cmbReaderAge.AddItem "4-8 Years"
    cmbReaderAge.AddItem "9-12 Years"
    cmbReaderAge.AddItem "Teen"
    
    cmbLanguage.AddItem "All Languages"
    cmbLanguage.AddItem "Spanish"
    
    cmbPublicationDate.AddItem "All Dates"
    cmbPublicationDate.AddItem "Before the year"
    cmbPublicationDate.AddItem "During the year"
    cmbPublicationDate.AddItem "After the year"
    
    cmbFormat.ListIndex = 0
    cmbReaderAge.ListIndex = 0
    cmbLanguage.ListIndex = 0
    cmbPublicationDate.ListIndex = 0

    lvwMatches.View = lvwReport
    lvwMatches.ColumnHeaders.Add , "Title", "Title"
    lvwMatches.ColumnHeaders.Add , "Authors", "Authors"
    lvwMatches.ColumnHeaders.Add , "Average Rating", "Average Rating"
    lvwMatches.ColumnHeaders.Add , "Cover", "Cover", 700
    lvwMatches.ColumnHeaders.Add , "Price", "Price", 450
End Sub

Private Sub Form_Resize()
Dim i As Integer
    ' Resize All Controls.
    
    i = 100 + lblSearch.Width + 400 + lblMatches.Width + 400 + lblAdvancedOptions.Width
    
    picTab.Top = 100
    picTab.Left = 100
    picTab.Width = Me.Width - 200
    picTab.Height = lblSearch.Height + lblSearch.Top + 200 'Line1.Y2 - Line1.Y1
    
    Line1.Y1 = 0
    Line1.Y2 = lblSearch.Height + 50
    Line2.Y1 = 0
    Line2.Y2 = lblSearch.Height + 50
    
    lblSearch.Left = (picTab.Width - i) / 2
    lblSearch.Top = (Line1.Y2 - Line1.Y1 - lblSearch.Height) / 2
    
    Line1.X1 = lblSearch.Left + lblSearch.Width + 200
    Line1.X2 = lblSearch.Left + lblSearch.Width + 200
    
    lblMatches.Left = Line1.X1 + 200
    lblMatches.Top = (Line1.Y2 - Line1.Y1 - lblMatches.Height) / 2
    
    Line2.X1 = lblMatches.Left + lblMatches.Width + 200
    Line2.X2 = lblMatches.Left + lblMatches.Width + 200
    
    lblAdvancedOptions.Left = Line2.X1 + 200
    lblAdvancedOptions.Top = (Line2.Y2 - Line2.Y1 - lblAdvancedOptions.Height) / 2
    
    picSearch.Top = picTab.Top + picTab.Height + 50
    picSearch.Left = 100
    picSearch.Width = Me.ScaleWidth - 200
    picSearch.Height = Me.ScaleHeight - picSearch.Top - 100
    
    lblAuthor.Top = 100
    lblAuthor.Left = 100
    cmbAuthor.Top = lblAuthor.Top + lblAuthor.Height + 50
    cmbAuthor.Left = lblAuthor.Left
    cmbAuthor.Width = picSearch.Width - cmbAuthor.Left - 300
    
    picAuthorOptions.Left = lblAuthor.Left
    picAuthorOptions.Top = cmbAuthor.Top + cmbAuthor.Height + 50
    picAuthorOptions.Width = optAuthor(0).Width + optAuthor(1).Width + optAuthor(2).Width
    picAuthorOptions.Height = optAuthor(0).Height
    
    optAuthor(0).Top = 0
    optAuthor(0).Left = 0
    optAuthor(1).Top = 0
    optAuthor(1).Left = optAuthor(0).Left + optAuthor(0).Width
    optAuthor(2).Top = 0
    optAuthor(2).Left = optAuthor(1).Left + optAuthor(1).Width

    lblTitle.Top = picAuthorOptions.Top + picAuthorOptions.Height + 50
    lblTitle.Left = lblAuthor.Left
    cmbTitle.Top = lblTitle.Top + lblTitle.Height + 50
    cmbTitle.Left = lblTitle.Left
    cmbTitle.Width = picSearch.Width - cmbTitle.Left - 300
    
    picTitleOptions.Left = lblTitle.Left
    picTitleOptions.Top = cmbTitle.Top + cmbTitle.Height + 50
    picTitleOptions.Width = optTitle(0).Width + optTitle(1).Width + optTitle(2).Width
    picTitleOptions.Height = optTitle(0).Height
    
    optTitle(0).Top = 0
    optTitle(0).Left = 0
    optTitle(1).Top = 0
    optTitle(1).Left = optTitle(0).Left + optTitle(0).Width
    optTitle(2).Top = 0
    optTitle(2).Left = optTitle(1).Left + optTitle(1).Width

    lblSubject.Top = picTitleOptions.Top + picSubjectOptions.Height + 50
    lblSubject.Left = lblTitle.Left
    cmbSubject.Top = lblSubject.Top + lblSubject.Height + 50
    cmbSubject.Left = lblSubject.Left
    cmbSubject.Width = picSearch.Width - cmbSubject.Left - 300
    
    picSubjectOptions.Left = lblSubject.Left
    picSubjectOptions.Top = cmbSubject.Top + cmbSubject.Height + 50
    picSubjectOptions.Width = optSubject(0).Width + optSubject(1).Width + optSubject(2).Width
    picSubjectOptions.Height = optSubject(0).Height
    
    optSubject(0).Top = 0
    optSubject(0).Left = 0
    optSubject(1).Top = 0
    optSubject(1).Left = optSubject(0).Left + optSubject(0).Width
    optSubject(2).Top = 0
    optSubject(2).Left = optSubject(1).Left + optSubject(1).Width

    lblISBN.Top = picSubjectOptions.Top + picSubjectOptions.Height + 50
    lblISBN.Left = lblSubject.Left
    txtISBN.Top = lblISBN.Top + lblISBN.Height + 50
    txtISBN.Left = lblSubject.Left
    txtISBN.Width = picSearch.Width - txtISBN.Left - 300

    lblPublisher.Top = txtISBN.Top + txtISBN.Height + 50
    lblPublisher.Left = lblISBN.Left
    txtPublisher.Top = lblPublisher.Top + lblPublisher.Height + 50
    txtPublisher.Left = lblISBN.Left
    txtPublisher.Width = picSearch.Width - txtPublisher.Left - 300
    
    cmdSearch.Top = txtPublisher.Top + txtPublisher.Height + 400
    cmdSearch.Left = picSearch.Width / 2 - (cmdSearch.Width + 100 + cmdReset.Width) / 2
    
    cmdReset.Top = txtPublisher.Top + txtPublisher.Height + 400
    cmdReset.Left = cmdSearch.Left + cmdSearch.Width + 100

    picAdvancedOptions.Top = picTab.Top + picTab.Height + 50
    picAdvancedOptions.Left = 100
    picAdvancedOptions.Width = Me.ScaleWidth - 200
    picAdvancedOptions.Height = Me.ScaleHeight - picAdvancedOptions.Top - 100
    
    lblFormat.Left = 100
    lblFormat.Top = 100
    
    cmbFormat.Left = lblFormat.Left
    cmbFormat.Top = lblFormat.Top + lblFormat.Height + 50
    cmbFormat.Width = 2055
    
    lblReaderAge.Left = lblFormat.Left
    lblReaderAge.Top = cmbFormat.Top + cmbFormat.Height + 50
    
    cmbReaderAge.Left = lblReaderAge.Left
    cmbReaderAge.Top = lblReaderAge.Top + lblReaderAge.Height + 50
    cmbReaderAge.Width = 2055

    lblLanguage.Left = lblReaderAge.Left
    lblLanguage.Top = cmbReaderAge.Top + cmbReaderAge.Height + 50
    
    cmbLanguage.Left = lblReaderAge.Left
    cmbLanguage.Top = lblLanguage.Top + lblLanguage.Height + 50
    cmbLanguage.Width = 2055

    lblPublicationDate.Left = lblReaderAge.Left
    lblPublicationDate.Top = cmbLanguage.Top + cmbLanguage.Height + 50
    
    cmbPublicationDate.Left = lblReaderAge.Left
    cmbPublicationDate.Top = lblPublicationDate.Top + lblPublicationDate.Height + 50
    cmbPublicationDate.Width = 2055
    
    txtPublicationDate.Left = cmbPublicationDate.Left + cmbPublicationDate.Width + 70
    txtPublicationDate.Top = cmbPublicationDate.Top
    txtPublicationDate.Width = 1575
    
    lblYear.Left = txtPublicationDate.Left
    lblYear.Top = lblPublicationDate.Top
    
    chkAdditionalInfo.Left = lblPublicationDate.Left
    chkAdditionalInfo.Top = txtPublicationDate.Top + txtPublicationDate.Height + 200

    picMatches.Top = picTab.Top + picTab.Height + 50
    picMatches.Left = 100
    picMatches.Width = Me.ScaleWidth - 200
    picMatches.Height = Me.ScaleHeight - picMatches.Top - 100
    
    lblStatus.Left = 100
    lblStatus.Top = 100
    
    lblStat.Left = lblStatus.Left
    lblStat.Top = lblStatus.Top + lblStatus.Height + 50
    
    lvwMatches.Top = lblStat.Top + lblStat.Height + 50
    lvwMatches.Left = lblStatus.Left
    lvwMatches.Width = picMatches.Width - lvwMatches.Left
    lvwMatches.Height = picMatches.Height - lvwMatches.Top

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Save the settings.
    SaveSettings (App.Path & "\" & SETTINGS_FILENAME)
End Sub

Private Sub lblAdvancedOptions_Click()
    lblSearch.FontBold = False
    lblMatches.FontBold = False
    lblAdvancedOptions.FontBold = True
    
    picSearch.Visible = False
    picMatches.Visible = False
    picAdvancedOptions.Visible = True
End Sub

Private Sub lblMatches_Click()
    lblSearch.FontBold = False
    lblMatches.FontBold = True
    lblAdvancedOptions.FontBold = False
    
    picSearch.Visible = False
    picMatches.Visible = True
    picAdvancedOptions.Visible = False
End Sub

Private Sub lblSearch_Click()
    lblSearch.FontBold = True
    lblMatches.FontBold = False
    lblAdvancedOptions.FontBold = False
    
    picSearch.Visible = True
    picMatches.Visible = False
    picAdvancedOptions.Visible = False
End Sub

Private Sub lvwMatches_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
' Used for determining the sort order (Ascending or Descending).
Static intSortOrder As Integer
Static intPreviousIndex As Integer
Static blnBeenClicked As Boolean

    If (intPreviousIndex <> ColumnHeader.Index - 1) And (blnBeenClicked = True) Then
        ' A different column was clicked.
        intSortOrder = 0
    End If

    ' Sort the Matches List.
    lvwMatches.SortKey = ColumnHeader.Index - 1
    lvwMatches.SortOrder = intSortOrder
    lvwMatches.Sorted = True
    
    intPreviousIndex = ColumnHeader.Index - 1
    intSortOrder = intSortOrder + (-1) ^ intSortOrder
    blnBeenClicked = True

End Sub

Private Sub lvwMatches_DblClick()
Dim i As Integer

    If Not (lvwMatches.SelectedItem Is Nothing) Then
        ' Show the Book's Additional Information.
        i = BookIndex(lvwMatches.SelectedItem.Text)
        
        If (i > 0) Then
            ' Found the Selected Book.
            Set frmBookInfo.gclsTargetBook = gclsBooks(i)
            
            Load frmBookInfo
            frmBookInfo.Show vbModal, Me
        End If
    End If
End Sub

Private Sub lvwMatches_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, c As Integer

    If ((KeyCode = vbKeyDelete) And (Not (lvwMatches.SelectedItem Is Nothing))) Then
        ' Ask the user if he's sure he wants to delete the item.
        If (MsgBox("Are you sure you want to remove '" & lvwMatches.SelectedItem.Text & "'?", vbYesNo Or vbQuestion, "Amazon Book Searcher") = vbNo) Then Exit Sub
        
        ' Delete the selected item.
        
        ' Find the book by its title and delete it.
        i = BookIndex(lvwMatches.SelectedItem.Text)
        
        If (i <> 0) Then
            ' Found the Selected Book.
            
            If (i < UBound(gclsBooks)) Then
                ' Remove the book.
                For c = i To UBound(gclsBooks) - 1
                    Set gclsBooks(c) = gclsBooks(c + 1)
                Next c
            End If
            
            ' Remove the last book.
            If (UBound(gclsBooks) = 1) Then
                ReDim gclsBooks(0 To 0)
            Else
                ReDim Preserve gclsBooks(1 To UBound(gclsBooks) - 1)
            End If
            
            ' Remove the book from the listview.
            lvwMatches.ListItems.Remove lvwMatches.SelectedItem.Index
        End If
    End If
End Sub


Private Sub mnuFileItem_Click(Index As Integer)
Dim i As Integer, c As Integer, s() As String
Dim itmBook As ListItem

    On Error GoTo ErrHandler
    
    Select Case Index
        Case 0
            ' Open a Booklist File.
            
            cmd1.CancelError = True
            cmd1.Filter = "All Files (*.*)|*.*|Amazon Booklist Files (*.abl)|*.abl"
            cmd1.FilterIndex = 2
            
            cmd1.ShowOpen
            
            Me.MousePointer = vbHourglass
            
            If (Dir$(cmd1.Filename) = "") Then
                ' File not found.
                MsgBox "The file '" + cmd1.Filename + "' wasn't found!", vbExclamation, "Amazon Book Searcher"
                Exit Sub
            End If
            
            OpenBooklist (cmd1.Filename)
            
            lvwMatches.ListItems.Clear
            For i = 1 To UBound(gclsBooks)
                If (gclsBooks(i).AuthorCount > 0) Then
                    ' Build the Book Author Array.
                    ReDim s(1 To gclsBooks(i).AuthorCount)
                    For c = 1 To gclsBooks(i).AuthorCount
                        s(c) = gclsBooks(i).Authors(c)
                    Next c
                End If
                
                ' Add the Book's Title.
                Set itmBook = lvwMatches.ListItems.Add(, gclsBooks(i).Title, gclsBooks(i).Title)
                
                ' Add the Book's Authors.
                itmBook.SubItems(1) = Join(s, ", ")
                ' Add the Book's Average Rating.
                If (gclsBooks(i).AverageRating > 0) Then
                    itmBook.SubItems(2) = gclsBooks(i).AverageRating & " out of 5"
                Else
                    itmBook.SubItems(2) = "N\A"
                End If
                If (gclsBooks(i).Cover <> "") Then
                    ' Add the Book's Cover Type.
                    itmBook.SubItems(3) = gclsBooks(i).Cover
                Else
                    itmBook.SubItems(3) = "N\A"
                End If
                ' Add the Book's Price.
                itmBook.SubItems(4) = "$" & gclsBooks(i).Price
            Next i
            
            lblStat.Caption = lvwMatches.ListItems.Count & " matches loaded."
            
            Me.MousePointer = vbDefault
            
            Call lblMatches_Click
            
            ' Enable the "Save Booklist" Menu item.
            mnuFileItem(1).Enabled = True
            
        Case 1
            ' Save a Booklist File.
            
            cmd1.CancelError = True
            cmd1.Filter = "All Files (*.*)|*.*|Amazon Booklist Files (*.abl)|*.abl"
            cmd1.FilterIndex = 2
            
            cmd1.ShowSave
            
            If (Dir(cmd1.Filename) <> "") Then
                ' File already exists - Prompt the user for the overwriting of the file.
                If (MsgBox("The file '" & cmd1.Filename & "' already exists, would you like to overwrite it?", vbYesNo Or vbQuestion, "Amazon Book Searcher") = vbNo) Then Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            SaveBooklist (cmd1.Filename)
            
            Me.MousePointer = vbDefault
        
        Case 2
            ' Merge 2 Booklists.
            
            Load frmMergeBooklists
            frmMergeBooklists.Show vbModal, Me
        
        Case 4
            ' Exit.
            Unload Me
    End Select
    
    Exit Sub
    
ErrHandler:
    ' The user pressed "Cancel".
    Exit Sub
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            ' Help.
            MsgBox "With the Amazon Book Searcher, you can search for books which are found on the Amazon.com website." & vbCrLf & _
            "This program has many features, including advanced searching and the ability to save and edit book lists."
            
        Case 2
            ' About.
            MsgBox "Amazon Book Searcher, " & vbCrLf & "Created by Yaron Budowski", vbInformation, "Amazon Book Searcher"
    End Select
End Sub

Private Sub txtPublicationDate_KeyPress(KeyAscii As Integer)
    If (Not ((KeyAscii >= 48) And (KeyAscii <= 57))) And (KeyAscii <> Asc(vbBack)) Then
        ' The Key isn't a number, ignore it.
        KeyAscii = 0
    End If
End Sub
