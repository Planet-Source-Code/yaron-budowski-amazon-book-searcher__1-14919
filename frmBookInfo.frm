VERSION 5.00
Begin VB.Form frmBookInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Information"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   ScaleHeight     =   4155
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   1320
      TabIndex        =   16
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblSalesRank2 
      AutoSize        =   -1  'True
      Caption         =   "50,236"
      Height          =   210
      Left            =   2040
      TabIndex        =   20
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblSalesRank 
      AutoSize        =   -1  'True
      Caption         =   "Sales Rank"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   2040
      TabIndex        =   19
      Top             =   2760
      Width           =   810
   End
   Begin VB.Label lblAverageRating2 
      AutoSize        =   -1  'True
      Caption         =   "4.5 out of 5"
      Height          =   210
      Left            =   2040
      TabIndex        =   18
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label lblAverageRating 
      AutoSize        =   -1  'True
      Caption         =   "Average Rating"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   2040
      TabIndex        =   17
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label lblDimensions2 
      AutoSize        =   -1  'True
      Caption         =   "2.74 x 9.20 x 7.08"
      Height          =   210
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1305
   End
   Begin VB.Label lblDimensions 
      AutoSize        =   -1  'True
      Caption         =   "Dimensions (in inches)"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1635
   End
   Begin VB.Label lblISBN2 
      AutoSize        =   -1  'True
      Caption         =   "1565921496"
      Height          =   210
      Left            =   2040
      TabIndex        =   13
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label lblISBN 
      AutoSize        =   -1  'True
      Caption         =   "ISBN"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   2040
      TabIndex        =   12
      Top             =   1800
      Width           =   345
   End
   Begin VB.Label lblPublishers2 
      AutoSize        =   -1  'True
      Caption         =   "O'Reilly && Associates"
      Height          =   210
      Left            =   2040
      TabIndex        =   11
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label lblPublishers 
      AutoSize        =   -1  'True
      Caption         =   "Publishers"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   2040
      TabIndex        =   10
      Top             =   1320
      Width           =   750
   End
   Begin VB.Label lblPublishingDate2 
      AutoSize        =   -1  'True
      Caption         =   "October 1996"
      Height          =   210
      Left            =   2040
      TabIndex        =   9
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lblPublishingDate 
      AutoSize        =   -1  'True
      Caption         =   "Publication Date"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lblPages2 
      AutoSize        =   -1  'True
      Caption         =   "652"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label lblPages 
      AutoSize        =   -1  'True
      Caption         =   "Pages"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label lblCover2 
      AutoSize        =   -1  'True
      Caption         =   "Paperback"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label lblCover 
      AutoSize        =   -1  'True
      Caption         =   "Cover"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label lblPrice2 
      AutoSize        =   -1  'True
      Caption         =   "$99"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label lblPrice 
      AutoSize        =   -1  'True
      Caption         =   "Price"
      ForeColor       =   &H80000011&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   360
   End
   Begin VB.Label lblAuthors 
      AutoSize        =   -1  'True
      Caption         =   "by Authors"
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   810
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Book Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmBookInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The Book shown by this form (can be changed
' from outside the form).
Public gclsTargetBook As clsBook

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim s() As String
Dim i As Integer

    lblTitle.MousePointer = vbCustom
    lblTitle.MouseIcon = LoadPicture(App.Path & "\hand.ico")

    ' Display the Target Book's Information.
    
    lblTitle.Caption = Replace(gclsTargetBook.Title, "&", "&&")
    
    If (gclsTargetBook.AuthorCount > 0) Then
        ReDim s(1 To gclsTargetBook.AuthorCount)
        For i = 1 To gclsTargetBook.AuthorCount
            s(i) = gclsTargetBook.Authors(i)
        Next i
        lblAuthors.Caption = "by " & Join(s, ", ")
    Else
        lblAuthors.Caption = ""
    End If
    
    If (gclsTargetBook.Price <> 0) Then
        lblPrice2.Caption = "$" & gclsTargetBook.Price
    Else
        lblPrice2.Caption = "N\A"
    End If
    
    If (gclsTargetBook.Cover <> "") Then
        lblCover2.Caption = gclsTargetBook.Cover
    Else
        lblCover2.Caption = "N\A"
    End If
    
    If (gclsTargetBook.Pages <> 0) Then
        lblPages2.Caption = gclsTargetBook.Pages
    Else
        lblPages2.Caption = "N\A"
    End If
    
    If (gclsTargetBook.Pages <> 0) Then
        lblDimensions2.Caption = gclsTargetBook.Dimensions(1) & " x " & gclsTargetBook.Dimensions(2) & " x " & gclsTargetBook.Dimensions(3)
    Else
        lblDimensions2.Caption = "N\A"
    End If
    
    If (gclsTargetBook.PublishingDate <> "") Then
        lblPublishingDate2.Caption = gclsTargetBook.PublishingDate
    Else
        lblPublishingDate2.Caption = "N\A"
    End If
    
    If (gclsTargetBook.Publishers <> "") Then
        lblPublishers2.Caption = Replace(gclsTargetBook.Publishers, "&", "&&")
    Else
        lblPublishers2.Caption = "N\A"
    End If
    
    If (gclsTargetBook.ISBN <> "") Then
        lblISBN2.Caption = gclsTargetBook.ISBN
    Else
        lblISBN2.Caption = "N\A"
    End If
    
    If (gclsTargetBook.AverageRating <> 0) Then
        lblAverageRating2.Caption = gclsTargetBook.AverageRating & " out of 5"
    Else
        lblAverageRating2.Caption = "N\A"
    End If
    
    If (gclsTargetBook.SalesRank <> 0) Then
        lblSalesRank2.Caption = gclsTargetBook.SalesRank
    Else
        lblSalesRank2.Caption = "N\A"
    End If
    
    If (lblTitle.Width > lblAuthors.Width) Then
        Me.Width = lblTitle.Width + lblTitle.Left + 200
    Else
        Me.Width = lblAuthors.Width + lblAuthors.Left + 200
    End If
    If (lblPublishers.Left + lblPublishers.Width > Me.Width) Then
        Me.Width = lblPublishers.Left + lblPublishers.Width + 200
    End If
    If (lblPublishingDate.Left + lblPublishingDate.Width > Me.Width) Then
        Me.Width = lblPublishingDate.Left + lblPublishingDate.Width + 200
    End If
    
    cmdOK.Left = Me.Width / 2 - cmdOK.Width / 2
End Sub

Private Sub lblTitle_Click()
    If (gclsTargetBook.URL <> "") Then
        ' Open the Book's URL in the default browser.
        OpenURL Me.hwnd, gclsTargetBook.URL
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTitle.ForeColor = vbRed
End Sub

Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTitle.ForeColor = &HC00000
End Sub
