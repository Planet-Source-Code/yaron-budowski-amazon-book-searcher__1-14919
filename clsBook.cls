VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Book Class (clsBook.cls)
'
' Part of the Amazon Book Searcher Project.
' --------------------------------------------
'
' Purpose: Holds the Properties of an Amazon Book.
'


'
' Private Variables.
'


' The Book's Title.
Private mstrTitle As String
' The Book's Authors.
Private mstrAuthors() As String
' The Book's Price.
Private msngPrice As Single
' The Book's Number of Pages.
Private mintPages As Integer
' The Book's Cover Type (Paperback, HardCover and etc)
Private mstrCover As String
' The Book's Publishing Date.
Private mstrPublishingDate As String
' The Book's ISBN.
Private mstrISBN As String
' The Book's Dimensions (In Inches) -
' 1 for Thickness, 2 for Height, 3 for Width.
Private msngDimensions() As Single
' The Book's Publishers.
Private mstrPublishers As String
' The Book's Sales Rank.
Private mlngSalesRank As Long
' The Book's Average Customer Rating.
Private msngAverageRating As Single
' The Book's Page URL Address.
Private mstrURL As String


'
' Public Properties.
'


'
' The Title Property.
'

Public Property Let Title(ByVal newTitle As String)
    mstrTitle = newTitle
End Property

Public Property Get Title() As String
    Title = mstrTitle
End Property


'
' The Authors Property.
'

Public Property Let Authors(Index As Integer, ByVal newAuthors As String)
    mstrAuthors(Index) = newAuthors
End Property

Public Property Get Authors(Index As Integer) As String
    Authors = mstrAuthors(Index)
End Property


'
' The Author Count Property (Read Only).
'

Public Property Get AuthorCount() As Integer
    AuthorCount = UBound(mstrAuthors)
End Property


'
' The Price Property.
'

Public Property Let Price(ByVal newPrice As Single)
    msngPrice = newPrice
End Property

Public Property Get Price() As Single
    Price = msngPrice
End Property


'
' The Pages Property.
'

Public Property Let Pages(ByVal newPages As Integer)
    mintPages = newPages
End Property

Public Property Get Pages() As Integer
    Pages = mintPages
End Property


'
' The Cover Property.
'

Public Property Let Cover(ByVal newCover As String)
    mstrCover = newCover
End Property

Public Property Get Cover() As String
    Cover = mstrCover
End Property


'
' The PublishingDate Property.
'

Public Property Let PublishingDate(ByVal newPublishingDate As String)
    mstrPublishingDate = newPublishingDate
End Property

Public Property Get PublishingDate() As String
    PublishingDate = mstrPublishingDate
End Property


'
' The ISBN Property.
'

Public Property Let ISBN(ByVal newISBN As String)
    mstrISBN = newISBN
End Property

Public Property Get ISBN() As String
    ISBN = mstrISBN
End Property


'
' The Dimensions Property.
'

Public Property Let Dimensions(Index As Integer, ByVal newDimensions As Single)
    msngDimensions(Index) = newDimensions
End Property

Public Property Get Dimensions(Index As Integer) As Single
    Dimensions = msngDimensions(Index)
End Property


'
' The Publishers Property.
'

Public Property Let Publishers(ByVal newPublishers As String)
    mstrPublishers = newPublishers
End Property

Public Property Get Publishers() As String
    Publishers = mstrPublishers
End Property


'
' The SalesRank Property.
'

Public Property Let SalesRank(ByVal newSalesRank As Long)
    mlngSalesRank = newSalesRank
End Property

Public Property Get SalesRank() As Long
    SalesRank = mlngSalesRank
End Property


'
' The AverageRating Property.
'

Public Property Let AverageRating(ByVal newAverageRating As Single)
    msngAverageRating = newAverageRating
End Property

Public Property Get AverageRating() As Single
    AverageRating = msngAverageRating
End Property


'
' The URL Property.
'

Public Property Let URL(ByVal newURL As String)
    mstrURL = newURL
End Property

Public Property Get URL() As String
    URL = mstrURL
End Property


'
' Public Subs.
'


'
' Adds an author the author list.
'
Public Sub AddAuthor(Name As String)
    If (UBound(mstrAuthors) = 0) Then
        ReDim mstrAuthors(1 To 1)
    Else
        ReDim Preserve mstrAuthors(1 To UBound(mstrAuthors) + 1)
    End If
    
    mstrAuthors(UBound(mstrAuthors)) = Name
End Sub


'
' Private Subs.
'


' Initialize the Entire Book Class.
Private Sub Class_Initialize()
    mstrTitle = ""
    ReDim mstrAuthors(0 To 0)
    msngPrice = 0
    mintPages = 0
    mstrCover = ""
    mstrPublishingDate = ""
    mstrISBN = ""
    ReDim msngDimensions(1 To 3)
    mstrPublishers = ""
    mlngSalesRank = 0
    msngAverageRating = 0
    mstrURL = ""
End Sub
