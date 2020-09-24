Attribute VB_Name = "modFileIO"
Option Explicit

'
' File I\O Module (modFileIO.bas)
'
' Part of the Amazon Book Searcher Project.
' --------------------------------------------
'
' Purpose: Holds the Subs and Functions related to
' file I\O (Opening, Saving and Merging Book List Files).
'

'
' Public Constants.
'

' The number of items to store in history (Meaning,
' the number of author, title and subject fields being
' stored in the history).
Public Const HISTORY_COUNT As Integer = 10
' The Settings filename.
Public Const SETTINGS_FILENAME As String = "Settings.dat"


'
' API Declarations.
'

' INI manipulation API's.
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


'
' Public Subs.
'


'
' Open the settings file.
'
Public Sub OpenSettings(Filename As String)
' A buffer variable used for opening the INI file.
Dim strBuffer As String
Dim s As String
Dim i As Integer

    If (Dir(Filename) = "") Then Exit Sub

    ' Create the Buffer.
    strBuffer = String(750, Chr(0))
    
    frmMain.Width = CInt(Left$(strBuffer, GetPrivateProfileString("WindowSettings", "Width", "7380", strBuffer, Len(strBuffer), Filename)))
    frmMain.Height = CInt(Left$(strBuffer, GetPrivateProfileString("WindowSettings", "Height", "6870", strBuffer, Len(strBuffer), Filename)))
    frmMain.Top = CInt(Left$(strBuffer, GetPrivateProfileString("WindowSettings", "Top", Str((Screen.Height - frmMain.Height) / 2), strBuffer, Len(strBuffer), Filename)))
    frmMain.Left = CInt(Left$(strBuffer, GetPrivateProfileString("WindowSettings", "Left", Str((Screen.Width - frmMain.Width) / 2), strBuffer, Len(strBuffer), Filename)))
    frmMain.WindowState = CInt(Left$(strBuffer, GetPrivateProfileString("WindowSettings", "WindowState", "0", strBuffer, Len(strBuffer), Filename)))
    
    For i = 1 To HISTORY_COUNT
        frmMain.cmbAuthor.AddItem Left$(strBuffer, GetPrivateProfileString("History", "Author" & i, "", strBuffer, Len(strBuffer), Filename))
    Next i

    For i = 1 To HISTORY_COUNT
        frmMain.cmbTitle.AddItem Left$(strBuffer, GetPrivateProfileString("History", "Title" & i, "", strBuffer, Len(strBuffer), Filename))
    Next i

    For i = 1 To HISTORY_COUNT
        frmMain.cmbSubject.AddItem Left$(strBuffer, GetPrivateProfileString("History", "Subject" & i, "", strBuffer, Len(strBuffer), Filename))
    Next i
End Sub


'
' Save the settings in a file.
'
Public Sub SaveSettings(Filename As String)
Dim s As String
Dim i As Integer
Dim F As Integer

    F = FreeFile
    Open Filename For Output Access Write As F
        Print #F, "[WindowSettings]"
        Print #F, "WindowState = " & frmMain.WindowState
        
        If (frmMain.WindowState = 0) Then
            Print #F, "Width = " & frmMain.Width
            Print #F, "Height = " & frmMain.Height
            Print #F, "Top = " & frmMain.Top
            Print #F, "Left = " & frmMain.Left
        End If
        
        Print #F, "[History]"
        For i = 0 To frmMain.cmbAuthor.ListCount - 1
            Print #F, "Author" & (i + 1) & " = " & frmMain.cmbAuthor.List(i)
        Next i
        
        For i = 0 To frmMain.cmbTitle.ListCount - 1
            Print #F, "Title" & (i + 1) & " = " & frmMain.cmbTitle.List(i)
        Next i
        
        For i = 0 To frmMain.cmbSubject.ListCount - 1
            Print #F, "Subject" & (i + 1) & " = " & frmMain.cmbSubject.List(i)
        Next i
    Close F
End Sub


'
' Merges 2 booklist files.
'
Public Sub MergeBooklists(FileName1 As String, FileName2 As String)
Dim i As Integer
Dim clsBooks() As clsBook
Dim clsOldBooks() As clsBook

    ' Exit the sub if one of the files doesn't exists.
    If ((Dir(FileName1) = "") Or (Dir(FileName2) = "")) Then Exit Sub
    
    If (UBound(gclsBooks) > 0) Then
        ' Save the old book list.
        ReDim clsOldBooks(1 To UBound(gclsBooks))
        
        For i = 1 To UBound(gclsBooks)
            Set clsOldBooks(i) = gclsBooks(i)
        Next i
    Else
        ReDim clsOldBooks(0 To 0)
    End If
    
    ' Open the first booklist file.
    OpenBooklist (FileName1)
    
    If (UBound(gclsBooks) = 0) Then Exit Sub
    
    ' Copy gclsBooks array into clsBooks array.
    ReDim clsBooks(1 To UBound(gclsBooks))
    
    For i = 1 To UBound(gclsBooks)
        Set clsBooks(i) = gclsBooks(i)
    Next i
    
    ' Open the second booklist file.
    OpenBooklist (FileName2)
    
    ' Append the first book list to the second booklist.
    For i = 1 To UBound(clsBooks)
        If (BookIndex(clsBooks(i).Title) = 0) Then
            ReDim Preserve gclsBooks(1 To UBound(gclsBooks) + 1)
            Set gclsBooks(UBound(gclsBooks)) = clsBooks(i)
        End If
    Next i
    
    ' Save the merged booklist into the second file.
    SaveBooklist (FileName2)
    
    ' Return to the original book list.
    If (UBound(clsOldBooks) > 0) Then
        ReDim gclsBooks(1 To UBound(clsOldBooks))
        For i = 1 To UBound(clsOldBooks)
            Set gclsBooks(i) = clsOldBooks(i)
        Next i
    End If
End Sub


'
' Saves the global variable array gclsBooks into
' a specific file.
'
Public Sub SaveBooklist(Filename As String)
Dim F As Integer
Dim i As Integer, c As Integer

    If (UBound(gclsBooks) = 0) Then Exit Sub
    
    F = FreeFile
    Open Filename For Output Access Write As F
        For i = 1 To UBound(gclsBooks)
            Print #F, "[Book" & i & "]"
            For c = 1 To gclsBooks(i).AuthorCount
                Print #F, "Author" & c & " = " & gclsBooks(i).Authors(c)
            Next c
            If (gclsBooks(i).AverageRating <> 0) Then Print #F, "AverageRating = " & gclsBooks(i).AverageRating
            If (gclsBooks(i).Cover <> "") Then Print #F, "Cover = " & gclsBooks(i).Cover
            If (gclsBooks(i).Dimensions(1) <> 0) Then
                Print #F, "Dimensions1 = " & gclsBooks(i).Dimensions(1)
                Print #F, "Dimensions2 = " & gclsBooks(i).Dimensions(2)
                Print #F, "Dimensions3 = " & gclsBooks(i).Dimensions(3)
            End If
            If (gclsBooks(i).ISBN <> "") Then Print #F, "ISBN = " & gclsBooks(i).ISBN
            If (gclsBooks(i).Pages <> 0) Then Print #F, "Pages = " & gclsBooks(i).Pages
            If (gclsBooks(i).Price <> 0) Then Print #F, "Price = " & gclsBooks(i).Price
            If (gclsBooks(i).Publishers <> "") Then Print #F, "Publishers = " & gclsBooks(i).Publishers
            If (gclsBooks(i).PublishingDate <> "") Then Print #F, "PublishingDate = " & gclsBooks(i).PublishingDate
            If (gclsBooks(i).SalesRank <> 0) Then Print #F, "SalesRank = " & gclsBooks(i).SalesRank
            If (gclsBooks(i).Title <> "") Then Print #F, "Title = " & gclsBooks(i).Title
            If (gclsBooks(i).URL <> "") Then Print #F, "URL = " & gclsBooks(i).URL
        Next i
    Close F
End Sub


'
' Opens a booklist file into the global variable
' array gclsBooks.
'
Public Sub OpenBooklist(Filename As String)
Dim F As Integer
Dim i As Integer, c As Integer
' A buffer variable used for opening the INI file.
Dim strBuffer As String
Dim s As String
Dim clsB As clsBook

    ' Exit the sub if the file doesn't exists.
    If (Dir$(Filename) = "") Then Exit Sub
    
    ReDim gclsBooks(0 To 0)
        
    i = 1
    
    Do
        ' Create a new instance of a Book class.
        Set clsB = New clsBook
        
        ' Create the Buffer.
        strBuffer = String(750, Chr(0))
        
        ' Retrieve the Authors from the INI file.
        c = 1
        Do
            s = Left$(strBuffer, GetPrivateProfileString("Book" & i, "Author" & c, "", strBuffer, Len(strBuffer), Filename))
            
            If (s <> "") Then
                ' Add the Author.
                clsB.AddAuthor s
            End If
            
            c = c + 1
        Loop Until (s = "")
        
        ' Retrieve the Average Rating from the INI file.
        clsB.AverageRating = CInt(Left$(strBuffer, GetPrivateProfileString("Book" & i, "AverageRating", "0", strBuffer, Len(strBuffer), Filename)))
        ' Retrieve the Cover from the INI file.
        clsB.Cover = Left$(strBuffer, GetPrivateProfileString("Book" & i, "Cover", "", strBuffer, Len(strBuffer), Filename))
        ' Retrieve the Dimensions from the INI file.
        clsB.Dimensions(1) = CInt(Left$(strBuffer, GetPrivateProfileString("Book" & i, "Dimensions1", "0", strBuffer, Len(strBuffer), Filename)))
        clsB.Dimensions(2) = CInt(Left$(strBuffer, GetPrivateProfileString("Book" & i, "Dimensions2", "0", strBuffer, Len(strBuffer), Filename)))
        clsB.Dimensions(3) = CInt(Left$(strBuffer, GetPrivateProfileString("Book" & i, "Dimensions3", "0", strBuffer, Len(strBuffer), Filename)))
        ' Retrieve the ISBN from the INI file.
        clsB.ISBN = Left$(strBuffer, GetPrivateProfileString("Book" & i, "ISBN", "", strBuffer, Len(strBuffer), Filename))
        ' Retrieve the Pages from the INI file.
        clsB.Pages = CInt(Left$(strBuffer, GetPrivateProfileString("Book" & i, "Pages", "0", strBuffer, Len(strBuffer), Filename)))
        ' Retrieve the Price from the INI file.
        clsB.Price = CInt(Left$(strBuffer, GetPrivateProfileString("Book" & i, "Price", "0", strBuffer, Len(strBuffer), Filename)))
        ' Retrieve the Publishers from the INI file.
        clsB.Publishers = Left$(strBuffer, GetPrivateProfileString("Book" & i, "Publishers", "", strBuffer, Len(strBuffer), Filename))
        ' Retrieve the Pusblishing Date from the INI file.
        clsB.PublishingDate = Left$(strBuffer, GetPrivateProfileString("Book" & i, "PublishingDate", "", strBuffer, Len(strBuffer), Filename))
        ' Retrieve the Sales Rank from the INI file.
        clsB.SalesRank = CInt(Left$(strBuffer, GetPrivateProfileString("Book" & i, "SalesRank", "0", strBuffer, Len(strBuffer), Filename)))
        ' Retrieve the Title from the INI file.
        clsB.Title = Left$(strBuffer, GetPrivateProfileString("Book" & i, "Title", "", strBuffer, Len(strBuffer), Filename))
        ' Retrieve the URL from the INI file.
        clsB.URL = Left$(strBuffer, GetPrivateProfileString("Book" & i, "URL", "", strBuffer, Len(strBuffer), Filename))
        
        ' Add the book to the book array.
        If (UBound(gclsBooks) = 0) Then
            ReDim gclsBooks(1 To 1)
        Else
            ReDim Preserve gclsBooks(1 To UBound(gclsBooks) + 1)
        End If
        
        Set gclsBooks(UBound(gclsBooks)) = clsB
        
        i = i + 1
    Loop Until (GetPrivateProfileString("Book" & i, "Title", "", strBuffer, Len(strBuffer), Filename) = 0)
End Sub


'
' Public functions.
'


'
' Returns the Index of a book in the gclsBooks
' array by its title. Returns zero if book wasn't found.
'
Public Function BookIndex(Title As String) As Integer
Dim i As Integer
    For i = 1 To UBound(gclsBooks)
        If (gclsBooks(i).Title = Title) Then
            ' Found the book.
            BookIndex = i
            Exit Function
        End If
    Next i
    
    BookIndex = 0
End Function
