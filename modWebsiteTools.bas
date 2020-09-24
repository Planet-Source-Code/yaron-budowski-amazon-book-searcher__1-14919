Attribute VB_Name = "modWebsiteTools"
Option Explicit

'
' Website Tools Module (modWebsiteTools.bas)
'
' Part of the Amazon Book Searcher Project.
' --------------------------------------------
'
' Purpose: Holds the Subs and Functions which are used as
' "Tools" in handling website activities (Such as loading url's,
' encoding url's and such).
'


'
' Public Cosntants.
'

' Results Per Page of a Book Results Page.
Public Const RESULTS_PER_PAGE As Integer = 25
' Various Constants related to the Search URL.
Public Const AMAZON_URL As String = "http://www.amazon.com"
Public Const SEARCH_URL_PREFIX As String = "/exec/obidos/search-handle-url/ix=books&rank=%2Bamzrank&fqp="
Public Const SEARCH_URL_SUFFIX As String = "&sz=25&pg=<Page Number>/ref=s_b_np/002-0029107-7786404"
Public Const BOOK_PAGE_URL As String = "http://www.amazon.com/exec/obidos/ASIN/<ISBN>/qid=979490648/sr=2-1/ref=sc_b_1/002-0029107-7786404"

Public Const SW_SHOW = 5
Public Const SYNCHRONIZE = &H100000
Public Const INFINITE = -1&


'
' API Declarations
'

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


'
' Public Variables.
'


' An Array of the Books found by the book search.
Public gclsBooks() As clsBook
' The Number of Matches found.
Public glngResultsCount As Long
' The Current Results Page and Book Page being downloaded.
Public gintCurrentPage As Long, glngCurrentBook As Integer
' The Number of Results Pages.
Public gintResultPageCount As Integer
' A Global flag variable used for stopping the search.
Public gblnStopSearch As Boolean


'
' Public Subs.
'


'
' Opens a URL in the default web browser.
'
Public Sub OpenURL(ParentHwnd As Long, ByVal URL As String)
    'Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, SW_SHOW)
    ShellExecute ParentHwnd, "open", URL, vbNullString, vbNullString, SW_SHOW
End Sub

'
' Parses an Amazon Book Search Results Page, and
' retrieves the info of the books found.
' These Books (and their info's) are saved in a
' global variable array.
'
Public Sub ParseResultPage(ByVal HTML As String)
' A List Item used when adding a book entry to
' the Book Results ListView.
Dim itmBook As ListItem
' The Tag Tokenizer used for parsing the search results page.
Dim clsTT As clsTagTokenizer
' Holds the number of the current resulting book
' and the number of first book of the result page.
Dim intCurrentBook As Integer, intStartBook As Integer
' Counter Variables.
Dim i As Integer, c As Integer
' Temp Variable used for parsing.
Dim s1 As String, s2() As String
' A Flag Variable used for telling if a book has
' already been found by the search.
Dim blnDuplicateBook As Boolean

    On Error GoTo ErrHandler
    
    ' See if we found any matches before even
    ' starting to parse the HTML code.
    i = InStr(1, HTML, "we were unable to find exact matches for your search for", vbTextCompare)
    
    If (i > 0) Then
        ' No matches were found.
        Exit Sub
    End If
    
    ' Skip a part of the HTML Code before starting
    ' to parse it (This makes the process faster) -
    ' We actually skip unnecessary information.
    i = InStr(1, HTML, "total matches for", vbTextCompare)
    
    If (i > 0) Then
        i = InStrRev(HTML, ">", i)
        
        If (i > 0) Then
            HTML = Mid$(HTML, i + 1)
        End If
    End If

    ' Load the Tag Tokenizer with the HTML Code.
    Set clsTT = New clsTagTokenizer
    clsTT.HTML = HTML
    
    Do
        ' Parse every Resulting Book's Info.
        Do
            ' Parse the Next Tag from the HTML Code.
            clsTT.NextTag
        Loop Until ((UCase$(clsTT.TagName) = "B") And (clsTT.ClosingTag = True) And _
            (((Trim$(clsTT.Text) = intCurrentBook & ".") And (intCurrentBook > 0)) Or ((Right$(Trim$(clsTT.Text), 1) = ".") And (intCurrentBook = 0))) Or _
            ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True) And (UCase$(Left$(Trim$(clsTT.Text), 2)) = "BY")) Or _
            ((UCase$(clsTT.TagName) = "IMG") And (UCase$(Trim$(clsTT.Text)) = "AVERAGE CUSTOMER REVIEW:")) Or _
            ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = False) And (UCase$(Trim$(clsTT.Text)) = "OUR PRICE:")) Or _
            ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True) And (InStr(1, clsTT.Text, "total matches for") > 0)) Or _
            (clsTT.Offset >= Len(HTML)))
        
        DoEvents
        
        If ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True) And (InStr(1, clsTT.Text, "total matches for") > 0)) Then
            ' Parse the Number of matches found.
            If (InStr(1, Trim$(clsTT.Text), " ") > 0) Then
                glngResultsCount = CLng(Left$(Trim$(clsTT.Text), InStr(1, Trim$(clsTT.Text), " ")))
            End If
        ElseIf ((UCase$(clsTT.TagName) = "B") And (clsTT.ClosingTag = True) And (((Trim$(clsTT.Text) = intCurrentBook & ".") And (intCurrentBook > 0)) Or ((Right$(Trim$(clsTT.Text), 1) = ".") And (intCurrentBook = 0)))) Then
            ' Parse the Book's URL, ISBN and Title.
            
            If (intCurrentBook = 0) Then
                ' Save the number of the first result in the result page.
                intCurrentBook = CInt(Left$(Trim$(clsTT.Text), InStr(1, Trim$(clsTT.Text), ".") - 1))
                intStartBook = intCurrentBook
            End If
            
            Do
                clsTT.NextTag
            Loop Until ((UCase$(clsTT.TagName) = "A") Or (clsTT.Offset >= Len(HTML)))
            
            DoEvents
            
            For i = 1 To clsTT.ParameterCount
                If ((UCase$(clsTT.Parameters(i).Name) = "HREF") And (clsTT.Parameters(i).Value <> "")) Then
                    ' Retrieve the Book's URL.
                    If (UBound(gclsBooks) = 0) Then
                        ReDim gclsBooks(1 To 1)
                    Else
                        ReDim Preserve gclsBooks(1 To UBound(gclsBooks) + 1)
                    End If
                    
                    Set gclsBooks(UBound(gclsBooks)) = New clsBook
                    gclsBooks(UBound(gclsBooks)).URL = AMAZON_URL & clsTT.Parameters(i).Value
                    
                    ' Retrieve the Book's ISBN.
                    s2() = Split(Mid$(Trim$(clsTT.Parameters(i).Value), 2), "/")
                    
                    If (UBound(s2) >= 3) Then
                        gclsBooks(UBound(gclsBooks)).ISBN = Trim$(s2(3))
                    End If
                    
                    ' Retrieve the Book's Title.
                    Do
                        clsTT.NextTag
                    Loop Until (((UCase$(clsTT.TagName) = "A") And (clsTT.ClosingTag = False)) Or (clsTT.Offset >= Len(HTML)))
                    
                    If (clsTT.ParameterCount > 0) Then
                        If ((UCase$(clsTT.Parameters(1).Name) = "HREF") And (AMAZON_URL & clsTT.Parameters(1).Value = gclsBooks(UBound(gclsBooks)).URL)) Then
                            clsTT.NextTag
                            
                            If ((UCase$(clsTT.TagName) = "A") And (clsTT.ClosingTag = True) And (Trim$(clsTT.Text) <> "")) Then
                                gclsBooks(UBound(gclsBooks)).Title = Trim$(clsTT.Text)
                            End If
                        End If
                    End If
                    
                    intCurrentBook = intCurrentBook + 1
                    
                    If ((intCurrentBook - intStartBook > RESULTS_PER_PAGE) And (intStartBook > 1)) Then
                        ' Since we've reached more than the expected results per page,
                        ' we can now safetly exit the parsing proccess.
                        Exit Do
                    End If
                End If
            Next i
            
        ElseIf ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True) And (UCase$(Left$(Trim$(clsTT.Text), 2)) = "BY")) Then
            ' Parse the Book's Authors, Cover and Publishing Date.
            
            ' Parse the Book's Authors.
            
            s2() = Split(Trim$(Mid$(clsTT.Text, 3)), ",")
            
            For i = 0 To UBound(s2)
                If (Trim$(s2(i)) <> "") Then
                    ' Add the Author to the List.
                    gclsBooks(UBound(gclsBooks)).AddAuthor (Trim$(s2(i)))
                End If
            Next i
            
            ' Parse the Book's Cover and Publishing Date.
            
            clsTT.NextTag
            
            If ((UCase$(clsTT.TagName) = "BR") And (Left$(Trim$(clsTT.Text), 1) = "(") And (Right$(Trim$(clsTT.Text), 1) = ")")) Then
                s2() = Split(Trim$(Mid$(Trim$(clsTT.Text), 2, Len(Trim$(clsTT.Text)) - 2)), "-")
                
                For i = 0 To UBound(s2)
                    If (i = 0) Then
                        gclsBooks(UBound(gclsBooks)).Cover = Trim$(s2(0))
                    ElseIf (i = 1) Then
                        gclsBooks(UBound(gclsBooks)).PublishingDate = Trim$(s2(1))
                    End If
                Next i
            End If
        
        ElseIf ((UCase$(clsTT.TagName) = "IMG") And (UCase$(Trim$(clsTT.Text)) = "AVERAGE CUSTOMER REVIEW:")) Then
            ' Parse the Book's Average Customer Rating.
            
            For i = 1 To clsTT.ParameterCount
                If (UCase$(clsTT.Parameters(i).Name) = "ALT") Then
                
                    If ((clsTT.Parameters(i).Value <> "") And (InStr(1, clsTT.Parameters(i).Value, " ") > 0)) Then
                        gclsBooks(UBound(gclsBooks)).AverageRating = CSng(Left$(clsTT.Parameters(i).Value, InStr(1, clsTT.Parameters(i).Value, " ")))
                    End If
                    
                    Exit For
                End If
            Next i
        
        ElseIf ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = False) And (UCase$(Trim$(clsTT.Text)) = "OUR PRICE:")) Then
            ' Parse the Book's Price.
            
            clsTT.NextTag
            
            If ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True)) Then
                gclsBooks(UBound(gclsBooks)).Price = CSng(Mid$(Trim$(clsTT.Text), 2))
            End If
            
            ' Since this is the last bit of info we can find on this book,
            ' we can now add the book to the Listview.
            
            If (gclsBooks(UBound(gclsBooks)).AuthorCount > 0) Then
                ' Build the Book Author Array.
                ReDim s2(1 To gclsBooks(UBound(gclsBooks)).AuthorCount)
                For i = 1 To gclsBooks(UBound(gclsBooks)).AuthorCount
                    s2(i) = gclsBooks(UBound(gclsBooks)).Authors(i)
                Next i
            End If
            
            ' Add the Book's Title.
            Set itmBook = frmMain.lvwMatches.ListItems.Add(, gclsBooks(UBound(gclsBooks)).Title, gclsBooks(UBound(gclsBooks)).Title)
            
            If (blnDuplicateBook = True) Then
                ' The Book has already been found by
                ' the search, delete it from the array.
                If (UBound(gclsBooks) = 1) Then
                    ReDim gclsBooks(0 To 0)
                Else
                    ReDim Preserve gclsBooks(1 To UBound(gclsBooks) - 1)
                End If
                
                blnDuplicateBook = False
            Else
                ' Add the Book's Authors.
                itmBook.SubItems(1) = Join(s2, ", ")
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
                
                frmMain.lvwMatches.Refresh
                
                DoEvents
            End If
            
        End If
        
        DoEvents
    Loop Until (clsTT.Offset >= Len(HTML))
    
    Exit Sub
    
ErrHandler:
    If (Err.Number = 35602) Then
        ' The Book has already been added.
        blnDuplicateBook = True
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
        
        Resume Next
    End If
    
End Sub


'
' Public Functions.
'


'
' Parses the HTML Code of an Amazon Book page,
' and returns the information about the specific book
' in a book class.
'
Public Function ParseBookPage(HTML As String) As clsBook
' The Tag Tokenizer used for parsing the Book Page.
Dim clsTT As clsTagTokenizer
' The Book Class to be returned.
Dim clsReturnBook As clsBook
' Temp variables.
Dim s() As String, s2 As String, s3() As String
Dim i As Integer, c As Integer

    ' Load the HTML Code onto the Tag Tokenizer.
    Set clsTT = New clsTagTokenizer
    clsTT.HTML = HTML
    
    Set clsReturnBook = New clsBook
    
    '
    ' Retrieve the Book's Information.
    '
    
    
    '
    ' An outer loop used for retrieving all necessary
    ' information about the book (Until reaching the
    ' end of the HTML Code).
    '
    Do
        
        '
        ' An inner loop used for parsing all tags
        ' until reaching an important piece of information
        ' about the book.
        '
        Do
            ' Parse the Next Tag from the HTML Code.
            clsTT.NextTag
        Loop Until ((UCase$(clsTT.TagName) = "TITLE") And (clsTT.ClosingTag = True)) Or _
            ((UCase$(clsTT.TagName) = "A") And (clsTT.ClosingTag = False) And (UCase$(Trim$(clsTT.Text)) = "BY")) Or _
            ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = False) And (UCase$(Trim$(clsTT.Text)) = "OUR PRICE:")) Or _
            ((UCase$(clsTT.TagName) = "B") And (clsTT.ClosingTag = True) And ((UCase$(Trim$(clsTT.Text)) = "HARDCOVER") Or _
            (UCase$(Trim$(clsTT.Text)) = "PAPERBACK") Or (UCase$(Trim$(clsTT.Text)) = "DIGITAL") Or (UCase$(Trim$(clsTT.Text)) = "AUDIO CASSETTE"))) Or _
            ((InStr(1, clsTT.Text, "ISBN") > 0) And (UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True)) Or _
            ((UCase$(clsTT.TagName) = "B") And (clsTT.ClosingTag = True) And (UCase$(Trim$(clsTT.Text)) = "AMAZON.COM SALES RANK:")) Or _
            ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True) And (UCase$(Trim$(clsTT.Text)) = "AVG. CUSTOMER RATING:")) Or _
            (clsTT.Offset >= Len(HTML))
    
        If ((UCase$(clsTT.TagName) = "TITLE") And (clsTT.ClosingTag = True)) Then
            ' Retrieve the Book's Title.
            If ((Left$(clsTT.Text, 24) = "Amazon.com: buying info:") And (Len(clsTT.Text) > 25)) Then
                clsReturnBook.Title = Trim$(Mid$(clsTT.Text, 25))
            ElseIf (Left$(clsTT.Text, 21) = "Amazon.com Error Page") Then
                ' We found a missing page.
                clsReturnBook.Title = ""
                Set ParseBookPage = clsReturnBook
                Exit Function
            End If
        
        ElseIf ((UCase$(clsTT.TagName) = "A") And (clsTT.ClosingTag = False) And (UCase$(Trim$(clsTT.Text)) = "BY")) Then
            ' Retrieve the Book's Authors.
            Do
                clsTT.NextTag
                
                If ((UCase$(clsTT.TagName) = "A") And (clsTT.ClosingTag = True)) Then
                    ' Add the Author's Name to the list.
                    
                    clsReturnBook.AddAuthor Trim$(clsTT.Text)
                ElseIf ((UCase$(clsTT.TagName) = "A") And (clsTT.ClosingTag = False) And (Right$(Trim$(clsTT.Text), 1) = ",")) Then
                    ' There are more authors.
                Else
                    ' We've reached the end of the author list - Exit the loop.
                    Exit Do
                End If
            Loop
        
        ElseIf ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = False) And (UCase$(Trim$(clsTT.Text)) = "OUR PRICE:")) Then
            ' Retrieve the Book's Price.
            clsTT.NextTag
            
            If ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True)) Then
                clsReturnBook.Price = CSng(Mid$(Trim$(clsTT.Text), 2))
            End If
        
        ElseIf ((UCase$(clsTT.TagName) = "B") And (clsTT.ClosingTag = True) And ((UCase$(Trim$(clsTT.Text)) = "HARDCOVER") Or _
            (UCase$(Trim$(clsTT.Text)) = "PAPERBACK") Or (UCase$(Trim$(clsTT.Text)) = "DIGITAL") Or (UCase$(Trim$(clsTT.Text)) = "AUDIO CASSETTE"))) Then
            ' Retrieve the Book's Cover Type.
            If (UCase$(Trim$(clsTT.Text)) = "HARDCOVER") Then
                ' Hardcover.
                clsReturnBook.Cover = "Hardcover"
            ElseIf (UCase$(Trim$(clsTT.Text)) = "PAPERBACK") Then
                ' Paperback.
                clsReturnBook.Cover = "Paperback"
            ElseIf (UCase$(Trim$(clsTT.Text)) = "DIGITAL") Then
                ' Digital E-Book.
                clsReturnBook.Cover = "Digital E-Book"
            ElseIf (UCase$(Trim$(clsTT.Text)) = "AUDIO CASSETTE") Then
                ' Audio Cassette.
                clsReturnBook.Cover = "Audio Cassette"
            End If
            
            ' Retrieve the other book information that
            ' comes right after the cover type.
            clsTT.NextTag
            
            If (UCase$(clsTT.TagName) = "BR") Then
                ' Retrieve the Number of Pages and
                ' Publishing Date.

                s() = Split(Trim$(clsTT.Text), " ")
                
                For i = 0 To UBound(s)
                    's(i) = Trim$(s(i))
                    
                    If ((UCase$(s(i)) = "PAGES") And (i >= 1)) Then
                        ' The Number of Pages.
                        clsReturnBook.Pages = CInt(s(i - 1))
                    ElseIf (Left$(s(i), 1) = "(") Then
                        ' The Publishing Date.
                        
                        For c = i To UBound(s)
                            If (Left$(s(c), 1) = "(") Then Mid$(s(c), 1) = " "
                            clsReturnBook.PublishingDate = clsReturnBook.PublishingDate & " " & s(c)
                            
                            If (Right$(s(c), 1) = ")") Then clsReturnBook.PublishingDate = Left$(clsReturnBook.PublishingDate, Len(clsReturnBook.PublishingDate) - 1): Exit For
                        Next c
                        
                        clsReturnBook.PublishingDate = Trim$(clsReturnBook.PublishingDate)
                    End If
                Next i
            End If
        
        ElseIf ((InStr(1, clsTT.Text, "ISBN") > 0) And (UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True)) Then
            ' Retrieve the Book's Publishers, ISBN and Dimensions.
            
            ' Replace any special characters in the text.
            s2 = Replace(clsTT.Text, "&amp;", "&", 1, -1, vbTextCompare)
            s2 = Replace(s2, "&lt;", "<", 1, -1, vbTextCompare)
            s2 = Replace(s2, "&gt;", ">", 1, -1, vbTextCompare)
            
            ' Retrieve the Different parts of the text.
            s() = Split(s2, ";")
            
            For i = 0 To UBound(s)
                s(i) = Trim$(s(i))
                
                If (i = 0) And (s(i) <> "") Then
                    ' The Book's Publishers.
                    clsReturnBook.Publishers = Trim$(s(i))
                ElseIf (Left$(s(i), 5) = "ISBN:") Then
                    ' The Book's ISBN.
                    clsReturnBook.ISBN = Trim$(Mid$(s(i), 6))
                ElseIf (UCase$(Left$(s(i), 23)) = "DIMENSIONS (IN INCHES):") Then
                    ' The Book's Dimensions.
                    s3() = Split(Mid$(s(i), 24), "x", -1, vbTextCompare)
                    clsReturnBook.Dimensions(1) = CSng(Trim$(s3(0)))
                    clsReturnBook.Dimensions(2) = CSng(Trim$(s3(1)))
                    clsReturnBook.Dimensions(3) = CSng(Trim$(s3(2)))
                End If
            Next i
        
        ElseIf ((UCase$(clsTT.TagName) = "B") And (clsTT.ClosingTag = True) And (UCase$(Trim$(clsTT.Text)) = "AMAZON.COM SALES RANK:")) Then
            ' Retrieve the Book's Sales Rank.
            
            clsTT.NextTag
            clsReturnBook.SalesRank = CLng(Trim$(clsTT.Text))
        
        ElseIf ((UCase$(clsTT.TagName) = "FONT") And (clsTT.ClosingTag = True) And (UCase$(Trim$(clsTT.Text)) = "AVG. CUSTOMER RATING:")) Then
            ' Retrieve the Book's Average Rating.
            
            Do
                clsTT.NextTag
            Loop Until ((UCase$(clsTT.TagName) = "IMG") Or (clsTT.Offset >= Len(HTML)))
            
            For i = 1 To clsTT.ParameterCount
                If (UCase$(clsTT.Parameters(i).Name) = "ALT") Then
                
                    If ((clsTT.Parameters(i).Value <> "") And (InStr(1, clsTT.Parameters(i).Value, " ") > 0)) Then
                        clsReturnBook.AverageRating = CSng(Left$(clsTT.Parameters(i).Value, InStr(1, clsTT.Parameters(i).Value, " ")))
                    End If
                    
                    Exit For
                End If
            Next i
            
            ' Since this is last bit of info neccessary
            ' for the book, we can stop scannig the HTML Code.
            Exit Do
        End If

    Loop Until (clsTT.Offset >= Len(HTML)) ' End of Outer Loop.
    
    ' Return the Parsed Book.
    Set ParseBookPage = clsReturnBook
    
End Function


'
' Checks if the user if connected to the internet.
'
Public Function IsConnected() As Boolean
Dim hKey As Long
Dim lpSubKey As String
Dim phkResult As Long
Dim lpValueName As String
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As Long
Dim lpcbData As Long
Dim ReturnCode As Long
Const ERROR_SUCCESS = 0&
Const APINULL = 0&
Const HKEY_LOCAL_MACHINE = &H80000002

    IsConnected = False
    lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
    ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, _
    phkResult)
    
    If ReturnCode = ERROR_SUCCESS Then
        hKey = phkResult
        lpValueName = "Remote Connection"
        lpReserved = APINULL
        lpType = APINULL
        lpData = APINULL
        lpcbData = APINULL
        ReturnCode = RegQueryValueEx(hKey, lpValueName, _
        lpReserved, lpType, ByVal lpData, lpcbData)
        lpcbData = Len(lpData)
        ReturnCode = RegQueryValueEx(hKey, lpValueName, _
        lpReserved, lpType, lpData, lpcbData)
        
        If ReturnCode = ERROR_SUCCESS Then
            If lpData = 0 Then
                IsConnected = False
            Else
                IsConnected = True
            End If
        End If
    RegCloseKey (hKey)
    End If

End Function


'
' Encodes a String to fit URL Standards.
'
Public Function EncodeURL(ByVal s As String) As String
Dim i As Integer
Dim c As String

    ' Replace any characters which are not alphabet
    ' or numeric ones with hexadecimal signatures.
    
    i = 1
    Do
        If (Mid$(s, i, 1) = " ") Then
            ' Replace the Space with a "+".
            Mid$(s, i, 1) = "+"
        ElseIf Not (((Asc(UCase$(Mid$(s, i, 1))) >= 48) And _
            (Asc(UCase$(Mid$(s, i, 1))) <= 57)) Or _
            ((Asc(UCase$(Mid$(s, i, 1))) >= 65) And _
            (Asc(UCase$(Mid$(s, i, 1))) <= 90))) Then
            ' Replace the Character.
            c = Mid$(Hex$(Asc(Mid$(s, i, 1))), 1)
            If (Len(c) = 1) Then c = "0" & c
            s = Left$(s, i - 1) & "%" & c & Mid$(s, i + 1)
            
            i = i + Len("%" & c) - 1
        End If
        i = i + 1
    Loop Until (i > Len(s))

    ' Return the Encoded URL.
    EncodeURL = s

End Function

Public Function UrlDecode(ByVal sEncoded As String) As String
'========================================================
' Accept url-encoded string
' Return decoded string
'========================================================

Dim pointer    As Long      ' sEncoded position pointer
Dim pos        As Long      ' position of InStr target

If sEncoded = "" Then Exit Function

' convert "+" to space
pointer = 1
Do
   pos = InStr(pointer, sEncoded, "+")
   If pos = 0 Then Exit Do
   Mid$(sEncoded, pos, 1) = " "
   pointer = pos + 1
Loop
    
' convert "%xx" to character
pointer = 1

On Error GoTo errorUrlDecode

Do
   pos = InStr(pointer, sEncoded, "%")
   If pos = 0 Then Exit Do
   
   Mid$(sEncoded, pos, 1) = Chr$("&H" & (Mid$(sEncoded, pos + 1, 2)))
   sEncoded = Left$(sEncoded, pos) _
             & Mid$(sEncoded, pos + 3)
   pointer = pos + 1
Loop
On Error GoTo 0     'reset error handling
UrlDecode = sEncoded
Exit Function

errorUrlDecode:
'--------------------------------------------------------------------
' If this function was mistakenly called with the following:
'    UrlDecode("100% natural")
' a type mismatch error would be raised when trying to convert
' the 2 characters after "%" from hex to character.
' Instead, a more descriptive error message will be generated.
'--------------------------------------------------------------------
If Err.Number = 13 Then      'Type Mismatch error
   Err.Clear
   Err.Raise 65001, , "Invalid data passed to UrlDecode() function."
Else
   Err.Raise Err.Number
End If
Resume Next
End Function


