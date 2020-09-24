VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTagTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'
' Tag Tokenizer Class (clsTagTokenzier.cls)
'
' Part of the Amazon Book Searcher Project.
' --------------------------------------------
'
' Purpose: Holds the Properties and Subs related
' to the Tag Tokenizing (Or Tag Parsing) of a specific
' HTML String.
'


'
' Private Variables.
'


' The HTML String to be parsed the Tag Tokenizer.
Private mstrHTML As String
' The Current Tag's Name.
Private mstrTagName As String
' The Text Between the Current and previous tags.
Private mstrText As String
' Is the Current Tag a Closing Tag?
Private mblnClosingTag As Boolean
' The Tag's Parameters Array.
Private mclsParameters() As clsParameter
' The Current Offset in the HTML String (This variable
' is increased every time a tag is parsed) - Read Only.
' Contains the Start position of the current tag.
Private mlngOffset As Long
' Saves the End Position of the Previous Tag.
Private mlngPreviousTagEnd As Long


'
' Public Properties.
'


'
' The HTML Property.
'

Public Property Let HTML(ByVal newHTML As String)
    mstrHTML = newHTML
End Property

Public Property Get HTML() As String
    HTML = mstrHTML
End Property

'
' The TagName Property.
'

Public Property Let TagName(ByVal newTagName As String)
    mstrTagName = newTagName
End Property

Public Property Get TagName() As String
    TagName = mstrTagName
End Property


'
' The Text Property.
'

Public Property Let Text(ByVal newText As String)
    mstrText = newText
End Property

Public Property Get Text() As String
    Text = mstrText
End Property


'
' The ClosingTag Property.
'

Public Property Let ClosingTag(ByVal newClosingTag As Boolean)
    mblnClosingTag = newClosingTag
End Property

Public Property Get ClosingTag() As Boolean
    ClosingTag = mblnClosingTag
End Property


'
' The Parameters Property.
'

Public Property Let Parameters(Index As Integer, ByVal newParameter As clsParameter)
    Set mclsParameters(Index) = newParameter
End Property

Public Property Get Parameters(Index As Integer) As clsParameter
    Set Parameters = mclsParameters(Index)
End Property


'
' The ParameterCount Property (Read Only).
'

Public Property Get ParameterCount() As Integer
    ParameterCount = UBound(mclsParameters)
End Property


'
' The Offset Property (Read Only).
'

Public Property Get Offset() As Long
    Offset = mlngOffset
End Property


'
' Public Functions.
'


'
' Retrieves the Next Tag from the HTML Code and saves its
' properties in public variables. It also returns
' the name of the tag (Although it can be obtained
' from the global variable "TagName").
'
Public Function NextTag() As String
' Tag Coordinations.
Dim TagStart As Long, TagEnd As Long, NextStart As Long, OldTagEnd As Long, CommentTagStart As Long
' Tag Parameters Coordinations.
Dim ParamStart As Long, ParamEnd As Long, ParamEnd2 As Long
' The Tag String itself.
Dim TagString As String, OldTagString As String

    ' Exit the Sub the HTML Code is Empty.
    If (Trim$(HTML) = "") Then
        ' Return a null Tag.
        mstrTagName = ""
        mstrText = ""
        mblnClosingTag = False
        mlngOffset = 0
        ReDim mclsParameters(0 To 0)
        Exit Function
    End If

    ' Load the Coordinations from the global variables.
    TagStart = mlngOffset
    OldTagEnd = mlngPreviousTagEnd
    
    ' A general-purpose loop used for for executing
    ' the code that comes right at the end of the sub
    ' and needs to be executed in all cases, including
    ' cases in which the "Exit Sub" statement was
    ' originally used.
    Do
    
        ' This loop is used to skip "bad" tags -
        ' <Bad Tag <Inside Tag> <Good Tag>.
        '
        ' This skips "Bad Tag" and "Inside Tag" and
        ' executes "Good Tag" (and Comment Tags as well).
        Do
        
            ' Seek for any Comment Tags.
            Do
                TagStart = InStr(TagStart + 1, mstrHTML, "<")
                If ((TagStart + 3 <= Len(mstrHTML)) And (TagStart > 0)) Then
                    ' See if it's a comment tag.
                    If (Mid$(mstrHTML, TagStart, 4) = "<!--") Then
                        ' Seek for the end of the comment tag.
                        'TagEnd = InStr(TagStart, mstrHTML, "-->")
                        
                        'If (TagEnd > 0) Then TagStart = TagEnd: OldTagEnd = TagEnd
                        
                        ' Skip to the next tag.
                        TagStart = InStr(TagStart + 1, mstrHTML, "<")
                    Else
                        Exit Do
                    End If
                End If
            Loop Until (TagStart = 0)
            
            ' Didn't find another tag.
            If (TagStart = 0) Then Exit Do
            
            TagEnd = InStr(TagStart, mstrHTML, ">")
            If (TagStart < Len(mstrHTML)) Then
                NextStart = InStr(TagStart + 1, mstrHTML, "<")
            Else
                NextStart = 0
            End If
            
            If (TagEnd > 0) And (TagEnd > NextStart) And (NextStart > 0) Then
                TagStart = TagEnd
            End If
            
        Loop Until ((TagEnd = 0) Or (TagEnd < NextStart) Or (TagEnd > 0) And (NextStart = 0))
        
        If (TagStart = 0) Then
            ' Didn't find another tag.
            mstrText = Mid$(mstrHTML, OldTagEnd + 1)
            ' Return a null Tag.
            mstrTagName = ""
            mblnClosingTag = False
            ReDim mclsParameters(0 To 0)
            mlngOffset = Len(mstrHTML)
            
            Exit Do
        ElseIf (TagEnd = 0) Then
            ' Didn't find the end of the tag.
            mstrText = Mid$(mstrHTML, OldTagEnd + 1)
            ' Return a null Tag.
            mstrTagName = ""
            mblnClosingTag = False
            ReDim mclsParameters(0 To 0)
            mlngOffset = Len(mstrHTML)
            
            Exit Do
        End If
        
        ' Save the Text between the Tags.
        mstrText = Mid$(mstrHTML, OldTagEnd + 1, TagStart - OldTagEnd - 1)
        ' Save the Tag String.
        TagString = Trim$(Mid$(mstrHTML, TagStart + 1, TagEnd - TagStart - 1))
        ' Delete Tabs, CrLf's and Double Spaces.
        TagString = Replace(TagString, vbTab, " ")
        TagString = Replace(TagString, vbCrLf, " ")
        
        OldTagString = "<" & TagString & ">"
        
        OldTagEnd = TagEnd
        ' Save the Coordinations in the global variables.
        mlngPreviousTagEnd = OldTagEnd
        mlngOffset = TagStart
        
        '
        ' Get the Tag's Name and Parameters.
        '
        
        mstrTagName = ""
        mblnClosingTag = False
        ReDim mclsParameters(0 To 0)
        
        '
        ' See if the Tag is a Closing Tag.
        '
        
        If (Left$(TagString, 1) = "/") Then
            ' Closing Tag.
            mblnClosingTag = True
            TagString = Trim$(Mid$(TagString, 2))
        Else
            mblnClosingTag = False
        End If
    
        '
        ' Find the Tag's Name.
        '
        
        ParamStart = InStr(1, TagString, " ")
        
        If (ParamStart = 0) Then
            mstrTagName = Mid$(TagString, 1)
            TagString = ""
        Else
            mstrTagName = Mid$(TagString, 1, ParamStart - 1)
            TagString = Trim$(Mid$(TagString, ParamStart))
        End If
        
        '
        ' Find the Tag's Parameters.
        '
        
        Do While (TagString <> "")
            ParamEnd = InStr(1, TagString, " ")
            ParamEnd2 = InStr(1, TagString, "=")
            
            If ((ParamEnd > ParamEnd2) Or (ParamEnd = 0) And (ParamEnd2 > 0)) Then
                ParamEnd = ParamEnd2
            End If
            
            If ParamEnd = 0 Then
                ' Didn't find another Param Delimiter.
                If (UBound(mclsParameters) = 0) Then
                    ReDim mclsParameters(1 To 1)
                Else
                    ReDim Preserve mclsParameters(1 To UBound(mclsParameters) + 1)
                End If
                
                Set mclsParameters(UBound(mclsParameters)) = New clsParameter
                mclsParameters(UBound(mclsParameters)).Name = Trim$(TagString)
                
                Exit Do
            End If
            
            ' Add the Parameter.
            If (UBound(mclsParameters) = 0) Then
                ReDim mclsParameters(1 To 1)
            Else
                ReDim Preserve mclsParameters(1 To UBound(mclsParameters) + 1)
            End If
            
            Set mclsParameters(UBound(mclsParameters)) = New clsParameter
            mclsParameters(UBound(mclsParameters)).Name = Left$(TagString, ParamEnd - 1)
            
            ' Find the Parameter's Value (if there is one).
            TagString = Trim$(Mid$(TagString, ParamEnd))
            
            If (Left$(TagString, 1) = "=") Then
                ' Parameter has a Value.
                TagString = Trim$(Mid$(TagString, 2))
                
                If (Left$(TagString, 1) = """") Then
                    ' The Value is a String.
                    
                    If (Len(TagString) > 1) Then
                        ParamEnd = InStr(2, TagString, """")
                    Else
                        ParamEnd = 0
                    End If
                    
                    If (ParamEnd = 0) Then
                        ' No Closing ".
                        ' Add the Tag String itself to the mstrText.
                        mstrText = mstrText & OldTagString
                        
                        Exit Do
                    End If
                    
                    ' Save the Parameter's Value.
                    mclsParameters(UBound(mclsParameters)).Value = Mid$(TagString, 2, ParamEnd - 2)
                    
                    If (ParamEnd = Len(TagString)) Then
                        TagString = ""
                    Else
                        TagString = Trim$(Mid$(TagString, ParamEnd + 1))
                    End If
                Else
                    ' The Value is one word long.
                    
                    If (Len(TagString) > 1) Then
                        ParamEnd = InStr(2, TagString, " ")
                    Else
                        ParamEnd = 0
                    End If
                    
                    If (ParamEnd = 0) Then
                        ' Save the Parameter's Value.
                        mclsParameters(UBound(mclsParameters)).Value = Left$(TagString, Len(TagString))
                        TagString = ""
                        Exit Do
                    End If
                    
                    ' Save the Parameter's Value.
                    mclsParameters(UBound(mclsParameters)).Value = Left$(TagString, ParamEnd - 1)
                    
                    TagString = Trim$(Mid$(TagString, ParamEnd))
                End If
            End If
            
        Loop ' End of Parameter Finding Loop.
        
        ' Exit the General-Purpose Loop (Because
        ' we want the general-purpose loop to run
        ' one time only).
        Exit Do
        
    Loop ' End of General-Purpose Loop.
    
    ' Delete any Comment tags within the text
    ' found between the tags.
    Do
        TagStart = InStr(1, mstrText, "<!--")
        
        If (TagStart > 0) Then
            ' Found a comment tag.
            
            ' Seek for the end of the Comment tag.
            TagEnd = InStr(TagStart, mstrText, "-->")
            
            If (TagEnd > 0) Then
                ' Delete the Comment Tag from the Text.
                mstrText = Left$(mstrText, TagStart - 1) & Mid$(mstrText, TagEnd + 3)
            Else
                ' The End of the Comment Tag wasn't found,
                ' Delete any text that comes afterwards.
                mstrText = Left$(mstrText, TagStart - 1)
            End If
        End If
    Loop Until (TagStart = 0)
    
    ' Delete Tabs, CrLf's and Double Spaces.
    mstrText = Replace(mstrText, vbTab, " ")
    mstrText = Replace(mstrText, vbCrLf, " ")
    mstrText = Replace(mstrText, vbCr, " ")
    mstrText = Replace(mstrText, vbLf, " ")
    mstrText = Replace(mstrText, "  ", " ")
    
    ' Return the Tag's Name.
    NextTag = mstrTagName
    
End Function ' End of NextTag Function.


'
' Private Subs.
'


' Initialize the Entire Tag Tokenizer Class.
Private Sub Class_Initialize()
    mstrHTML = ""
    mstrTagName = ""
    mstrText = ""
    mblnClosingTag = False
    ReDim mclsParameters(0 To 0)
    mlngOffset = 0
    mlngPreviousTagEnd = 0
End Sub
