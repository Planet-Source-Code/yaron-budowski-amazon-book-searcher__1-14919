VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Parameter Class (clsParameter.cls)
'
' Part of the Amazon Book Searcher Project.
' --------------------------------------------
'
' Purpose: Holds the Properties related to the Parameter
' of a Tag (This class is used by the Tag Tokenizer Class).
'


'
' Private Variables.
'


' The Parameter's Name.
Private mstrName As String
' The Parameter's Value.
Private mstrValue As String


'
' Public Properties.
'


'
' The Name Property.
'

Public Property Let Name(ByVal newName As String)
    mstrName = newName
End Property

Public Property Get Name() As String
    Name = mstrName
End Property

'
' The Value Property.
'

Public Property Let Value(ByVal newValue As String)
    mstrValue = newValue
End Property

Public Property Get Value() As String
    Value = mstrValue
End Property

