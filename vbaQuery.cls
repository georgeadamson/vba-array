VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbaQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit

Private stack As New VBAArray

Private Sub Class_Initialize()

End Sub


Public Property Get find(selector As String, Optional context As Variant)

    'Attribute find.VB_UserMemId = 0
    Attribute find.VB_UserMemId = 0

    Dim contexts As VBAArray
    Dim matches As New VBAArray
    Dim selectors() As String
    Dim r As Range
    Dim s As Variant
    Dim i As Long
    'Dim reSplit As Object
    'Set reSplit = CreateObject("vbscript.regexp")
    'reSplit.Global = True
    'reSplit.Pattern = ""
    If isNothing(context) Then Set context = defaultContext()

    selectors = Split(selector, ",")
    Set contexts = makeArray(context)

    'Search for each selector in each context:
    'TODO: Make VBAArray enumerable.
    For i = 0 To contexts.Count - 1
        Set r = contexts.Item(i)
        For Each s In selectors
            matches.Merge findNode(s, r)
        Next
    Next

    stack.Push matches
    Set find = stack.Last()

End Property



' Simply find one type of item in the document:
Private Function findNode(ByVal node As String, Optional context As Range) As VBAArray
        
    Dim objRow As Row
    Dim objTable As Table

    Dim matches As New VBAArray
    node = UCase(node)
    If isNothing(context) Then Set context = defaultContext()

    Select Case node: Case "TABLE", "*"

        For Each objTable In context.Tables
            matches.Push objTable
        Next

    End Select

    Select Case node: Case "ROW", "TR", "*"

        For Each objTable In context.Tables
            For Each objRow In objTable.Rows
                matches.Push objRow
            Next
        Next

    End Select

    Select Case node: Case "CELL", "TD", "*"

        Dim objCell As Cell
        For Each objTable In context.Tables
            For Each objRow In objTable.Rows
                For Each objCell In objRow.Cells
                    matches.Push objCell
                Next
            Next
        Next

    End Select

    Select Case node: Case "WORD", "*"

        Dim objWord As Range
        For Each objWord In context.Words
            matches.Push objWord
        Next

    End Select

    Select Case node: Case "FIELD", "*"

        Dim objField As Field
        For Each objField In context.Fields
            matches.Push objField
        Next

    End Select

    Select Case node: Case "HYPERLINK", "A", "*"

        Dim objHyperlink As Hyperlink
        For Each objHyperlink In context.Hyperlinks
            matches.Push objHyperlink
        Next

    End Select

    Select Case node: Case "LISTPARAGRAPH", "LI", "*"

        Dim objListParagraph As Paragraph
        For Each objHyperlink In context.ListParagraphs
            matches.Push objListParagraph
        Next

    End Select

    Select Case node: Case "LISTPARAGRAPH", "LI", "*"

        Dim objParagraph As Paragraph
        For Each objHyperlink In context.Paragraphs
            matches.Push objParagraph
        Next

    End Select
    
    'Return matches
    Set findNode = matches

End Function


Public Function makeArray(context As Variant) As VBAArray

    If isNothing(context) Then
    
        Set makeArray = New VBAArray

    ElseIf TypeName(context) = "Range" Then
    
        Dim r As Range
        Set r = context
        Set makeArray = New VBAArray
        makeArray.Push r
        
    Else
    
        Set makeArray = context
        
    End If


End Function


'Helper for returning the default context for search queries: (ie the range that represents the whole document)
Public Function defaultContext() As Range
    Set defaultContext = Application.ActiveDocument.Range
End Function


'Helper to return true if value is Missing or Nothing or Null:
Private Function isNothing(value As Variant)

    isNothing = True

    If IsMissing(value) Then
        isNothing = True
    ElseIf IsObject(value) Then
        If Not value Is Nothing Then isNothing = False
    ElseIf IsArray(value) Then
        isNothing = False
    ElseIf Not IsNull(value) Then
        isNothing = False
    End If

End Function

'Helper to return true unless value is Missing or Nothing or Null:
Private Function isSomething(value As Variant)

    isSomething = Not isNothing(value)

End Function

