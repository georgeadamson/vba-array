VERSION 5.00
Begin VB.Form frmEnumeratorFun 
   AutoRedraw      =   -1  'True
   Caption         =   "Enumerator Object Fun"
   ClientHeight    =   6375
   ClientLeft      =   2340
   ClientTop       =   1575
   ClientWidth     =   5625
   Icon            =   "frmEnumeratorFun.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   Begin VB.CommandButton cmdClone 
      Caption         =   "Clone The Enumerator After Item 5"
      Height          =   360
      Left            =   1042
      TabIndex        =   3
      Top             =   5910
      Width           =   3540
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset After Item 4 The First Time Through"
      Height          =   360
      Left            =   1042
      TabIndex        =   2
      Top             =   5445
      Width           =   3540
   End
   Begin VB.CommandButton cmdSkip 
      Caption         =   "Skip Odd Numbered Items"
      Height          =   360
      Left            =   1042
      TabIndex        =   1
      Top             =   4980
      Width           =   3540
   End
   Begin VB.TextBox txDisplay 
      Height          =   4800
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   5520
   End
End
Attribute VB_Name = "frmEnumeratorFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
  ' demo project showing how to use a typelibrary to access the
  ' Skip, Reset and Clone methods of a NewEnum enumerator object
  ' by Bryan Stafford of New Vision Software® - newvision@mvps.org
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.
  
  Private m_oCol As Collection

Private Sub Form_Load()

  Dim i&, oCls As cSimpleClass

  ' add some objects to our collection
  For i = 0 To 10
    Set oCls = New cSimpleClass
    With oCls
      .Name = "Item #" & CStr(i)
      .Number = i
    End With
    
    m_oCol.Add oCls, oCls.Name
    
    Set oCls = Nothing
  Next
  
  
  ' now, we can enumerate the collection using the same technique that the VB
  ' Fore...Each statement uses under the hood
  With txDisplay
    .Text = "Initial iteration of the collection" & vbCrLf
  End With

  ' oEnumer is an implementation of a VB friendly version of the IEnumVARIANT
  ' interface.  because of the way the parameters are passed in the standard
  ' definition of the IEnumVARIANT interface, VB cannot implement the interface.
  Dim oEnumer As IEnumVARIANTReDef
  Dim vVarRet As Variant, nFetched&
  
  ' get an enumerator object from the collection
  Set oEnumer = m_oCol.[_NewEnum]
  
  ' call the Next method until no more items are returned
  Do
    oEnumer.Next 1, vVarRet, VarPtr(nFetched)
    ' the variant "vVarRet" is passed into the Next call and the object at the
    ' current position in the collection is placed in it.  if a value greater than one
    ' were passed in the first param, an array of objects would be returned.

    ' nFetched tells us how many items were returned.  display the name from the class
    If nFetched Then
      With txDisplay
        .Text = .Text & vVarRet.Name & vbCrLf
      End With
    End If
    
    ' be sure to clean up properly
    Set vVarRet = Nothing
    
  Loop Until nFetched = 0
  
  Set oEnumer = Nothing


  ' show the form
  Show
  
  ' scroll to the bottom of the textbox
  With txDisplay
    .Text = .Text & vbCrLf
    .SelStart = Len(.Text)
  End With

End Sub
  
Private Sub cmdSkip_Click()

  Dim oEnumer As IEnumVARIANTReDef
  Dim i&, vVarRet As Variant, nFetched&

  ' display what we are doing....
  With txDisplay
    .Text = .Text & "Use the Skip method to skip all odd items" & vbCrLf
  End With

  ' get an enumerator object from the collection
  Set oEnumer = m_oCol.[_NewEnum]
    
  ' loop through the collection.  see the form_load event for further details
  Do
    oEnumer.Next 1, vVarRet, VarPtr(nFetched)
  
    If nFetched Then
      With txDisplay
        .Text = .Text & vVarRet.Name & vbCrLf
      End With
      
      ' if this is an even numbered item, skip the next item
      If (i Mod 2) = 0 Then
        oEnumer.Skip 1
        
        ' adjust i to account for the skip
        i = i + 1
      End If
      
      i = i + 1
    End If
    
    Set vVarRet = Nothing
    
  Loop Until nFetched = 0
  
  Set oEnumer = Nothing

  ' scroll to the bottom of the textbox
  With txDisplay
    .Text = .Text & vbCrLf
    .SelStart = Len(.Text)
  End With
  
End Sub

Private Sub cmdReset_Click()

  Dim oEnumer As IEnumVARIANTReDef
  Dim i&, vVarRet As Variant, nFetched&

  
  ' display what we are doing....
  With txDisplay
    .Text = .Text & "Use the Reset method to start from the beginning while in an iteration" & vbCrLf
  End With

  ' get an enumerator object from the collection
  Set oEnumer = m_oCol.[_NewEnum]
    
  ' loop through the collection.  see the form_load event for further details
  Do
    oEnumer.Next 1, vVarRet, VarPtr(nFetched)
  
    If nFetched Then
      With txDisplay
        .Text = .Text & vVarRet.Name & vbCrLf
      End With
      
      ' if this is item 4, reset the enumeration
      If i = 4 Then
        oEnumer.Reset
      
        With txDisplay
          .Text = .Text & "Resetting here" & vbCrLf
        End With
      End If
      
      i = i + 1
    End If
    
    Set vVarRet = Nothing
    
  Loop Until nFetched = 0
  
  Set oEnumer = Nothing


  ' scroll to the bottom of the textbox
  With txDisplay
    .Text = .Text & vbCrLf
    .SelStart = Len(.Text)
  End With

End Sub

Private Sub cmdClone_Click()

  Dim oEnumer As IEnumVARIANTReDef, oCloneEnumer As IEnumVARIANTReDef
  Dim i&, vVarRet As Variant, nFetched&

  
  ' display what we are doing....
  With txDisplay
    .Text = .Text & "Clone an enumeration after item number 5" & vbCrLf
  End With

  ' get an enumerator object from the collection
  Set oEnumer = m_oCol.[_NewEnum]
    
  ' loop through the collection.  see the form_load event for further details
  Do
    oEnumer.Next 1, vVarRet, VarPtr(nFetched)
  
    If nFetched Then
      With txDisplay
        .Text = .Text & vVarRet.Name & vbCrLf
      End With
      
      ' if this is item 5, clone the enumeration.  we pass in an enumerator object
      ' and the Clone method fills it for us.  we can then pick up from this exact
      ' point in the enumeration whenever we want to
      If i = 5 Then oEnumer.Clone oCloneEnumer
      
      i = i + 1
    End If
    
    Set vVarRet = Nothing
    
  Loop Until nFetched = 0
  
  Set oEnumer = Nothing


  With txDisplay
    .Text = .Text & "Now we enumerate the clone" & vbCrLf
  End With


  ' this time, we are enumerating the clone tht we created above.  the enumeration will
  ' begin with item number 6 and iterate through the rest of the items
  Do
    oCloneEnumer.Next 1, vVarRet, VarPtr(nFetched)
  
    If nFetched Then
      With txDisplay
        .Text = .Text & vVarRet.Name & vbCrLf
      End With
    End If
    
    Set vVarRet = Nothing
    
  Loop Until nFetched = 0
  
  Set oCloneEnumer = Nothing


  ' scroll to the bottom of the textbox
  With txDisplay
    .Text = .Text & vbCrLf
    .SelStart = Len(.Text)
  End With

End Sub

Private Sub Form_Initialize()
  ' instanciate our collection
  Set m_oCol = New Collection
End Sub

Private Sub Form_Terminate()
  ' destroy our collection
  Set m_oCol = Nothing
End Sub

