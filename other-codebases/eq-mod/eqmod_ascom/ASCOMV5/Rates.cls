VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Rate"
Attribute VB_Ext_KEY = "Member0" ,"Rate"
' -----------------------------------------------------------------------------'
' ==================
'   Rates.CLS
' ==================
'
' Implementation of the ASCOM Rates Class
'
' Written: Chris Rowland
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' ???       cr      Initial edit
' 10-Sep-03 jab     cut out "keys", changed varient to long, minor cleanup
' 04-Nov-03 cdr     Change name to Rates and use Rate instead of Range for V2
' -----------------------------------------------------------------------------'

Implements IAxisRates

'local variable to hold collection
Private mCol As Collection

Friend Function Add(Maximum As Double, Minimum As Double) As rate

    'create a new object
    Dim objNewMember As rate
    Set objNewMember = New rate
    'set the properties passed into the method
    objNewMember.Maximum = Maximum
    objNewMember.Minimum = Minimum
    mCol.Add objNewMember
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(ByVal Index As Long) As rate
Attribute Item.VB_UserMemId = 0

    'used when referencing an element in the collection
    'Index contains either the Index to the collection
    'Syntax: Set foo = x.Item(5)
    If Index <= mCol.Count Then
        Set Item = mCol.Item(1)
    End If
    
End Property

Public Property Get Count() As Long

    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
    
End Property

Friend Sub Remove(Index As Long)

    'used when removing an element from the collection
    'Index contains the Index
    'Syntax: x.Remove(5)
    mCol.Remove Index
    
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
    
End Property

Private Sub Class_Initialize()

    'creates the collection when this class is created
    Set mCol = New Collection
    
End Sub

Private Sub Class_Terminate()

    'destroys collection when this class is terminated
    Set mCol = Nothing
    
End Sub

' ============================
' IMPLEMENTATION OF IAxisRates
' ============================

' Get the number of elements in the collection.
' Example: Debug.Print x.Count

Private Property Get IAxisRates_Count() As Long
    IAxisRates_Count = mCol.Count
End Property

' Get an element in the collection
' Example: Set foo = x.Item(5)

Private Property Get IAxisRates_Item(ByVal Index As Long) As IRate
    Set IAxisRates_Item = mCol(Index)
End Property

' this property allows you to enumerate
' this collection with the For...Each syntax

Private Property Get IAxisRates_NewEnum() As stdole.IEnumVARIANT
    Set IAxisRates_NewEnum = mCol.[_NewEnum]
End Property


