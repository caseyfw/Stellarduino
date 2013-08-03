VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' -----------------------------------------------------------------------------'
' ==================
'   Rate.CLS
' ==================
'
' Implementation of the ASCOM Rate Class
'
' Written: Chris Rowland
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' ???       cr      Initial edit
' 04-Nov-03 cdr     Change name to Rate for V2
' -----------------------------------------------------------------------------'

'local variable(s) to hold property value(s)
Private m_Maximum As Double 'local copy
Private m_Minimum As Double

Public Property Let Maximum(ByVal max As Double)

    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Maximum = 5
    m_Maximum = max
    
End Property

Public Property Get Maximum() As Double
    
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Maximum
    Maximum = m_Maximum

End Property

Public Property Let Minimum(ByVal min As Double)
    
    m_Minimum = min

End Property

Public Property Get Minimum() As Double

    Minimum = m_Minimum

End Property
