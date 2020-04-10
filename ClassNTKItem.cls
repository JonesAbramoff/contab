VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNTKItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public name As String
Public qtd As Double
Public price_cents As Long
Public subtotal_cents As Long
Public tipo  As String 'type: Item, Complement
Public complements As New Collection ' de ClassNTKItem
Public obs  As String
Public discount_percentage As Double
Public discounted_price_cents As Long
Public codigo As String
