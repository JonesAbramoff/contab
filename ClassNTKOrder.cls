VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNTKOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public id As String
Public order_number As String
Public status As Integer
Public status_description As String
Public payment_type As String 'Dinheiro, Cartão, ONLINE (pré-pago no app), PAGAMENTO PARCIAL
Public cpf As String
Public troco As Long
Public total As Long
Public delivery_fee As Long
Public items_description As String
Public merchant_id As String
Public order_date As Date
Public order_time As Double
Public customer_name As String
Public customer_phone As String
Public address As ClassNTKAddress
Public items As New Collection ' de ClassNTKItem
Public payments As New Collection ' de ClassNTKPayment
