VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqAmexPgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsReg As String

Private mvarsEstabelecimento As String
Private mvardtDataCredito As Date
Private mvardValorCredito As Double
Private mvardValorBruto As Double
Private mvardValorComissao As Double
Private mvariBanco As Integer
Private mvarsAgencia As String
Private mvarsConta As String
Private mvariPagtoFechado As Integer
Private mvarsSeqPgto As String

Public lNumIntDoc As Long
Public lNumRegArq As Long
Public iCodConta As Integer
Public iFilEmpCCI As Integer

Public colSOC As New Collection

Public Property Let sReg(ByVal vData As String)
Dim sToken As String, iPOS As Integer, iTokenNum As Integer, iTokenIni As Integer
Dim sRegAux As String

    mvarsReg = vData
    
    iTokenNum = 0
    iTokenIni = 0
    iPOS = 1
    sRegAux = Trim(mvarsReg) & ","
    Do While iPOS <= Len(sRegAux)
    
        If Mid(sRegAux, iPOS, 1) = "," Then
        
            If ((iTokenIni + 1) <= (iPOS - 1)) Then
                sToken = Mid(sRegAux, iTokenIni + 1, iPOS - iTokenIni - 1)
            Else
                sToken = ""
            End If
            
            Select Case iTokenNum
            
                Case 0
                    mvarsEstabelecimento = sToken
                
                Case 1
                    mvardtDataCredito = AAAAMMDD_ParaDate(sToken)
                    
                Case 2
                    mvarsSeqPgto = sToken
                    
                Case 7
                    mvardValorCredito = StrParaDbl(sToken)
                
                Case 8
                    mvariBanco = StrParaInt(sToken)
                
                Case 9
                    mvarsAgencia = CStr(StrParaLong(sToken))
                
                Case 10 'a conta vem com o DV
                    mvarsConta = CStr(StrParaLong(Mid(sToken, 1, Len(sToken) - 1)))
                    
                Case 14
                    mvardValorBruto = StrParaDbl(sToken)
                
                Case 15
                    mvardValorComissao = StrParaDbl(sToken)
                
                Case 19
                    If sToken = "P" Then
                        mvariPagtoFechado = 1
                    Else
                        mvariPagtoFechado = 0
                    End If
            
            End Select
            
            iTokenNum = iTokenNum + 1
            iTokenIni = iPOS
            
        End If
        
        iPOS = iPOS + 1
    
    Loop
    
End Property

Public Property Get sReg() As String
    sReg = mvarsReg
End Property

Public Property Get sEstabelecimento() As String
    sEstabelecimento = mvarsEstabelecimento
End Property

Public Property Get dtDataCredito() As Date
    dtDataCredito = mvardtDataCredito
End Property

Public Property Get dValorCredito() As Double
    dValorCredito = mvardValorCredito
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Get iPagtoFechado() As Integer
    iPagtoFechado = mvariPagtoFechado
End Property

Public Property Get dValorBruto() As Double
    dValorBruto = mvardValorBruto
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
End Property

Public Property Get sSeqPgto() As String
    sSeqPgto = mvarsSeqPgto
End Property


