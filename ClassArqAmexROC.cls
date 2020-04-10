VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqAmexROC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumRegArq As Long
Public sNumSOC As String

Private mvarsReg As String

Private mvarsEstabelecimento As String
Private mvardtDataPagto As Date
Private mvardtDataVenda As Date
Private mvarsAutorizacao As String
Private mvarsNumCartao As String
Private mvardValorTotal As Double
Private mvardValorPrimParc As Double
Private mvardValorOutrasParc As Double
Private mvariQtdeParcelas As Integer
Private mvariNumParcela As Integer
Private mvarsNSU As String
Private mvarsNSUInt As String
Private mvarsNumFatROC As String

Public Function PreencherMovDet(ByVal objMovDet As ClassAdmExtFinMovDet) As Long

    With objMovDet
        .iTipo = ADMEXTFIN_TIPOMOVDET_BXCARTAO
        Select Case mvariNumParcela
            Case 0
                .dValor = mvardValorTotal
                
            Case 1
                .dValor = mvardValorPrimParc
                            
            Case Else
                If mvardValorOutrasParc > DELTA_VALORMONETARIO Or mvariQtdeParcelas <> 2 Then
                    .dValor = mvardValorOutrasParc
                Else
                    .dValor = mvardValorTotal - mvardValorPrimParc
                End If
            
        End Select
        .sNumCartao = left(mvarsNumCartao, 15)
        .dtDataCompra = mvardtDataVenda
        .iNumParcela = mvariNumParcela
        .sAutorizacao = mvarsAutorizacao
        .sRO = sNumSOC
        .sNSU = mvarsNSU
    End With
       
End Function

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
                    mvardtDataPagto = AAAAMMDD_ParaDate(sToken)
                    
                Case 7
                    mvardtDataVenda = AAAAMMDD_ParaDate(sToken)
                    
                Case 9
                    mvarsAutorizacao = sToken
                    
                Case 10
                    mvarsNumCartao = sToken
                    
                Case 11
                    mvardValorTotal = StrParaDbl(sToken)
                
                Case 12
                    mvardValorPrimParc = StrParaDbl(sToken)
                
                Case 13
                    mvardValorOutrasParc = StrParaDbl(sToken)
                
                Case 14
                    mvariQtdeParcelas = StrParaInt(sToken)
                
                Case 15
                    mvariNumParcela = StrParaInt(sToken)
                    
                Case 18
                    mvarsNSU = sToken
                    
                Case 19
                    mvarsNSUInt = sToken
                    
                Case 20
                    mvarsNumFatROC = sToken
                
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

Public Property Get dtDataPagto() As Date
    dtDataPagto = mvardtDataPagto
End Property

Public Property Get dtDataVenda() As Date
    dtDataVenda = mvardtDataVenda
End Property

Public Property Get sAutorizacao() As String
    sAutorizacao = mvarsAutorizacao
End Property

Public Property Get sNumCartao() As String
    sNumCartao = mvarsNumCartao
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Get dValorPrimParc() As Double
    dValorPrimParc = mvardValorPrimParc
End Property

Public Property Get dValorOutrasParc() As Double
    dValorOutrasParc = mvardValorOutrasParc
End Property

Public Property Get iQtdeParcelas() As Integer
    iQtdeParcelas = mvariQtdeParcelas
End Property

Public Property Get iNumParcela() As Integer
    iNumParcela = mvariNumParcela
End Property

Public Property Get sNSU() As String
    sNSU = mvarsNSU
End Property

Public Property Get sNSUInt() As String
    sNSUInt = mvarsNSUInt
End Property

Public Property Get sNumFatROC() As String
    sNumFatROC = mvarsNumFatROC
End Property


