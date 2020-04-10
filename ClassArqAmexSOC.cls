VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqAmexSOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsReg As String

Private mvarsEstabelecimento As String
Private mvardtDataPagto As Date
Private mvardtDataVenda As Date
Private mvarsNumSOC As String
Private mvardValorLiqSemAntec As Double
Private mvardtDataPagto24hs As Date
Private mvardValorLiqComAntec As Double

Public lNumIntDoc As Long
Public lNumRegArq As Long
Public colROC As New Collection

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
                    
                Case 8
                    mvarsNumSOC = sToken
                
                Case 14
                    mvardValorLiqSemAntec = StrParaDbl(sToken)
                
                Case 21
                    mvardtDataPagto24hs = AAAAMMDD_ParaDate(sToken)
                
                Case 24
                    mvardValorLiqComAntec = StrParaDbl(sToken)
                
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

Public Property Get dValorLiqSemAntec() As Double
    dValorLiqSemAntec = mvardValorLiqSemAntec
End Property

Public Property Get sNumSOC() As String
    sNumSOC = mvarsNumSOC
End Property

Public Property Get dValorLiqComAntec() As Double
    dValorLiqComAntec = mvardValorLiqComAntec
End Property

Public Property Get dtDataPagto24hs() As Date
    dtDataPagto24hs = mvardtDataPagto24hs
End Property


