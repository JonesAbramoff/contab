VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqAmexAjuste"
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
Private mvarsNumCartao As String
Private mvarsSOC As String
Private mvardValorBruto As Double
Private mvardValorDesconto As Double

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
                    mvarsSOC = CStr(StrParaLong(sToken))
                    
                Case 8
                    mvardValorBruto = StrParaDbl(sToken)
                
                Case 9
                    mvardValorDesconto = StrParaDbl(sToken)
                
                Case 13
                    mvarsNumCartao = sToken
                    
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

Public Property Get sNumCartao() As String
    sNumCartao = mvarsNumCartao
End Property

Public Function Processa(colPagtos As Collection) As Long
Dim objPgto As ClassArqAmexPgto, objSOC As ClassArqAmexSOC, objROC As ClassArqAmexROC
Dim iItem As Integer

    For Each objPgto In colPagtos
    
        For Each objSOC In objPgto.colSOC
        
            If mvarsSOC = objSOC.sNumSOC Then
            
                For iItem = objSOC.colROC.Count To 1 Step -1
                
                    Set objROC = objSOC.colROC(iItem)
                    If objROC.sNumCartao = mvarsNumCartao And Abs(objROC.dValorTotal + mvardValorBruto) < 0.02 Then
                    
                        objSOC.colROC.Remove (iItem)
                        Exit For
                        
                    End If
                
                Next
        
            End If
        
        Next
    
    Next
    
End Function
