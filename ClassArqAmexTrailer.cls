VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqAmexTrailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mvarsReg As String

Private mvardtDataCriacao As Date
Private mvarsEstabelecimento As String
Private mvarlSeq As Long
Private mvarlQtdeRegsArq As Long

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
                
                Case 7
                    mvardtDataCriacao = AAAAMMDD_ParaDate(sToken)
                    
                Case 9
                    mvarlSeq = StrParaLong(sToken)
                    
                Case 12
                    mvarlQtdeRegsArq = StrParaLong(sToken)
                
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

Public Property Get dtDataCriacao() As Date
    dtDataCriacao = mvardtDataCriacao
End Property

Public Property Get sEstabelecimento() As String
    sEstabelecimento = mvarsEstabelecimento
End Property

Public Property Get lSeq() As Long
    lSeq = mvarlSeq
End Property

Public Property Get lQtdeRegsArq() As Long
    lQtdeRegsArq = mvarlQtdeRegsArq
End Property



