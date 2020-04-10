VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'classe auxiliar a relatorios que percisem de dados do contas a receber

Private mvarbCalcSldDev As Boolean 'se calculo de saldo de devolucoes/creditos já foi executado
Private iFilEmpUltDev As Integer
Private mvardSldDev As Double
Private mvardQtdDev As Double

Private mvarbCalcSldAdiant As Boolean 'se calculo de saldo de adiantamentos já foi executado
Private iFilEmpUltAdiant As Integer
Private mvardSldAdiant As Double
Private mvardQtdAdiant As Double

Public Function ObterSldDevCR(dSldDev As Double, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo em valor de devolucoes/creditos em aberto no contas a receber

Dim lErro As Long, dSaldo As Double, lQtde As Long

On Error GoTo Erro_ObterSldDevCR

    If mvarbCalcSldDev = False Or iFilEmpUltDev <> iFilialEmpresa Then
    
        lErro = CF("DebitosRec_ObtemSaldo",dSaldo, lQtde, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 59367
        
        mvarbCalcSldDev = True
        iFilEmpUltDev = iFilialEmpresa
        mvardSldDev = dSaldo
        mvardQtdDev = lQtde
        
    End If
    
    dSldDev = mvardSldDev
    
    ObterSldDevCR = SUCESSO
     
    Exit Function
    
Erro_ObterSldDevCR:

    ObterSldDevCR = Err
     
    Select Case Err
          
        Case 59367
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153065)
     
    End Select
     
    Exit Function
    
End Function

Public Function ObterQtdDevCR(dQtdDev As Double, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde de devolucoes/creditos em aberto no contas a receber

Dim lErro As Long, dSaldo As Double, lQtde As Long

On Error GoTo Erro_ObterQtdDevCR

    If mvarbCalcSldDev = False Or iFilEmpUltDev <> iFilialEmpresa Then
    
        lErro = CF("DebitosRec_ObtemSaldo",dSaldo, lQtde, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 59368
        
        mvarbCalcSldDev = True
        iFilEmpUltDev = iFilialEmpresa
        mvardSldDev = dSaldo
        mvardQtdDev = lQtde
    
    End If
    
    dQtdDev = mvardQtdDev
        
    ObterQtdDevCR = SUCESSO
     
    Exit Function
    
Erro_ObterQtdDevCR:

    ObterQtdDevCR = Err
     
    Select Case Err
          
        Case 59368
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153066)
     
    End Select
     
    Exit Function
    
End Function

Public Function ObterSldAdiantCR(dSldAdiant As Double, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo em valor de adiantamentos em aberto no contas a receber

Dim lErro As Long, dSaldo As Double, lQtde As Long

On Error GoTo Erro_ObterSldAdiantCR

    If mvarbCalcSldAdiant = False Or iFilEmpUltAdiant <> iFilialEmpresa Then
    
        lErro = CF("AntecipRec_ObtemSaldo",dSaldo, lQtde, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 59369
        
        mvarbCalcSldAdiant = True
        iFilEmpUltAdiant = iFilialEmpresa
        mvardSldAdiant = dSaldo
        mvardQtdAdiant = lQtde
        
    End If
    
    dSldAdiant = mvardSldAdiant
    
    ObterSldAdiantCR = SUCESSO
     
    Exit Function
    
Erro_ObterSldAdiantCR:

    ObterSldAdiantCR = Err
     
    Select Case Err
          
        Case 59369
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153067)
     
    End Select
     
    Exit Function
    
End Function

Public Function ObterQtdAdiantCR(dQtdAdiant As Double, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde de adiantamentos em aberto no contas a receber

Dim lErro As Long, dSaldo As Double, lQtde As Long

On Error GoTo Erro_ObterQtdAdiantCR

    If mvarbCalcSldAdiant = False Or iFilEmpUltAdiant <> iFilialEmpresa Then
    
        lErro = CF("AntecipRec_ObtemSaldo",dSaldo, lQtde, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 59370
        
        mvarbCalcSldAdiant = True
        iFilEmpUltAdiant = iFilialEmpresa
        mvardSldAdiant = dSaldo
        mvardQtdAdiant = lQtde
        
    End If
    
    dQtdAdiant = mvardQtdAdiant
    
    ObterQtdAdiantCR = SUCESSO
     
    Exit Function
    
Erro_ObterQtdAdiantCR:

    ObterQtdAdiantCR = Err
     
    Select Case Err
          
        Case 59370
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153068)
     
    End Select
     
    Exit Function
    
End Function

Private Sub Class_Initialize()
    
    mvarbCalcSldDev = False
    mvarbCalcSldAdiant = False
    
End Sub
