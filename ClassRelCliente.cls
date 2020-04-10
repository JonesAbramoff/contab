VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelCliente"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Criada para atender Clientes

Private lComando As Long
Private objUltimoFilCliFilEmp As New ClassFilCliFilEmp

Public Function Obter_PrimeiraCompra_Cliente(dtPrimeiraCompra As Date, lCliente As Long, iFilialCliente As Integer) As Long
'Le a Data da Primeira Compra do Cliente

Dim lErro As Long

On Error GoTo Erro_Obter_PrimeiraCompra_Cliente

    If lCliente <> objUltimoFilCliFilEmp.lCodCliente Or iFilialCliente <> objUltimoFilCliFilEmp.iCodFilial Then
        
        objUltimoFilCliFilEmp.lCodCliente = lCliente
        objUltimoFilCliFilEmp.iFilialEmpresa = giFilialEmpresa
        objUltimoFilCliFilEmp.iCodFilial = iFilialCliente
        
        lErro = CF("CliFilEmp_Le_Rel",objUltimoFilCliFilEmp, lComando)
        If lErro <> SUCESSO Then Error 60458
    
    End If
    
    dtPrimeiraCompra = objUltimoFilCliFilEmp.dtDataPrimeiraCompra
    
    Obter_PrimeiraCompra_Cliente = SUCESSO
    
    Exit Function
    
Erro_Obter_PrimeiraCompra_Cliente:
        
    Obter_PrimeiraCompra_Cliente = Err
    
    Select Case Err
        
        Case 60458 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153053)
    
    End Select

    Exit Function
    
End Function

Public Function Obter_UltimaCompra_Cliente(dtUltimaCompra As Date, lCliente As Long, iFilialCliente As Integer) As Long
'Le a ultima data de Compra do Cliente

Dim lErro As Long

On Error GoTo Erro_Obter_UltimaCompra_Cliente

    If lCliente <> objUltimoFilCliFilEmp.lCodCliente Or iFilialCliente <> objUltimoFilCliFilEmp.iCodFilial Then
        
        objUltimoFilCliFilEmp.lCodCliente = lCliente
        objUltimoFilCliFilEmp.iFilialEmpresa = giFilialEmpresa
        objUltimoFilCliFilEmp.iCodFilial = iFilialCliente
        
        lErro = CF("CliFilEmp_Le_Rel",objUltimoFilCliFilEmp, lComando)
        If lErro <> SUCESSO Then Error 60459
    
    End If
        
    dtUltimaCompra = objUltimoFilCliFilEmp.dtDataUltimaCompra
    
    Obter_UltimaCompra_Cliente = SUCESSO
    
    Exit Function
    
Erro_Obter_UltimaCompra_Cliente:
        
    Obter_UltimaCompra_Cliente = Err
    
    Select Case Err
        
        Case 60459 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153054)
    
    End Select

    Exit Function
    
End Function

Public Function Obter_NumeroCompras_Cliente(lNumeroDeCompras As Long, lCliente As Long, iFilialCliente As Integer) As Long
'Le Numero de Compras do Cliente

Dim lErro As Long

On Error GoTo Erro_Obter_NumeroCompras_Cliente

    If lCliente <> objUltimoFilCliFilEmp.lCodCliente Or iFilialCliente <> objUltimoFilCliFilEmp.iCodFilial Then
        
        objUltimoFilCliFilEmp.lCodCliente = lCliente
        objUltimoFilCliFilEmp.iFilialEmpresa = giFilialEmpresa
        objUltimoFilCliFilEmp.iCodFilial = iFilialCliente
        
        lErro = CF("CliFilEmp_Le_Rel",objUltimoFilCliFilEmp, lComando)
        If lErro <> SUCESSO Then Error 60460
    
    End If
        
    lNumeroDeCompras = objUltimoFilCliFilEmp.lNumeroCompras
    
    Obter_NumeroCompras_Cliente = SUCESSO
    
    Exit Function
    
Erro_Obter_NumeroCompras_Cliente:
        
    Obter_NumeroCompras_Cliente = Err
    
    Select Case Err
        
        Case 60460 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153055)
    
    End Select

    Exit Function
    
End Function


Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

