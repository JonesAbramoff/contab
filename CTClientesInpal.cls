VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTClientesInpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Frame10(ByVal objCT As Object) As Object
     Set Frame10 = objCT.objUserControl.Controls("Frame10")
End Property

Public Property Get JurosPadrao(ByVal objCT As Object) As Object
     Set JurosPadrao = objCT.objUserControl.Controls("JurosPadrao")
End Property

Public Property Get JurosEspec(ByVal objCT As Object) As Object
     Set JurosEspec = objCT.objUserControl.Controls("JurosEspec")
End Property

Public Property Get Juros(ByVal objCT As Object) As Object
     Set Juros = objCT.objUserControl.Controls("Juros")
End Property

Public Property Get FinPadrao(ByVal objCT As Object) As Object
     Set FinPadrao = objCT.objUserControl.Controls("FinPadrao")
End Property

Public Property Get FinEspec(ByVal objCT As Object) As Object
     Set FinEspec = objCT.objUserControl.Controls("FinEspec")
End Property

Public Property Get DespesaFinanceira(ByVal objCT As Object) As Object
     Set DespesaFinanceira = objCT.objUserControl.Controls("DespesaFinanceira")
End Property

Public Sub JurosPadrao_Click(ByVal objCT As Object)
    Juros(objCT).Enabled = False
    Juros(objCT).Text = ""
End Sub

Public Sub JurosEspec_Click(ByVal objCT As Object)
    Juros(objCT).Enabled = True
End Sub

Public Sub FinPadrao_Click(ByVal objCT As Object)
    DespesaFinanceira(objCT).Enabled = False
    DespesaFinanceira(objCT).Text = ""
End Sub

Public Sub FinEspec_Click(ByVal objCT As Object)
    DespesaFinanceira(objCT).Enabled = True
End Sub

Public Sub Juros_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Juros_Validate

    'Se os Juros tiver sido preenchido
    If Len(Trim(Juros(objCT).Text)) > 0 Then
        
        'Critica o valor
        lErro = Porcentagem_Critica(Juros(objCT).Text)
        If lErro <> SUCESSO Then gError 75492
        
    End If
    
    Exit Sub
    
Erro_Juros_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 75492
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select
    
    Exit Sub
    
End Sub

Public Sub DespesaFinanceira_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DespesaFinanceira_Validate

    'Se os DespesaFinanceira tiver sido preenchido
    If Len(Trim(DespesaFinanceira(objCT).Text)) > 0 Then
        
        'Critica o valor
        lErro = Porcentagem_Critica(DespesaFinanceira(objCT).Text)
        If lErro <> SUCESSO Then gError 75493
        
    End If
    
    Exit Sub
    
Erro_DespesaFinanceira_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 75493
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select
    
    Exit Sub
    
End Sub

Public Function CLI_Grava_Registro(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CLI_Grava_Registro

    'Se a Despesa financeira for específica, critica o preenchimento
    If FinEspec(objCT).Value = True And Len(Trim(DespesaFinanceira(objCT).Text)) = 0 Then gError 75494
    
    'Se o Juros for específico, critica o preenchimento
    If JurosEspec(objCT).Value = True And Len(Trim(Juros(objCT).Text)) = 0 Then gError 75495
    
    CLI_Grava_Registro = SUCESSO
     
    Exit Function
    
Erro_CLI_Grava_Registro:

    CLI_Grava_Registro = gErr
     
    Select Case gErr
    
        Case 75494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESPESAFINANCEIRA_NAO_PREENCHIDA", gErr)
        
        Case 75495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_JUROS_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192964)
     
    End Select
     
    Exit Function

End Function

Public Function CLI_Limpa_Tela_Cliente(ByVal objCT As Object) As Long

    JurosPadrao(objCT).Value = True
    FinPadrao(objCT).Value = True
    
    CLI_Limpa_Tela_Cliente = SUCESSO
    
End Function

Public Function CLI_Traz_Cliente_Tela(ByVal objCT As Object, ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim objClienteInpal As ClassClienteInpal

On Error GoTo Erro_CLI_Traz_Cliente_Tela

    lErro = CF("Cliente_Le_Customizado", objCliente)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192978
    
    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteInpal = objCliente.objInfoUsu
 
        If objClienteInpal.iPadraoTaxaFin = Padrao Then
            FinPadrao(objCT).Value = True
            DespesaFinanceira(objCT).Text = ""
        Else
            FinEspec(objCT).Value = True
            DespesaFinanceira(objCT).Text = objClienteInpal.dTaxaFinanceira * 100
        End If
        
        If objClienteInpal.iPadraoJuros = Padrao Then
            JurosPadrao(objCT).Value = True
            Juros(objCT).Text = ""
        Else
            JurosEspec(objCT).Value = True
            Juros(objCT).Text = objClienteInpal.dJuros * 100
        End If
        
    End If
    
    CLI_Traz_Cliente_Tela = SUCESSO
     
    Exit Function
    
Erro_CLI_Traz_Cliente_Tela:

    CLI_Traz_Cliente_Tela = gErr
     
    Select Case gErr
    
        Case 192978
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192979)
     
    End Select
     
    Exit Function

End Function

Public Function CLI_Move_Cliente_Memoria(ByVal objCT As Object, ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim objClienteInpal As New ClassClienteInpal

On Error GoTo Erro_CLI_Move_Cliente_Memoria

    If JurosPadrao(objCT).Value = True Then
        objClienteInpal.iPadraoJuros = Padrao
    Else
        objClienteInpal.iPadraoJuros = NAO_PADRAO
        objClienteInpal.dJuros = StrParaDbl(Juros(objCT).Text) / 100
    End If
    
    If FinPadrao(objCT).Value = True Then
        objClienteInpal.iPadraoTaxaFin = Padrao
    Else
        objClienteInpal.iPadraoTaxaFin = NAO_PADRAO
        objClienteInpal.dTaxaFinanceira = StrParaDbl(DespesaFinanceira(objCT).Text) / 100
    End If

    Set objCliente.objInfoUsu = objClienteInpal
    
    CLI_Move_Cliente_Memoria = SUCESSO
     
    Exit Function
    
Erro_CLI_Move_Cliente_Memoria:

    CLI_Move_Cliente_Memoria = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192977)
     
    End Select
     
    Exit Function

End Function

Public Function CLI_Form_Load(ByVal objCT As Object) As Long
    
    Select Case gsUsuario
    
        Case "Supervisor", "LIMA", "L.F.M.", "Milton"
        
        Case Else
            Frame10(objCT).Enabled = False
    
    End Select
    
    CLI_Form_Load = SUCESSO
    
End Function


