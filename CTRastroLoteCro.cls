VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTRastroLoteCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private gobjCT As Object

Public Property Get Localizacao(ByVal objCT As Object) As Object
     Set Localizacao = objCT.objUserControl.Controls("Localizacao")
End Property

Public Property Get Fornecedor(ByVal objCT As Object) As Object
     Set Fornecedor = objCT.objUserControl.Controls("Fornecedor")
End Property

Public Property Get Filial(ByVal objCT As Object) As Object
     Set Filial = objCT.objUserControl.Controls("Filial")
End Property

Function RL_Move_Tela_Memoria(ByVal objCT As Object, objRastroLote As ClassRastreamentoLote) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objRastroLoteCro As New ClassRastreamentoLoteCro

On Error GoTo Erro_RL_Move_Tela_Memoria

    objRastroLoteCro.sLocalizacao = Localizacao(objCT).Text
    
    '############################################
    'Inserido por Wagner 23/02/2007
    'Se o Fornecedor estiver preenchido lê o código no BD
    If Len(Trim(Fornecedor(objCT).ClipText)) > 0 Then

        Set objFornecedor = New ClassFornecedor

        objFornecedor.sNomeReduzido = Fornecedor(objCT).Text
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 185907
        If lErro = 6681 Then gError 185908 'Se nao encontrar --> erro

        objRastroLoteCro.lFornecedor = objFornecedor.lCodigo

    End If

    'Recolhe os demais campos da Tela
    objRastroLoteCro.iFilialForn = Codigo_Extrai(Filial(objCT).Text)
    
    Set objRastroLote.objInfoUsu = objRastroLoteCro
    '############################################

    RL_Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_RL_Move_Tela_Memoria:

    RL_Move_Tela_Memoria = gErr
    
    Select Case gErr

        Case 75737, 185907
        
        Case 185908
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function RL_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long

    Set objEventoFornecedor = Nothing
    Set gobjCT = Nothing
    
    RL_Form_Unload = SUCESSO

End Function

Public Function RL_Traz_RastroLote_Tela(ByVal objCT As Object, objRastroLote As ClassRastreamentoLote) As Long
'Traz dados do RastreamentoLote para a tela

Dim lErro As Long
Dim objRastroLoteCro As ClassRastreamentoLoteCro
Dim objFornecedor As New ClassFornecedor
Dim iFilialForn As Integer

On Error GoTo Erro_RL_Traz_RastroLote_Tela

    lErro = CF("RastreamentoLote_Le_Customizado", objRastroLote)
    If lErro <> SUCESSO Then gError 185909

    Set objRastroLoteCro = objRastroLote.objInfoUsu
    Localizacao(objCT).Text = objRastroLoteCro.sLocalizacao
    
    lErro = CF("RelOPCarga_Le_Fornecedor_Lote", objRastroLote.sProduto, objRastroLote.sCodigo, objFornecedor, 0, iFilialForn)
    If lErro <> SUCESSO Then gError 185909
    
    If objFornecedor.lCodigo <> 0 Then
    
        Fornecedor(objCT).Text = objFornecedor.lCodigo
        Call Fornecedor_Validate(objCT, bSGECancelDummy)
        
        Filial(objCT).Text = iFilialForn
        Call Filial_Validate(objCT, bSGECancelDummy)
    
    End If

    RL_Traz_RastroLote_Tela = SUCESSO

    Exit Function

Erro_RL_Traz_RastroLote_Tela:

    RL_Traz_RastroLote_Tela = gErr
    
    Select Case gErr

        Case 75722, 185909

        Case 75723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objRastroLote.iFilialOP)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function RL_Limpa_Tela_RastroLote(ByVal objCT As Object) As Long

    Filial(objCT).Clear

    RL_Limpa_Tela_RastroLote = SUCESSO
    
End Function

Public Function RL_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer) As Long

    If objCT.ActiveControl Is Fornecedor(objCT) Then
        Call FornecedorLabel_Click(objCT)
    End If

    RL_UserControl_KeyDown = SUCESSO
    
End Function

Public Sub Localizacao_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

'##########################################################
'Inserido por Wagner 23/02/2007
Public Sub Fornecedor_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Fornecedor_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Se alterou o fornecedor
    If objCT.gobjInfoUsu.sFornAnt <> UCase(Fornecedor(objCT).Text) Then
    
        'Verifica se fornecedor esta preenchido
        If Len(Trim(Fornecedor(objCT).Text)) > 0 Then
    
            'Tenta ler o Fornecedor
            lErro = TP_Fornecedor_Le(Fornecedor(objCT), objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 185912
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO And lErro <> 6698 Then gError 185913
    
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial(objCT), colCodigoNome)
    
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial(objCT), iCodFilial)
            
        Else
    
            Filial(objCT).Clear
    
        End If
        
        objCT.gobjInfoUsu.sFornAnt = UCase(Fornecedor(objCT).Text)
        
    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 185912, 185913

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub FornecedorLabel_Click(ByVal objCT As Object)

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    Set objEventoFornecedor = New AdmEvento
    Set gobjCT = objCT
    
    'Preenche objFornecedor com NomeReduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor(objCT).Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Public Sub Filial_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Filial_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Se a filial nao estiver preenchida => sai da rotina
    If Len(Trim(Filial(objCT).Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial(objCT).ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial(objCT), iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 185916

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor(objCT).Text)) = 0 Then gError 185917

        sFornecedor = Fornecedor(objCT).Text
        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 185918

        If lErro = 18272 Then gError 185919

        'coloca na tela
        Filial(objCT).Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 185920

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 185916, 185918

        Case 185919
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor(objCT).Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 185917
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 185920
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial(objCT).Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor(gobjCT).Text = objFornecedor.sNomeReduzido

    'Dispara Validate de Fornecedor
    Call Fornecedor_Validate(gobjCT, bSGECancelDummy)

    gobjCT.Show

    Exit Sub

End Sub
'##################################################################



