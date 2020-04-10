VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTKitCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get TipoCarga(ByVal objCT As Object) As Object
     Set TipoCarga = objCT.objUserControl.Controls("TipoCarga")
End Property

Public Property Get GrupoPesagem(ByVal objCT As Object) As Object
     Set GrupoPesagem = objCT.objUserControl.Controls("GrupoPesagem")
End Property

Public Function KT_Extrai_ProdutoKit(ByVal objCT As Object, ByVal objProdutoKit As ClassProdutoKit) As Long

Dim objProdutoKitCro As New ClassProdutoKitCro

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    objProdutoKitCro.iTipoCarga = Codigo_Extrai(TipoCarga(objCT).Text)
    '#########################################
    
    '#########################################
    'Inserido por Wagner - CROMATON 17/07/07
    objProdutoKitCro.sGrupoPesagem = GrupoPesagem(objCT).Text
    
    Set objProdutoKit.objInfoUsu = objProdutoKitCro

    KT_Extrai_ProdutoKit = SUCESSO

End Function

Public Function KT_Preenche_ProdutoKit(ByVal objCT As Object, ByVal objProdutoKit As ClassProdutoKit) As Long

Dim objProdutoKitCro As ClassProdutoKitCro

    '#########################################
    'Inserido por Wagner - CROMATON 17/07/07
    GrupoPesagem(objCT).ListIndex = -1
    If Not (objProdutoKit.objInfoUsu Is Nothing) Then
        
        Set objProdutoKitCro = objProdutoKit.objInfoUsu
        
        '#########################################
        'Inserido por Wagner - CROMATON 03/11/04
        If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
            If objProdutoKitCro.iTipoCarga = 0 Then
                TipoCarga(objCT).Text = "1"
            Else
                TipoCarga(objCT).Text = objProdutoKitCro.iTipoCarga
            End If
        Else
            TipoCarga(objCT).Text = "3"
        End If
        Call TipoCarga_Validate(objCT, bSGECancelDummy)
        '#########################################
    
        Call CF("SCombo_Seleciona2", GrupoPesagem(objCT), objProdutoKitCro.sGrupoPesagem)
    
    End If
    '#########################################

    KT_Preenche_ProdutoKit = SUCESSO
    
End Function
    
Public Function KT_CampoSel_Limpa(ByVal objCT As Object) As Long

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'Coloca Padrão preenchido
    TipoCarga(objCT).Text = PRODUTOKIT_PARTECARGA
    Call TipoCarga_Validate(objCT, bSGECancelDummy)
    '#########################################
    
    GrupoPesagem(objCT).ListIndex = -1
    
    KT_CampoSel_Limpa = SUCESSO

End Function
    
Public Function KT_Limpa_Tela_Kit(ByVal objCT As Object) As Long
    
    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'Coloca Padrão preenchido
    TipoCarga(objCT).Text = PRODUTOKIT_PARTECARGA
    Call TipoCarga_Validate(objCT, bSGECancelDummy)
    '#########################################
    
    GrupoPesagem(objCT).ListIndex = -1

    KT_Limpa_Tela_Kit = SUCESSO
    
End Function

Public Function KT_BotaoAlterar_Click(ByVal objCT As Object, ByVal objProdutoKitNovo As ClassProdutoKit, ByVal objProdutoKit As ClassProdutoKit) As Long

Dim lErro As Long

On Error GoTo Erro_KT_BotaoAlterar_Click

    '########################################
    'Inserido por Wagner
    If gobjEST.iValidaProdBase = VALIDA_PRODUTO_BASE_CARGA Then
        If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
            If objProdutoKitNovo.objInfoUsu.iTipoCarga <> PRODUTOKIT_NAOCARGA Then gError 131456
        End If
    End If
    '########################################

    KT_BotaoAlterar_Click = SUCESSO
    
    Exit Function
    
Erro_KT_BotaoAlterar_Click:

    KT_BotaoAlterar_Click = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184303)

    End Select
    
    Exit Function

End Function

Public Function KT_GravarRegistro(ByVal objCT As Object, ByVal objKit As ClassKit) As Long

Dim lErro As Long, iIndice As Integer
'#########################################
'Inserido por Wagner - CROMATON 03/11/04
Dim objProdutoKit As ClassProdutoKit
Dim bAchou As Boolean
Dim vbMsgRes As VbMsgBoxResult
'#########################################

On Error GoTo Erro_KT_GravarRegistro

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    iIndice = 0
    If gobjEST.iValidaProdBase = VALIDA_PRODUTO_BASE_CARGA Then
        For Each objProdutoKit In objKit.colComponentes
            If objProdutoKit.objInfoUsu.iTipoCarga = PRODUTOKIT_BASECARGA And objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
                iIndice = iIndice + 1
            End If
        Next
        If iIndice = 0 Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SEM_BASEPARACARGA", objKit.sProdutoRaiz, objKit.sVersao)
            If vbMsgRes = vbNo Then gError 131069
        End If
        If iIndice > 1 Then gError 131455
    End If
    '#########################################

    KT_GravarRegistro = SUCESSO
    
    Exit Function
    
Erro_KT_GravarRegistro:

    KT_GravarRegistro = gErr

    Select Case gErr

        Case 131455
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_BASE_CARGA_REPETIDO", gErr)
        
        Case 131069
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184304)

    End Select
    
    Exit Function

End Function

Public Function KT_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_KT_Form_Load

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    lErro = CarregaCombo_TipoCarga(TipoCarga(objCT))
    If lErro <> SUCESSO Then gError 131067
    
    'Coloca Padrão preenchido
    TipoCarga(objCT).Text = PRODUTOKIT_PARTECARGA
    Call TipoCarga_Validate(objCT, bSGECancelDummy)
    '#########################################

    KT_Form_Load = SUCESSO
    
    Exit Function
    
Erro_KT_Form_Load:

    KT_Form_Load = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184305)

    End Select
    
    Exit Function

End Function

Public Function KT_Limpa_Componente_Kit(ByVal objCT As Object) As Long
    
    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'Coloca Padrão preenchido
    TipoCarga(objCT).Text = PRODUTOKIT_PARTECARGA
    Call TipoCarga_Validate(objCT, bSGECancelDummy)
    '#########################################

    KT_Limpa_Componente_Kit = SUCESSO
    
End Function

'#########################################
'Inserido por Wagner - CROMATON 03/11/04
Private Function CarregaCombo_TipoCarga(objCombo As ComboBox) As Long

On Error GoTo Erro_CarregaCombo_TipoCarga

    objCombo.AddItem (PRODUTOKIT_PARTECARGA & SEPARADOR & PRODUTOKIT_STRING_PARTECARGA)
    objCombo.ItemData(objCombo.NewIndex) = PRODUTOKIT_PARTECARGA
    
    objCombo.AddItem (PRODUTOKIT_BASECARGA & SEPARADOR & PRODUTOKIT_STRING_BASECARGA)
    objCombo.ItemData(objCombo.NewIndex) = PRODUTOKIT_BASECARGA
    
    objCombo.AddItem (PRODUTOKIT_NAOCARGA & SEPARADOR & PRODUTOKIT_STRING_NAOCARGA)
    objCombo.ItemData(objCombo.NewIndex) = PRODUTOKIT_NAOCARGA

    CarregaCombo_TipoCarga = SUCESSO

    Exit Function
    
Erro_CarregaCombo_TipoCarga:

    CarregaCombo_TipoCarga = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184306)
    
    End Select
    
    Exit Function

End Function

Public Sub TipoCarga_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TipoCarga_Validate

    If Len(Trim(TipoCarga(objCT).Text)) = 0 Then Exit Sub

    If TipoCarga(objCT).Text = TipoCarga(objCT).List(TipoCarga(objCT).ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(TipoCarga(objCT), iCodigo)
    If lErro <> SUCESSO Then gError 131068

    Exit Sub

Erro_TipoCarga_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131068
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCARGA_NAO_CADASTRADO", gErr, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184307)

    End Select

    Exit Sub

End Sub

Public Sub TipoCarga_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub
'#########################################

Public Function KT_Cria_Raiz(ByVal objCT As Object, ByVal objProduto As ClassProduto, ByVal objProdutoKit As ClassProdutoKit) As Long

Dim objProdutoKitCro As New ClassProdutoKitCro

    objProdutoKitCro.iTipoCarga = PRODUTOKIT_NAOCARGA
    objProdutoKitCro.sGrupoPesagem = ""
    
    Set objProdutoKit.objInfoUsu = objProdutoKitCro
    
    KT_Cria_Raiz = SUCESSO
    
End Function

