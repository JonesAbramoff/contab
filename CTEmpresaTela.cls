VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTEmpresaTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objTela As Object
Public objCTFilial As CTFilialEmpresa
Dim m_objUserControl As Object

Event Unload()

Dim iFrameAtual As Integer
Dim glCodigoAnterior As Long

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Modulos = 2

'--- inicio dos properties get dos controles da tela

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objTela.BotaoGravar
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objTela.BotaoExcluir
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objTela.BotaoLimpar
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objTela.BotaoFechar
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objTela.Frame1
End Property

Public Property Get Empresas() As Object
     Set Empresas = objTela.Empresas
End Property

Public Property Get Nome() As Object
     Set Nome = objTela.Nome
End Property

Public Property Get Codigo() As Object
     Set Codigo = objTela.Codigo
End Property

Public Property Get NomeReduzido() As Object
     Set NomeReduzido = objTela.NomeReduzido
End Property

Public Property Get BotaoMarcarTodos() As Object
     Set BotaoMarcarTodos = objTela.BotaoMarcarTodos
End Property

Public Property Get BotaoDesmarcarTodos() As Object
     Set BotaoDesmarcarTodos = objTela.BotaoDesmarcarTodos
End Property

Public Property Get Modulos() As Object
     Set Modulos = objTela.Modulos
End Property

Public Property Get EmpresaLabel() As Object
     Set EmpresaLabel = objTela.EmpresaLabel
End Property

Public Property Get Opcao() As Object
     Set Opcao = objTela.Opcao
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoDesmarcarTodos_Click()
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Error_BotaoDesmarcarTodos_Click
    
    'percorre todas as linhas da ListBox de Módulos
    For iIndice = 0 To Modulos.ListCount - 1
    
        'marca cada checkbox da ListBox
        Modulos.Selected(iIndice) = False
    
    Next
    
    Exit Sub
    
Error_BotaoDesmarcarTodos_Click:
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155700)
            
    End Select
    
    Exit Sub
End Sub
Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objEmpresa As New ClassDicEmpresa
Dim lCodigo As Long

On Error GoTo Erro_BotaoExcluir_Click
    
    'Verifica se Código da Empresa foi informado
    If Len(Trim(Codigo.Text)) = 0 Then Error 6207
    objEmpresa.lCodigo = CLng(Codigo.Text)
    
    'Verifica se a Empresa existe
    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO And lErro <> 6179 Then Error 6209 'Erro na leitura da Empresa
    If lErro = 6179 Then Error 6208 'Empresa não cadastrada
    
    'Exclui a Empresa
    lErro = Empresa_Exclui(objEmpresa.lCodigo, objEmpresa.sNome)
    If lErro <> SUCESSO Then Error 6210
    
    'Exclui a Empresa da ListBox
    Call EmpresasList_Remove(objEmpresa)
    
    'Limpa a Tela
'    If giTipoVersao = VERSAO_FULL Then
        
        lErro = Limpa_Tela_Empresa()
        If lErro <> SUCESSO Then Error 25927
    
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        lErro = objCTFilial.Limpa_Tela_FilialEmpresa()
'        If lErro <> SUCESSO Then Error 25928
'
'    End If
    
    'Carrega listBox de Modulos
    lErro = Carrega_ListModulos()
    If lErro <> SUCESSO Then Error 25510
    
    'Gera número automático
    lErro = Empresa_Automatica(lCodigo)
    If lErro Then Error 6211
    Codigo.Text = CStr(lCodigo)

Exit Sub

Erro_BotaoExcluir_Click:
    
    Select Case Err
        
        Case 6207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_EMPRESA_NAO_INFORMADO", Err)
 
        Case 6208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_CADASTRADA", Err, objEmpresa.lCodigo)
 
        Case 6209, 6210, 6211, 25510, 25927, 25928  'Tratados nas rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 155701)
    End Select
    Exit Sub
End Sub
Public Sub BotaoFechar_Click()
    Unload Me
End Sub
Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim objEmpresa As New ClassDicEmpresa
Dim lCodigo As Long
Dim iIndice As Integer
Dim objModulo As New AdmModulo
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_BotaoGravar_Click
    
    'Torna Mouse Pointer ampulheta
    BotaoGravar.MousePointer = vbHourglass
    
    'Verifica se dados da Empresa foram informados
    If Len(Codigo.Text) = 0 Then Error 6194
    If CLng(Codigo.Text) = 0 Then Error 25521
    If Len(Trim(Nome.Text)) = 0 Then Error 6195
    If Len(Trim(NomeReduzido.Text)) = 0 Then Error 25472
    
    'Verifica se outros dados obrigatórios foram informados
'    If giTipoVersao = VERSAO_LIGHT Then
'        lErro = objCTFilial.Verifica_Dados_Obrigatorios()
'        If lErro <> SUCESSO Then Error 25933
'    End If
    
    'Passa dados da Empresa para objEmpresa
    objEmpresa.lCodigo = CLng(Codigo.Text)
    objEmpresa.sNome = Trim(Nome.Text)
    objEmpresa.sNomeReduzido = Trim(NomeReduzido.Text)
    
    'Pega na listBox Modulos as siglas dos Modulos ativos
    objEmpresa.colModulos.Add MODULO_ADM
    
    For iIndice = 0 To Modulos.ListCount - 1
        If Modulos.Selected(iIndice) = True Then
            objModulo.sSigla = Trim(left(Modulos.List(iIndice), InStr(Modulos.List(iIndice), SEPARADOR) - 1))
            objEmpresa.colModulos.Add objModulo.sSigla
        End If
    Next
    
'    If giTipoVersao = VERSAO_FULL Then
    
        'Grava Empresa no banco de dados
        lErro = Empresa_Grava(objEmpresa)
        If lErro Then Error 6196
    
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        'Passa dados de Filial para objFilialEmpresa
'        lErro = objCTFilial.Traz_Dados_Memoria(objFilialEmpresa, objEmpresa)
'        If lErro <> SUCESSO Then Error 25924
'
'        'Grava Empresa (e única Filial) no banco de dados
'        lErro = Empresa_Grava(objEmpresa, objFilialEmpresa)
'        If lErro <> SUCESSO Then Error 25934
'
'    End If
    
    'Remove e insere Empresa na ListBox
    Call EmpresasList_Remove(objEmpresa)
    Call EmpresasList_Adiciona(objEmpresa)
    
    'Limpa a Tela
'    If giTipoVersao = VERSAO_FULL Then
        Call Limpa_Tela_Empresa
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        Call objCTFilial.Limpa_Tela_FilialEmpresa
'    End If
    
    'Carrega listBox de Modulos
    lErro = Carrega_ListModulos()
    If lErro <> SUCESSO Then Error 25511
    
    'Gera número automático.
    lErro = Empresa_Automatica(lCodigo)
    If lErro Then Error 6197
    Codigo.Text = CStr(lCodigo)

    BotaoGravar.MousePointer = vbDefault

Exit Sub

Erro_BotaoGravar_Click:
    
    BotaoGravar.MousePointer = vbDefault
    
    Select Case Err
        
        Case 6194, 25521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_EMPRESA_NAO_INFORMADO", Err)
        
        Case 6195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_EMPRESA_NAO_INFORMADO", Err)
        
        Case 6196, 6197, 25511, 25923, 25924, 25933, 25934  'Tratado na rotina chamada
        
        Case 25472
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMERED_EMPRESA_NAO_INFORMADO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155702)
     
     End Select
     
     Exit Sub

End Sub
Private Sub EmpresasList_Remove(objEmpresa As ClassDicEmpresa)
'Percorre a ListBox Empresas para remover a Empresa caso ela exista na List
Dim iIndice As Integer
For iIndice = 0 To Empresas.ListCount - 1
    If Empresas.ItemData(iIndice) = objEmpresa.lCodigo Then
        Empresas.RemoveItem iIndice
        Exit For
    End If
Next
End Sub
Private Sub EmpresasList_Adiciona(objEmpresa As ClassDicEmpresa)
'Inclui Empresa na List
    Empresas.AddItem objEmpresa.sNomeReduzido
    Empresas.ItemData(Empresas.NewIndex) = objEmpresa.lCodigo
End Sub
Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long
Dim colModulos As New Collection
Dim objModulo As AdmModulo

On Error GoTo Erro_BotaoLimpar_Click
    
    'Limpa a Tela
'    If giTipoVersao = VERSAO_FULL Then
    
        lErro = Limpa_Tela_Empresa()
        If lErro <> SUCESSO Then Error 25470
    
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        lErro = objCTFilial.Limpa_Tela_FilialEmpresa()
'        If lErro <> SUCESSO Then Error 25926
'
'    End If
    
    'Gera número automático.
    lErro = Empresa_Automatica(lCodigo)
    If lErro Then Error 6193
    
    'Coloca na tela
    Codigo.Text = CStr(lCodigo)
    
    'Carrega list de Modulos (com todos)
    lErro = Carrega_ListModulos()
    If lErro <> SUCESSO Then Error 25503

Exit Sub

Erro_BotaoLimpar_Click:
    
    Select Case Err
        
        Case 6193, 25470, 25503, 25926 'tratados nas rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155703)
     
     End Select
     
     Exit Sub

End Sub
Private Sub BotaoSair_Click()
End Sub
Public Sub BotaoMarcarTodos_Click()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Error_BotaoMarcarTodos_Click
    
    'percorre todas as linhas da ListBox de Módulos
    For iIndice = 0 To Modulos.ListCount - 1
    
        'marca cada checkbox da ListBox
        Modulos.Selected(iIndice) = True
    
    Next
        
    Modulos.Refresh
    
    Exit Sub
    
Error_BotaoMarcarTodos_Click:
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155704)
            
    End Select
    
    Exit Sub

End Sub
Public Sub Codigo_GotFocus()
    
    If Len(Trim(Codigo.Text)) > 0 Then
        glCodigoAnterior = CLng(Codigo.Text)
    Else
        glCodigoAnterior = 0
    End If
    Call MaskEdBox_TrataGotFocus(Codigo)

End Sub
Public Sub Codigo_Validate(Cancel As Boolean)

Dim objEmpresa As New ClassDicEmpresa
Dim lCodigo As Long
Dim colModulos As New Collection
Dim objModulo As AdmModulo
Dim lErro As Long
Dim bAchou As Boolean

On Error GoTo Erro_Codigo_Validate
    
    If Len(Trim(Codigo.Text)) > 0 Then
        'Critica se é inteiro positivo
        lErro = Inteiro_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 25531
        
        lCodigo = CLng(Codigo.Text)
    
    Else
        lCodigo = 0
    
    End If
    
    'Se mudou o código acerta os MÓDULOS para corresponderem
    If lCodigo <> glCodigoAnterior Then
        
        If lCodigo <> 0 Then
            objEmpresa.lCodigo = lCodigo
            'Tenta ler a Empresa com esse codigo no BD
            lErro = Empresa_Le(objEmpresa)
            If lErro <> SUCESSO And lErro <> 6179 Then Error 25478
            If lErro = SUCESSO Then bAchou = True Else bAchou = False
        End If
        
        If lCodigo = 0 Or bAchou = False Then
            'Carrega todos os módulos independente de Empresa
            lErro = Carrega_ListModulos()
            If lErro <> SUCESSO Then Error 25476
         
         Else 'lCodigo<>0 And bAchou=True
            'Carrega módulos associados à Empresa
            lErro = Carrega_ListModulosEmpresa(objEmpresa)
            If lErro Then Error 25477
        End If
    
    End If
    
    Exit Sub

Erro_Codigo_Validate:
    
    Cancel = True
    
    Select Case Err
        Case 25476, 25477, 25478
        Case 25531
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155705)
    End Select
    Exit Sub
End Sub
Public Sub Empresas_DblClick()

Dim lErro As Long
Dim objEmpresa As New ClassDicEmpresa

On Error GoTo Erro_Empresas_DblClick
    
    'Se não há Empresa selecionada sai da rotina
    If Empresas.ListIndex = -1 Then Exit Sub
    
    'Pega o Codigo do Item selecionado
    objEmpresa.lCodigo = Empresas.ItemData(Empresas.ListIndex)
    
    'Verifica se a Empresa existe
    lErro = Empresa_Le(objEmpresa)
    If lErro <> 6179 And lErro <> SUCESSO Then Error 6192
    If lErro = 6179 Then Error 25502
    
    'Traz Empresa para a Tela
    lErro = Traz_Empresa_Tela(objEmpresa)
    Exit Sub

Erro_Empresas_DblClick:
    
    Select Case Err
        
        Case 6192
        
        Case 25502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_CADASTRADA", Err, objEmpresa.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155706)
    End Select
    
    Exit Sub

End Sub
Public Sub Form_Unload(Cancel As Integer)

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objEmpresa As New ClassDicEmpresa
Dim lCodigo As Long
Dim colEmpresa As New Collection

On Error GoTo Erro_Empresa_Form_Load
    
    EmpresaTela.HelpContextID = IDH_EMPRESA_ID
    
    iFrameAtual = 1
    
    'Lê dados das Empresas ATIVAS no BD
    lErro = Empresas_Le_Todas(colEmpresa)
    If lErro <> SUCESSO And lErro <> 6179 Then Error 6189
    
    'Encontrou empresa(s)
    If lErro = SUCESSO Then
        
        'Preenche a ListBox com Empresas existentes no BD
        For Each objEmpresa In colEmpresa
            Empresas.AddItem objEmpresa.sNomeReduzido
            Empresas.ItemData(Empresas.NewIndex) = objEmpresa.lCodigo
        Next
    
    Else 'Não encontrou empresas
        
        lErro = Empresa_Automatica(lCodigo)
        If lErro Then Error 6191
        Codigo.Text = CStr(lCodigo)
    
    End If
    
    'Carrega List de Modulos
    lErro = Carrega_ListModulos()
    If lErro <> SUCESSO Then Error 25501
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Empresa_Form_Load:
    
    lErro_Chama_Tela = Err
    
    Select Case Err
        
        Case 6189, 6191, 25501, 25931, 25932 'tratada na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155707)
    
    End Select
      
    Exit Sub

End Sub
Function Carrega_ListModulosEmpresa(objEmpresa As ClassDicEmpresa) As Long

Dim lErro As Long
Dim colModulos As New Collection
Dim objModulo As AdmModulo

On Error GoTo Erro_Carrega_ListModulosEmpresa
    
    'Limpa a list
    Modulos.Clear
    
    'Lê módulos associados a Empresa
    lErro = Modulos_Le_Empresa1(objEmpresa.lCodigo, colModulos)
    If lErro <> SUCESSO Then Error 25469
    
    'Coloca os módulos na ListBox Modulos
    For Each objModulo In colModulos
        
        If objModulo.sSigla <> MODULO_ADM Then
            
            'Adiciona na listBox
            Modulos.AddItem objModulo.sSigla & SEPARADOR & objModulo.sNome
    
            'Define status de seleção
            If objModulo.iAtivo = 1 Then
                Modulos.Selected(Modulos.NewIndex) = True
            Else
                Modulos.Selected(Modulos.NewIndex) = False
            End If
            
        End If
    Next
    
    Carrega_ListModulosEmpresa = SUCESSO
    
    Exit Function

Erro_Carrega_ListModulosEmpresa:
    
    Carrega_ListModulosEmpresa = Err
    
    Select Case Err
        
        Case 25469  'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155708)
    
    End Select
    
    Exit Function

End Function
Function Carrega_ListModulos() As Long

Dim lErro As Long
Dim colModulos As New Collection
Dim objModulo As AdmModulo

On Error GoTo Erro_Carrega_ListModulos
    
    'Limpa a list
    Modulos.Clear
    
    'Lê todos os módulos
    lErro = Modulos_Le2(colModulos)
    If lErro <> SUCESSO Then Error 25474
    
    'Coloca os módulos (todos, independente de Empresa) na ListBox Modulos
    For Each objModulo In colModulos
        
        If objModulo.sSigla <> MODULO_ADM Then
        
            'Adiciona na listBox
            Modulos.AddItem objModulo.sSigla & SEPARADOR & objModulo.sNome
            Modulos.Selected(Modulos.NewIndex) = True
        
        End If
        
    Next
    
    Carrega_ListModulos = SUCESSO
    
    Exit Function

Erro_Carrega_ListModulos:
    
    Carrega_ListModulos = Err
    
    Select Case Err
        
        Case 25474
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155709)
    
    End Select
    
    Exit Function

End Function
Function Traz_Empresa_Tela(objEmpresa As ClassDicEmpresa) As Long
'Traz para a tela os dados correspondentes a objEmpresa

Dim lErro As Long
Dim colModulos As New Collection
Dim objModulo As AdmModulo
Dim objFilialEmpresa As New AdmFiliais
Dim iIndice As Integer

On Error GoTo Erro_Traz_Empresa_Tela
    
    
'    If giTipoVersao = VERSAO_FULL Then
    
        'Limpa a Tela
        lErro = Limpa_Tela_Empresa
        If lErro <> SUCESSO Then Error 25468
        
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        'Para poder fazer transacao no bd da empresa
'        lErro = Sistema_DefEmpresa(objEmpresa.sNome, objEmpresa.lCodigo, EMPRESA_TODA_NOME, EMPRESA_TODA)
'        If lErro <> AD_BOOL_TRUE Then Error 25920
'
'        'Limpa a Tela
'        lErro = objCTFilial.Limpa_Tela_FilialEmpresa()
'        If lErro <> SUCESSO Then Error 25918
'
'    End If
    
    'Preenche campos da Tela
    Codigo.Text = CStr(objEmpresa.lCodigo)
    Nome.Text = objEmpresa.sNome
    NomeReduzido.Text = objEmpresa.sNomeReduzido
    
'    If giTipoVersao = VERSAO_FULL Then
        
        EmpresaLabel.Caption = objEmpresa.sNomeReduzido
    
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        For iIndice = EmpresaLabel.LBound To EmpresaLabel.UBound
'            EmpresaLabel(iIndice) = objEmpresa.sNomeReduzido
'        Next
'
'    End If
    
    'Carrega list de Modulos com Modulos da Empresa
    lErro = Carrega_ListModulosEmpresa(objEmpresa)
    If lErro <> SUCESSO Then Error 25503
    Traz_Empresa_Tela = SUCESSO
    
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        'Preenche dados chave de objFilialEmpresa
'        objFilialEmpresa.iCodFilial = FILIAL_MATRIZ
'        objFilialEmpresa.sNome = MATRIZ
'        objFilialEmpresa.lCodEmpresa = objEmpresa.lCodigo
'
'        lErro = objCTFilial.Traz_Filial_Tela(objFilialEmpresa)
'        If lErro <> SUCESSO Then Error 25919
'
'    End If
    
    Exit Function

Erro_Traz_Empresa_Tela:
    
    Traz_Empresa_Tela = Err
    
    Select Case Err
        
        Case 25468, 25503, 25918, 25919, 25920
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155710)
    
    End Select
    
    Exit Function

End Function
Function Limpa_Tela_Empresa() As Long
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Empresa
    
    'Limpa a Tela
    lErro = Limpa_Tela(Me)

    EmpresaLabel.Caption = ""
    
    Limpa_Tela_Empresa = SUCESSO
    
    Exit Function

Erro_Limpa_Tela_Empresa:
    
    Limpa_Tela_Empresa = Err
    Select Case Err
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155711)
    End Select
    Exit Function
End Function
Public Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then
        If Not IniciaLetra(NomeReduzido.Text) Then Error 57826
    End If
        
'    If giTipoVersao = VERSAO_FULL Then
        
        EmpresaLabel.Caption = Trim(NomeReduzido.Text)
    
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        For iIndice = EmpresaLabel.LBound To EmpresaLabel.UBound
'            EmpresaLabel(iIndice).Caption = Trim(NomeReduzido.Text)
'        Next
'
'    End If
        
    Exit Sub
Erro_NomeReduzido_Validate:
    
    Cancel = True
    
    Select Case Err
    
        Case 57826
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155712)
    
    End Select
    
    Exit Sub
End Sub
Public Sub Opcao_Click()
    
    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then
        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me, 0) <> SUCESSO Then Exit Sub
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
    
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                EmpresaTela.HelpContextID = IDH_EMPRESA_ID
                
            Case TAB_Modulos
                EmpresaTela.HelpContextID = IDH_EMPRESA_MODULOS
                        
        End Select

    End If

End Sub
Public Property Get objTela() As Object
    Set objTela = m_objTela
End Property

Public Property Set objTela(ByVal vData As Object)
    Set m_objTela = vData
End Property

Public Property Get Controls() As Object
    Set Controls = objTela.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objTela.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objTela.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objTela.Enabled = New_Enabled
End Property

Public Sub Show()
    objTela.Show
    objTela.SetFocus
End Sub

Public Sub Unload(objMe As Object)
    RaiseEvent Unload
End Sub

Public Property Get Name() As String
    Name = objTela.Name
End Property

