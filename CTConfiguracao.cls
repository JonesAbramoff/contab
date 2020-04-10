VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTConfiguracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Const CONTA_RESULTADO = "1"
Const CONTA_TRANSFERENCIA = "2"
Const CONTA_PRODUCAO = "3"

Dim iAlterado As Integer
Dim iFrameAtual As Integer

'Constantes públicas dos tabs
Private Const TAB_Inicializacao = 1
Private Const TAB_Ccl = 2
Private Const TAB_ValoresIniciais = 3
Private Const TAB_Contas = 4


Function Trata_Parametros() As Long
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
End Function

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get DocPorExercicio() As Object
     Set DocPorExercicio = objUserControl.Controls("DocPorExercicio")
End Property

Public Property Get DocPorPeriodo() As Object
     Set DocPorPeriodo = objUserControl.Controls("DocPorPeriodo")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get LotePorExercicio() As Object
     Set LotePorExercicio = objUserControl.Controls("LotePorExercicio")
End Property

Public Property Get LotePorPeriodo() As Object
     Set LotePorPeriodo = objUserControl.Controls("LotePorPeriodo")
End Property

Public Property Get ContaResultado() As Object
     Set ContaResultado = objUserControl.Controls("ContaResultado")
End Property

Public Property Get TvwContas() As Object
     Set TvwContas = objUserControl.Controls("TvwContas")
End Property

Public Property Get ContaTransferencia() As Object
     Set ContaTransferencia = objUserControl.Controls("ContaTransferencia")
End Property

Public Property Get ContaProducao() As Object
     Set ContaProducao = objUserControl.Controls("ContaProducao")
End Property

Public Property Get LabelContas() As Object
     Set LabelContas = objUserControl.Controls("LabelContas")
End Property

Public Property Get LabelContaProducao() As Object
     Set LabelContaProducao = objUserControl.Controls("LabelContaProducao")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get SemCcl() As Object
     Set SemCcl = objUserControl.Controls("SemCcl")
End Property

Public Property Get CclContabil() As Object
     Set CclContabil = objUserControl.Controls("CclContabil")
End Property

Public Property Get CclExtra() As Object
     Set CclExtra = objUserControl.Controls("CclExtra")
End Property

Public Property Get TipoConta() As Object
     Set TipoConta = objUserControl.Controls("TipoConta")
End Property

Public Property Get Natureza() As Object
     Set Natureza = objUserControl.Controls("Natureza")
End Property

Public Property Get TipoDaConta() As Object
     Set TipoDaConta = objUserControl.Controls("TipoDaConta")
End Property

Public Property Get Nat() As Object
     Set Nat = objUserControl.Controls("Nat")
End Property

Public Property Get BotaoCancela() As Object
     Set BotaoCancela = objUserControl.Controls("BotaoCancela")
End Property

Public Property Get BotaoOk() As Object
     Set BotaoOk = objUserControl.Controls("BotaoOk")
End Property

Public Property Get Opcoes() As Object
     Set Opcoes = objUserControl.Controls("Opcoes")
End Property

Public Property Get HistoricoObrigatorio() As Object
     Set HistoricoObrigatorio = objUserControl.Controls("HistoricoObrigatorio")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoCancela_Click()

    Unload Me
    
End Sub

Public Sub ContaResultado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaResultado_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaResultado_Validate

    TvwContas.Tag = CONTA_RESULTADO

    'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Conta_Critica", ContaResultado.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 5700 Then Error 44669
            
    'conta não cadastrada
    If lErro = 5700 Then Error 44670

    Exit Sub

Erro_ContaResultado_Validate:

    Cancel = True

    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err
    
        Case 44669
        
        Case 44670
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaResultado.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Chama_Tela("PlanoConta", objPlanoConta)

            End If
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155612)
    
    End Select
    
    Exit Sub

End Sub

Public Sub ContaTransferencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaTransferencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaTransferencia_Validate

    TvwContas.Tag = CONTA_TRANSFERENCIA

    'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Conta_Critica", ContaTransferencia.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 5700 Then Error 44939
            
    'conta não cadastrada
    If lErro = 5700 Then Error 44940

    Exit Sub

Erro_ContaTransferencia_Validate:

    Cancel = True

    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err
    
        Case 44939
            
        Case 44940
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaTransferencia.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Chama_Tela("PlanoConta", objPlanoConta)

            End If
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155613)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objConfiguracao As New ClassConfiguracao
Dim iIndice As Integer
Dim objCTBConfig As ClassCTBConfig
Dim sContaEnxuta As String

On Error GoTo Erro_Form_Load
    
    'Le o registro da tabela Configuracao
    lErro = CF("Configuracao_Le", objConfiguracao)
    If lErro <> SUCESSO Then Error 12160

    'marca as opcoes na parte do Lote item Inicializacao
    If objConfiguracao.iLotePorPeriodo = LOTE_INICIALIZADO_POR_PERIODO Then
        LotePorPeriodo.Value = True
    Else
        LotePorExercicio.Value = True
    End If
    
    'marca as opcoes na parte do Documento no item Inicializacao
    If objConfiguracao.iDocPorPeriodo = DOC_INICIALIZADO_POR_PERIODO Then
        DocPorPeriodo = True
    Else
        DocPorExercicio = True
    End If

    'marca a opcao referente no item Centro de Custo/Lucro
    Select Case objConfiguracao.iUsoCcl
    
        Case CCL_NAO_USA
                SemCcl.Value = True
        Case CCL_USA_CONTABIL
                CclContabil = True
        Case CCL_USA_EXTRACONTABIL
                CclExtra = True
        
    End Select
    
    'inicializar os tipos de conta
    For iIndice = 1 To gobjColTipoConta.Count
        TipoConta.AddItem gobjColTipoConta.Item(iIndice).sDescricao
    Next
    
    'inicializar as naturezas de conta
    For iIndice = 1 To gobjColNaturezaConta.Count
        Natureza.AddItem gobjColNaturezaConta.Item(iIndice).sDescricao
    Next
    
    'mostra o TipoConta que esta na tabela Configuracao
    For iIndice = 0 To gobjColTipoConta.Count - 1
        TipoConta.ListIndex = iIndice
        If TipoConta.Text = gobjColTipoConta.Descricao(objConfiguracao.iTipoContaDefault) Then Exit For
    Next

    'mostra a Natureza que esta na tabela Confuguracao
    For iIndice = 0 To gobjColNaturezaConta.Count - 1
        Natureza.ListIndex = iIndice
        If Natureza.Text = gobjColNaturezaConta.Descricao(objConfiguracao.iNaturezaDefault) Then Exit For
    Next
    
    'Inicializa as Mascaras das contas
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 20712

    Set objCTBConfig = New ClassCTBConfig
    
    objCTBConfig.sCodigo = CONTA_RESULTADO_EXERCICIO
    objCTBConfig.iFilialEmpresa = giFilialEmpresa
            
    'Preenche a conta Resultado
    lErro = Preenche_Contas(objCTBConfig, ContaResultado)
    If lErro <> SUCESSO Then Error 52795
    
    Set objCTBConfig = New ClassCTBConfig
    
    objCTBConfig.sCodigo = CONTA_TRANSFERENCIA_FILIAL
    objCTBConfig.iFilialEmpresa = giFilialEmpresa
               
    'Preenche a Conta Transferência
    lErro = Preenche_Contas(objCTBConfig, ContaTransferencia)
    If lErro <> SUCESSO Then Error 52796
            
    Set objCTBConfig = New ClassCTBConfig
    
    objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
    objCTBConfig.iFilialEmpresa = giFilialEmpresa
    
    'Preenche a Conta Produção
    lErro = Preenche_Contas(objCTBConfig, ContaProducao)
    If lErro <> SUCESSO Then Error 52797
     
    'Inicializa a Lista de Plano de Contas
    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
    If lErro <> SUCESSO Then Error 20716
    
    '##################################################
    'Inserido por Wagner 13/03/2006
    Set objCTBConfig = New ClassCTBConfig
    
    objCTBConfig.sCodigo = HISTORICO_CTB_OBRIGATORIO
    
    lErro = CF("CTBConfig_Le", objCTBConfig)
    If lErro <> SUCESSO And lErro <> 9755 Then Error 20716
    
    HistoricoObrigatorio.Value = StrParaInt(objCTBConfig.sConteudo)
    '##################################################

    
    TvwContas.Tag = CONTA_RESULTADO
    
    iAlterado = 0
    iFrameAtual = 0

    lErro_Chama_Tela = SUCESSO
            
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 12160, 20712, 20716, 52795, 52796, 52797
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155614)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub Opcoes_Click()

    If Opcoes.SelectedItem.Index - 1 <> iFrameAtual Then
    
        If TabStrip_PodeTrocarTab(iFrameAtual + 1, Opcoes, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcoes.SelectedItem.Index - 1).Visible = True
        Frame1(iFrameAtual).Visible = False
        iFrameAtual = Opcoes.SelectedItem.Index - 1
        
        Select Case iFrameAtual
        
            Case TAB_Inicializacao
                Parent.HelpContextID = IDH_CONFIGURACAO_INICIALIZACAO
                
            Case TAB_Ccl
                Parent.HelpContextID = IDH_CONFIGURACAO_CENTRO_CUSTO_LUCRO
                        
            Case TAB_ValoresIniciais
                Parent.HelpContextID = IDH_CONFIGURACAO_VALORES_INICIAIS
                
            Case TAB_Contas
                Parent.HelpContextID = IDH_CONFIGURACAO_CONTAS
                
        End Select
    
    End If

End Sub

Public Sub TipoConta_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoConta_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Natureza_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Natureza_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub BotaoOk_Click()

    Call Gravar_Registro

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objConfiguracao As New ClassConfiguracao
Dim colConfigs As New Collection
Dim objCTBConfig As ClassCTBConfig
Dim sContaResultado As String
Dim sContaTransferencia As String
Dim sContaProducao As String
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    objConfiguracao.iTipoContaDefault = gobjColTipoConta.TipoConta(TipoConta.Text)
    objConfiguracao.iNaturezaDefault = gobjColNaturezaConta.NaturezaConta(Natureza.Text)
    
    'Grava os registros na tabela Configuracao com os dados de objConfiguracao
    lErro = CF("Configuracao_Altera", objConfiguracao)
    If lErro <> SUCESSO Then Error 12161
    
    Set objCTBConfig = New ClassCTBConfig
    
    objCTBConfig.sCodigo = CONTA_RESULTADO_EXERCICIO
    objCTBConfig.iFilialEmpresa = giFilialEmpresa
    objCTBConfig.iTipo = CONFIG_TIPO_TEXTO
    
    If Len(ContaResultado.ClipText) = 0 Then
        objCTBConfig.sConteudo = ""
    Else
    
        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaResultado.Text, sContaResultado, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 44671
    
        'conta não cadastrada
        If lErro = 5700 Then Error 44672
        
         objCTBConfig.sConteudo = sContaResultado
         
    End If
    
    colConfigs.Add objCTBConfig
    
    Set objCTBConfig = New ClassCTBConfig
    
    objCTBConfig.sCodigo = CONTA_TRANSFERENCIA_FILIAL
    objCTBConfig.iFilialEmpresa = giFilialEmpresa
    objCTBConfig.iTipo = CONFIG_TIPO_TEXTO
    
    If Len(ContaTransferencia.ClipText) = 0 Then
        objCTBConfig.sConteudo = ""
    Else
    
        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaTransferencia.Text, sContaTransferencia, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 44941
    
        'conta não cadastrada
        If lErro = 5700 Then Error 44942
        
         objCTBConfig.sConteudo = sContaTransferencia
         
    End If
    
    colConfigs.Add objCTBConfig
        
    Set objCTBConfig = New ClassCTBConfig
    
    objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
    objCTBConfig.iFilialEmpresa = giFilialEmpresa
    objCTBConfig.iTipo = CONFIG_TIPO_TEXTO
    
    If Len(ContaProducao.ClipText) = 0 Then
        objCTBConfig.sConteudo = ""
    Else
    
        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaProducao.Text, sContaProducao, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 52870
    
        'conta não cadastrada
        If lErro = 5700 Then Error 52871
        
         objCTBConfig.sConteudo = sContaProducao
         
    End If
    
    colConfigs.Add objCTBConfig
    
    '##################################################
    'Inserido por Wagner 10/03/2006
    Set objCTBConfig = New ClassCTBConfig
    
    objCTBConfig.sCodigo = HISTORICO_CTB_OBRIGATORIO
    objCTBConfig.iTipo = CONFIG_TIPO_INTEIRO
    
    objCTBConfig.sConteudo = CStr(HistoricoObrigatorio.Value)
    
    colConfigs.Add objCTBConfig
    '##################################################
    
    lErro = CF("CTBConfig_Grava_Configs", colConfigs)
    If lErro <> SUCESSO Then Error 44665
    
    giHistoricoObrigatorio = HistoricoObrigatorio.Value 'Inserido por Wagner
    
    GL_objMDIForm.MousePointer = vbDefault
    
    iAlterado = 0
    
    Unload Me
        
    Exit Function
    
Erro_Gravar_Registro:
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 12161, 44665, 44671, 44941, 52870
            
        Case 44672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaResultado.Text)
            
        Case 44942
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaTransferencia.Text)
        
        Case 52871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaProducao.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155615)

    End Select

    Exit Function
    
End Function

Private Function Inicializa_Mascaras() As Long
'inicializa a mascara de conta resultado

Dim sMascaraConta As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 20715
    
    ContaResultado.Mask = sMascaraConta
    ContaTransferencia.Mask = sMascaraConta
    ContaProducao.Mask = sMascaraConta

    Inicializa_Mascaras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err
    
    Select Case Err
    
        Case 20715
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155616)
        
    End Select

    Exit Function
    
End Function

Public Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 44666
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 44666
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155617)
        
    End Select
        
    Exit Sub
    
End Sub

Public Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String
Dim sContaMascarada As String
Dim cControl As Control
Dim iLinha As Integer

On Error GoTo Erro_TvwContas_NodeClick

    sCaracterInicial = Left(Node.Key, 1)

    If sCaracterInicial <> "A" Then Error 44667
    
    sConta = Right(Node.Key, Len(Node.Key) - 1)
    
    sContaEnxuta = String(STRING_CONTA, 0)

    'volta mascarado apenas os caracteres preenchidos
    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 44668

    If TvwContas.Tag = CONTA_RESULTADO Then

        ContaResultado.PromptInclude = False
        ContaResultado.Text = sContaEnxuta
        ContaResultado.PromptInclude = True
        
    ElseIf TvwContas.Tag = CONTA_TRANSFERENCIA Then

        ContaTransferencia.PromptInclude = False
        ContaTransferencia.Text = sContaEnxuta
        ContaTransferencia.PromptInclude = True
        
    ElseIf TvwContas.Tag = CONTA_PRODUCAO Then
        
        ContaProducao.PromptInclude = False
        ContaProducao.Text = sContaEnxuta
        ContaProducao.PromptInclude = True
    
    End If

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 44667

        Case 44668
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155618)

    End Select

    Exit Sub

End Sub

Public Sub ContaProducao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaProducao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaProducao_Validate

    TvwContas.Tag = CONTA_PRODUCAO

    'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Conta_Critica", ContaProducao.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 5700 Then Error 52790
            
    'conta não cadastrada
    If lErro = 5700 Then Error 52791

    Exit Sub

Erro_ContaProducao_Validate:

    Cancel = True

    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err
    
        Case 52790
        
        Case 52791
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaProducao.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Chama_Tela("PlanoConta", objPlanoConta)

            End If
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155619)
    
    End Select
    
    Exit Sub

End Sub

Function Preenche_Contas(objCTBConfig As ClassCTBConfig, objCampoConta As Object) As Long
'Preenche a conta no campo passado como parametro

Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_Contas

    lErro = CF("CTBConfig_Le", objCTBConfig)
    If lErro <> SUCESSO And lErro <> 9755 Then Error 20713
            
    If lErro = SUCESSO And Len(objCTBConfig.sConteudo) > 0 Then
    
        sContaEnxuta = String(STRING_CONTA, 0)
    
        lErro = Mascara_RetornaContaEnxuta(objCTBConfig.sConteudo, sContaEnxuta)
        If lErro <> SUCESSO Then Error 20714
        
        objCampoConta.PromptInclude = False
        objCampoConta.Text = sContaEnxuta
        objCampoConta.PromptInclude = True
        
    End If
    
    Preenche_Contas = SUCESSO
       
    Exit Function
    
Erro_Preenche_Contas:
    
    Preenche_Contas = Err
    
    Select Case Err
        
        Case 20713
        
        Case 20714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objCTBConfig.sConteudo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155620)
    
    End Select
    
    Exit Function
          
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONFIGURACAO_INICIALIZACAO
    Set Form_Load_Ocx = Me
    Caption = "Configuração"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Configuracao"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'obs.: nao apague pois é chamado da tela (usercontrol)
Public Sub Form_Unload(Cancel As Integer)

End Sub

'***** fim do trecho a ser copiado ******

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

