VERSION 5.00
Begin VB.Form EmpresaFilial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresa e Filial"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "EmpresaFilial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ComboEmpresa 
      Height          =   315
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   3435
   End
   Begin VB.ComboBox ComboFilial 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1245
      Width           =   3435
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   660
      Picture         =   "EmpresaFilial.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1875
      Width           =   975
   End
   Begin VB.CommandButton BotaoCancel 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2055
      Picture         =   "EmpresaFilial.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1875
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   5
      Top             =   150
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Filial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   1005
      Width           =   1785
   End
End
Attribute VB_Name = "EmpresaFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Já Declaradas em PrincipalNovo
Const MENU_CTB_CAD_ASSOCCCL = 4
Const MENU_CTB_CAD_ASSOCCCLCTB = 5
Const MENU_CTB_CAD_RATEIOOFF = 10

Dim gobjTelaPrincipal As Form

Private Sub BotaoCancel_Click()

    Unload Me

End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long
Dim objFilialEmpresa As New ClassFilialEmpresa
Dim iOrdemStrCmp As Integer

On Error GoTo Erro_BotaoOk_Click

    'Verifica se Empresa foi preenchida
    If ComboEmpresa.ListIndex = -1 Then Error 43653

    'Verifica se Filial foi preenchida
    If ComboFilial.ListIndex = -1 Then Error 43654

    objFilialEmpresa.sNomeEmpresa = ComboEmpresa.Text
    objFilialEmpresa.lCodEmpresa = ComboEmpresa.ItemData(ComboEmpresa.ListIndex)
    objFilialEmpresa.sNomeFilial = ComboFilial.Text
    objFilialEmpresa.iCodFilial = ComboFilial.ItemData(ComboFilial.ListIndex)
    
    'se nao mudou empresa nem filial nao precisa fazer nada
    If objFilialEmpresa.lCodEmpresa <> glEmpresa Or objFilialEmpresa.iCodFilial <> giFilialEmpresa Then
        
        lErro = Sistema_Reseta_Modulos
        '?? trocar número de erro
        If lErro <> SUCESSO Then Error 111
        
        Set gobjTributacao = Nothing
        
        Call GL_objMDIForm.objAdmSeta.ComandoSeta_Fechar2
        With GL_objMDIForm.objAdmSeta
            Set .gobj_ST_TelaAtiva = Nothing
            gs_ST_TelaTabela = ""
            .gs_ST_TelaIndice = ""
            .gs_ST_TelaSetaClick = ""
            .gi_ST_SetaIgnoraClick = 1
        End With
        
        'Configura Empresa e Filial inclusive conexão
        lErro = Empresa_Filial_Configura(objFilialEmpresa)
        If lErro <> SUCESSO Then Error 25876
        
        Call Reset_Fest
        Call Reset_Contab
                
        'como só depende do grupo do usuario nao precisa recarregar
    '    Call Rotinas_Recarrega_Objetos
        
        PrincipalNovo.Caption = TITULO_TELA_PRINCIPAL & " - " & gsNomeEmpresa & " - " & gsNomeFilialEmpresa
    
        Set gcolModulo = New AdmColModulo
        
        'Carrega em gcolModulo módulos indicando atividade p/ FilialEmpresa
        lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
        If lErro <> SUCESSO Then Error 57974
                
        'Carrega combo com módulos ativos p/ essa Filial e com permissão (alguma tela ou rotina) p/ Usuário
        lErro = Carrega_ComboModulo()
        If lErro <> SUCESSO Then Error 57814

    End If
    
    Unload Me

    Exit Sub

Erro_BotaoOk_Click:

    Select Case Err

        Case 25876, 57814, 57974  'tratados nas rotinas chamadas
        
        Case 43653
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_PREENCHIDA", Err)

        Case 43654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159461)

    End Select

    Exit Sub

End Sub

Private Sub ComboEmpresa_Click()

Dim lErro As Long
Dim colFilialEmpresa As New Collection
Dim objUsuarioEmpresa As ClassUsuarioEmpresa
Dim lCodEmpresa As Long
Dim sCodUsuario As String

On Error GoTo Erro_ComboEmpresa_Click

    If ComboEmpresa.ListIndex = -1 Then Exit Sub

    'Limpar a ComboFilial
    ComboFilial.Clear

    'Ler o Código da Empresa
    lCodEmpresa = ComboEmpresa.ItemData(ComboEmpresa.ListIndex)

    sCodUsuario = gsUsuario

    'Carregar todas as filiais da empresa selecionada para os quais o usuário está autorizado a acessar
    lErro = FiliaisEmpresa_Le_Usuario(sCodUsuario, lCodEmpresa, colFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 50172 Then Error 43658

    'Se não houverem filiais para empresa/usuário em questão ==> erro
    If lErro = 50172 Then Error 43659

    If giTipoVersao = VERSAO_FULL Then
        For Each objUsuarioEmpresa In colFilialEmpresa
            
            If objUsuarioEmpresa.iCodFilial = EMPRESA_TODA Then
                objUsuarioEmpresa.sNomeFilial = EMPRESA_TODA_NOME
            End If
            
            ComboFilial.AddItem objUsuarioEmpresa.sNomeFilial
            ComboFilial.ItemData(ComboFilial.NewIndex) = objUsuarioEmpresa.iCodFilial
        Next
    ElseIf giTipoVersao = VERSAO_LIGHT Then
        For Each objUsuarioEmpresa In colFilialEmpresa
            If objUsuarioEmpresa.iCodFilial <> EMPRESA_TODA Then
                ComboFilial.AddItem objUsuarioEmpresa.sNomeFilial
                ComboFilial.ItemData(ComboFilial.NewIndex) = objUsuarioEmpresa.iCodFilial
            End If
        Next
    End If
    
    If ComboFilial.ListCount >= 1 Then ComboFilial.ListIndex = 0
        
    Exit Sub

Erro_ComboEmpresa_Click:

    Select Case Err

        Case 43658

        Case 43659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_SEM_FILIAIS", Err, sCodUsuario)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 159462)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim iDiferencaAltura As Integer

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        Label1.left = POSICAO_FORA_TELA
        ComboFilial.left = POSICAO_FORA_TELA
        ComboFilial.TabStop = False
        iDiferencaAltura = BotaoOk.top - Label1.top
        BotaoOk.top = Label1.top
        BotaoCancel.top = Label1.top
        EmpresaFilial.Height = EmpresaFilial.Height - iDiferencaAltura
    End If
    
    'Carrega as Empresas
    lErro = Carrega_Empresa()
    If lErro <> SUCESSO Then Error 43646

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 43646

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159463)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Empresa() As Long

Dim lErro As Long
Dim colEmpresas As New Collection
Dim objUsuarios As New ClassUsuarios
Dim objEmpresa As ClassDicEmpresa

On Error GoTo Erro_Carrega_Empresa

   'Limpa a ComboEmpresa
   ComboEmpresa.Clear

    objUsuarios.sCodUsuario = gsUsuario

   'Carregar as Empresas que o usuário está autorizado a acessar
   lErro = Empresas_Le_Usuario(objUsuarios.sCodUsuario, colEmpresas)
   If lErro <> SUCESSO And lErro <> 50183 Then Error 43657

   'Não há empresa cadastrada para o usuário
   If lErro = 50183 Then Error 43658

   For Each objEmpresa In colEmpresas
       ComboEmpresa.AddItem objEmpresa.sNome
       ComboEmpresa.ItemData(ComboEmpresa.NewIndex) = objEmpresa.lCodigo
       If glEmpresa = objEmpresa.lCodigo Then ComboEmpresa.ListIndex = ComboEmpresa.NewIndex
   Next

    Carrega_Empresa = SUCESSO

    Exit Function

Erro_Carrega_Empresa:

    Carrega_Empresa = Err

    Select Case Err

        Case 43657

        Case 43658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_SEM_EMPRESA", Err, objUsuarios.sCodUsuario)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159464)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objTelaPrincipal1 As Form) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTelaPrincipal1 Is Nothing) Then Error 43652

    Set gobjTelaPrincipal = objTelaPrincipal1

    Unload Me

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 43652

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159465)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboModulo() As Long
'Carrega combo com módulos ativos p/ essa Filial e com permissão (alguma tela ou rotina) p/ Usuário

Dim lErro As Long
Dim collCodigoNome As New AdmCollCodigoNome
Dim objlCodigoNome As AdmlCodigoNome
Dim objUsuarioModulo As New ClassUsuarioModulo

On Error GoTo Erro_Carrega_ComboModulo

    PrincipalNovo.ComboModulo.Clear
    
    'Preenche objUsuarioModulo
    objUsuarioModulo.sCodUsuario = gsUsuario
    objUsuarioModulo.lCodEmpresa = glEmpresa
    objUsuarioModulo.iCodFilial = giFilialEmpresa
    objUsuarioModulo.dtDataValidade = Date
    
    'Lê os Módulos
    lErro = CF("UsuarioModulos_Le", objUsuarioModulo, collCodigoNome)
    If lErro <> SUCESSO Then Error 43630

    For Each objlCodigoNome In collCodigoNome

       'Insere na combo de Módulos
       PrincipalNovo.ComboModulo.AddItem objlCodigoNome.sNome
       PrincipalNovo.ComboModulo.ItemData(PrincipalNovo.ComboModulo.NewIndex) = objlCodigoNome.lCodigo

    Next

    'ativa/desativa a opção de menu que acessa associacao de conta com ccl (contabil ou extra-contabil)
    Call MenuCadCTB_Contabil_ExtraContabil
    
    Carrega_ComboModulo = SUCESSO

    Exit Function

Erro_Carrega_ComboModulo:

    Carrega_ComboModulo = Err

    Select Case Err

        Case 43630

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159466)

    End Select

    Exit Function

End Function

Private Sub MenuCadCTB_Contabil_ExtraContabil()
'torna visivel/invisivel as opcoes do menu de cadastros do CTB relativo a associacao de conta x centro de custo contabil/extra-contabil

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
            PrincipalNovo.mnuCTBCad(MENU_CTB_CAD_ASSOCCCL).Visible = True
            PrincipalNovo.mnuCTBCad(MENU_CTB_CAD_ASSOCCCLCTB).Visible = False
            PrincipalNovo.mnuCTBCad(MENU_CTB_CAD_RATEIOOFF).Visible = True
        ElseIf giSetupUsoCcl = CCL_USA_CONTABIL Then
            PrincipalNovo.mnuCTBCad(MENU_CTB_CAD_ASSOCCCL).Visible = False
            PrincipalNovo.mnuCTBCad(MENU_CTB_CAD_ASSOCCCLCTB).Visible = True
            PrincipalNovo.mnuCTBCad(MENU_CTB_CAD_RATEIOOFF).Visible = True
        Else
            PrincipalNovo.mnuCTBCad(MENU_CTB_CAD_ASSOCCCL).Visible = False
            PrincipalNovo.mnuCTBCad(MENU_CTB_CAD_ASSOCCCLCTB).Visible = False
            PrincipalNovo.mnuCTBCad(MENU_CTB_CAD_RATEIOOFF).Visible = False
        End If
    
    End If
    
End Sub


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

