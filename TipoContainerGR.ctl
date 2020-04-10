VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl TipoContainer 
   Appearance      =   0  'Flat
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   ScaleHeight     =   4185
   ScaleWidth      =   6225
   Begin VB.TextBox TextDescricao 
      Height          =   345
      Left            =   1140
      MaxLength       =   100
      TabIndex        =   11
      Top             =   750
      Width           =   4980
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   315
      Left            =   1680
      Picture         =   "TipoContainerGR.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Numeração Automática"
      Top             =   255
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "TipoContainerGR.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1080
         Picture         =   "TipoContainerGR.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "TipoContainerGR.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoContainerGR.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.ListBox ListTipos 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   2460
      Width           =   6000
   End
   Begin MSMask.MaskEdBox MaskCodigo 
      Height          =   315
      Left            =   1140
      TabIndex        =   12
      Top             =   255
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskValor 
      Height          =   345
      Left            =   1140
      TabIndex        =   13
      Top             =   1275
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskISO 
      Height          =   315
      Left            =   1140
      TabIndex        =   15
      Top             =   1800
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ISO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   1
      Left            =   645
      TabIndex        =   14
      Top             =   1875
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   5
      Left            =   555
      TabIndex        =   5
      Top             =   1335
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   405
      TabIndex        =   4
      Top             =   285
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   825
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   2205
      Width           =   480
   End
End
Attribute VB_Name = "TipoContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'variaveis globais
Dim iAlterado As Integer

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 97045
    
    Unload Me
    
    Exit Sub

Erro_BotaoFechar_Click:
    
    Select Case gErr

        Case 97045

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a funcao que ira efetuar a gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 97063

    'limpa a tela apos a gravacao
    Call Limpa_Tela(Me)
    
    'faz com que o registro seja marcado como nao alterado apos a gravacao
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 97063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objTipoContainer As ClassTipoContainer) As Long
'Move os dados da tela para a memoria...

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
    
    'copiando o codigo da tela para a memoria..
    objTipoContainer.iTipo = StrParaInt(MaskCodigo.ClipText)
    objTipoContainer.dValor = StrParaDbl(MaskValor.ClipText)
    objTipoContainer.dISO = StrParaDbl(MaskISO.ClipText)
    objTipoContainer.sDescricao = TextDescricao.Text
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoContainer As New ClassTipoContainer

On Error GoTo Erro_Gravar_Registro

     'Coloca o MouseIcon de Ampulheta durante a gravação
     GL_objMDIForm.MousePointer = vbHourglass

     If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 97046

     If Len(Trim(TextDescricao.Text)) = 0 Then gError 97047
     
     If Len(Trim(MaskValor.ClipText)) = 0 Then gError 97048
     
     lErro = Move_Tela_Memoria(objTipoContainer)
     If lErro <> SUCESSO Then gError 97049

     lErro = Trata_Alteracao(objTipoContainer, objTipoContainer.iTipo)
     If lErro <> SUCESSO Then gError 97050

     lErro = CF("TipoContainer_Grava", objTipoContainer)
     If lErro <> SUCESSO Then gError 97051

     'fechando comando de setas
     Call ComandoSeta_Fechar(Me.Name)

     'atualizando a list
     Call ListTipos_Exclui(objTipoContainer.iTipo)
     Call ListTipos_Adicionar(objTipoContainer)

     'Coloca o MouseIcon de setinha
     GL_objMDIForm.MousePointer = vbDefault

     Gravar_Registro = SUCESSO
     
     Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 97046
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 97047
             Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 97048
             Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
        
        Case 97049 To 97051

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'Coloca o MouseIcon de setinha
    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 97052

    'limpando a tela
    Call Limpa_Tela(Me)
        
    'fechando comando de setas
    Call ComandoSeta_Fechar(Me.Name)
        
    iAlterado = 0
       
    Exit Sub
    
Erro_Botaolimpar_Click:

    Select Case gErr
    
        Case 97052
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCod As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = TipoContainer_Codigo_Automatico(iCod)
    If lErro <> SUCESSO Then gError 97053

    MaskCodigo.PromptInclude = False
    MaskCodigo.Text = iCod
    MaskCodigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 97053

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function TipoContainer_Codigo_Automatico(iCod As Integer) As Long
'funcao que gera o codigo automatico
'terminado -> falta incluir no fatconfig

Dim lErro As Long

On Error GoTo Erro_TipoContainer_Codigo_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_Obter_Inteiro_Automatico", "FatConfig", "NUM_PROX_TIPOCONTAINER", "TipoContainer", "Tipo", iCod)
    If lErro <> SUCESSO Then gError 97054

    TipoContainer_Codigo_Automatico = SUCESSO

    Exit Function

Erro_TipoContainer_Codigo_Automatico:

    TipoContainer_Codigo_Automatico = gErr

    Select Case gErr

        Case 97054
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objTipoContainer As New ClassTipoContainer
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    objTipoContainer.iTipo = colCampoValor.Item("Tipo").vValor

    If objTipoContainer.iTipo <> 0 Then

        'Carrega o obj com dados de colcampovalor
        objTipoContainer.sDescricao = colCampoValor.Item("Descricao").vValor
        objTipoContainer.dValor = colCampoValor.Item("Valor").vValor

        'Traz para tela os dados carregados...
        lErro = Traz_TipoContainer_Tela(objTipoContainer)
        If lErro <> SUCESSO And lErro <> 97093 Then gError 97055
        
        If lErro = 97093 Then gError 92819
        
        iAlterado = 0

    End If
    
    Exit Sub
    
Erro_Tela_Preenche:

    Select Case gErr
    
        Case 92819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCONTAINTER_NAO_CADASTRADO", gErr, objTipoContainer.iTipo)
    
        Case 97055
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD
'terminado

Dim objTipoContainer As New ClassTipoContainer

    'Informa tabela associada à Tela
     sTabela = "TipoContainer"

    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Trim(MaskCodigo.ClipText)) > 0 Then
        objTipoContainer.iTipo = StrParaInt(MaskCodigo.ClipText)
    Else
        objTipoContainer.iTipo = 0
    End If

    objTipoContainer.sDescricao = TextDescricao.Text
    objTipoContainer.dISO = StrParaDbl(MaskISO.Text)
    
    If Len(Trim(MaskValor.ClipText)) > 0 Then
        objTipoContainer.dValor = StrParaDbl(MaskValor.ClipText)
    Else
        objTipoContainer.dValor = 0
    End If
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Tipo", objTipoContainer.iTipo, 0, "Tipo"
    colCampoValor.Add "Descricao", objTipoContainer.sDescricao, STRING_TIPOCONTAINER_DESCRICAO, "Descricao"
    colCampoValor.Add "Valor", objTipoContainer.dValor, 0, "Valor"
    colCampoValor.Add "ISO", objTipoContainer.dISO, 0, "ISO"

End Sub

Private Sub ListTipos_Adicionar(objTipoContainer As ClassTipoContainer)

Dim iInd As Integer
    
    For iInd = 0 To ListTipos.ListCount - 1
        If ListTipos.ItemData(iInd) > objTipoContainer.iTipo Then
            Exit For
        End If
    Next
    
   'Concatena o código com a descrição e adiciona na list...
    ListTipos.AddItem objTipoContainer.iTipo & SEPARADOR & objTipoContainer.sDescricao, iInd
    ListTipos.ItemData(ListTipos.NewIndex) = objTipoContainer.iTipo

End Sub

Private Sub ListTipos_Exclui(iCod As Integer)

Dim iIndice As Integer

    For iIndice = 0 To ListTipos.ListCount - 1

        If ListTipos.ItemData(iIndice) = iCod Then

            ListTipos.RemoveItem (iIndice)
            Exit For

        End If

    Next

End Sub

Private Sub BotaoExcluir_Click()
'terminado
'cadastrar EXCLUSAO_TIPOCONTAINER
'          ERRO_TIPOCONTAINER_NAO_CADASTRADO

Dim lErro As Long
Dim objTipoContainer As New ClassTipoContainer
Dim vbMsgRet As VbMsgBoxResult
Dim iCod As Integer

On Error GoTo Erro_BotaoExcluir_Click

    'coloca o cursor no formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Codigo do Tipo de Container foi informado
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 97056
    
    objTipoContainer.iTipo = StrParaInt(MaskCodigo.ClipText)

    'Verifica se o Tipo com o codigo em questao existe
    lErro = CF("TipoContainer_Le", objTipoContainer)
    If lErro <> SUCESSO And lErro <> 97086 Then gError 97057

    'tipo não está cadastrado
    If lErro = 97086 Then gError 97058

    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TIPOCONTAINER", objTipoContainer.iTipo)

    If vbMsgRet = vbYes Then

        'exclui o tipo de container
        lErro = CF("TipoContainer_Exclui", objTipoContainer)
        If lErro <> SUCESSO Then gError 97059

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        'Exclui o tipo de container da ListBox
        Call ListTipos_Exclui(objTipoContainer.iTipo)

        'Limpa a Tela
        Call Limpa_Tela(Me)

        iAlterado = 0

    End If

    'coloca o cursor com formato de setinha
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    'coloca o cursor com formato de setinha
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 97056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 97057, 97059

        Case 97058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCONTAINER_NAO_CADASTRADO", gErr, objTipoContainer.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colTiposContainer As New Collection
Dim objTipoContainer As ClassTipoContainer

On Error GoTo Erro_Form_Load

    'Preenche a ListBox com os tipos de container existentes no BD
    lErro = CF("TipoContainer_Le_Todos", colTiposContainer)
    If lErro <> SUCESSO And lErro <> 97091 Then gError 97095

    For Each objTipoContainer In colTiposContainer

        ListTipos.AddItem objTipoContainer.iTipo & SEPARADOR & objTipoContainer.sDescricao
        ListTipos.ItemData(ListTipos.NewIndex) = objTipoContainer.iTipo

    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 97095

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTipoContainer As ClassTipoContainer) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um tipo de container selecionado, exibir seus dados
    If Not (objTipoContainer Is Nothing) Then

        'Verifica se Tipo de Container existe
        lErro = Traz_TipoContainer_Tela(objTipoContainer)
        If lErro <> SUCESSO And lErro <> 97093 Then gError 97096

        If lErro <> SUCESSO Then
            'tipo de container não está cadastrado
            
            'Limpa a Tela
            Call Limpa_Tela(Me)

            MaskCodigo.PromptInclude = False
            MaskCodigo.Text = objTipoContainer.iTipo
            MaskCodigo.PromptInclude = True
                   
       End If
          
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 97096

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function Traz_TipoContainer_Tela(objTipoContainer As ClassTipoContainer) As Long
    
Dim lErro As Long

On Error GoTo Erro_Traz_TipoContainer_Tela
    
    Call Limpa_Tela(Me)
    
    lErro = CF("TipoContainer_Le", objTipoContainer)
    If lErro <> SUCESSO And lErro <> 97086 Then gError 97092
    
    'se nao achou
    If lErro = 97086 Then gError 97093
    
    'Preenchendo o codigo
    MaskCodigo.PromptInclude = False
    MaskCodigo.Text = objTipoContainer.iTipo
    MaskCodigo.PromptInclude = True
    
    'Preenchendo o ISO
    If objTipoContainer.dISO <> 0 Then
        MaskISO.PromptInclude = False
        MaskISO.Text = objTipoContainer.dISO
        MaskISO.PromptInclude = True
    End If
    
    'Preenchendo a descricao
    TextDescricao.Text = objTipoContainer.sDescricao
        
    'Preenchendo o valor
    MaskValor.PromptInclude = False
    MaskValor.Text = objTipoContainer.dValor
    MaskValor.PromptInclude = True

    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    Traz_TipoContainer_Tela = SUCESSO
            
    Exit Function
    
Erro_Traz_TipoContainer_Tela:

    Traz_TipoContainer_Tela = gErr
    
    Select Case gErr
    
        Case 97092, 97093
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub


Private Sub ListTipos_DblClick()

Dim lErro As Long
Dim objTipoContainer As New ClassTipoContainer

On Error GoTo Erro_ListTipos_DblClick
    
    'Pega no ItemData o código do Tipo de Container selecionado
    objTipoContainer.iTipo = ListTipos.ItemData(ListTipos.ListIndex)

    'Traz p/ a tela os dados do tipo
    lErro = Traz_TipoContainer_Tela(objTipoContainer)
    If lErro <> SUCESSO And lErro <> 97093 Then gError 97060

    'nao encontrou o tipo...
    'significa q ele esta na lista, mas nao esta no bd.. logo, deve ser
    'removido da list para manter a consistencia list x bd
    If lErro = 97093 Then
        Call ListTipos_Exclui(objTipoContainer.iTipo)
        gError 92818
    End If
    
    iAlterado = 0
    
    Exit Sub

Erro_ListTipos_DblClick:

    Select Case gErr

        Case 92818
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCONTAINER_NAO_CADASTRADO", gErr, objTipoContainer.iTipo)

        Case 97060
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ListTipos_KeyPress(KeyAscii As Integer)

    If ListTipos.ListIndex <> -1 Then
        If KeyAscii = ENTER_KEY Then
            Call ListTipos_DblClick
        End If
    End If

End Sub

Private Sub MaskCodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskCodigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodigo, iAlterado)

End Sub

Private Sub MaskCodigo_Validate(Cancel As Boolean)
'terminado

Dim lErro As Long

On Error GoTo Erro_MaskCodigo_Validade

    If Len(Trim(MaskCodigo.ClipText)) > 0 Then
        lErro = Inteiro_Critica(MaskCodigo.ClipText)
        If lErro <> SUCESSO Then gError 97062
        
    End If
    
    Cancel = False
    
    Exit Sub
    
Erro_MaskCodigo_Validade:

    Cancel = True
    
    Select Case gErr
    
        Case 97062
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
  
    End Select
    
    Exit Sub

End Sub



Private Sub MaskISO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskISO_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskISO, iAlterado)

End Sub

Private Sub MaskISO_Validate(Cancel As Boolean)
'terminado

Dim lErro As Long

On Error GoTo Erro_MaskISO_Validade

    If Len(Trim(MaskISO.ClipText)) > 0 Then
        lErro = Inteiro_Critica(MaskISO.ClipText)
        If lErro <> SUCESSO Then gError 99343
        
    End If
    
    Cancel = False
    
    Exit Sub
    
Erro_MaskISO_Validade:

    Cancel = True
    
    Select Case gErr
    
        Case 99343
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
  
    End Select
    
    Exit Sub

End Sub

Private Sub MaskValor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskValor_Validate

    'Verifica se há um valor digitado
    If Len(Trim(MaskValor.ClipText)) > 0 Then
    
        'Critica o valor digitado
        lErro = Valor_Positivo_Critica(MaskValor.ClipText)
        If lErro <> SUCESSO Then gError 97098
     
    End If

    Cancel = False

    Exit Sub

Erro_MaskValor_Validate:

    Cancel = True

    Select Case gErr

        Case 97098

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub TextDescricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskValor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskValor_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskValor, iAlterado)

End Sub


''**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_HISTORICO_PADRAO
    Set Form_Load_Ocx = Me
    Caption = "Tipo de Container"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TipoContainer"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
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

'***** fim do trecho a ser copiado ******
