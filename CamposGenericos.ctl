VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CamposGenericosOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.CommandButton CorrigeBD 
      Caption         =   "Corrige BD"
      Height          =   375
      Left            =   315
      TabIndex        =   19
      Top             =   135
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame FramePrincipal 
      Caption         =   "Principal"
      Height          =   5415
      Left            =   120
      TabIndex        =   5
      Top             =   555
      Width           =   9315
      Begin MSMask.MaskEdBox Complemento1 
         Height          =   315
         Left            =   1305
         TabIndex        =   20
         Top             =   2655
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.CheckBox CodigoAutomatico 
         Caption         =   "Gera Código Automático"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   555
         Width           =   2415
      End
      Begin MSMask.MaskEdBox CodValor 
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.CheckBox Padrao 
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   16
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Complemento5 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   13
         Top             =   4200
         Width           =   1800
      End
      Begin VB.TextBox Complemento4 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   4200
         Width           =   1800
      End
      Begin VB.TextBox Complemento3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   3720
         Width           =   1800
      End
      Begin VB.TextBox Complemento2 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   3120
         Width           =   1800
      End
      Begin VB.TextBox Valor 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   9
         Top             =   2040
         Width           =   3255
      End
      Begin MSFlexGridLib.MSFlexGrid GridValores 
         Height          =   3015
         Left            =   360
         TabIndex        =   8
         Top             =   795
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin VB.ComboBox Campo 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   8130
      End
      Begin VB.Label Comentarios 
         BorderStyle     =   1  'Fixed Single
         Height          =   1200
         Left            =   360
         TabIndex        =   15
         Top             =   4110
         Width           =   8760
      End
      Begin VB.Label LabelComentarios 
         AutoSize        =   -1  'True
         Caption         =   "Comentários:"
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
         Left            =   360
         TabIndex        =   14
         Top             =   3855
         Width           =   1110
      End
      Begin VB.Label LabelCampo 
         AutoSize        =   -1  'True
         Caption         =   "Campo:"
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
         Left            =   360
         TabIndex        =   6
         Top             =   255
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7335
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CamposGenericos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CamposGenericos.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CamposGenericos.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CamposGenericos.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "CamposGenericosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'SUGESTÕES PARA IMPLEMENTAÇÕES FUTURAS:
'Criar uma tela onde o usuário possa cadastrar os campos para os quais ele deseja
'cadastrar valores. Nessa tela o usuário poderia indicar o tipo de dados que o campo receberá
'e essa tela deveria validar os campos conforme o cadastro feito pelo usuário.
'Ao cadastrar os campos, o usuário poderia informar se ele tem complementos, quantos tem e quais os títulos
'de cada complemento. A carga dessa tela deveria ser feita conforme esse cadastro.
'Todas as telas deveriam ser alteradas para que, ao serem carregadas, chamem uma função
'genérica que verifica se a tela tem campos com valores cadastrados a partir da tela CamposGenericos
'Caso tenha, a função deve carregar os campos conforme o cadastro feito

'************** VARIAVEIS GLOBAIS A TELA ************
Const NUM_MAXIMO_VALORES = 600 'Constante para inicialização do grid
Dim iAlterado As Integer 'Controla se houve alguma alteração na tela
Dim gcolCampos As Collection 'Guarda uma coleção com os campos que são carregados na combo
Dim glCampoAtual As Long ' Guarda o código do campo selecionado atualmente. Usado no tratamento de Campo_Click
'****************************************************

'****** GRIDVALORES ********************
'Obj do grid
Dim objGridValores As AdmGrid

'Grid Valores
Public iGrid_Padrao_Col As Integer
Public iGrid_CodValor_Col As Integer
Public iGrid_Valor_Col As Integer
Public iGrid_Complemento1_Col As Integer
Public iGrid_Complemento2_Col As Integer
Public iGrid_Complemento3_Col As Integer
Public iGrid_Complemento4_Col As Integer
Public iGrid_Complemento5_Col As Integer
'****************************************

'************** CARREGAMENTO DA TELA************
Public Function Trata_Parametros(Optional objCamposGenericos As ClassCamposGenericos, Optional sNovoValor As String) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há um item selecionado, exibir seus dados
    If Not (objCamposGenericos Is Nothing) Then

        'Se não foi passado o código do campo em questão => erro
        If objCamposGenericos.lCodigo = 0 Then gError 102287
        
        'Para cada campo na combo
        For iIndice = 0 To Campo.ListCount
        
            'Se o código do campo for igual ao código recebido como parâemtro
            If Campo.ItemData(iIndice) = objCamposGenericos.lCodigo Then
            
                'Seleciona o campo
                Campo.ListIndex = iIndice
                
                'Sai do For
                Exit For
            
            End If
        
        Next
        
        'Não encontrou o campo
        If iIndice > Campo.ListCount Then gError 102289
        
        'Faz o obj recebido como parâmetro apontar para o obj global correspondente ao campo em questão
        Set objCamposGenericos = gcolCampos(Campo.ListIndex + 1)
                
'??? remover
'        'Lê os dados do campo em questão
'        lErro = CF("CamposGenericos_Le", objCamposGenericos)
'        If lErro <> SUCESSO And lErro <> 102295 Then gError 102291
'
'        'Se não encontrou o campo => erro
'        If lErro = 102295 Then Error 102289
'
'        'Lê os valores que estão cadastrados para o campo em questão
'        lErro = CF("CamposGenericosValores_Le_CodCampo", objCamposGenericos)
'        If lErro <> SUCESSO And lErro <> 102300 Then gError 102288
'
'        'Preenche a tela com os valores para o campo em questão
'        lErro = Traz_CamposGenericos_Tela(objCamposGenericos)
'        If lErro <> SUCESSO Then gError 102290
'??? fim

        'O iAlterado tem que ser posicionado aqui, pois caso o usuário tenha passado algum valor
        'esse valor será adicionado ao grid e iAlterado será setado
        iAlterado = 0
        
        'Se foi passado um novo valor para se adicionar ao grid
        If Len(Trim(sNovoValor)) > 0 Then
        
            'Se o valor passado é numérico
            If IsNumeric(sNovoValor) Then
                
                'Preenche o código do novo valor a ser criado
                GridValores.TextMatrix(objGridValores.iLinhasExistentes + 1, iGrid_CodValor_Col) = StrParaLong(sNovoValor)
                
                'Verifica se não exista outra linha do grid com o mesmo código
                lErro = Verifica_Codigo_Repetido(objGridValores.iLinhasExistentes + 1)
                If lErro <> SUCESSO Then gError 102407
                
            'Senão
            Else
                
                'Preenche o conteúdo do novo valor a ser criado
                GridValores.TextMatrix(objGridValores.iLinhasExistentes + 1, iGrid_Valor_Col) = sNovoValor
                
                'Verifica se não exista outra linha do grid com o mesmo código
                lErro = Verifica_Valor_Repetido(objGridValores.iLinhasExistentes + 1)
                If lErro <> SUCESSO Then gError 102408
                
                'Gera o próximo código automático para o valor a ser criado
                GridValores.TextMatrix(objGridValores.iLinhasExistentes + 1, iGrid_CodValor_Col) = objCamposGenericos.lProxCodValor
                
                'Verifica se não exista outra linha do grid com o mesmo código
                lErro = Verifica_Codigo_Repetido(objGridValores.iLinhasExistentes + 1)
                If lErro <> SUCESSO Then gError 102409
            
            End If
            
            'Atualiza o próximo código a ser utilizado
            objCamposGenericos.lProxCodValor = objCamposGenericos.lProxCodValor + 1
            
            'Aumenta o número de linhas do grid
            objGridValores.iLinhasExistentes = objGridValores.iLinhasExistentes + 1
            
            iAlterado = REGISTRO_ALTERADO
        
        End If

    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 102288, 102290, 102291
        
        Case 102407 To 102409
        
            'Limpa o código do novo valor que seria criado
            GridValores.TextMatrix(objGridValores.iLinhasExistentes + 1, iGrid_CodValor_Col) = ""
        
            'Limpa o conteúdo do novo valor que seria criado
            GridValores.TextMatrix(objGridValores.iLinhasExistentes + 1, iGrid_Valor_Col) = ""
        
        Case 102287
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOGENERICO_TRATAPARAMETRO_SEM_CODIGO", gErr, Error)
        
        Case 102289
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOGENERICO_NAO_ENCONTRADO", gErr, objCamposGenericos.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144113)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Instancia a coleção que guardará os campos da combo
    Set gcolCampos = New Collection
    
    'Carrega a combo Campo com os campos que podem ter valores cadastrados nessa tela...
    lErro = Carrega_Campos()
    If lErro <> SUCESSO Then gError 102302
    
    'Instancia o objeto do grid
    Set objGridValores = New AdmGrid
    
    'Inicializa o grid valores
    lErro = Inicializa_GridValores(objGridValores)
    If lErro <> SUCESSO Then gError 102303

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 102302, 102303

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144114)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_Campos() As Long
'Lê os campos genéricos que podem ter valores cadastrados nessa tela
'Guarda o resultado da leitura em uma coleção global

Dim objCamposGenericos As ClassCamposGenericos
Dim lErro As Long

On Error GoTo Erro_Carrega_Campos

    'Lê o código e a descrição de todos os campos cadastrados na tabela CamposGenericos
    lErro = CF("CamposGenericos_Le_Todos", gcolCampos)
    If lErro <> SUCESSO And lErro <> 102308 Then gError 102310
    
    'Se não encontrou campos => erro
    If lErro = 102308 Then gError 102309

    'Para cada campo lido
    For Each objCamposGenericos In gcolCampos

        'Adiciona o código e a descrição do campo na combo
        Campo.AddItem CInt(objCamposGenericos.lCodigo) & SEPARADOR & objCamposGenericos.sDescricao
        Campo.ItemData(Campo.NewIndex) = objCamposGenericos.lCodigo

    Next
    
    Carrega_Campos = SUCESSO
    
    Exit Function
    
Erro_Carrega_Campos:

    Carrega_Campos = gErr
    
    Select Case gErr
    
        Case 102310
        
        Case 102309
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOSGENERICOS_VAZIA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144115)
            
    End Select

End Function

Private Function Inicializa_GridValores(objGridInt As AdmGrid) As Long
'Inicializa o Grid GridValores

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Padrão")
    objGridInt.colColuna.Add ("Código")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Complemento 1")
    objGridInt.colColuna.Add ("Complemento 2")
    objGridInt.colColuna.Add ("Complemento 3")
    objGridInt.colColuna.Add ("Complemento 4")
    objGridInt.colColuna.Add ("Complemento 5")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Padrao.Name)
    objGridInt.colCampo.Add (CodValor.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (Complemento1.Name)
    objGridInt.colCampo.Add (Complemento2.Name)
    objGridInt.colCampo.Add (Complemento3.Name)
    objGridInt.colCampo.Add (Complemento4.Name)
    objGridInt.colCampo.Add (Complemento5.Name)

    'Colunas do Grid
    iGrid_Padrao_Col = 1
    iGrid_CodValor_Col = 2
    iGrid_Valor_Col = 3
    iGrid_Complemento1_Col = 4
    iGrid_Complemento2_Col = 5
    iGrid_Complemento3_Col = 6
    iGrid_Complemento4_Col = 7
    iGrid_Complemento5_Col = 8

    'Grid do GridInterno
    objGridInt.objGrid = GridValores

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_VALORES

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridValores.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Indica a existência da rotina grid enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridValores = SUCESSO

    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set gcolCampos = Nothing
End Sub
'*************** FIM DO CARREGAMENTO DA TELA *************

'*************** EVENTOS DA TELA **********************
Private Sub CorrigeBD_Click()

Dim lErro As Long

On Error GoTo Erro_CorrigeBD_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("CamposGenericos_CorrigeBD_MarcaEspecie", CAMPOSGENERICOS_VOLUMEESPECIE)
    If lErro <> SUCESSO Then gError 102384
    
    lErro = CF("CamposGenericos_CorrigeBD_MarcaEspecie", CAMPOSGENERICOS_VOLUMEMARCA)
    If lErro <> SUCESSO Then gError 102385
    
    Call Rotina_Aviso(vbOKOnly, "Correção executada com sucesso.")
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_CorrigeBD_Click:
    
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 102384, 102385
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144116)
    
    End Select
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Executa a gravação do registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 102325

    'Limpa a tela
    Call Limpa_Tela_CamposGenericos

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 102325

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144117)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCamposGenericos As ClassCamposGenericos
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o campo não foi selecionado => erro
    If Campo.ListIndex = -1 Then gError 102383
    
    'Instancia objCamposGenericos apontando para o obj global correspondente ao campo selecionado
    Set objCamposGenericos = gcolCampos(Campo.ListIndex + 1)
    
    'Verifica se existem valores para esse campo
    lErro = CF("CamposGenericosValores_Le_CodCampo", objCamposGenericos)
    If lErro <> SUCESSO And lErro <> 102300 Then gError 102384
    
    'Se não encontrou valores para o campo => erro
    If lErro = 102300 Then gError 102385
    
    'Pede a confirmação da exclusão dos campos
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CAMPOSGENERICOSVALORES", objCamposGenericos.sDescricao)

    'Se o usuário confirmou a exclusão
    If vbMsgRes = vbYes Then

        'exclui os valores para o campo selecionado
        lErro = CF("CamposGenericosValores_Exclui", objCamposGenericos)
        If lErro <> SUCESSO Then gError 102386
        
        'Limpa tela
        Call Limpa_Tela_CamposGenericos

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 102383
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOGENERICO_NAO_SELECIONADO", gErr)
            
        Case 102384, 102386
        
        Case 102385
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOSGENERICOSVALORES_NAOENCONTRADO", gErr, objCamposGenericos.sDescricao)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144118)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 102326

    'Limpa a tela
    Call Limpa_Tela_CamposGenericos
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 102326

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144119)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Campo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Campo_Click()

Dim objCamposGenericos As New ClassCamposGenericos
Dim lErro As Long

On Error GoTo Erro_Campo_Click
    
    'Se não há nenhum campo selecionado => sai da função
    If Campo.ListIndex = -1 Then Exit Sub
    
    'Se o campo que estava selecionado é igual ao novo campo clicado => sai da função
    If glCampoAtual = Campo.ItemData(Campo.ListIndex) Then Exit Sub
    
    'Atualiza o código do campo atual
    glCampoAtual = Campo.ItemData(Campo.ListIndex)
    
    'Verifica se o usuário deseja salva as alteraçãoes feitas para o Campo atual
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 102379
    
    'Limpa o grid valores
    Call Grid_Limpa(objGridValores)
    
    'Habilita o campo código
    CodValor.Enabled = True
    
    'Guarda o código do campo selecionado
    objCamposGenericos.lCodigo = Campo.ItemData(Campo.ListIndex)
    
    'Exibe o comentário relativo ao campo
    Comentarios.Caption = gcolCampos(Campo.ListIndex + 1).sComentarios
    
    'Lê os valores cadastrados para o campo em questão
    lErro = CF("CamposGenericosValores_Le_CodCampo", objCamposGenericos)
    If lErro <> SUCESSO And lErro <> 102300 Then gError 102324
    
    If objCamposGenericos.lCodigo = CAMPOSGENERICOS_PROD_FABR Then
        GridValores.TextMatrix(0, iGrid_Complemento1_Col) = "CNPJ"
        Complemento1.Mask = "##############"
        Complemento1.Format = "00\.000\.000\/0000-00; ; ; "
    Else
        GridValores.TextMatrix(0, iGrid_Complemento1_Col) = "Complemento 1"
        Complemento1.Mask = ""
        Complemento1.Format = ""
    End If
    
    'Exibe na tela os valores lidos no bd
    lErro = Traz_CamposGenericosValores_Tela(objCamposGenericos.colCamposGenericosValores)
    If lErro <> SUCESSO Then gError 102366
    
    iAlterado = 0
    
    Exit Sub
    
Erro_Campo_Click:

    Select Case gErr
    
        Case 102324, 102366, 102379
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144120)
            
    End Select

End Sub
'*************** FIM DOS EVENTOS DA TELA **********************

'*************** FUNCIONAMENTO DO GRIDVALORES ************

'***** EVENTOS DO GRID *******
Private Sub GridValores_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridValores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridValores, iAlterado)
    End If

End Sub

Private Sub GridValores_EnterCell()
    Call Grid_Entrada_Celula(objGridValores, iAlterado)
End Sub

Private Sub GridValores_GotFocus()
    Call Grid_Recebe_Foco(objGridValores)
End Sub

Private Sub GridValores_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridValores)
    
End Sub

Private Sub GridValores_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridValores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridValores, iAlterado)
    End If
    
End Sub

Private Sub GridValores_LeaveCell()
    Call Saida_Celula(objGridValores)
End Sub

Private Sub GridValores_RowColChange()
    Call Grid_RowColChange(objGridValores)
End Sub

Private Sub GridValores_Scroll()
    Call Grid_Scroll(objGridValores)
End Sub

Private Sub GridValores_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridValores)
End Sub
'******* FIM DOS EVENTOS DO GRID **************

'**** EVENTOS DOS CONTROLES DO GRID *********
Private Sub Padrao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridValores)
End Sub

Private Sub Padrao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridValores)
End Sub

Private Sub Padrao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridValores.objControle = Padrao
    lErro = Grid_Campo_Libera_Foco(objGridValores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CodValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodValor_GotFocus()

Dim lCodigo As Long
Dim iLinha As Integer
Dim colCodValor As New Collection

    Call Grid_Campo_Recebe_Foco(objGridValores)
    
    If CodigoAutomatico.Value = MARCADO And Len(Trim(CodValor.Text)) = 0 Then
    
        lCodigo = gcolCampos(Campo.ListIndex + 1).lProxCodValor
        
        colCodValor.Add StrParaLong(GridValores.TextMatrix(1, iGrid_CodValor_Col))
        
        For iLinha = 1 To objGridValores.iLinhasExistentes
        
            If colCodValor(iLinha) > StrParaLong(GridValores.TextMatrix(iLinha, iGrid_CodValor_Col)) Then
            
                colCodValor.Add StrParaLong(GridValores.TextMatrix(iLinha, iGrid_CodValor_Col)), , iLinha
            
            Else
            
                colCodValor.Add StrParaLong(GridValores.TextMatrix(iLinha, iGrid_CodValor_Col))
            
            End If
            
        Next
        
        For iLinha = 1 To colCodValor.Count
        
            If lCodigo = colCodValor(iLinha) Then lCodigo = lCodigo + 1
        
        Next
        
        CodValor.PromptInclude = False
        CodValor.Text = lCodigo
        CodValor.PromptInclude = True
        
        gcolCampos(Campo.ListIndex + 1).lProxCodValor = lCodigo + 1
    
    End If
    
End Sub

Private Sub CodValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridValores)
End Sub

Private Sub CodValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridValores.objControle = CodValor
    lErro = Grid_Campo_Libera_Foco(objGridValores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridValores)
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridValores)
End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridValores.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridValores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Complemento1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Complemento1_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridValores)
End Sub

Private Sub Complemento1_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridValores)
End Sub

Private Sub Complemento1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridValores.objControle = Complemento1
    lErro = Grid_Campo_Libera_Foco(objGridValores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Complemento2_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Complemento2_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridValores)
End Sub

Private Sub Complemento2_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridValores)
End Sub

Private Sub Complemento2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridValores.objControle = Complemento2
    lErro = Grid_Campo_Libera_Foco(objGridValores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Complemento3_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Complemento3_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridValores)
End Sub

Private Sub Complemento3_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridValores)
End Sub

Private Sub Complemento3_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridValores.objControle = Complemento3
    lErro = Grid_Campo_Libera_Foco(objGridValores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Complemento4_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Complemento4_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridValores)
End Sub

Private Sub Complemento4_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridValores)
End Sub

Private Sub Complemento4_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridValores.objControle = Complemento4
    lErro = Grid_Campo_Libera_Foco(objGridValores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Complemento5_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Complemento5_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridValores)
End Sub

Private Sub Complemento5_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridValores)
End Sub

Private Sub Complemento5_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridValores.objControle = Complemento5
    lErro = Grid_Campo_Libera_Foco(objGridValores)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'**** FIM DOS EVENTOS DOS CONTROLES DO GRID *********

'**** SAÍDA DE CÉLULA DO GRID E DOS CONTROLES ******
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridInt.objGrid.Col

            'Se for a coluna Padrão
            Case iGrid_Padrao_Col
                lErro = Saida_Celula_Padrao(objGridInt)
                If lErro <> SUCESSO Then gError 102311

            'Se for a coluna código
            Case iGrid_CodValor_Col
                lErro = Saida_Celula_CodValor(objGridInt)
                If lErro <> SUCESSO Then gError 102338

            'Se for a coluna Valor
            Case iGrid_Valor_Col
                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then gError 102312
            
            'Se for uma das colunas Complemento
            Case iGrid_Complemento1_Col, iGrid_Complemento2_Col, iGrid_Complemento3_Col, iGrid_Complemento4_Col, iGrid_Complemento5_Col
                lErro = Saida_Celula_Complemento(objGridInt, objGridInt.objGrid.Col)
                If lErro <> SUCESSO Then gError 102313

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 102314

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 102311 To 102313, 102338

        Case 102314
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144121)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Padrao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = Padrao

    'Faz o tratamento para item padrão para evitar que vários itens fiquem marcados
    'como padrão
    lErro = Trata_Item_Padrao()
    If lErro <> SUCESSO Then gError 102266
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 102267

    Saida_Celula_Padrao = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 102266, 102267
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144122)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CodValor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CodValor

    Set objGridInt.objControle = CodValor
    
    'Se o código foi preenchido
    If Len(Trim(CodValor.Text)) > 0 Then
        
        'Verifica se é um valor positivo
        lErro = Long_Critica(Trim(CodValor.Text))
        If lErro <> SUCESSO Then gError 102404
    
    End If
    
    'Verifica se não exista outra linha do grid com o mesmo código
    lErro = Verifica_Codigo_Repetido()
    If lErro <> SUCESSO Then gError 102336
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 102337

    'se for ultima linha do grid habilitada e o campo estiver preenchido
    If GridValores.Row - GridValores.FixedRows = objGridValores.iLinhasExistentes And Len(Trim(GridValores.TextMatrix(GridValores.Row, iGrid_CodValor_Col))) > 0 Then
        
        'inclui a proxima linha
        objGridValores.iLinhasExistentes = objGridValores.iLinhasExistentes + 1

    End If

    Saida_Celula_CodValor = SUCESSO

    Exit Function
    
Erro_Saida_Celula_CodValor:

    Saida_Celula_CodValor = gErr

    Select Case gErr

        Case 102336, 102337, 102404
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144123)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor
    
    'Verifica se não exista outra linha do grid com o mesmo valor
    lErro = Verifica_Valor_Repetido()
    If lErro <> SUCESSO Then gError 102268
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 102269

    Saida_Celula_Valor = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 102268, 102269
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144124)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Complemento(objGridInt As AdmGrid, iColuna As Integer) As Long
'Faz a crítica da célula Complemento do grid que está deixando de ser a corrente

Dim lErro As Long, sCNPJ As String

On Error GoTo Erro_Saida_Celula_Complemento

    Set objGridInt.objControle = Me.Controls("Complemento" & CStr(iColuna - 3))
    
    If Campo.ListIndex <> -1 Then
    
        If Campo.ItemData(Campo.ListIndex) = CAMPOSGENERICOS_PROD_FABR And iColuna = iGrid_Complemento1_Col Then
    
            Call Formata_String_Numero(Complemento1.Text, sCNPJ)
            
            If Len(Trim(sCNPJ)) <> 0 Then
            
                If Len(Trim(sCNPJ)) <> STRING_CGC Then gError 12318
        
                lErro = Cgc_Critica(sCNPJ)
                If lErro <> SUCESSO Then gError 12317
    
            End If
    
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 102271

    Saida_Celula_Complemento = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Complemento:

    Saida_Celula_Complemento = gErr

    Select Case gErr

        Case 102271
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 12317

        Case 12318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144125)

    End Select

    Exit Function

End Function
'**** FIM DA SAÍDA DE CÉLULA DO GRID E DOS CONTROLES ******

'*** ROTINA_GRID_ENABLE ************
Public Sub Rotina_Grid_Enable(ByVal iLinha As Integer, ByVal objControl As Object, ByVal iLocalChamada As Integer)

Dim lErro As Long
Dim iIndex As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Seleciona o controle atual
    Select Case objControl.Name

        'se for o campo código
        Case CodValor.Name
            
            'Se o campo código não estiver preenchido e houver uma campo selecionado
            If Len(Trim(GridValores.TextMatrix(GridValores.Row, iGrid_CodValor_Col))) = 0 And Campo.ListIndex <> -1 Then
            
                'habilita o controle
                objControl.Enabled = True
            
            'Senão
            Else
                
                'Desabilita o controle
                objControl.Enabled = False
            
            End If
        
        
        'Se for qualquer campo do grid diferente do campo código
        Case Padrao.Name, Valor.Name, Complemento1.Name, Complemento2.Name, Complemento3.Name, Complemento4.Name, Complemento5.Name
        
            'Se o campo código estiver preenchido
            If Len(Trim(GridValores.TextMatrix(GridValores.Row, iGrid_CodValor_Col))) > 0 Then
            
                'habilita o controle
                objControl.Enabled = True
            
            'Senão
            Else
                
                'Desabilita o controle
                objControl.Enabled = False
            
            End If
            
        
    End Select
    
    Exit Sub
    
Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144126)
            
    End Select
    
    Exit Sub

End Sub
'*****************

'*** TRATAMENTO DO CLICK DO CAMPO 'PADRAO' ******
Private Sub Padrao_Click()
    Call Trata_Item_Padrao
End Sub
'*** FIM DO TRATAMENTO DO CLICK DO CAMPO 'PADRAO' ******

'**** TRATAMENTO DO SISTEMA DE SETAS ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCamposGenericos As New ClassCamposGenericos
Dim objCampoValor As AdmCampoValor
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CamposGenericos"

    'Guarda no obj o código do campo
    objCamposGenericos.lCodigo = LCodigo_Extrai(Campo.Text)
    'objCamposGenericos.sDescricao = SCodigo_Extrai(Campo.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCamposGenericos.lCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objCamposGenericos.sDescricao, STRING_CAMPOSGENERICOS_DESCRICAO, "Descricao"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
    
        Case 102315
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144127)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD
'Esse Tela_Preenche foge do padrão, pois selecionar o campo
'na combo Campo é suficiente para que o restante da tela seja preenchido

Dim lErro As Long
Dim objCamposGenericos As New ClassCamposGenericos
Dim iIndice As Integer

On Error GoTo Erro_Tela_Preenche

    'Para cada campo na combo
    For iIndice = 0 To Campo.ListCount
    
        'Se o código do campo for igual ao código recebido como parâemtro
        If Campo.ItemData(iIndice) = colCampoValor.Item("Codigo").vValor Then
        
            'Seleciona o campo
            Campo.ListIndex = iIndice
            
            'Sai do For
            Exit For
        
        End If
    
    Next
    
'Comentado por Luiz Nogueira
'    'Guarda o código do campo em questão no obj
'    objCamposGenericos.lCodigo = colCampoValor.Item("Codigo").vValor
'
'    'Lê os valores que estão cadastrados para o campo em questão
'    lErro = CF("CamposGenericos_Le_Completo", objCamposGenericos)
'    If lErro <> SUCESSO And lErro <> 102318 And lErro <> 102320 Then gError 102316
'
'    'Se não encontrou o campo => erro
'    If lErro = 102318 Then gError 102321
'
'    'Se não encontrou valores para o campo
'    If lErro = 102320 Then gError 102322
'
'    'Preenche a tela com os valores para o campo em questão
'    lErro = Traz_CamposGenericos_Tela(objCamposGenericos)
'    If lErro <> SUCESSO Then gError 102323
        
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 102316, 102322, 102323
        
        Case 102321
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOGENERICO_NAO_ENCONTRADO", gErr, objCamposGenericos.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144128)

    End Select

    Exit Sub

End Sub
'**** FIM DO TRATAMENTO DO SISTEMA DE SETAS ****

'**** OUTRAS FUNÇÕES DE APOIO À TELA ****
Private Function Trata_Item_Padrao() As Long
'Impede que 2 ou mais itens sejam configurados como item padrão

Dim iLinha As Integer

On Error GoTo Erro_Trata_Item_Padrao

    'Para cada item do Grid
    For iLinha = 1 To objGridValores.iLinhasExistentes
    
        'Se o item estiver configurado como padrão e não for o item da linha atual
        If StrParaInt(GridValores.TextMatrix(iLinha, iGrid_Padrao_Col)) = MARCADO And iLinha <> GridValores.Row Then
        
            'desmarca a opção padrão para esse item, pois apenas um item pode ser considerado padrão
            GridValores.TextMatrix(iLinha, iGrid_Padrao_Col) = DESMARCADO
        
        End If
    
    Next
    
    'Faz um refresh no grid para atualizar as figuras de marcado / desmarcado
    Call Grid_Refresh_Checkbox(objGridValores)

    Trata_Item_Padrao = SUCESSO
    
    Exit Function

Erro_Trata_Item_Padrao:

    Trata_Item_Padrao = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144129)

    End Select

End Function

Private Function Verifica_Codigo_Repetido(Optional iLinha As Integer) As Long
'Verifica se existe 2 ou mais itens com o mesmo código preenchido

Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_Verifica_Codigo_Repetido

    'Se iLinha não foi passado como parâmetro => faz apontar para a linha atual do grid
    If iLinha = 0 Then iLinha = GridValores.Row
    
    'Se o campo no grid está preenhcido
    If StrParaLong(GridValores.TextMatrix(iLinha, iGrid_CodValor_Col)) > 0 Then
        
        'Usa o conteúdo do grid para fazer a verificação
        lCodigo = StrParaLong(GridValores.TextMatrix(iLinha, iGrid_CodValor_Col))
    
    'Senão, se o controle código está preenchido
    ElseIf StrParaLong(CodValor.ClipText) > 0 Then
        
        'Usa o conteúdo do controle para fazer a verificação
        lCodigo = StrParaLong(CodValor.ClipText)
    
    End If
    
    'Se o código não foi preenchido => sai da função
    If lCodigo = 0 Then Exit Function
    
    'Para cada item do Grid
    For iIndice = 1 To objGridValores.iLinhasExistentes
    
        'Se o item com iIndice estiver com o mesmo código do item da linha atual e iIndice não for a linha atual => erro
        If StrParaLong(GridValores.TextMatrix(iIndice, iGrid_CodValor_Col)) = lCodigo And iIndice <> iLinha Then gError 102335
    
    Next
    
    Verifica_Codigo_Repetido = SUCESSO
    
    Exit Function

Erro_Verifica_Codigo_Repetido:

    Verifica_Codigo_Repetido = gErr
    
    Select Case gErr
    
        Case 102335
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_REPETIDO", gErr, lCodigo, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144130)
    
    End Select
    
End Function


Private Function Verifica_Valor_Repetido(Optional iLinha As Integer) As Long
'Verifica se existe 2 ou mais itens com o mesmo valor preenchido

Dim iIndice As Integer
Dim sValor As String

On Error GoTo Erro_Verifica_Valor_Repetido

    'Se iLinha não foi passado como parâmetro => faz apontar para a linha atual do grid
    If iLinha = 0 Then iLinha = GridValores.Row
    
    'Se o campo no grid está preenhcido
    If Len(Trim((GridValores.TextMatrix(iLinha, iGrid_Valor_Col)))) > 0 Then
        
        'Usa o conteúdo do grid para fazer a verificação
        sValor = Trim(GridValores.TextMatrix(iLinha, iGrid_Valor_Col))
    
    'Senão, se o controle código está preenchido
    ElseIf Len(Trim(Valor.Text)) > 0 Then
        
        'Usa o conteúdo do controle para fazer a verificação
        sValor = Trim(Valor.Text)
    
    End If
    
    'Se o valor não foi preenchido => sai da função
    If Len(Trim(sValor)) = 0 Then Exit Function
    
    'Para cada item do Grid
    For iIndice = 1 To objGridValores.iLinhasExistentes
    
        'Se o item da linha iLinha estiver com o mesmo valor do item da linha atual e iLinha não for a linha atual => erro
        If UCase(Trim(GridValores.TextMatrix(iIndice, iGrid_Valor_Col))) = UCase(sValor) And iIndice <> iLinha Then gError 102270
    
    Next
    
    Verifica_Valor_Repetido = SUCESSO
    
    Exit Function

Erro_Verifica_Valor_Repetido:

    Verifica_Valor_Repetido = gErr
    
    Select Case gErr
    
        Case 102270
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_REPETIDO", gErr, Trim(Valor.Text), iLinha)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144131)
    
    End Select
    
End Function

Private Function Move_Tela_Memoria(ByVal objCamposGenericos As ClassCamposGenericos) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Se não selecionou um campo => erro
    If Campo.ListIndex = -1 Then gError 102272
    
    'Guarda no obj o código do campo
    objCamposGenericos.lCodigo = LCodigo_Extrai(Campo.Text)
    
    'Move para memória os valores informados para o campo em questão
    lErro = Move_GridValores_Memoria(objCamposGenericos.colCamposGenericosValores)
    If lErro <> SUCESSO Then gError 102329
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 102328, 102329
        
        Case 102272
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOGENERICO_NAO_SELECIONADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144132)
    
    End Select

End Function

Private Function Move_GridValores_Memoria(ByVal colCamposGenericosValores As Collection) As Long

Dim iLinha As Integer
Dim objCamposGenericosValores As ClassCamposGenericosValores

On Error GoTo Erro_Move_GridValores_Memoria

    'Se nenhum valor foi preenchido no grid
    If objGridValores.iLinhasExistentes = 0 Then gError 102327
    
    'Para cada linha do grid
    For iLinha = 1 To objGridValores.iLinhasExistentes
    
        'Instancia um novo obj
        Set objCamposGenericosValores = New ClassCamposGenericosValores
        
        'Transfere os dados da linha atual para o obj
        With objCamposGenericosValores
            
            .iPadrao = StrParaInt(GridValores.TextMatrix(iLinha, iGrid_Padrao_Col))
            .lCodValor = StrParaLong(GridValores.TextMatrix(iLinha, iGrid_CodValor_Col))
            .sValor = Trim(GridValores.TextMatrix(iLinha, iGrid_Valor_Col))
            
            'Se o valor não foi preenchido => erro
            If Len(Trim(.sValor)) = 0 Then gError 102378
            
            .sComplemento1 = Trim(GridValores.TextMatrix(iLinha, iGrid_Complemento1_Col))
            .sComplemento2 = Trim(GridValores.TextMatrix(iLinha, iGrid_Complemento2_Col))
            .sComplemento3 = Trim(GridValores.TextMatrix(iLinha, iGrid_Complemento3_Col))
            .sComplemento4 = Trim(GridValores.TextMatrix(iLinha, iGrid_Complemento4_Col))
            .sComplemento5 = Trim(GridValores.TextMatrix(iLinha, iGrid_Complemento5_Col))
        
        End With
        
        'Guarda o obj na coleção de valores do campo
        colCamposGenericosValores.Add objCamposGenericosValores
        
    Next
    
    Move_GridValores_Memoria = SUCESSO
    
    Exit Function

Erro_Move_GridValores_Memoria:

    Move_GridValores_Memoria = gErr
    
    Select Case gErr
    
        Case 102327
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)
        
        Case 102378
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO_GRID", gErr, iLinha)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144133)
    
    End Select

End Function

Private Function Traz_CamposGenericos_Tela(ByVal objCamposGenericos As ClassCamposGenericos) As Long
'Exibe na tela os dados do campo e seus valores passados em objCamposGenericos

Dim iIndice As Integer
Dim lErro As Integer

On Error GoTo Erro_Traz_CamposGenericos_Tela

    'Limpa o campo 'Campo'
    Campo.ListIndex = -1
    
    'Seleciona o campo que está sendo carregado
    For iIndice = 0 To Campo.ListCount - 1
    
        'Se o conteúdo do item data para o campo em questão é o código do campo a ser carregado
        If Campo.ItemData(iIndice) = objCamposGenericos.lCodigo Then
            'Seleciona o campo
            Campo.ListIndex = iIndice
            'Sai do For
            Exit For
        End If
    
    Next
    
    'Exibe o comentário para o campo em questão
    Comentarios.Caption = objCamposGenericos.sComentarios
    
    'Carrega os valores do campo em questão
    lErro = Traz_CamposGenericosValores_Tela(objCamposGenericos.colCamposGenericosValores)
    If lErro <> SUCESSO Then gError 102301
    
    Traz_CamposGenericos_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_CamposGenericos_Tela:

    Traz_CamposGenericos_Tela = gErr
    
    Select Case gErr
    
        Case 102301
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144134)
    
    End Select
    
End Function

Private Function Traz_CamposGenericosValores_Tela(ByVal colCamposGenericosValores As Collection) As Long
'Exibe na tela os valores do campo em questão

Dim lErro As Integer
Dim iLinha As Integer

On Error GoTo Erro_Traz_CamposGenericosValores_Tela

    'Limpa o Grid
    Call Grid_Limpa(objGridValores)

    'Para cada valor na coleção
    For iLinha = 1 To colCamposGenericosValores.Count
    
        'Exibe o valor e os complementos no grid
        GridValores.TextMatrix(iLinha, iGrid_Padrao_Col) = colCamposGenericosValores(iLinha).iPadrao
        GridValores.TextMatrix(iLinha, iGrid_CodValor_Col) = colCamposGenericosValores(iLinha).lCodValor
        GridValores.TextMatrix(iLinha, iGrid_Valor_Col) = colCamposGenericosValores(iLinha).sValor
        GridValores.TextMatrix(iLinha, iGrid_Complemento1_Col) = colCamposGenericosValores(iLinha).sComplemento1
        GridValores.TextMatrix(iLinha, iGrid_Complemento2_Col) = colCamposGenericosValores(iLinha).sComplemento2
        GridValores.TextMatrix(iLinha, iGrid_Complemento3_Col) = colCamposGenericosValores(iLinha).sComplemento3
        GridValores.TextMatrix(iLinha, iGrid_Complemento4_Col) = colCamposGenericosValores(iLinha).sComplemento4
        GridValores.TextMatrix(iLinha, iGrid_Complemento5_Col) = colCamposGenericosValores(iLinha).sComplemento5
    
    Next

    'Atualiza o número de linhas existentes
    objGridValores.iLinhasExistentes = iLinha - 1
    
    'Atualiza os desenhos das checkboxes no grid
    Call Grid_Refresh_Checkbox(objGridValores)
    
    Traz_CamposGenericosValores_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_CamposGenericosValores_Tela:

    Traz_CamposGenericosValores_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144135)
    
    End Select
    
End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCamposGenericos As ClassCamposGenericos

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Instancia objCamposGenericos apontando para o obj global correspondente ao campo selecionado
    Set objCamposGenericos = gcolCampos(Campo.ListIndex + 1)
    
    'Limpa a coleção de valores desse objeto
    Set objCamposGenericos.colCamposGenericosValores = New Collection
    
    'Move para a memória os dados a serem gravados
    lErro = Move_Tela_Memoria(objCamposGenericos)
    If lErro <> SUCESSO Then gError 102330
    
    'Verifica se já existem valores gravados para o campo e exibe alerta
    lErro = Trata_Alteracao(objCamposGenericos.colCamposGenericosValores(1), objCamposGenericos.colCamposGenericosValores(1).lCodCampo)
    If lErro <> SUCESSO Then gError 102331
    
    'Grava o campo genérico
    lErro = CF("CamposGenericos_Grava", objCamposGenericos)
    If lErro <> SUCESSO Then gError 102332

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 102330 To 102332
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144136)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_CamposGenericos()

On Error GoTo Erro_Limpa_Tela_CamposGenericos

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    'Limpa o grid
    Call Grid_Limpa(objGridValores)
    
    'Limpa o campo Campo
    Campo.ListIndex = -1
    
    'Limpa o campo comentários
    Comentarios.Caption = ""
    
    glCampoAtual = 0
    iAlterado = 0

    Exit Sub
    
Erro_Limpa_Tela_CamposGenericos:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144137)
    
    End Select

End Sub
'**** FIM DE OUTRAS FUNÇÕES DE APOIO À TELA ****

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID = '???
    Set Form_Load_Ocx = Me
    Caption = "Campos Genéricos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CamposGenericos"
    
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

Private Sub LabelCampo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCampo, Button, Shift, X, Y)
End Sub

Private Sub LabelCampo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCampo, Source, X, Y)
End Sub

Private Sub LabelComentarios_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelComentarios, Button, Shift, X, Y)
End Sub

Private Sub LabelComentarios_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelComentarios, Source, X, Y)
End Sub



