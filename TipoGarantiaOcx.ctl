VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TipoGarantiaOcx 
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7545
   Begin VB.Frame Frame7 
      Caption         =   "Serviços/Peças"
      Height          =   3450
      Left            =   135
      TabIndex        =   18
      Top             =   1635
      Width           =   7260
      Begin VB.CheckBox GarantiaTotal 
         Caption         =   "Todos c/exceção dos listados abaixo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   420
         TabIndex        =   4
         Top             =   285
         Width           =   3615
      End
      Begin VB.TextBox DescricaoServico 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1890
         MaxLength       =   250
         TabIndex        =   8
         Top             =   1200
         Width           =   3000
      End
      Begin VB.CommandButton BotaoServicos 
         Caption         =   "Serviços/Peças"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5235
         TabIndex        =   10
         Top             =   3045
         Width           =   1740
      End
      Begin MSMask.MaskEdBox PrazoValidade 
         Height          =   225
         Left            =   4935
         TabIndex        =   9
         Top             =   1200
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Servico 
         Height          =   225
         Left            =   630
         TabIndex        =   7
         Top             =   1125
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridServicos 
         Height          =   1875
         Left            =   270
         TabIndex        =   6
         Top             =   705
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   3307
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox GarantiaTotalPrazo 
         Height          =   315
         Left            =   6330
         TabIndex        =   5
         Top             =   300
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Prazo (em dias):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   4875
         TabIndex        =   19
         Top             =   345
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5220
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoGarantiaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoGarantiaOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "TipoGarantiaOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoGarantiaOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   765
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   300
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PrazoPadraoValidade 
      Height          =   315
      Left            =   2235
      TabIndex        =   3
      Top             =   1215
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "Prazo Padrão(em dias):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   17
      Top             =   1260
      Width           =   1980
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
      Index           =   0
      Left            =   660
      TabIndex        =   15
      Top             =   810
      Width           =   930
   End
   Begin VB.Label LblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      Left            =   1140
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   345
      Width           =   450
   End
End
Attribute VB_Name = "TipoGarantiaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iTipoAlterado As Integer
Dim iAlterado As Integer


'Eventos de browser
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1

Dim objGridServico As AdmGrid

Dim iGrid_Servico_Col As Integer
Dim iGrid_ServicoDesc_Col As Integer
Dim iGrid_PrazoValidade_Col As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Private Function Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa eventos de browser
    Set objEventoCodigo = New AdmEvento
    Set objEventoServico = New AdmEvento
    
    Set objGridServico = New AdmGrid
    
    Call Inicializa_Grid_Servico(objGridServico)
    
    'Inicializa a Máscara de Servico
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Servico)
    If lErro <> SUCESSO Then gError 186021
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Function
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 186021
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186022)
    
    End Select
    
End Function

Private Function Inicializa_Grid_Servico(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Desc. Serviço")
    objGridInt.colColuna.Add ("Prazo Validade")

    objGridInt.colCampo.Add (Servico.Name)
    objGridInt.colCampo.Add (DescricaoServico.Name)
    objGridInt.colCampo.Add (PrazoValidade.Name)

    'Controles que participam do Grid
    iGrid_Servico_Col = 1
    iGrid_ServicoDesc_Col = 2
    iGrid_PrazoValidade_Col = 3

    objGridInt.objGrid = GridServicos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_GARANTIA_SERVICOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridServicos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Servico = SUCESSO

End Function

Public Function Trata_Parametros(Optional ByVal objTipoGarantia As ClassTipoGarantia) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se recebeu um objeto com dados de um tipo de garantia
    If Not (objTipoGarantia Is Nothing) Then
    
        'Lê e traz os dados do tipo de garantia
        lErro = Traz_TipoGarantia_Tela(objTipoGarantia)
        If lErro <> SUCESSO Then gError 186023
        
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 186023
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186024)
    
    End Select
    
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodigo = Nothing
    Set objEventoServico = Nothing

    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 186025
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 186025
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186026)

    End Select

    Exit Sub

End Sub

Private Sub GarantiaTotal_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GarantiaTotalPrazo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GarantiaTotalPrazo_GotFocus()
    Call MaskEdBox_TrataGotFocus(GarantiaTotalPrazo, iAlterado)
End Sub

Private Sub GarantiaTotalPrazo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_GarantiaTotalPrazo_Validate

    If Len(Trim(GarantiaTotalPrazo.ClipText)) = 0 Then Exit Sub

    lErro = Inteiro_Critica(GarantiaTotalPrazo.Text)
    If lErro <> SUCESSO Then gError 186118

    
    Exit Sub

Erro_GarantiaTotalPrazo_Validate:

    Cancel = True

    Select Case gErr

        Case 186118
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186119)

    End Select

    Exit Sub

End Sub

Private Sub PrazoPadraoValidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrazoPadraoValidade_GotFocus()

    Call MaskEdBox_TrataGotFocus(PrazoPadraoValidade, iAlterado)

End Sub

Public Sub PrazoPadraoValidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PrazoPadraoValidade_Validate

    If Len(Trim(PrazoPadraoValidade.ClipText)) = 0 Then Exit Sub

    lErro = Inteiro_Critica(PrazoPadraoValidade.Text)
    If lErro <> SUCESSO Then gError 186085

    
    Exit Sub

Erro_PrazoPadraoValidade_Validate:

    Cancel = True

    Select Case gErr

        Case 186085
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186086)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Descricao_GotFocus()

    Call MaskEdBox_TrataGotFocus(Descricao, iAlterado)

End Sub

Private Sub LblTipo_Click()

Dim objTipoGarantia As New ClassTipoGarantia
Dim colSelecao As New Collection

    objTipoGarantia.lCodigo = StrParaLong(Codigo.Text)
    
    Call Chama_Tela("TipoGarantiaLista", colSelecao, objTipoGarantia, objEventoCodigo)
    
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objTipoGarantia As ClassTipoGarantia
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTipoGarantia = obj1
    
    'Traz para a tela o tipo de garantia com código passado pelo browser
    lErro = Traz_TipoGarantia_Tela(objTipoGarantia)
    If lErro <> SUCESSO Then gError 186027
        
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 186027
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186028)
    
    End Select

End Sub

Public Sub BotaoServicos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoServicos_Click

    If Me.ActiveControl Is Servico Then
    
        sProduto1 = Servico.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridServicos.Row = 0 Then gError 186028

        sProduto1 = GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 186029
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto
    
    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoServico)

    Exit Sub
        
Erro_BotaoServicos_Click:
    
    Select Case gErr
        
        Case 186028
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 186029
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186030)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridServicos.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 186031

    Servico.PromptInclude = False
    Servico.Text = sProduto
    Servico.PromptInclude = True

    If Not (Me.ActiveControl Is Servico) Then
    
        GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col) = Servico.Text
    
        'Faz o Tratamento do produto
        lErro = Traz_Servico_Tela()
        If lErro <> SUCESSO Then gError 186032

    End If
    
    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 186031
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 186032
            GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col) = ""
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186033)

    End Select

    Exit Sub

End Sub

Private Function Traz_Servico_Tela() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Traz_Servico_Tela

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial", Servico.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 186034
    
    If lErro = 51381 Then gError 186035

    'Descricao Servico
    GridServicos.TextMatrix(GridServicos.Row, iGrid_ServicoDesc_Col) = objProduto.sDescricao

    'Acrescenta uma linha no Grid se for o caso
    If GridServicos.Row - GridServicos.FixedRows = objGridServico.iLinhasExistentes Then
        
        objGridServico.iLinhasExistentes = objGridServico.iLinhasExistentes + 1

    End If

    Traz_Servico_Tela = SUCESSO

    Exit Function

Erro_Traz_Servico_Tela:

    Traz_Servico_Tela = gErr

    Select Case gErr

        Case 186034

        Case 186035
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Servico.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186036)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 186037

    'Limpa a Tela
    Call Limpa_TipoGarantia

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 186037

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186038)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim objTipoGarantia As New ClassTipoGarantia
Dim lErro As Long
Dim sAviso As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Se o código não foi preenchido => erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 186039

    'Guarda no obj, código da garantia e filial empresa
    'Essas informações são necessárias para excluir a garantia
    objTipoGarantia.lCodigo = StrParaLong(Codigo.Text)

    'Lê a garantia
    lErro = CF("TipoGarantia_Le", objTipoGarantia)
    If lErro <> SUCESSO And lErro <> 183849 Then gError 186040
    
    'Se não encontrou => erro
    If lErro <> SUCESSO Then gError 186041
    
    'Pede a confirmação da exclusão da garantia
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TIPOGARANTIA")
    
    If vbMsgRes = vbYes Then

        'Faz a exclusão da Solicitacao
        lErro = CF("TipoGarantia_Exclui", objTipoGarantia)
        If lErro <> SUCESSO Then gError 186042
    
        'Limpa a Tela
        Call Limpa_TipoGarantia
        
        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 186039
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 186040, 186042

        Case 186041
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_NAO_ENCONTRADA", gErr, objTipoGarantia.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186043)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 186052

    'Limpa a Tela
    Call Limpa_TipoGarantia
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 186052

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186053)

    End Select

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

'*** TRATAMENTO DO EVENTO KEYDOWN  - INÍCIO ***
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LblTipo_Click
        ElseIf Me.ActiveControl Is Servico Then
            Call BotaoServicos_Click
        End If
    
    End If

End Sub


'***************************************************
'Trecho de codigo comum as telas
'***************************************************

Public Function Form_Load_Ocx() As Object
'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Tipo de Garantia"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "TipoGarantia"
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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
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
    RaiseEvent Unload
End Sub

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - INÍCIO ***
Private Sub LblTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTipo, Source, X, Y)
End Sub

Private Sub LblTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTipo, Button, Shift, X, Y)
End Sub

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - FIM ***


Private Function Traz_TipoGarantia_Tela(ByVal objTipoGarantia As ClassTipoGarantia) As Long
'Traz pra tela os dados da garantia passado como parâmetro

Dim lErro As Long
Dim bCancel As Boolean
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim iAchou As Integer

On Error GoTo Erro_Traz_TipoGarantia_Tela

    'Limpa a tela
    Call Limpa_TipoGarantia
    
    'Lê no BD os dados da garantia em questao
    lErro = CF("TipoGarantia_Le", objTipoGarantia)
    If lErro <> SUCESSO And lErro <> 183849 Then gError 186055
    
    Codigo.PromptInclude = False
    Codigo.Text = objTipoGarantia.lCodigo
    Codigo.PromptInclude = True
    
    'Se não encontrou a garantia => erro
    If lErro = SUCESSO Then

        Descricao.Text = objTipoGarantia.sDescricao
        
        If objTipoGarantia.iPrazoPadrao <> 0 Then PrazoPadraoValidade.Text = objTipoGarantia.iPrazoPadrao
        
        GarantiaTotal.Value = objTipoGarantia.iGarantiaTotal
        
        If objTipoGarantia.iGarantiaTotalPrazo <> 0 Then GarantiaTotalPrazo.Text = objTipoGarantia.iGarantiaTotalPrazo
        
        lErro = Carrega_Grid_Servicos(objTipoGarantia)
        If lErro <> SUCESSO Then gError 186057
    
    End If
    
    iAlterado = 0
    
    Traz_TipoGarantia_Tela = SUCESSO

    Exit Function

Erro_Traz_TipoGarantia_Tela:

    Traz_TipoGarantia_Tela = gErr

    Select Case gErr

        Case 186055, 186057
        
        Case 186056
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOGARANTIA_NAO_CADASTRADA", gErr, objTipoGarantia.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186058)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_Servicos(objTipoGarantia As ClassTipoGarantia) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sServicoEnxuto As String
Dim objTipoGarantiaProduto As ClassTipoGarantiaProduto
Dim objProduto As New ClassProduto

On Error GoTo Erro_Carrega_Grid_Servicos

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridServico)

    For iIndice = 1 To objTipoGarantia.colTipoGarantiaProduto.Count
       
        Set objTipoGarantiaProduto = objTipoGarantia.colTipoGarantiaProduto(iIndice)
       
        lErro = Mascara_RetornaProdutoEnxuto(objTipoGarantiaProduto.sProduto, sServicoEnxuto)
        If lErro <> SUCESSO Then gError 186059

        'Mascara o produto enxuto
        Servico.PromptInclude = False
        Servico.Text = sServicoEnxuto
        Servico.PromptInclude = True

        GridServicos.TextMatrix(iIndice, iGrid_Servico_Col) = Servico.Text
        
        objProduto.sCodigo = objTipoGarantiaProduto.sProduto
        
        'Lê o Servico
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 186060
        
        If lErro = SUCESSO Then
            GridServicos.TextMatrix(iIndice, iGrid_ServicoDesc_Col) = objProduto.sDescricao
        End If
        
        If objTipoGarantiaProduto.iPrazo <> 0 Then GridServicos.TextMatrix(iIndice, iGrid_PrazoValidade_Col) = objTipoGarantiaProduto.iPrazo
        
    Next

    'Atualiza o número de linhas existentes
    objGridServico.iLinhasExistentes = objTipoGarantia.colTipoGarantiaProduto.Count

    Carrega_Grid_Servicos = SUCESSO

    Exit Function

Erro_Carrega_Grid_Servicos:

    Carrega_Grid_Servicos = gErr

    Select Case gErr

        Case 186059, 186060
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186061)

    End Select

    Exit Function

End Function

Private Sub GridServicos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServico, iAlterado)

    End If

End Sub

Private Sub GridServicos_EnterCell()

    Call Grid_Entrada_Celula(objGridServico, iAlterado)

End Sub

Private Sub GridServicos_GotFocus()

    Call Grid_Recebe_Foco(objGridServico)

End Sub

Private Sub GridServicos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServico, iAlterado)
    End If


End Sub

Private Sub GridServicos_LeaveCell()

    Call Saida_Celula(objGridServico)

End Sub

Private Sub GridServicos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridServico)

End Sub

Private Sub GridServicos_Scroll()

    Call Grid_Scroll(objGridServico)

End Sub

Private Sub GridServicos_RowColChange()

    Call Grid_RowColChange(objGridServico)

End Sub

Private Sub GridServicos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridServico)

End Sub

Public Sub Servico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Servico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub Servico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub Servico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = Servico
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrazoValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrazoValidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub PrazoValidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub PrazoValidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = PrazoValidade
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col

            'Se for a de Servico
            Case iGrid_Servico_Col
                lErro = Saida_Celula_Servico(objGridInt)
                If lErro <> SUCESSO Then gError 186062
    
            'Se for a de Prazo
            Case iGrid_PrazoValidade_Col
                lErro = Saida_Celula_PrazoValidade(objGridInt)
                If lErro <> SUCESSO Then gError 186063
        
        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 186064
    
    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 186062 To 186064

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186065)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Servico(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim iIndice As Integer
Dim sServicoEnxuto As String
Dim sServico As String
Dim iPreenchido As Integer


On Error GoTo Erro_Saida_Celula_Servico

    Set objGridInt.objControle = Servico

    If Len(Trim(Servico.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica", Servico.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 186066

        'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
        If lErro <> SUCESSO Then gError 186067
                
        lErro = CF("Produto_Formata", Servico.Text, sServico, iPreenchido)
        If lErro <> SUCESSO Then gError 186087
    
    
        'Mascara o produto enxuto
        Servico.PromptInclude = False
        Servico.Text = sServico
        Servico.PromptInclude = True
                
        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridServico.iLinhasExistentes
            If iIndice <> GridServicos.Row Then
                If GridServicos.TextMatrix(iIndice, iGrid_Servico_Col) = Servico.Text Then gError 186122
            End If
        Next
                
                
    Else
        
        GridServicos.TextMatrix(GridServicos.Row, iGrid_ServicoDesc_Col) = ""
        GridServicos.TextMatrix(GridServicos.Row, iGrid_PrazoValidade_Col) = ""
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186068

    If Len(Trim(Servico.ClipText)) <> 0 Then

        GridServicos.TextMatrix(GridServicos.Row, iGrid_ServicoDesc_Col) = objProduto.sDescricao
        
        If Len(PrazoPadraoValidade.Text) <> 0 And Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_PrazoValidade_Col))) = 0 Then GridServicos.TextMatrix(GridServicos.Row, iGrid_PrazoValidade_Col) = PrazoPadraoValidade.Text
    
        If GridServicos.Row - GridServicos.FixedRows = objGridServico.iLinhasExistentes Then
            
            objGridServico.iLinhasExistentes = objGridServico.iLinhasExistentes + 1
    
        End If
    
    End If

    Saida_Celula_Servico = SUCESSO

    Exit Function

Erro_Saida_Celula_Servico:

    Saida_Celula_Servico = gErr

    Select Case gErr

        Case 186066, 186068
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186067
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Servico.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Servico.Text
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 186122
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_NO_GRID", gErr, Servico.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186069)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrazoValidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Garantia está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PrazoValidade

    Set objGridInt.objControle = PrazoValidade

    If Len(Trim(PrazoValidade.Text)) > 0 Then

        lErro = Inteiro_Critica(PrazoValidade.Text)
        If lErro <> SUCESSO Then gError 186070
        
    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186071
    
    Saida_Celula_PrazoValidade = SUCESSO

    Exit Function

Erro_Saida_Celula_PrazoValidade:

    Saida_Celula_PrazoValidade = gErr

    Select Case gErr

        Case 186070, 186071
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186072)

    End Select

    Exit Function

End Function

Private Sub Limpa_TipoGarantia()

Dim iIndice As Integer

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    GarantiaTotal.Value = 0
    
    Call Grid_Limpa(objGridServico)
    
    iAlterado = 0
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoGarantia As New ClassTipoGarantia

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os campos obrigatórios estão preenchidos
    lErro = Valida_Gravacao()
    If lErro <> SUCESSO Then gError 186073

    'Move os dados da tela para o objTipoGarantia
    lErro = Move_TipoGarantia_Memoria(objTipoGarantia)
    If lErro <> SUCESSO Then gError 186074

    'Verifica se esse tipo já existe no BD
    'e, em caso positivo, alerta ao usuário que está sendo feita uma alteração
    lErro = Trata_Alteracao(objTipoGarantia, objTipoGarantia.lCodigo)
    If lErro <> SUCESSO Then gError 186075
    
    'Grava no BD
    lErro = CF("TipoGarantia_Grava", objTipoGarantia)
    If lErro <> SUCESSO Then gError 186076

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 186073 To 186076
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186077)

    End Select

    Exit Function

End Function

Private Function Valida_Gravacao() As Long
'Verifica se os dados da tela são válidos para a gravação do registro

Dim lErro As Long
Dim iIndice As Integer
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Gravacao

    'Se o código não estiver preenchido => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 186078
    
    'Se o código não estiver preenchido => erro
    If Len(Trim(Descricao.Text)) = 0 Then gError 186079
    
    If GarantiaTotal.Value = 1 And Len(Trim(GarantiaTotalPrazo.Text)) = 0 Then gError 207375
    
    For iIndice = 1 To objGridServico.iLinhasExistentes

        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Servico_Col))) = 0 Then gError 186080
        
        If GarantiaTotal.Value = 0 Then
        
            If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_PrazoValidade_Col))) = 0 Then gError 186081
            
        End If
        
        
        
    Next
    
    Valida_Gravacao = SUCESSO

    Exit Function

Erro_Valida_Gravacao:

    Valida_Gravacao = gErr
    
    Select Case gErr
    
        Case 186078
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 186079
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            
        Case 186080
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case 186081
            Call Rotina_Erro(vbOKOnly, "ERRO_PRAZO_NAO_PREENCHIDO_GRID", gErr, iIndice)
        
        Case 207375
            Call Rotina_Erro(vbOKOnly, "ERRO_PRAZO_GARANTIATOTAL_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186082)

    End Select

End Function

Private Function Move_TipoGarantia_Memoria(objTipoGarantia As ClassTipoGarantia) As Long
'Move os dados da tela para objGarantia

Dim lErro As Long

On Error GoTo Erro_Move_TipoGarantia_Memoria

    objTipoGarantia.lCodigo = StrParaLong(Codigo.Text)

    objTipoGarantia.sDescricao = Descricao.Text

    objTipoGarantia.iPrazoPadrao = StrParaInt(PrazoPadraoValidade.Text)

    objTipoGarantia.iGarantiaTotal = GarantiaTotal.Value

    objTipoGarantia.iGarantiaTotalPrazo = StrParaInt(GarantiaTotalPrazo.Text)
    
    Set objTipoGarantia.objTela = Me

    'Move Grid Itens para memória
    lErro = Move_GridServico_Memoria(objTipoGarantia)
    If lErro <> SUCESSO Then gError 186083

    Move_TipoGarantia_Memoria = SUCESSO

    Exit Function

Erro_Move_TipoGarantia_Memoria:

    Move_TipoGarantia_Memoria = gErr

    Select Case gErr

        Case 186083

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186084)

    End Select

    Exit Function

End Function

Private Function Move_GridServico_Memoria(objTipoGarantia As ClassTipoGarantia) As Long
'Recolhe do Grid os dados

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objTipoGarantiaProduto As ClassTipoGarantiaProduto
Dim iIndice As Integer

On Error GoTo Erro_Move_GridServico_Memoria

    For iIndice = 1 To objGridServico.iLinhasExistentes

        Set objTipoGarantiaProduto = New ClassTipoGarantiaProduto
    
        'Formata o produto
        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Servico_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 186087
    
        If iPreenchido = PRODUTO_VAZIO Then gError 186088
    
        objTipoGarantiaProduto.sProduto = sProduto
        objTipoGarantiaProduto.iPrazo = StrParaInt(GridServicos.TextMatrix(iIndice, iGrid_PrazoValidade_Col))
    
        objTipoGarantia.colTipoGarantiaProduto.Add objTipoGarantiaProduto
    
    Next
    
    Move_GridServico_Memoria = SUCESSO

    Exit Function

Erro_Move_GridServico_Memoria:

    Move_GridServico_Memoria = gErr

    Select Case gErr

        Case 186087

        Case 186088
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186089)

    End Select

    Exit Function

End Function

'**** TRATAMENTO DO SISTEMA DE SETAS - INÍCIO ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objTipoGarantia As New ClassTipoGarantia
Dim objCampoValor As AdmCampoValor
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TipoGarantia"

    'Guarda no obj os dados que serão usados para identifica o registro a ser exibido
    objTipoGarantia.lCodigo = StrParaLong(Trim(Codigo.Text))
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTipoGarantia.lCodigo, 0, "Codigo"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186090)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTipoGarantia As New ClassTipoGarantia

On Error GoTo Erro_Tela_Preenche

    'Guarda o código do campo em questão no obj
    objTipoGarantia.lCodigo = colCampoValor.Item("Codigo").vValor

    lErro = Traz_TipoGarantia_Tela(objTipoGarantia)
    If lErro <> SUCESSO Then gError 186091

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 186091
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186092)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub
'**** FIM DO TRATAMENTO DO SISTEMA DE SETAS ****


