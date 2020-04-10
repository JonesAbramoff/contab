VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TiposDetRetCobrOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.ComboBox CodigoMovto 
      Height          =   315
      Left            =   2145
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   750
      Width           =   5730
   End
   Begin VB.ComboBox Banco 
      Height          =   315
      Left            =   2145
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   315
      Width           =   2340
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7815
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   135
      Width           =   1560
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TiposDetRetCobr.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   540
         Picture         =   "TiposDetRetCobr.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1020
         Picture         =   "TiposDetRetCobr.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.Frame FrameGridDetalhe 
      Caption         =   "Tipos de Ocorrências"
      Height          =   4800
      Left            =   120
      TabIndex        =   8
      Top             =   1110
      Width           =   9315
      Begin VB.CommandButton BotaoTiposDif 
         Caption         =   "Tipos de Diferença"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4290
         UseMaskColor    =   -1  'True
         Width           =   2085
      End
      Begin VB.ComboBox DetAcao 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4305
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1620
         Width           =   1275
      End
      Begin VB.ComboBox DetAcaoManual 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4305
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2100
         Width           =   1275
      End
      Begin MSMask.MaskEdBox DetTipoDifDescricao 
         Height          =   315
         Left            =   4245
         TabIndex        =   14
         Top             =   3135
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DetDescricao 
         Height          =   315
         Left            =   4230
         TabIndex        =   13
         Top             =   1095
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DetCodigo 
         Height          =   315
         Left            =   4200
         TabIndex        =   11
         Top             =   660
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DetTipoDif 
         Height          =   315
         Left            =   4290
         TabIndex        =   12
         Top             =   2610
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   4
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridDetalhe 
         Height          =   3345
         Left            =   75
         TabIndex        =   2
         Top             =   210
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5900
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.Label LabelBanco 
      Alignment       =   1  'Right Justify
      Caption         =   "Banco:"
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
      Height          =   315
      Left            =   570
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   9
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label LabelCodigoMovto 
      Alignment       =   1  'Right Justify
      Caption         =   "Código do Movimento:"
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
      Height          =   315
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   780
      Width           =   1950
   End
End
Attribute VB_Name = "TiposDetRetCobrOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim lBancoAnt As Long
Dim iMovtoAnt As Integer

Private WithEvents objEventoTiposDifParcRec As AdmEvento
Attribute objEventoTiposDifParcRec.VB_VarHelpID = -1

Dim objGridDetalhe As AdmGrid
Dim iGrid_DetCodigo_Col As Integer
Dim iGrid_DetDescricao_Col As Integer
Dim iGrid_DetAcao_Col As Integer
Dim iGrid_DetAcaoManual_Col As Integer
Dim iGrid_DetTipoDif_Col As Integer
Dim iGrid_DetTipoDifDescricao_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Detalhamento dos Tipos de Movto do Arquivo de Retorno"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TiposDetRetCobr"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is DetTipoDif Then
            Call BotaoTiposDif_Click
        End If
    End If
End Sub

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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridDetalhe = Nothing
    Set objEventoTiposDifParcRec = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177771)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoTiposDifParcRec = New AdmEvento
    
    lErro = Carrega_Bancos
    If lErro <> SUCESSO Then gError 177748

    lErro = Carrega_TiposDif(DetAcao)
    If lErro <> SUCESSO Then gError 177749

    lErro = Carrega_TiposDif(DetAcaoManual)
    If lErro <> SUCESSO Then gError 177757

    lErro = Inicializa_GridDetalhe(objGridDetalhe)
    If lErro <> SUCESSO Then gError 177623

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 177623, 177748, 177749, 177757

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177772)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTiposDetRetCobr As ClassTiposDetRetCobr) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTiposDetRetCobr Is Nothing) Then

        lErro = Traz_TiposDetRetCobr_Tela(objTiposDetRetCobr)
        If lErro <> SUCESSO Then gError 177624

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 177624

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177773)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objTiposDetRetCobr As ClassTiposDetRetCobr, ByVal colTiposDetRetCobr As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objTiposDetRetCobr.lBanco = LCodigo_Extrai(Banco.Text)
    objTiposDetRetCobr.iCodigoMovto = Codigo_Extrai(CodigoMovto.Text)
    
    lErro = Move_GridDetalhe_Memoria(objTiposDetRetCobr, colTiposDetRetCobr)
    If lErro <> SUCESSO Then gError 177750

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 177750

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177774)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objTiposDetRetCobr As New ClassTiposDetRetCobr
Dim colTiposDetRetCobr As New Collection

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TiposDetRetCobr"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objTiposDetRetCobr, colTiposDetRetCobr)
    If lErro <> SUCESSO Then gError 177625

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Banco", objTiposDetRetCobr.lBanco, 0, "Banco"
    colCampoValor.Add "CodigoMovto", objTiposDetRetCobr.iCodigoMovto, 0, "CodigoMovto"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 177625

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177775)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTiposDetRetCobr As New ClassTiposDetRetCobr

On Error GoTo Erro_Tela_Preenche

    objTiposDetRetCobr.lBanco = colCampoValor.Item("Banco").vValor
    objTiposDetRetCobr.iCodigoMovto = colCampoValor.Item("CodigoMovto").vValor

    If objTiposDetRetCobr.lBanco <> 0 And objTiposDetRetCobr.iCodigoMovto <> 0 Then
        
        lErro = Traz_TiposDetRetCobr_Tela(objTiposDetRetCobr)
        If lErro <> SUCESSO Then gError 177626
    
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 177626

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177776)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTiposDetRetCobr As New ClassTiposDetRetCobr
Dim colTiposDetRetCobr As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Banco.Text)) = 0 Then gError 177627
    If Len(Trim(CodigoMovto.Text)) = 0 Then gError 177628

    'Preenche o objTiposDetRetCobr
    lErro = Move_Tela_Memoria(objTiposDetRetCobr, colTiposDetRetCobr)
    If lErro <> SUCESSO Then gError 177629

    lErro = Trata_Alteracao(objTiposDetRetCobr, objTiposDetRetCobr.lBanco, objTiposDetRetCobr.iCodigoMovto)
    If lErro <> SUCESSO Then gError 177630

    'Grava o/a TiposDetRetCobr no Banco de Dados
    lErro = CF("TiposDetRetCobr_Grava", colTiposDetRetCobr)
    If lErro <> SUCESSO Then gError 177631

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 177627
            Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_PREENCHIDO", gErr)
            Banco.SetFocus

        Case 177628
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOMOVTO_NAO_PREENCHIDO", gErr)
            CodigoMovto.SetFocus

        Case 177629, 177630, 177631

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177777)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TiposDetRetCobr() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TiposDetRetCobr

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridDetalhe)
    
    CodigoMovto.ListIndex = -1
    
    iAlterado = 0

    Limpa_Tela_TiposDetRetCobr = SUCESSO

    Exit Function

Erro_Limpa_Tela_TiposDetRetCobr:

    Limpa_Tela_TiposDetRetCobr = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177778)

    End Select

    Exit Function

End Function

Function Traz_TiposDetRetCobr_Tela(objTiposDetRetCobr As ClassTiposDetRetCobr) As Long

Dim lErro As Long
Dim colTiposDetRetCobr As New Collection

On Error GoTo Erro_Traz_TiposDetRetCobr_Tela

    'Lê o TiposDetRetCobr que está sendo Passado
    lErro = CF("TiposDetRetCobr_Le", objTiposDetRetCobr, colTiposDetRetCobr)
    If lErro <> SUCESSO And lErro <> 177604 Then gError 177632

    If lErro = SUCESSO Then

        Call Combo_Seleciona_ItemData(Banco, objTiposDetRetCobr.lBanco)
        
        Call Combo_Seleciona_ItemData(CodigoMovto, objTiposDetRetCobr.iCodigoMovto)
        
        lErro = Preenche_GridDetalhe_Tela(colTiposDetRetCobr)
        If lErro <> SUCESSO Then gError 177751

    End If

    iAlterado = 0

    Traz_TiposDetRetCobr_Tela = SUCESSO

    Exit Function

Erro_Traz_TiposDetRetCobr_Tela:

    Traz_TiposDetRetCobr_Tela = gErr

    Select Case gErr

        Case 177632, 177751

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177779)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 177633

    'Limpa Tela
    Call Limpa_Tela_TiposDetRetCobr

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 177633

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177780)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177781)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 177634

    Call Limpa_Tela_TiposDetRetCobr

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 177634

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177782)

    End Select

    Exit Sub

End Sub

Private Sub Banco_Change()

    iAlterado = REGISTRO_ALTERADO
    
    Call Carrega_TiposMovto
    
End Sub

Private Sub Banco_Click()

    iAlterado = REGISTRO_ALTERADO
    
    Call Carrega_TiposMovto
    
End Sub

Private Sub CodigoMovto_Change()

Dim objTiposDetRetCobr As New ClassTiposDetRetCobr
    
    iAlterado = REGISTRO_ALTERADO
    
    If iMovtoAnt <> Codigo_Extrai(CodigoMovto.Text) Then
    
        objTiposDetRetCobr.lBanco = LCodigo_Extrai(Banco.Text)
        objTiposDetRetCobr.iCodigoMovto = Codigo_Extrai(CodigoMovto.Text)
    
        Call Traz_TiposDetRetCobr_Tela(objTiposDetRetCobr)
    
        iMovtoAnt = Codigo_Extrai(CodigoMovto.Text)
    
    End If
    
End Sub

Private Sub CodigoMovto_Click()

Dim objTiposDetRetCobr As New ClassTiposDetRetCobr
    
    iAlterado = REGISTRO_ALTERADO
    
    If iMovtoAnt <> Codigo_Extrai(CodigoMovto.Text) Then
    
        objTiposDetRetCobr.lBanco = LCodigo_Extrai(Banco.Text)
        objTiposDetRetCobr.iCodigoMovto = Codigo_Extrai(CodigoMovto.Text)
    
        Call Traz_TiposDetRetCobr_Tela(objTiposDetRetCobr)
    
        iMovtoAnt = Codigo_Extrai(CodigoMovto.Text)
    
    End If
    
End Sub

Private Function Inicializa_GridDetalhe(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Código")
    objGrid.colColuna.Add ("Descricao")
    objGrid.colColuna.Add ("Ação Padrão")
    objGrid.colColuna.Add ("Ação Desejada")
    objGrid.colColuna.Add ("Tipo Dif")
    objGrid.colColuna.Add ("Descrição da Diferença")

    'Controles que participam do Grid
    objGrid.colCampo.Add (DetCodigo.Name)
    objGrid.colCampo.Add (DetDescricao.Name)
    objGrid.colCampo.Add (DetAcao.Name)
    objGrid.colCampo.Add (DetAcaoManual.Name)
    objGrid.colCampo.Add (DetTipoDif.Name)
    objGrid.colCampo.Add (DetTipoDifDescricao.Name)

    'Colunas do Grid
    iGrid_DetCodigo_Col = 1
    iGrid_DetDescricao_Col = 2
    iGrid_DetAcao_Col = 3
    iGrid_DetAcaoManual_Col = 4
    iGrid_DetTipoDif_Col = 5
    iGrid_DetTipoDifDescricao_Col = 6

    objGrid.objGrid = GridDetalhe

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridDetalhe.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridDetalhe = SUCESSO

End Function

Private Sub GridDetalhe_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridDetalhe, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridDetalhe, iAlterado)
        End If
End Sub

Private Sub GridDetalhe_GotFocus()
    Call Grid_Recebe_Foco(objGridDetalhe)
End Sub

Private Sub GridDetalhe_EnterCell()
    Call Grid_Entrada_Celula(objGridDetalhe, iAlterado)
End Sub

Private Sub GridDetalhe_LeaveCell()
    Call Saida_Celula(objGridDetalhe)
End Sub

Private Sub GridDetalhe_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDetalhe, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDetalhe, iAlterado)
    End If

End Sub

Private Sub GridDetalhe_RowColChange()
    Call Grid_RowColChange(objGridDetalhe)
End Sub

Private Sub GridDetalhe_Scroll()
    Call Grid_Scroll(objGridDetalhe)
End Sub

Private Sub GridDetalhe_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridDetalhe)
End Sub

Private Sub GridDetalhe_LostFocus()
    Call Grid_Libera_Foco(objGridDetalhe)
End Sub

Private Sub DetCodigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DetCodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDetalhe)
End Sub

Private Sub DetCodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDetalhe)
End Sub

Private Sub DetCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDetalhe.objControle = DetCodigo
    lErro = Grid_Campo_Libera_Foco(objGridDetalhe)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DetDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DetDescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDetalhe)
End Sub

Private Sub DetDescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDetalhe)
End Sub

Private Sub DetDescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDetalhe.objControle = DetDescricao
    lErro = Grid_Campo_Libera_Foco(objGridDetalhe)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DetAcao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DetAcao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDetalhe)
End Sub

Private Sub DetAcao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDetalhe)
End Sub

Private Sub DetAcao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDetalhe.objControle = DetAcao
    lErro = Grid_Campo_Libera_Foco(objGridDetalhe)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DetAcaoManual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DetAcaoManual_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDetalhe)
End Sub

Private Sub DetAcaoManual_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDetalhe)
End Sub

Private Sub DetAcaoManual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDetalhe.objControle = DetAcaoManual
    lErro = Grid_Campo_Libera_Foco(objGridDetalhe)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DetTipoDif_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DetTipoDif_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDetalhe)
End Sub

Private Sub DetTipoDif_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDetalhe)
End Sub

Private Sub DetTipoDif_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDetalhe.objControle = DetTipoDif
    lErro = Grid_Campo_Libera_Foco(objGridDetalhe)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DetTipoDifDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DetTipoDifDescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDetalhe)
End Sub

Private Sub DetTipoDifDescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDetalhe)
End Sub

Private Sub DetTipoDifDescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDetalhe.objControle = DetTipoDifDescricao
    lErro = Grid_Campo_Libera_Foco(objGridDetalhe)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_DetAcaoManual(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DetAcaoManual

    Set objGridInt.objControle = DetAcaoManual

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 177647

    Saida_Celula_DetAcaoManual = SUCESSO

    Exit Function

Erro_Saida_Celula_DetAcaoManual:

    Saida_Celula_DetAcaoManual = gErr

    Select Case gErr

        Case 177647
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177783)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_DetTipoDif(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objTipoDif As New ClassTiposDifParcRec

On Error GoTo Erro_Saida_Celula_DetTipoDif

    Set objGridInt.objControle = DetTipoDif
    
    If Len(Trim(DetTipoDif.Text)) <> 0 Then
    
        lErro = Inteiro_Critica(DetTipoDif.Text)
        If lErro <> SUCESSO Then gError 177757
        
        objTipoDif.iCodigo = StrParaInt(DetTipoDif.Text)
        
        lErro = CF("TiposDifParcRec_Le", objTipoDif)
        If lErro <> SUCESSO And lErro <> 177657 Then gError 177758
        
        If lErro <> SUCESSO Then gError 177759
        
        GridDetalhe.TextMatrix(GridDetalhe.Row, iGrid_DetTipoDifDescricao_Col) = objTipoDif.sDescricao
    Else
        GridDetalhe.TextMatrix(GridDetalhe.Row, iGrid_DetTipoDifDescricao_Col) = ""
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 177648

    Saida_Celula_DetTipoDif = SUCESSO

    Exit Function

Erro_Saida_Celula_DetTipoDif:

    Saida_Celula_DetTipoDif = gErr

    Select Case gErr

        Case 177648, 177757, 177758
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 177759
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDIFPARCREC_NAO_CADASTRADO", gErr, objTipoDif.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177784)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'GridDetalhe
        If objGridInt.objGrid.Name = GridDetalhe.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_DetAcaoManual_Col

                    lErro = Saida_Celula_DetAcaoManual(objGridInt)
                    If lErro <> SUCESSO Then gError 177651

                Case iGrid_DetTipoDif_Col

                    lErro = Saida_Celula_DetTipoDif(objGridInt)
                    If lErro <> SUCESSO Then gError 177652

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 177653

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 177651 To 177652

        Case 177653
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177785)

    End Select

    Exit Function

End Function

Function Preenche_GridDetalhe_Tela(ByVal colTiposDetRetCobr As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTiposDetRetCobr As ClassTiposDetRetCobr
Dim objTipoDif As ClassTiposDifParcRec

On Error GoTo Erro_Preenche_GridDetalhe_Tela

    Call Grid_Limpa(objGridDetalhe)

    For Each objTiposDetRetCobr In colTiposDetRetCobr

        iIndice = iIndice + 1
        
        Call Combo_Seleciona_ItemData(DetAcao, objTiposDetRetCobr.iAcao)
    
        GridDetalhe.TextMatrix(iIndice, iGrid_DetAcao_Col) = DetAcao.Text
        
        Call Combo_Seleciona_ItemData(DetAcaoManual, objTiposDetRetCobr.iAcaoManual)
        
        GridDetalhe.TextMatrix(iIndice, iGrid_DetAcaoManual_Col) = DetAcaoManual.Text
        
        GridDetalhe.TextMatrix(iIndice, iGrid_DetCodigo_Col) = CStr(objTiposDetRetCobr.iCodigoDetalhe)
        GridDetalhe.TextMatrix(iIndice, iGrid_DetDescricao_Col) = objTiposDetRetCobr.sDescricao
        
        If objTiposDetRetCobr.iCodTipoDiferenca <> 0 Then
            GridDetalhe.TextMatrix(iIndice, iGrid_DetTipoDif_Col) = CStr(objTiposDetRetCobr.iCodTipoDiferenca)
        Else
            GridDetalhe.TextMatrix(iIndice, iGrid_DetTipoDif_Col) = ""
        End If
        
        Set objTipoDif = New ClassTiposDifParcRec
        
        objTipoDif.iCodigo = objTiposDetRetCobr.iCodTipoDiferenca
        
        lErro = CF("TiposDifParcRec_Le", objTipoDif)
        If lErro <> SUCESSO And lErro <> 177657 Then gError 177754
        
        GridDetalhe.TextMatrix(iIndice, iGrid_DetTipoDifDescricao_Col) = objTipoDif.sDescricao

    Next

    objGridDetalhe.iLinhasExistentes = iIndice

    Preenche_GridDetalhe_Tela = SUCESSO

    Exit Function

Erro_Preenche_GridDetalhe_Tela:

    Preenche_GridDetalhe_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177786)

    End Select

    Exit Function

End Function

Function Move_GridDetalhe_Memoria(ByVal objTiposDetRetCobr As ClassTiposDetRetCobr, ByVal colTiposDetRetCobr As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTiposDetRetCobrDet As ClassTiposDetRetCobr

On Error GoTo Erro_Move_GridDetalhe_Memoria

    For iIndice = 1 To objGridDetalhe.iLinhasExistentes
    
        Set objTiposDetRetCobrDet = New ClassTiposDetRetCobr
        
        objTiposDetRetCobrDet.lBanco = objTiposDetRetCobr.lBanco
        objTiposDetRetCobrDet.iCodigoMovto = objTiposDetRetCobr.iCodigoMovto
        objTiposDetRetCobrDet.iCodigoDetalhe = StrParaInt(GridDetalhe.TextMatrix(iIndice, iGrid_DetCodigo_Col))
        
        objTiposDetRetCobrDet.iAcaoManual = Codigo_Extrai(GridDetalhe.TextMatrix(iIndice, iGrid_DetAcaoManual_Col))
        objTiposDetRetCobrDet.iCodTipoDiferenca = StrParaInt(GridDetalhe.TextMatrix(iIndice, iGrid_DetTipoDif_Col))
    
        colTiposDetRetCobr.Add objTiposDetRetCobrDet
    
    Next

    Move_GridDetalhe_Memoria = SUCESSO

    Exit Function

Erro_Move_GridDetalhe_Memoria:

    Move_GridDetalhe_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177787)

    End Select

    Exit Function

End Function

Private Function Carrega_Bancos() As Long

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Bancos

    'leitura dos codigos e descricoes das ListaCodConta de venda no BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then gError 177747

   'preenche ComboBox com código e nome dos CodBancos
    For iIndice = 1 To colCodigoNome.Count
        Set objCodigoNome = colCodigoNome(iIndice)
        Banco.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Banco.ItemData(Banco.NewIndex) = objCodigoNome.iCodigo
    Next

    'Seleciona uma dos Bancos
    Banco.Text = Banco.List(PRIMEIRA_CONTA)
    
    lBancoAnt = LCodigo_Extrai(Banco.Text)
    
    Call Carrega_TiposMovto

    Carrega_Bancos = SUCESSO

    Exit Function

Erro_Carrega_Bancos:

    Carrega_Bancos = gErr

    Select Case gErr

        Case 177747

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177746)

    End Select

    Exit Function

End Function

Private Function Carrega_TiposDif(ByVal objCombo As ComboBox) As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_TiposDif

    objCombo.AddItem CStr(TIPODIF_ACAO_AUTOMATICA) & SEPARADOR & STRING_TIPODIF_ACAO_AUTOMATICA
    objCombo.ItemData(objCombo.NewIndex) = TIPODIF_ACAO_AUTOMATICA

    objCombo.AddItem CStr(TIPODIF_ACAO_INFORMATIVA) & SEPARADOR & STRING_TIPODIF_ACAO_INFORMATIVA
    objCombo.ItemData(objCombo.NewIndex) = TIPODIF_ACAO_INFORMATIVA

    objCombo.AddItem CStr(TIPODIF_ACAO_SOMA) & SEPARADOR & STRING_TIPODIF_ACAO_SOMA
    objCombo.ItemData(objCombo.NewIndex) = TIPODIF_ACAO_SOMA

    objCombo.AddItem CStr(TIPODIF_ACAO_SUBTRAI) & SEPARADOR & STRING_TIPODIF_ACAO_SUBTRAI
    objCombo.ItemData(objCombo.NewIndex) = TIPODIF_ACAO_SUBTRAI

    'Seleciona uma dos Bancos
    objCombo.Text = objCombo.List(0)

    Carrega_TiposDif = SUCESSO

    Exit Function

Erro_Carrega_TiposDif:

    Carrega_TiposDif = gErr

    Select Case gErr

        Case 177755

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177756)

    End Select

    Exit Function

End Function

Private Function Carrega_TiposMovto() As Long

Dim lErro As Long
Dim objTiposMovRetCobr As New ClassTiposMovRetCobr
Dim colTiposMovRetCobr As New Collection

On Error GoTo Erro_Carrega_TiposMovto

    If lBancoAnt <> LCodigo_Extrai(Banco.Text) Then
    
        CodigoMovto.Clear
    
        objTiposMovRetCobr.lBanco = LCodigo_Extrai(Banco.Text)
        
        lErro = CF("TiposMovRetCobr_Le", objTiposMovRetCobr, colTiposMovRetCobr)
        If lErro <> SUCESSO Then gError 177764
        
        For Each objTiposMovRetCobr In colTiposMovRetCobr
    
            CodigoMovto.AddItem CStr(objTiposMovRetCobr.iCodigoMovto) & SEPARADOR & objTiposMovRetCobr.sDescricao
            CodigoMovto.ItemData(CodigoMovto.NewIndex) = objTiposMovRetCobr.iCodigoMovto
    
        Next
        
        lBancoAnt = LCodigo_Extrai(Banco.Text)
        
    End If

    Carrega_TiposMovto = SUCESSO

    Exit Function

Erro_Carrega_TiposMovto:

    Carrega_TiposMovto = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177756)

    End Select

    Exit Function

End Function

Private Sub BotaoTiposDif_Click()

Dim lErro As Long
Dim objTiposDifParcRec As New ClassTiposDifParcRec
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoItemContrato_Click

    If GridDetalhe.Row = 0 Then gError 177768
    
    If Not (Me.ActiveControl Is DetTipoDif) Then
        objTiposDifParcRec.iCodigo = StrParaInt(GridDetalhe.TextMatrix(GridDetalhe.Row, iGrid_DetTipoDif_Col))
    Else
        objTiposDifParcRec.iCodigo = StrParaInt(DetTipoDif.Text)
    End If
    
    Call Chama_Tela("TiposDifParcRecLista", colSelecao, objTiposDifParcRec, objEventoTiposDifParcRec)

    Exit Sub
    
Erro_BotaoItemContrato_Click:
    
    Select Case gErr
    
        Case 177768
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177788)

    End Select
    
    Exit Sub


End Sub

Private Sub objEventoTiposDifParcRec_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDifParcRec As ClassTiposDifParcRec

On Error GoTo Erro_objEventoTiposDifParcRec_evSelecao

    Set objTiposDifParcRec = obj1

    GridDetalhe.TextMatrix(GridDetalhe.Row, iGrid_DetTipoDif_Col) = CStr(objTiposDifParcRec.iCodigo)
    
    If Not (Me.ActiveControl Is DetTipoDif) Then
        GridDetalhe.TextMatrix(GridDetalhe.Row, iGrid_DetTipoDifDescricao_Col) = objTiposDifParcRec.sDescricao
    Else
        DetTipoDif.Text = CStr(objTiposDifParcRec.iCodigo)
    End If

    Me.Show

    Exit Sub

Erro_objEventoTiposDifParcRec_evSelecao:

    Select Case gErr

        Case 177769

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177770)

    End Select

    Exit Sub

End Sub
