VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Dan_OS 
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   7710
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5415
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "Dan_OS.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "Dan_OS.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "Dan_OS.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "Dan_OS.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox OS 
      Height          =   315
      Left            =   2000
      TabIndex        =   0
      Top             =   300
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   9
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Modelo 
      Height          =   315
      Left            =   2000
      TabIndex        =   2
      Top             =   1200
      Width           =   5500
      _ExtentX        =   9710
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumSerie 
      Height          =   315
      Left            =   2000
      TabIndex        =   4
      Top             =   1650
      Width           =   2200
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Frame FrameGridItens 
      Caption         =   "Itens"
      Height          =   3180
      Left            =   360
      TabIndex        =   14
      Top             =   2115
      Width           =   7170
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "Produtos"
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
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Abre o Browse de Produtos"
         Top             =   2790
         Width           =   1380
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   1395
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   315
         Left            =   2340
         TabIndex        =   7
         Top             =   1380
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   315
         Left            =   5070
         TabIndex        =   8
         Top             =   1395
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   2325
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   4101
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   1995
      TabIndex        =   1
      Top             =   780
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label LabelCliente 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
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
      Left            =   1275
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   18
      Top             =   825
      Width           =   660
   End
   Begin VB.Label LabelOS 
      Alignment       =   1  'Right Justify
      Caption         =   "OS:"
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
      Left            =   375
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   325
      Width           =   1500
   End
   Begin VB.Label LabelModelo 
      Alignment       =   1  'Right Justify
      Caption         =   "Modelo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   375
      TabIndex        =   16
      Top             =   1225
      Width           =   1500
   End
   Begin VB.Label LabelNumSerie 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de Serie:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   375
      TabIndex        =   17
      Top             =   1675
      Width           =   1500
   End
End
Attribute VB_Name = "Dan_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGridItens As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Quantidade_Col As Integer

Private WithEvents objEventoOS As AdmEvento
Attribute objEventoOS.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "OS"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Dan_OS"

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

Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_UnLoad

    Set objGridItens = Nothing

    Set objEventoOS = Nothing
    Set objEventoCliente = Nothing
    Set objEventoProduto = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184798)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoOS = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoProduto = New AdmEvento

    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 184799
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 198271

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 184799, 198271

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184800)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objDan_OS As ClassDan_OS) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objDan_OS Is Nothing) Then

        lErro = Traz_Dan_OS_Tela(objDan_OS)
        If lErro <> SUCESSO Then gError 184801

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 184801

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184802)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objDan_ItensOS As ClassDan_ItensOS
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCliente As New ClassCliente

On Error GoTo Erro_Move_Tela_Memoria

    objDan_OS.sOS = OS.Text
    objDan_OS.sModelo = Modelo.Text
    objDan_OS.sNumSerie = NumSerie.Text
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objDan_ItensOS = New ClassDan_ItensOS
        objDan_OS.colItens.Add objDan_ItensOS
        
        objDan_ItensOS.iItem = iIndice
        
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 198273
        
        objDan_ItensOS.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objDan_ItensOS.sProduto = sProdutoFormatado

    Next

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 198274

        If lErro = SUCESSO Then
            objDan_OS.lCliente = objCliente.lCodigo
        End If
            
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 198273, 198274

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184803)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objDan_OS As New ClassDan_OS

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Dan_OS"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objDan_OS)
    If lErro <> SUCESSO Then gError 184804

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "OS", objDan_OS.sOS, STRING_DAN_OS, "OS"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 184804

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184805)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objDan_OS As New ClassDan_OS

On Error GoTo Erro_Tela_Preenche

    objDan_OS.sOS = colCampoValor.Item("OS").vValor

    If Len(Trim(objDan_OS.sOS)) > 0 Then

        lErro = Traz_Dan_OS_Tela(objDan_OS)
        If lErro <> SUCESSO Then gError 184806

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 184806

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184807)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objDan_OS As New ClassDan_OS

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(OS.Text)) = 0 Then gError 184808
    If Len(Trim(Cliente.Text)) = 0 Then gError 198275
    '#####################

    'Preenche o objDan_OS
    lErro = Move_Tela_Memoria(objDan_OS)
    If lErro <> SUCESSO Then gError 184809

    lErro = Trata_Alteracao(objDan_OS, objDan_OS.sOS)
    If lErro <> SUCESSO Then gError 184810

    'Grava o/a Dan_OS no Banco de Dados
    lErro = CF("Dan_OS_Grava", objDan_OS)
    If lErro <> SUCESSO Then gError 184811

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 184808
            Call Rotina_Erro(vbOKOnly, "ERRO_OS_DAN_OS_NAO_PREENCHIDO", gErr)
            OS.SetFocus

        Case 184809, 184810, 184811
        
        Case 198275
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            Cliente.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184812)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Dan_OS() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Dan_OS

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Call Grid_Limpa(objGridItens)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_Dan_OS = SUCESSO

    Exit Function

Erro_Limpa_Tela_Dan_OS:

    Limpa_Tela_Dan_OS = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184813)

    End Select

    Exit Function

End Function

Function Traz_Dan_OS_Tela(objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim objDan_ItensOS As ClassDan_ItensOS
Dim objProdutos As ClassProduto
Dim sProdutoMascarado As String
Dim iIndice As Integer

On Error GoTo Erro_Traz_Dan_OS_Tela
    
    Call Limpa_Tela_Dan_OS
    
    OS.Text = objDan_OS.sOS

    'Lê o Dan_OS que está sendo Passado
    lErro = CF("Dan_OS_Le", objDan_OS)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 184814

    If lErro = SUCESSO Then

        OS.Text = objDan_OS.sOS

        If objDan_OS.lCliente <> 0 Then
            Cliente.PromptInclude = False
            Cliente.Text = CStr(objDan_OS.lCliente)
            Cliente.PromptInclude = True
            Call Cliente_Validate(bSGECancelDummy)
        End If

        Modelo.Text = objDan_OS.sModelo
        NumSerie.Text = objDan_OS.sNumSerie

    End If

    iIndice = 0
    For Each objDan_ItensOS In objDan_OS.colItens
    
        iIndice = iIndice + 1
        
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objDan_ItensOS.sProduto
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 198276
        
        lErro = Mascara_RetornaProdutoTela(objProdutos.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 198277
                                
        'Insere no Grid MaquinasInsumos
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        GridItens.TextMatrix(iIndice, iGrid_Descricao_Col) = objProdutos.sDescricao
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objDan_ItensOS.dQuantidade)

    Next
    
    objGridItens.iLinhasExistentes = objDan_OS.colItens.Count

    iAlterado = 0

    Traz_Dan_OS_Tela = SUCESSO

    Exit Function

Erro_Traz_Dan_OS_Tela:

    Traz_Dan_OS_Tela = gErr

    Select Case gErr

        Case 184814, 198276, 198277

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184815)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 184816

    'Limpa Tela
    Call Limpa_Tela_Dan_OS

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 184816

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184817)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184818)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 184819

    Call Limpa_Tela_Dan_OS

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 184819

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184820)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objDan_OS As New ClassDan_OS
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(OS.Text)) = 0 Then gError 184821
    '#####################

    objDan_OS.sOS = OS.Text

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_DAN_OS", objDan_OS.sOS)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("Dan_OS_Exclui", objDan_OS)
        If lErro <> SUCESSO Then gError 184822

        'Limpa Tela
        Call Limpa_Tela_Dan_OS

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 184821
            Call Rotina_Erro(vbOKOnly, "ERRO_OS_DAN_OS_NAO_PREENCHIDO", gErr)
            OS.SetFocus

        Case 184822

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184823)

    End Select

    Exit Sub

End Sub

Private Sub OS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OS_Validate

    'Verifica se OS está preenchida
    If Len(Trim(OS.Text)) <> 0 Then

       '#######################################
       'CRITICA OS
       '#######################################

    End If

    Exit Sub

Erro_OS_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184824)

    End Select

    Exit Sub

End Sub

Private Sub OS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Modelo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Modelo_Validate

    'Verifica se Modelo está preenchida
    If Len(Trim(Modelo.Text)) <> 0 Then

       '#######################################
       'CRITICA Modelo
       '#######################################

    End If

    Exit Sub

Erro_Modelo_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184827)

    End Select

    Exit Sub

End Sub

Private Sub Modelo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumSerie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumSerie_Validate

    'Verifica se NumSerie está preenchida
    If Len(Trim(NumSerie.Text)) <> 0 Then

       '#######################################
       'CRITICA NumSerie
       '#######################################

    End If

    Exit Sub

Erro_NumSerie_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184828)

    End Select

    Exit Sub

End Sub

Private Sub NumSerie_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoOS_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDan_OS As ClassDan_OS

On Error GoTo Erro_objEventoOS_evSelecao

    Set objDan_OS = obj1

    'Mostra os dados do Dan_OS na tela
    lErro = Traz_Dan_OS_Tela(objDan_OS)
    If lErro <> SUCESSO Then gError 184829

    Me.Show

    Exit Sub

Erro_objEventoOS_evSelecao:

    Select Case gErr

        Case 184829


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184830)

    End Select

    Exit Sub

End Sub

Private Sub LabelOS_Click()

Dim lErro As Long
Dim objDan_OS As New ClassDan_OS
Dim colSelecao As New Collection

On Error GoTo Erro_LabelOS_Click

    'Verifica se o OS foi preenchido
    If Len(Trim(OS.Text)) <> 0 Then

        objDan_OS.sOS = OS.Text

    End If

    Call Chama_Tela("Dan_OSCliLista", colSelecao, objDan_OS, objEventoOS)

    Exit Sub

Erro_LabelOS_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184831)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descricao")
    objGrid.colColuna.Add ("Quantidade")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (Descricao.Name)
    objGrid.colCampo.Add (Quantidade.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_Quantidade_Col = 3

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 200 + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_LostFocus()
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Descricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Descricao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade
    
    'Se o campo foi preenchido
    If Len(Trim(Quantidade.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 198250
        
        Quantidade.Text = Formata_Estoque(StrParaDbl(Quantidade.Text))
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184836

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 184836, 198250
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184837)
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

        'GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Produto_Col

                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 184838

                Case iGrid_Quantidade_Col

                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 184840

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 184841

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 184838 To 184840

        Case 184841
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184842)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Guardo o valor do Codigo do Item
    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 198272
    
    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name
    
        Case Produto.Name
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Quantidade.Name
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        Case Else
            objControl.Enabled = False

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 198273

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198273)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 198248

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 198248

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198249)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCodProduto As String
Dim iLinha As Integer
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto
                
    sCodProduto = Produto.Text
        
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 198257
    
    'Se o campo foi preenchido
    If Len(sProdutoFormatado) > 0 Then

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 198258
                
        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
                
        'Verifica se há algum produto repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridItens.Row Then
                                                    
                If GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = sProdutoMascarado Then
                    Produto.PromptInclude = False
                    Produto.Text = ""
                    Produto.PromptInclude = True
                    gError 198259
                    
                End If
                    
            End If
                           
        Next
        
        Set objProdutos = New ClassProduto

        objProdutos.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 198260
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProdutos.sDescricao

        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 198261

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 198259
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", gErr, sProdutoMascarado, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 198257, 198258, 198260, 198261
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 198262)

    End Select

    Exit Function

End Function

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sProduto1 As String

On Error GoTo Erro_LabelProduto_Click
    
    If Me.ActiveControl Is Produto Then
        sProduto1 = Produto.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 198268

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 198269
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr
        
        Case 198268
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 198269

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198270)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim sUnidadeMed As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
        
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 198263

    'Verifica se há algum produto repetido no grid
    For iLinha = 1 To objGridItens.iLinhasExistentes
        
        If iLinha <> GridItens.Row Then
                                                
            If GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = sProdutoMascarado Then
                Produto.PromptInclude = False
                Produto.Text = ""
                Produto.PromptInclude = True
                gError 198264
                
            End If
                
        End If
                       
    Next
    
    Produto.PromptInclude = False
    Produto.Text = sProdutoMascarado
    Produto.PromptInclude = True

    If Not (Me.ActiveControl Is Produto) Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = sProdutoMascarado
        GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If
        
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 198263
        
        Case 198264
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", gErr, sProdutoMascarado, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198265)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is OS Then Call LabelOS_Click
    
        If Me.ActiveControl Is Produto Then Call BotaoProdutos_Click
    
        If Me.ActiveControl Is Cliente Then Call LabelCliente_Click
        
    End If
    
End Sub
