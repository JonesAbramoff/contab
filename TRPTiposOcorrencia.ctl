VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRPTiposOcorrencia 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame6 
      Caption         =   "Incide Sobre"
      Height          =   450
      Left            =   405
      TabIndex        =   26
      Top             =   1800
      Width           =   5535
      Begin VB.OptionButton IncideFAT 
         Caption         =   "Faturável s\ CMA"
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
         Left            =   3585
         TabIndex        =   27
         Top             =   180
         Width           =   1845
      End
      Begin VB.OptionButton IncideCMA 
         Caption         =   "CMA"
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
         Left            =   2460
         TabIndex        =   4
         Top             =   180
         Width           =   1785
      End
      Begin VB.OptionButton IncideBruto 
         Caption         =   "Valor Bruto"
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
         Left            =   780
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.Frame FrameCE 
      Caption         =   "Comissão Externa"
      Height          =   1755
      Left            =   405
      TabIndex        =   24
      Top             =   2805
      Width           =   5535
      Begin VB.CheckBox ConsideraComisExt 
         Caption         =   "Deve alterar os valor das comissões externas"
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
         Left            =   150
         TabIndex        =   6
         Top             =   225
         Width           =   4725
      End
      Begin VB.Frame FrameQuais 
         Caption         =   "Quais ?"
         Height          =   1230
         Left            =   795
         TabIndex        =   25
         Top             =   465
         Width           =   4380
         Begin VB.ListBox AlteraComiVou 
            Height          =   960
            ItemData        =   "TRPTiposOcorrencia.ctx":0000
            Left            =   300
            List            =   "TRPTiposOcorrencia.ctx":0010
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   210
            Width           =   3810
         End
      End
   End
   Begin VB.Frame FrameCI 
      Caption         =   "Comissão Interna"
      Height          =   525
      Left            =   405
      TabIndex        =   23
      Top             =   2250
      Width           =   5535
      Begin VB.CheckBox ConsideraComisInt 
         Caption         =   "É considerado no comissionamento interno"
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
         Left            =   120
         TabIndex        =   5
         Top             =   195
         Width           =   4380
      End
   End
   Begin VB.Frame FrameA 
      Caption         =   "Investimentos"
      Height          =   510
      Left            =   405
      TabIndex        =   22
      Top             =   4560
      Width           =   5535
      Begin VB.CheckBox EstornaAporteVou 
         Caption         =   "Estorna integralmente ou parcialmente aportes do voucher"
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
         Left            =   135
         TabIndex        =   8
         Top             =   195
         Width           =   5370
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Restrições"
      Height          =   780
      Left            =   405
      TabIndex        =   21
      Top             =   5070
      Width           =   5535
      Begin VB.CheckBox AceitaVlrPositivo 
         Caption         =   "O valor da ocorrência pode ser positivo"
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
         TabIndex        =   9
         Top             =   210
         Width           =   4365
      End
      Begin VB.CheckBox AceitaVlrNegativo 
         Caption         =   "O valor da ocorrência pode ser negativo"
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
         Left            =   135
         TabIndex        =   10
         Top             =   420
         Width           =   4620
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   945
      Left            =   405
      TabIndex        =   18
      Top             =   855
      Width           =   5535
      Begin VB.TextBox Descricao 
         Height          =   315
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   2
         Top             =   540
         Width           =   3930
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2085
         Picture         =   "TRPTiposOcorrencia.ctx":0089
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   210
         Width           =   300
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1215
         TabIndex        =   0
         Top             =   180
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         PromptChar      =   " "
      End
      Begin VB.Label LabelDescricao 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   135
         TabIndex        =   20
         Top             =   570
         Width           =   1020
      End
      Begin VB.Label LabelCodigo 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   210
         Width           =   1020
      End
   End
   Begin VB.ListBox Tipos 
      Height          =   4935
      Left            =   6180
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   930
      Width           =   3210
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7290
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   75
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TRPTiposOcorrencia.ctx":0173
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "TRPTiposOcorrencia.ctx":02CD
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "TRPTiposOcorrencia.ctx":0457
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "TRPTiposOcorrencia.ctx":0989
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.Label Label13 
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
      Height          =   210
      Left            =   6180
      TabIndex        =   17
      Top             =   705
      Width           =   2190
   End
End
Attribute VB_Name = "TRPTiposOcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1


Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tipos de ocorrências"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRPTiposOcorrencia"

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

Private Sub ConsideraComisExt_Click()

    If ConsideraComisExt.Value = vbChecked Then
        FrameQuais.Enabled = True
    Else
        FrameQuais.Enabled = False
        AlteraComiVou.Selected(0) = False
        AlteraComiVou.Selected(1) = False
        AlteraComiVou.Selected(2) = False
        AlteraComiVou.Selected(3) = False
        'AlteraComiVou.Selected(4) = False
    End If

End Sub

Private Sub IncideBruto_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Boqueia_Funcionalidades
End Sub

Private Sub IncideCMA_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Boqueia_Funcionalidades
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
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

    Set objEventoCodigo = Nothing
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198028)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As AdmColCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento

    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê o Código e a Descrição de cada Tipo de Mão-de-Obra
    lErro = CF("Cod_Nomes_Le", "TRPTiposOcorrencia", "Codigo", "Descricao", STRING_TRP_TIPOOCR_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 137558

    'preenche a ListBox Tipos com os objetos da colecao
    For Each objCodigoDescricao In colCodigoDescricao
        Tipos.AddItem objCodigoDescricao.sNome
        Tipos.ItemData(Tipos.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 137558

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198029)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTipo As ClassTRPTiposOcorrencia) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTipo Is Nothing) Then

        lErro = Traz_TRPTiposOcorrencia_Tela(objTipo)
        If lErro <> SUCESSO Then gError 198030

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 198030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198031)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objTipo As ClassTRPTiposOcorrencia) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objTipo.iCodigo = StrParaInt(Codigo.Text)
    objTipo.sDescricao = Descricao.Text
    
    
    objTipo.iEstornaAporteVou = Conv_Check_em_Integer_Marcado(EstornaAporteVou.Value)
    objTipo.iConsideraComisInt = Conv_Check_em_Integer_Marcado(ConsideraComisInt.Value)
    objTipo.iAlteraComiVou = Conv_Check_em_Integer_Marcado(ConsideraComisExt.Value)
    objTipo.iAceitaVlrPositivo = Conv_Check_em_Integer_Marcado(AceitaVlrPositivo.Value)
    objTipo.iAceitaVlrNegativo = Conv_Check_em_Integer_Marcado(AceitaVlrNegativo.Value)
    
    objTipo.iAlteraCMCC = Conv_Boolean_em_Integer_Marcado(AlteraComiVou.Selected(0))
    objTipo.iAlteraCMC = Conv_Boolean_em_Integer_Marcado(AlteraComiVou.Selected(1))
    objTipo.iAlteraCMR = Conv_Boolean_em_Integer_Marcado(AlteraComiVou.Selected(2))
    objTipo.iAlteraOVER = Conv_Boolean_em_Integer_Marcado(AlteraComiVou.Selected(3))
    'objTipo.iAlteraCMA = Conv_Boolean_em_Integer_Marcado(AlteraComiVou.Selected(4))
    
    If IncideBruto.Value Then
        objTipo.iIncideSobre = TRP_TIPO_OCR_INCIDE_BRUTO
    ElseIf IncideFAT.Value Then
        objTipo.iIncideSobre = TRP_TIPO_OCR_INCIDE_FAT
    Else
        objTipo.iIncideSobre = TRP_TIPO_OCR_INCIDE_CMA
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198032)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objTipo As New ClassTRPTiposOcorrencia

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRPTiposOcorrencia"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objTipo)
    If lErro <> SUCESSO Then gError 198033

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTipo.iCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 198033

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198034)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTipo As New ClassTRPTiposOcorrencia

On Error GoTo Erro_Tela_Preenche

    objTipo.iCodigo = colCampoValor.Item("Codigo").vValor

    If objTipo.iCodigo <> 0 Then

        lErro = Traz_TRPTiposOcorrencia_Tela(objTipo)
        If lErro <> SUCESSO Then gError 198035

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 198035

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198036)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipo As New ClassTRPTiposOcorrencia

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 198037
    '#####################

    'Preenche o objTipo
    lErro = Move_Tela_Memoria(objTipo)
    If lErro <> SUCESSO Then gError 198038

    lErro = Trata_Alteracao(objTipo, objTipo.iCodigo)
    If lErro <> SUCESSO Then gError 198039

    'Grava o/a TRPTiposOcorrencia no Banco de Dados
    lErro = CF("TRPTiposOcorrencia_Grava", objTipo)
    If lErro <> SUCESSO Then gError 198040

    'Remove o item da lista de Tipos
    Call Tipos_Exclui(objTipo.iCodigo)

    'Insere o item na lista de Tipos
    Call Tipos_Adiciona(objTipo)
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198037
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRPTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 198038, 198039, 198040

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198041)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TRPTiposOcorrencia() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRPTiposOcorrencia

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    
    EstornaAporteVou.Value = vbUnchecked
    ConsideraComisInt.Value = vbUnchecked
    ConsideraComisExt.Value = vbUnchecked
    AceitaVlrPositivo.Value = vbUnchecked
    AceitaVlrNegativo.Value = vbUnchecked
        
    AlteraComiVou.Selected(0) = False
    AlteraComiVou.Selected(1) = False
    AlteraComiVou.Selected(2) = False
    AlteraComiVou.Selected(3) = False
    'AlteraComiVou.Selected(4) = False
    
    IncideBruto.Value = True

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_TRPTiposOcorrencia = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRPTiposOcorrencia:

    Limpa_Tela_TRPTiposOcorrencia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198042)

    End Select

    Exit Function

End Function

Function Traz_TRPTiposOcorrencia_Tela(objTipo As ClassTRPTiposOcorrencia) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_TRPTiposOcorrencia_Tela
    Call Limpa_Tela_TRPTiposOcorrencia

        If objTipo.iCodigo <> 0 Then
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objTipo.iCodigo)
            Codigo.PromptInclude = True
        End If


    'Lê o TRPTiposOcorrencia que está sendo Passado
    lErro = CF("TRPTiposOcorrencia_Le", objTipo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198043

    If lErro = SUCESSO Then


        If objTipo.iCodigo <> 0 Then
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objTipo.iCodigo)
            Codigo.PromptInclude = True
        End If

        Descricao.Text = objTipo.sDescricao

        EstornaAporteVou.Value = Traz_Tela_Marcado_Para_Check(objTipo.iEstornaAporteVou)
        ConsideraComisInt.Value = Traz_Tela_Marcado_Para_Check(objTipo.iConsideraComisInt)
        ConsideraComisExt.Value = Traz_Tela_Marcado_Para_Check(objTipo.iAlteraComiVou)
        AceitaVlrPositivo.Value = Traz_Tela_Marcado_Para_Check(objTipo.iAceitaVlrPositivo)
        AceitaVlrNegativo.Value = Traz_Tela_Marcado_Para_Check(objTipo.iAceitaVlrNegativo)

        AlteraComiVou.Selected(0) = Traz_Tela_Marcado_Para_Boolean(objTipo.iAlteraCMCC)
        AlteraComiVou.Selected(1) = Traz_Tela_Marcado_Para_Boolean(objTipo.iAlteraCMC)
        AlteraComiVou.Selected(2) = Traz_Tela_Marcado_Para_Boolean(objTipo.iAlteraCMR)
        AlteraComiVou.Selected(3) = Traz_Tela_Marcado_Para_Boolean(objTipo.iAlteraOVER)
        'AlteraComiVou.Selected(4) = Traz_Tela_Marcado_Para_Boolean(objTipo.iAlteraCMA)
        
        If objTipo.iIncideSobre = TRP_TIPO_OCR_INCIDE_BRUTO Then
            IncideBruto.Value = True
        ElseIf objTipo.iIncideSobre = TRP_TIPO_OCR_INCIDE_FAT Then
            IncideFAT.Value = True
        Else
            IncideCMA.Value = True
        End If
        
        Call Boqueia_Funcionalidades

    End If

    iAlterado = 0

    Traz_TRPTiposOcorrencia_Tela = SUCESSO

    Exit Function

Erro_Traz_TRPTiposOcorrencia_Tela:

    Traz_TRPTiposOcorrencia_Tela = gErr

    Select Case gErr

        Case 198043

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198044)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 198045

    'Limpa Tela
    Call Limpa_Tela_TRPTiposOcorrencia

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 198045

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198046)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198047)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 198048

    Call Limpa_Tela_TRPTiposOcorrencia

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 198048

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198049)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTipo As New ClassTRPTiposOcorrencia
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 198050
    '#####################

    objTipo.iCodigo = StrParaInt(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TRPTIPOSOCORRENCIA", objTipo.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("TRPTiposOcorrencia_Exclui", objTipo)
        If lErro <> SUCESSO Then gError 198051

        Call Tipos_Exclui(objTipo.iCodigo)

        'Limpa Tela
        Call Limpa_Tela_TRPTiposOcorrencia

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198050
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRPTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 198051

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198052)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Long_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError 198053

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 198053

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198054)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Verifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then

       '#######################################
       'CRITICA Descricao
       '#######################################

    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198055)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EstornaAporteVou_click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AlteraComisExt_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ConsideraComisInt_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AlteraComiVou_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AceitaVlrPositivo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AceitaVlrNegativo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipo As ClassTRPTiposOcorrencia

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTipo = obj1

    'Mostra os dados do TRPTiposOcorrencia na tela
    lErro = Traz_TRPTiposOcorrencia_Tela(objTipo)
    If lErro <> SUCESSO Then gError 198076

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 198076


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198077)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objTipo As New ClassTRPTiposOcorrencia
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objTipo.iCodigo = StrParaInt(Codigo.Text)

    End If

    Call Chama_Tela("TRPTiposOcorrenciaLista", colSelecao, objTipo, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198078)

    End Select

    Exit Sub

End Sub

Private Function Conv_Boolean_em_Integer_Marcado(ByVal bFlag As Boolean) As Integer
    If bFlag Then
        Conv_Boolean_em_Integer_Marcado = MARCADO
    Else
        Conv_Boolean_em_Integer_Marcado = DESMARCADO
    End If
End Function

Private Function Conv_Check_em_Integer_Marcado(ByVal iFlag As Integer) As Integer
    If iFlag = vbChecked Then
        Conv_Check_em_Integer_Marcado = MARCADO
    Else
        Conv_Check_em_Integer_Marcado = DESMARCADO
    End If
End Function

Private Function Traz_Tela_Marcado_Para_Boolean(ByVal iFlag As Integer) As Boolean
    If iFlag = MARCADO Then
        Traz_Tela_Marcado_Para_Boolean = True
    Else
        Traz_Tela_Marcado_Para_Boolean = False
    End If
End Function

Private Function Traz_Tela_Marcado_Para_Check(ByVal iFlag As Integer) As Integer
    If iFlag = MARCADO Then
        Traz_Tela_Marcado_Para_Check = vbChecked
    Else
        Traz_Tela_Marcado_Para_Check = vbUnchecked
    End If
End Function

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("Config_ObterAutomatico", "TRPConfig", "NUM_PROX_TRPTIPOSOCORRENCIA", "TRPTiposOcorrencia", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 197034
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 197034

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197035)
    
    End Select

    Exit Sub
    
End Sub

Public Sub Boqueia_Funcionalidades()

Dim bFlag As Boolean

    If IncideBruto.Value Then
        bFlag = True
    Else
        bFlag = False
        ConsideraComisInt.Value = vbChecked
        ConsideraComisExt.Value = vbUnchecked
        AlteraComiVou.Selected(0) = False
        AlteraComiVou.Selected(1) = False
        AlteraComiVou.Selected(2) = False
        AlteraComiVou.Selected(3) = False
    End If
    
    FrameCI.Enabled = bFlag
    FrameCE.Enabled = bFlag
    FrameA.Enabled = bFlag
    
End Sub

Private Sub Tipos_DblClick()

Dim lErro As Long
Dim objTipo As New ClassTRPTiposOcorrencia

On Error GoTo Erro_Tipos_DblClick

    'Guarda o valor do codigo do Tipo da Mão-de-Obra selecionado na ListBox Tipos
    objTipo.iCodigo = Tipos.ItemData(Tipos.ListIndex)

    'Mostra os dados do TiposDeMaodeObra na tela
    lErro = Traz_TRPTiposOcorrencia_Tela(objTipo)
    If lErro <> SUCESSO Then gError 137557

    Me.Show
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_Tipos_DblClick:

    Tipos.SetFocus

    Select Case gErr

    Case 137557
        'erro tratado na rotina chamada
    
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174962)

    End Select

    Exit Sub

End Sub

Private Sub Tipos_Adiciona(objTipos As ClassTRPTiposOcorrencia)

    Tipos.AddItem objTipos.sDescricao
    Tipos.ItemData(Tipos.NewIndex) = objTipos.iCodigo

End Sub

Private Sub Tipos_Exclui(iCodigo As Integer)

Dim iIndice As Integer

    For iIndice = 0 To Tipos.ListCount - 1

        If Tipos.ItemData(iIndice) = iCodigo Then

            Tipos.RemoveItem iIndice
            Exit For

        End If

    Next

End Sub

