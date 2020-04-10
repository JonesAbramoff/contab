VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl Colecao 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8100
   Begin VB.CommandButton BotaoPintura 
      Caption         =   "Pinturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2220
      TabIndex        =   17
      Top             =   5400
      Width           =   1680
   End
   Begin VB.CommandButton BotaoCor 
      Caption         =   "Cores\Variações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   16
      Top             =   5400
      Width           =   1680
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5820
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "Colecao.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "Colecao.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "Colecao.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "Colecao.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1230
      TabIndex        =   6
      Top             =   135
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   5
      Format          =   "00000"
      Mask            =   "#####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1230
      TabIndex        =   8
      Top             =   645
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   25
      PromptChar      =   " "
   End
   Begin VB.Frame FrameGridItens 
      Caption         =   "Itens"
      Height          =   4275
      Left            =   210
      TabIndex        =   5
      Top             =   1065
      Width           =   7725
      Begin MSMask.MaskEdBox DescPintura 
         Height          =   240
         Left            =   105
         TabIndex        =   15
         Top             =   345
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DescCor 
         Height          =   240
         Left            =   4485
         TabIndex        =   14
         Top             =   1935
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cor 
         Height          =   240
         Left            =   1695
         TabIndex        =   10
         Top             =   1140
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   2
         Format          =   "00"
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Variacao 
         Height          =   240
         Left            =   1995
         TabIndex        =   11
         Top             =   1890
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   3
         Format          =   "000"
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Pintura 
         Height          =   240
         Left            =   2460
         TabIndex        =   12
         Top             =   480
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   2
         Format          =   "00"
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   3885
         Left            =   255
         TabIndex        =   13
         Top             =   240
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   6853
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
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
      Left            =   15
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   165
      Width           =   1155
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
      Left            =   15
      TabIndex        =   9
      Top             =   690
      Width           =   1155
   End
End
Attribute VB_Name = "Colecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const NUM_MAX_ITENS_COLECAO = 999

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGridItens As AdmGrid
Dim iGrid_Cor_Col As Integer
Dim iGrid_Variacao_Col As Integer
Dim iGrid_DescCor_Col As Integer
Dim iGrid_Pintura_Col As Integer
Dim iGrid_DescPintura_Col As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCor As AdmEvento
Attribute objEventoCor.VB_VarHelpID = -1
Private WithEvents objEventoPintura As AdmEvento
Attribute objEventoPintura.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Coleções"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Colecao"

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

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_UnLoad

    Set objGridItens = Nothing

    Set objEventoCodigo = Nothing
    Set objEventoCor = Nothing
    Set objEventoPintura = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187286)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoCor = New AdmEvento
    Set objEventoPintura = New AdmEvento

    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 187287

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 187287

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187288)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objColecao As ClassColecao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objColecao Is Nothing) Then

        lErro = Traz_Colecao_Tela(objColecao)
        If lErro <> SUCESSO Then gError 187289

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 187289

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187290)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objColecao As ClassColecao) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objColecao.lCodigo = StrParaLong(Codigo.Text)
    objColecao.sDescricao = Descricao.Text
    
    lErro = Move_GridItens_Memoria(objColecao)
    If lErro <> SUCESSO Then gError 187374

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 187374

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187291)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objColecao As New ClassColecao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Colecao"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objColecao)
    If lErro <> SUCESSO Then gError 187292

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objColecao.lCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 187292

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187293)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objColecao As New ClassColecao

On Error GoTo Erro_Tela_Preenche

    objColecao.lCodigo = colCampoValor.Item("Codigo").vValor

    If objColecao.lCodigo <> 0 Then
        lErro = Traz_Colecao_Tela(objColecao)
        If lErro <> SUCESSO Then gError 187294
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 187294

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187295)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objColecao As New ClassColecao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 187296
    If Len(Trim(Descricao.Text)) = 0 Then gError 187366
    '#####################

    'Preenche o objColecao
    lErro = Move_Tela_Memoria(objColecao)
    If lErro <> SUCESSO Then gError 187297
    
    If objColecao.colItens.Count = 0 Then gError 187384

    lErro = Trata_Alteracao(objColecao, objColecao.lCodigo)
    If lErro <> SUCESSO Then gError 187298

    'Grava o/a Colecao no Banco de Dados
    lErro = CF("Colecao_Grava", objColecao)
    If lErro <> SUCESSO Then gError 187299

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 187296
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COLECAO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 187297, 187298, 187299

        Case 187366
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            Descricao.SetFocus
            
        Case 187384
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_ITENS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187300)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Colecao() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Colecao

    Call Grid_Limpa(objGridItens)

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_Colecao = SUCESSO

    Exit Function

Erro_Limpa_Tela_Colecao:

    Limpa_Tela_Colecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187301)

    End Select

    Exit Function

End Function

Function Traz_Colecao_Tela(objColecao As ClassColecao) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Colecao_Tela

    Call Limpa_Tela_Colecao

    'Lê o Colecao que está sendo Passado
    lErro = CF("Colecao_Le", objColecao)
    If lErro <> SUCESSO And lErro <> 187267 Then gError 187302

    If objColecao.lCodigo <> 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = CStr(objColecao.lCodigo)
        Codigo.PromptInclude = True
    End If
    
    If lErro = SUCESSO Then

        Descricao.Text = objColecao.sDescricao
        
        lErro = Preenche_GridItens_Tela(objColecao)
        If lErro <> SUCESSO Then gError 187367

    End If

    iAlterado = 0

    Traz_Colecao_Tela = SUCESSO

    Exit Function

Erro_Traz_Colecao_Tela:

    Traz_Colecao_Tela = gErr

    Select Case gErr

        Case 187302, 187367

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187303)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 187304

    'Limpa Tela
    Call Limpa_Tela_Colecao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 187304

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187305)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187306)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 187307

    Call Limpa_Tela_Colecao

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 187307

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187308)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objColecao As New ClassColecao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 187309
    '#####################

    objColecao.lCodigo = StrParaLong(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_COLECAO", objColecao.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("Colecao_Exclui", objColecao)
        If lErro <> SUCESSO Then gError 187310

        'Limpa Tela
        Call Limpa_Tela_Colecao

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 187309
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COLECAO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 187310

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187311)

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
       If lErro <> SUCESSO Then gError 187312

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 187312

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187376)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187375)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objColecao As ClassColecao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objColecao = obj1

    'Mostra os dados do Colecao na tela
    lErro = Traz_Colecao_Tela(objColecao)
    If lErro <> SUCESSO Then gError 187313

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 187313


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187374)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objColecao As New ClassColecao
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objColecao.lCodigo = Codigo.Text

    End If

    Call Chama_Tela("ColecaoLista", colSelecao, objColecao, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187373)

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
    objGrid.colColuna.Add ("Cor")
    objGrid.colColuna.Add ("Variação")
    objGrid.colColuna.Add ("Desc. Cor\Variação")
    objGrid.colColuna.Add ("Pintura")
    objGrid.colColuna.Add ("Descrição Pintura")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Cor.Name)
    objGrid.colCampo.Add (Variacao.Name)
    objGrid.colCampo.Add (DescCor.Name)
    objGrid.colCampo.Add (Pintura.Name)
    objGrid.colCampo.Add (DescPintura.Name)

    'Colunas do Grid
    iGrid_Cor_Col = 1
    iGrid_Variacao_Col = 2
    iGrid_DescCor_Col = 3
    iGrid_Pintura_Col = 4
    iGrid_DescPintura_Col = 5

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_COLECAO + 1

    objGrid.iLinhasVisiveis = 14

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

Private Sub Cor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Cor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Cor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Cor
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Variacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Variacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Variacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Variacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Variacao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Pintura_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Pintura_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Pintura_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Pintura_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Pintura
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Cor(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCor As New ClassCorVariacao

On Error GoTo Erro_Saida_Celula_Cor

    Set objGridInt.objControle = Cor
    
    If Len(Trim(Cor.Text)) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Variacao_Col))) > 0 Then

        objCor.iCor = StrParaInt(Cor.Text)
        objCor.iVariacao = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_Variacao_Col))

        'Lê o CorVariacao que está sendo Passado
        lErro = CF("CorVariacao_Le", objCor)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187381
    
        If lErro <> SUCESSO Then gError 187382
        
        GridItens.TextMatrix(GridItens.Row, iGrid_DescCor_Col) = objCor.sDescricao
        
        If (GridItens.Row - GridItens.FixedRows) = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else
    
        GridItens.TextMatrix(GridItens.Row, iGrid_DescCor_Col) = ""
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187314

    Saida_Celula_Cor = SUCESSO

    Exit Function

Erro_Saida_Celula_Cor:

    Saida_Celula_Cor = gErr

    Select Case gErr

        Case 187314, 187381
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 187382
            Call Rotina_Erro(vbOKOnly, "ERRO_CORVARIACAO_NAO_CADASTRADO", gErr, objCor.iCor, objCor.iVariacao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187372)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_Variacao(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCor As New ClassCorVariacao

On Error GoTo Erro_Saida_Celula_Variacao

    Set objGridInt.objControle = Variacao

    If Len(Trim(Variacao.Text)) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Cor_Col))) > 0 Then

        objCor.iVariacao = StrParaInt(Variacao.Text)
        objCor.iCor = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_Cor_Col))

        'Lê o CorVariacao que está sendo Passado
        lErro = CF("CorVariacao_Le", objCor)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187383
    
        If lErro <> SUCESSO Then gError 187384
        
        GridItens.TextMatrix(GridItens.Row, iGrid_DescCor_Col) = objCor.sDescricao
        
        If (GridItens.Row - GridItens.FixedRows) = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else
    
        GridItens.TextMatrix(GridItens.Row, iGrid_DescCor_Col) = ""
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187315

    Saida_Celula_Variacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Variacao:

    Saida_Celula_Variacao = gErr

    Select Case gErr

        Case 187315, 187383
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 187384
            Call Rotina_Erro(vbOKOnly, "ERRO_CORVARIACAO_NAO_CADASTRADO", gErr, objCor.iCor, objCor.iVariacao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187371)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_Pintura(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objPintura As New ClassPintura

On Error GoTo Erro_Saida_Celula_Pintura

    Set objGridInt.objControle = Pintura
    
    If Len(Trim(Pintura.Text)) > 0 Then

        objPintura.iCodigo = StrParaInt(Pintura.Text)
    
        'Lê o Pintura que está sendo Passado
        lErro = CF("Pintura_Le", objPintura)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187379
    
        If lErro <> SUCESSO Then gError 187380
            
        GridItens.TextMatrix(GridItens.Row, iGrid_DescPintura_Col) = objPintura.sDescricao
            
        If (GridItens.Row - GridItens.FixedRows) = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else
        
        GridItens.TextMatrix(GridItens.Row, iGrid_DescPintura_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187316

    Saida_Celula_Pintura = SUCESSO

    Exit Function

Erro_Saida_Celula_Pintura:

    Saida_Celula_Pintura = gErr

    Select Case gErr

        Case 187316, 187379
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 187380
            Call Rotina_Erro(vbOKOnly, "ERRO_PINTURA_NAO_CADASTRADO", gErr, objPintura.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187370)
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


                Case iGrid_Cor_Col

                    lErro = Saida_Celula_Cor(objGridInt)
                    If lErro <> SUCESSO Then gError 187317

                Case iGrid_Variacao_Col

                    lErro = Saida_Celula_Variacao(objGridInt)
                    If lErro <> SUCESSO Then gError 187318

                Case iGrid_Pintura_Col

                    lErro = Saida_Celula_Pintura(objGridInt)
                    If lErro <> SUCESSO Then gError 187319

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 187320

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 187317 To 187319

        Case 187320
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187369)

    End Select

    Exit Function

End Function

Function Preenche_GridItens_Tela(ByVal objColecao As ClassColecao) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objItensColecao As ClassItensColecao
Dim objPintura As New ClassPintura
Dim objCor As New ClassCorVariacao

On Error GoTo Erro_Preenche_GridItens_Tela

    If objColecao.colItens.Count = 0 Then
        lErro = CF("Colecao_Le_Itens", objColecao)
        If lErro <> SUCESSO Then gError 187388
    End If

    Call Grid_Limpa(objGridItens)

    iLinha = 0
    For Each objItensColecao In objColecao.colItens
    
        iLinha = iLinha + 1
    
        objPintura.iCodigo = objItensColecao.iPintura
    
        'Lê o Pintura que está sendo Passado
        lErro = CF("Pintura_Le", objPintura)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187375
    
        If lErro <> SUCESSO Then gError 187376
    
        objCor.iCor = objItensColecao.iCor
        objCor.iVariacao = objItensColecao.iVariacao

        'Lê o CorVariacao que está sendo Passado
        lErro = CF("CorVariacao_Le", objCor)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187377
    
        If lErro <> SUCESSO Then gError 187378
        
        GridItens.TextMatrix(iLinha, iGrid_Cor_Col) = Format(objCor.iCor, "00")
        GridItens.TextMatrix(iLinha, iGrid_DescCor_Col) = objCor.sDescricao
        GridItens.TextMatrix(iLinha, iGrid_DescPintura_Col) = objPintura.sDescricao
        GridItens.TextMatrix(iLinha, iGrid_Pintura_Col) = Format(objPintura.iCodigo, "00")
        GridItens.TextMatrix(iLinha, iGrid_Variacao_Col) = Format(objCor.iVariacao, "000")
        
    Next
    
    objGridItens.iLinhasExistentes = objColecao.colItens.Count
 
    Preenche_GridItens_Tela = SUCESSO

    Exit Function

Erro_Preenche_GridItens_Tela:

    Preenche_GridItens_Tela = gErr

    Select Case gErr
    
        Case 187375, 187377, 187388
        
        Case 187376
            Call Rotina_Erro(vbOKOnly, "ERRO_PINTURA_NAO_CADASTRADO", gErr, objPintura.iCodigo)
        
        Case 187378
            Call Rotina_Erro(vbOKOnly, "ERRO_CORVARIACAO_NAO_CADASTRADO", gErr, objItensColecao.iCor, objItensColecao.iVariacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187321)

    End Select

    Exit Function

End Function

Function Move_GridItens_Memoria(ByVal objColecao As ClassColecao) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndiceAux As Integer
Dim objItensColecao As ClassItensColecao
Dim objItensColecaoAux As ClassItensColecao

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objItensColecao = New ClassItensColecao
    
        objItensColecao.iSeq = iIndice
        objItensColecao.iCor = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Cor_Col))
        objItensColecao.iPintura = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Pintura_Col))
        objItensColecao.iVariacao = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Variacao_Col))
        objItensColecao.lColecao = objColecao.lCodigo
        
        If objItensColecao.iPintura = 0 Then gError 187385
        If objItensColecao.iVariacao = 0 Then gError 187386
        If objItensColecao.iCor = 0 Then gError 187387
        
        objColecao.colItens.Add objItensColecao

    Next
    
    iIndice = 0
    For Each objItensColecao In objColecao.colItens
        iIndice = iIndice + 1
        iIndiceAux = 0
        For Each objItensColecaoAux In objColecao.colItens
            iIndiceAux = iIndiceAux + 1
            If iIndiceAux <> iIndice Then
                If objItensColecao.iCor = objItensColecaoAux.iCor And _
                objItensColecao.iPintura = objItensColecaoAux.iPintura And _
                objItensColecao.iVariacao = objItensColecaoAux.iVariacao Then gError 187391
            End If
        Next
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr
    
        Case 187385
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PINTURA_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case 187386
            Call Rotina_Erro(vbOKOnly, "ERRO_VARIACAO_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 187387
            Call Rotina_Erro(vbOKOnly, "ERRO_COR_NAO_PREENCHIDA_GRID", gErr, iIndice)
            
        Case 187391
            Call Rotina_Erro(vbOKOnly, "ERRO_ITENS_REPETIDOS_GRID", gErr, iIndice, iIndiceAux)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187322)

    End Select

    Exit Function

End Function

Private Sub BotaoCor_Click()

Dim lErro As Long
Dim objCorVariacao As New ClassCorVariacao
Dim colSelecao As New Collection
Dim sOrdem  As String

On Error GoTo Erro_BotaoCor_Click

    If Me.ActiveControl Is Cor Then
        objCorVariacao.iCor = StrParaInt(Cor.Text)
        sOrdem = "Cor"
    ElseIf Me.ActiveControl Is Variacao Then
        objCorVariacao.iVariacao = StrParaInt(Variacao.Text)
        sOrdem = "Variação"
    Else
            'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 187371
        
        objCorVariacao.iCor = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_Cor_Col))
        objCorVariacao.iVariacao = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_Variacao_Col))
        
        sOrdem = "Cor\Variação"
        
    End If

    Call Chama_Tela("CorVariacaoLista", colSelecao, objCorVariacao, objEventoCor, "", sOrdem)

    Exit Sub

Erro_BotaoCor_Click:

    Select Case gErr
    
        Case 187371
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187372)

    End Select

    Exit Sub

End Sub

Private Sub BotaoPintura_Click()

Dim lErro As Long
Dim objPintura As New ClassPintura
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPintura_Click

    If Me.ActiveControl Is Pintura Then
        objPintura.iCodigo = StrParaInt(Pintura.Text)
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 187373
        
        objPintura.iCodigo = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_Pintura_Col))
       
    End If


    Call Chama_Tela("PinturaLista", colSelecao, objPintura, objEventoPintura)

    Exit Sub

Erro_BotaoPintura_Click:

    Select Case gErr

        Case 187373
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187374)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
        If Me.ActiveControl Is Cor Then Call BotaoCor_Click
        If Me.ActiveControl Is Variacao Then Call BotaoCor_Click
        If Me.ActiveControl Is Pintura Then Call BotaoPintura_Click
    
    End If
    
End Sub

Private Sub objEventoCor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim iLinha As Integer
Dim objCor As ClassCorVariacao

On Error GoTo Erro_objEventoCor_evSelecao

    Set objCor = obj1
    
    Cor.PromptInclude = False
    Cor.Text = CStr(objCor.iCor)
    Cor.PromptInclude = True
    
    Variacao.PromptInclude = False
    Variacao.Text = CStr(objCor.iVariacao)
    Variacao.PromptInclude = True

    If Not (Me.ActiveControl Is Cor) Then
        GridItens.TextMatrix(GridItens.Row, iGrid_Cor_Col) = Format(objCor.iCor, Cor.Format)
    End If
    If Not (Me.ActiveControl Is Variacao) Then
        GridItens.TextMatrix(GridItens.Row, iGrid_Variacao_Col) = Format(objCor.iVariacao, Variacao.Format)
    End If
    
    If ((Not (Me.ActiveControl Is Cor)) And (Not (Me.ActiveControl Is Variacao))) Then
        GridItens.TextMatrix(GridItens.Row, iGrid_DescCor_Col) = objCor.sDescricao
    End If
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoCor_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187389)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPintura_evSelecao(obj1 As Object)

Dim lErro As Long
Dim iLinha As Integer
Dim objPintura As ClassPintura

On Error GoTo Erro_objEventoPintura_evSelecao

    Set objPintura = obj1

    Pintura.PromptInclude = False
    Pintura.Text = CStr(objPintura.iCodigo)
    Pintura.PromptInclude = True

    If Not (Me.ActiveControl Is Pintura) Then
        GridItens.TextMatrix(GridItens.Row, iGrid_DescPintura_Col) = objPintura.sDescricao
        GridItens.TextMatrix(GridItens.Row, iGrid_Pintura_Col) = Format(objPintura.iCodigo, Pintura.Format)
    End If
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoPintura_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187390)

    End Select

    Exit Sub

End Sub

