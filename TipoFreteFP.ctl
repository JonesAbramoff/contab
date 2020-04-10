VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TipoFreteFPOcx 
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   KeyPreview      =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   6255
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2160
      Picture         =   "TipoFreteFP.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Numeração Automática"
      Top             =   315
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3990
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoFreteFP.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TipoFreteFP.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoFreteFP.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoFreteFP.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1635
      TabIndex        =   6
      Top             =   795
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1635
      TabIndex        =   7
      Top             =   300
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Preco 
      Height          =   315
      Left            =   1635
      TabIndex        =   8
      Top             =   1860
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1635
      TabIndex        =   9
      Top             =   1335
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   25
      PromptChar      =   " "
   End
   Begin VB.Label DataAtualizacao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1635
      TabIndex        =   15
      Top             =   2370
      Width           =   1695
   End
   Begin VB.Label LabelAtualizado 
      AutoSize        =   -1  'True
      Caption         =   "Atualizado Em:"
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
      Left            =   285
      TabIndex        =   14
      Top             =   2415
      Width           =   1275
   End
   Begin VB.Label LabelPreco 
      AutoSize        =   -1  'True
      Caption         =   "Preço:"
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
      Left            =   1005
      TabIndex        =   13
      Top             =   1905
      Width           =   570
   End
   Begin VB.Label LabelDescricao 
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
      Left            =   630
      TabIndex        =   12
      Top             =   870
      Width           =   930
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   900
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   360
      Width           =   660
   End
   Begin VB.Label LabelNomeReduzido 
      AutoSize        =   -1  'True
      Caption         =   "Nome Reduzido:"
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
      Left            =   150
      TabIndex        =   10
      Top             =   1395
      Width           =   1410
   End
End
Attribute VB_Name = "TipoFreteFPOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer

Private WithEvents objEventoTipoFrete As AdmEvento
Attribute objEventoTipoFrete.VB_VarHelpID = -1

Private Sub BotaoProxNum_Click()
'Gera Numero automatico para código

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("TipoFreteFP_Automatico", iCodigo)
    If lErro <> SUCESSO Then gError 116914

    Codigo.PromptInclude = False
    Codigo.Text = CStr(iCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 116914
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174873)
    
    End Select

    Exit Sub

End Sub

Sub Traz_Tipo_Tela(objTipo As ClassTipoFreteFP)

    'Mostra dados do Tipo de Frete na tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objTipo.iCodigo)
    Codigo.PromptInclude = True
    NomeReduzido.Text = objTipo.sNomeReduzido
    Descricao.Text = objTipo.sDescricao
    Preco.Text = objTipo.dPreco
    DataAtualizacao.Caption = objTipo.dtDataAtualizacao
    
    iAlterado = 0

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objTipo As New ClassTipoFreteFP
Dim colSelecao As Collection

On Error GoTo Erro_LabelCodigo_Click

    objTipo.iCodigo = StrParaInt(Codigo.Text)

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("TipoFreteLista", colSelecao, objTipo, objEventoTipoFrete)

    Exit Sub
    
Erro_LabelCodigo_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174874)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then gError 116915

    End If
        
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 116915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", gErr, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174875)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Aciona Rotinas de exclusão

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim objTipo As New ClassTipoFreteFP

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then gError 116916

    objTipo.iCodigo = CInt(Codigo.Text)
    'Lê o Tipo De Frete
    lErro = CF("TipoFreteFP_Le", objTipo)
    If lErro <> SUCESSO And lErro <> 116913 Then gError 116917

    'Se não achou o Tipo De Frete --> Erro
    If lErro <> SUCESSO Then gError 116918
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TIPOFRETE", objTipo.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Tipo de Frete
        lErro = CF("TipoFreteFP_Exclui", objTipo)
        If lErro <> SUCESSO Then gError 116919

        lErro = Limpa_Tela_Tipo
        If lErro <> SUCESSO Then gError 116920

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 116918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOFRETE_NAO_CADASTRADO", gErr, objTipo.iCodigo, giFilialEmpresa)

        Case 19163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOFRETE_EXCLUSAO", gErr, objTipo.iCodigo)

        Case 116916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 116917, 116919, 116920

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174876)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) > 0 Then

        'Verifica se codigo é numérico
        If Not IsNumeric(Codigo.Text) Then gError 116921

        'Verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then gError 116922

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
        
        Case 116921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", gErr, Codigo.Text)

        Case 116922
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", gErr, Codigo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174877)

    End Select

    Exit Sub
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoTipoFrete = Nothing
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub


Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'Aciona Rotinas de gravação

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 116923

    lErro = Limpa_Tela_Tipo
    If lErro <> SUCESSO Then gError 116924

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116923, 116924

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174878)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 116925

    lErro = Limpa_Tela_Tipo
    If lErro <> SUCESSO Then gError 116926

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 116925, 116926

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174879)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoTipoFrete = New AdmEvento
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174880)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objTipo As ClassTipoFreteFP) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de TipoFreteFP

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objTipo Is Nothing) Then

        lErro = CF("TipoFreteFP_Le", objTipo)
        If lErro <> SUCESSO And lErro <> 116913 Then gError 116927

        If lErro = SUCESSO Then

            Call Traz_Tipo_Tela(objTipo)
        
        Else
            Codigo.PromptInclude = False
            Codigo.Text = objTipo.iCodigo
            Codigo.PromptInclude = True
            
        End If
                    
    Else

        'Limpa a Tela
        lErro = Limpa_Tela_Tipo
        If lErro <> SUCESSO Then gError 116928

    End If
    
    iAlterado = 0

   Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 116927, 116928

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174881)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Function Gravar_Registro() As Long
'Grava os registros

Dim lErro As Long
Dim iIndice As Integer
Dim objTipo As New ClassTipoFreteFP

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then gError 116929

    'verifica preenchimento do nome
    If Len(Trim(Descricao.Text)) = 0 Then gError 116930

    'verifica preenchimento do nome reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 116931

    'verifica preenchimento do campo preço
    If Len(Trim(Preco.ClipText)) = 0 Then gError 116932
    
    'preenche objtipo
    objTipo.iCodigo = CInt(Codigo.Text)
    objTipo.sDescricao = Descricao.Text
    objTipo.sNomeReduzido = NomeReduzido.Text
    objTipo.dPreco = StrParaDbl(Preco.Text)
    objTipo.iFilialEmpresa = giFilialEmpresa
    objTipo.dtDataAtualizacao = gdtDataHoje
    
    lErro = Trata_Alteracao(objTipo, objTipo.iCodigo)
    If lErro <> SUCESSO Then gError 116933

    lErro = CF("TipoFreteFP_Grava", objTipo)
    If lErro <> SUCESSO Then gError 116934

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 116929
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 116930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 116931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)
            
        Case 116932
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRECO_NAO_PREENCHIDO", gErr)

        Case 116934, 116933

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174882)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Limpa_Tela_Tipo() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Tipo

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)

    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True
    
    DataAtualizacao.Caption = ""

    'Zerar iAlterado
    iAlterado = 0

    Limpa_Tela_Tipo = SUCESSO

    Exit Function

Erro_Limpa_Tela_Tipo:

    Limpa_Tela_Tipo = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174883)

    End Select

    Exit Function

End Function

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim iIndice As Integer

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    Codigo.PromptInclude = False
    Codigo.Text = CStr(colCampoValor.Item("Codigo").vValor)
    Codigo.PromptInclude = True
    Descricao.Text = colCampoValor.Item("Descricao").vValor
    NomeReduzido.Text = colCampoValor.Item("NomeReduzido").vValor
    Preco.Text = colCampoValor.Item("Preco").vValor
    DataAtualizacao.Caption = colCampoValor.Item("DataAtualizacao").vValor
    
    iAlterado = 0

End Sub

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

'Dim Geral
Dim objCampoValor As AdmCampoValor
'Dim específicos
Dim iCodigo As Integer

    'Informa tabela associada à Tela
    sTabela = "TipoFreteFP"

    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Trim(Codigo.Text)) <> 0 Then iCodigo = CInt(Codigo.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", Descricao.Text, STRING_TIPO_FRETE_DESCRICAO, "Descricao"
    colCampoValor.Add "NomeReduzido", NomeReduzido.Text, STRING_TIPO_FRETE_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Preco", StrParaDbl(Preco.Text), 0, "Preco"
    colCampoValor.Add "DataAtualizacao", StrParaDate(DataAtualizacao.Caption), 0, "DataAtualizacao"
        
   'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
       
End Sub


Private Sub objEventoTipoFrete_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipo As ClassTipoFreteFP

On Error GoTo Erro_objEventoTipoFrete_evSelecao

    Set objTipo = obj1

    'Lê o Produto
    lErro = CF("TipoFreteFP_Le", objTipo)
    If lErro <> SUCESSO And lErro <> 116913 Then gError 116914

    'Se não achou o Produto --> erro
    If lErro = 116913 Then gError 116915

    Call Traz_Tipo_Tela(objTipo)
    
    Me.Show

    Exit Sub

Erro_objEventoTipoFrete_evSelecao:

    Select Case gErr

        Case 116914

        Case 116915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INEXISTENTE", gErr, objTipo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174884)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TIPOS_BLOQUEIO
    Set Form_Load_Ocx = Me
    Caption = "Tipo de Frete"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TipoFreteFP"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        
        Call BotaoProxNum_Click
    
    ElseIf KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        End If
    
    End If
    
End Sub

Private Sub LabelDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescricao, Source, X, Y)
End Sub

Private Sub LabelDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescricao, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReduzido, Source, X, Y)
End Sub

Private Sub LabelNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelPreco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPreco, Source, X, Y)
End Sub

Private Sub LabelPreco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPreco, Button, Shift, X, Y)
End Sub

Private Sub LabelAtualizado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAtualizado, Source, X, Y)
End Sub

Private Sub LabelAtualizado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAtualizado, Button, Shift, X, Y)
End Sub

