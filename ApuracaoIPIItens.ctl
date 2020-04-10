VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ApuracaoIPIItensOcx 
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   KeyPreview      =   -1  'True
   ScaleHeight     =   2445
   ScaleWidth      =   6315
   Begin VB.TextBox Descricao 
      Height          =   300
      Left            =   1215
      MaxLength       =   255
      TabIndex        =   1
      Top             =   825
      Width           =   4845
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3945
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ApuracaoIPIItens.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "ApuracaoIPIItens.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "ApuracaoIPIItens.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ApuracaoIPIItens.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoItens 
      Caption         =   "Lan�amentos j� cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4290
      TabIndex        =   5
      Top             =   1350
      Width           =   1785
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   1215
      TabIndex        =   3
      Top             =   1890
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Tipo 
      Height          =   300
      Left            =   1215
      TabIndex        =   0
      Top             =   285
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   2235
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1365
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   1215
      TabIndex        =   2
      Top             =   1365
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label LabelTipo 
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
      Left            =   690
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   14
      Top             =   345
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descri��o:"
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
      Left            =   210
      TabIndex        =   13
      Top             =   885
      Width           =   930
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
      Left            =   630
      TabIndex        =   12
      Top             =   1950
      Width           =   510
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
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
      Left            =   660
      TabIndex        =   11
      Top             =   1425
      Width           =   480
   End
End
Attribute VB_Name = "ApuracaoIPIItensOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

'Eventos dos Browses
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1
Private WithEvents objEventoBotaoItens As AdmEvento
Attribute objEventoBotaoItens.VB_VarHelpID = -1

Function Trata_Parametros(Optional objRegApuracaoItem As ClassRegApuracaoItem) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    'Se foi passado um Item como par�metro
    If Not objRegApuracaoItem Is Nothing Then
    
        'Se os campos que identificam o Item veio preenchido
        If objRegApuracaoItem.iTipoReg > 0 And objRegApuracaoItem.dtData <> DATA_NULA And Len(Trim(objRegApuracaoItem.sDescricao)) > 0 Then
        
            'L� o Item de apura��o IPI a partir do Tipo, Descri��o e Data
            lErro = CF("RegApuracaoIPIItens_Le",objRegApuracaoItem)
            If lErro <> SUCESSO And lErro <> 79063 Then gError 79087
            
            'Se n�o encontrou, erro
            If lErro = 79063 Then gError 79088
                        
            'Traz os dados do Item de apura��o para a tela
            Call Traz_ApuracaoIPIItens_Tela(objRegApuracaoItem)
        
        End If
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 79087
        
        Case 79088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPIITEM_NAO_CADASTRADA", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143080)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoTipo = New AdmEvento
    Set objEventoBotaoItens = New AdmEvento
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143081)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

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

Dim lErro As Long

    'Libera vari�veis globais
    Set objEventoTipo = Nothing
    Set objEventoBotaoItens = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objRegApuracaoItem As New ClassRegApuracaoItem

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada � Tela
    sTabela = "RegApuracaoIPIItem"

    'Le os dados da tela
    Call Move_Tela_Memoria(objRegApuracaoItem)

    'Preenche a cole��o colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objRegApuracaoItem.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "NumIntDocApuracao", objRegApuracaoItem.lNumIntDoc, 0, "NumIntDocApuracao"
    colCampoValor.Add "TipoReg", objRegApuracaoItem.iTipoReg, 0, "TipoReg"
    colCampoValor.Add "Descricao", objRegApuracaoItem.sDescricao, STRING_DESCRICAO_APURACAO, "Descricao"
    colCampoValor.Add "Data", objRegApuracaoItem.dtData, 0, "Data"
    colCampoValor.Add "Valor", objRegApuracaoItem.dValor, 0, "Valor"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143082)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objRegApuracaoItem As New ClassRegApuracaoItem
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Carrega objRegApuracaoItem com os dados passados em colCampoValor
    objRegApuracaoItem.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objRegApuracaoItem.lNumIntDocApuracao = colCampoValor.Item("NumIntDocApuracao").vValor
    objRegApuracaoItem.iTipoReg = colCampoValor.Item("TipoReg").vValor
    objRegApuracaoItem.sDescricao = colCampoValor.Item("Descricao").vValor
    objRegApuracaoItem.dtData = colCampoValor.Item("Data").vValor
    objRegApuracaoItem.dValor = colCampoValor.Item("Valor").vValor
    
    'Verifica se o C�digo est� preenchido
    If objRegApuracaoItem.iTipoReg <> 0 Then

        'Traz os dados dos itens de apura��o IPI para a tela tela
        Call Traz_ApuracaoIPIItens_Tela(objRegApuracaoItem)

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143083)

    End Select

    Exit Sub

End Sub

Sub Traz_ApuracaoIPIItens_Tela(objRegApuracaoItem As ClassRegApuracaoItem)
'Traz os dados dos itens de apura��o IPI para a tela tela

    Tipo.Text = objRegApuracaoItem.iTipoReg
    Descricao.Text = objRegApuracaoItem.sDescricao
    
    Call DateParaMasked(Data, objRegApuracaoItem.dtData)
    
    Valor.Text = Format(objRegApuracaoItem.dValor, "Standard")
    
    iAlterado = 0
    
End Sub

Sub Move_Tela_Memoria(objRegApuracaoItem As ClassRegApuracaoItem)
'Move dados da tela para a mem�ria

    objRegApuracaoItem.iTipoReg = StrParaInt(Tipo.Text)
    objRegApuracaoItem.sDescricao = Descricao.Text
    objRegApuracaoItem.dtData = StrParaDate(Data.Text)
    objRegApuracaoItem.dValor = StrParaDbl(Valor.Text)

End Sub

Private Sub BotaoItens_Click()

Dim colSelecao As New Collection
Dim objApuracaoIPIItem As New ClassRegApuracaoItem

    Call Chama_Tela("ApuracaoIPIItensLista", colSelecao, objApuracaoIPIItem, objEventoBotaoItens)

End Sub

Private Sub objEventoBotaoItens_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRegApuracaoItem As ClassRegApuracaoItem

On Error GoTo Erro_objEventoBotaoItens_evSelecao

    Set objRegApuracaoItem = obj1

    Call Traz_ApuracaoIPIItens_Tela(objRegApuracaoItem)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoBotaoItens_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143084)

    End Select

    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Se a Data est� preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 79089

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True
    
    Select Case gErr

        Case 79089

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143085)

    End Select

    Exit Sub

End Sub

Private Sub LabelTipo_Click()

Dim colSelecao As New Collection
Dim objTipoReg As New ClassTiposRegApuracao

    Call Chama_Tela("TiposRegApuracaoIPILista", colSelecao, objTipoReg, objEventoTipo)

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Se a data est� preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 79090

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 79090

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143086)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Se a data est� preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 79091

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 79091

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143087)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Tipo, iAlterado)

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoRegApuracao As New ClassTiposRegApuracao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Tipo_Validate
            
    'Se o tipo n�o foi preenchido, sai da rotina
    If Len(Trim(Tipo.ClipText)) = 0 Then Exit Sub
    
    'Guarda o c�digo do Tipo
    objTipoRegApuracao.iCodigo = CInt(Tipo.Text)
    
    'Verifica se o Tipo de Apura��o est� cadastrado
    lErro = CF("TipoRegApuracaoIPI_Le",objTipoRegApuracao)
    If lErro <> SUCESSO And lErro <> 79024 Then gError 79092
    
    'Se o Tipo de Apura��o n�o est� cadastrado, pergunta se desja criar
    If lErro = 79024 Then gError 79093
    
    Exit Sub
    
Erro_Tipo_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 79092
        
        Case 79093
            'Pergunta se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOAPURACAOIPI", objTipoRegApuracao.iCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Tipo de Apura��o IPI
                Call Chama_Tela("TiposRegApuracaoIPI", objTipoRegApuracao)
            End If
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143088)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoRegApuracaoIPI As ClassTiposRegApuracao
 
On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoRegApuracaoIPI = obj1
    
    'Coloca c�digo do Tipo na tela
    Tipo.Text = objTipoRegApuracaoIPI.iCodigo
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoTipo_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143089)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Se o campo foi preenchido
    If Len(Trim(Valor.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 79094
    
    End If
    
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 79094
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143090)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava um Item de apura�ao IPI
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 79095

    'Limpa a tela
    Call Limpa_Tela_ApuracaoIPIItens

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 79095

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143091)

    End Select

Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRegApuracaoItem As New ClassRegApuracaoItem

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o tipo esta preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 79096

    'Verifica se a descri��o est� preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 79097

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 79098
    
    'Verifica se o valor foi preenchido
    If Len(Trim(Valor.Text)) = 0 Then gError 79099
    
    'Recolhe os dados da tela
    Call Move_Tela_Memoria(objRegApuracaoItem)

    'Grava um tipo de apura��o
    lErro = CF("RegApuracaoIPIItem_Grava",objRegApuracaoItem)
    If lErro <> SUCESSO Then gError 79100

    'Limpa a tela
    Call Limpa_Tela_ApuracaoIPIItens

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 79096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPURACAO_NAO_PREENCHIDA", gErr)

        Case 79097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 79098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 79099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
            
        Case 79100

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143092)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegApuracaoItem As New ClassRegApuracaoItem

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o tipo esta preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 79101

    'Verifica se a descri��o est� preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 79102

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 79103

    'Move os dqado da tela para a mem�ria
    Call Move_Tela_Memoria(objRegApuracaoItem)
    
    'L� o tipo de registro para apura��o IPI
    lErro = CF("RegApuracaoIPIItens_Le",objRegApuracaoItem)
    If lErro <> SUCESSO And lErro <> 79063 Then gError 79104

    'Se n�o encontrou, erro
    If lErro = 79063 Then gError 79105
    
    'Pede a confirma��o da exclus�o do Item de apura��o IPI
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REGAPURACAOIPIITEM", objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui o tipo de apura��o
    lErro = CF("RegApuracaoIPIItem_Exclui",objRegApuracaoItem)
    If lErro <> SUCESSO Then gError 79106

    'Limpa a tela
    Call Limpa_Tela_ApuracaoIPIItens

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 79101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPURACAO_NAO_PREENCHIDA", gErr)

        Case 79102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 79103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 79104, 79106

        Case 79105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPIITEM_NAO_CADASTRADA", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143093)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se h� altera��es e quer salv�-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 79107

    'Limpa a tela
    Call Limpa_Tela_ApuracaoIPIItens

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 79107

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143094)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Tipo Then
            Call LabelTipo_Click
        End If

    End If

End Sub

Private Sub Limpa_Tela_ApuracaoIPIItens()
'Limpa a Tela de Apuracao IPI

Dim lErro As Long

    Tipo.Text = ""
    Descricao.Text = ""
    
    Data.PromptInclude = False
    Data.Text = ""
    Data.PromptInclude = True
    
    Valor.Text = ""
        
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
        
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Lan�amentos para Apura��o de IPI"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ApuracaoIPIItens"

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

'**** fim do trecho a ser copiado *****


Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipo, Source, X, Y)
End Sub

Private Sub LabelTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipo, Button, Shift, X, Y)
End Sub


