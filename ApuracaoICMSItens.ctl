VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ApuracaoICMSItensOcx 
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2445
   ScaleWidth      =   6315
   Begin VB.CommandButton BotaoItens 
      Caption         =   "Lançamentos já cadastrados"
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
      Left            =   4320
      TabIndex        =   4
      Top             =   1350
      Width           =   1785
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   1245
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
      Left            =   1245
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
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3975
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ApuracaoICMSItens.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "ApuracaoICMSItens.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "ApuracaoICMSItens.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ApuracaoICMSItens.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox Descricao 
      Height          =   300
      Left            =   1245
      MaxLength       =   255
      TabIndex        =   1
      Top             =   825
      Width           =   4845
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   2265
      TabIndex        =   14
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
      Left            =   1245
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
      Left            =   690
      TabIndex        =   13
      Top             =   1418
      Width           =   480
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
      Left            =   660
      TabIndex        =   12
      Top             =   1943
      Width           =   510
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
      Left            =   240
      TabIndex        =   10
      Top             =   885
      Width           =   930
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
      Left            =   720
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   9
      Top             =   345
      Width           =   450
   End
End
Attribute VB_Name = "ApuracaoICMSItensOcx"
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
    
    'Se foi passado um Item como parâmetro
    If Not objRegApuracaoItem Is Nothing Then
    
        'Se os campos que identificam o Item veio preenchido
        If objRegApuracaoItem.iTipoReg > 0 And objRegApuracaoItem.dtData <> DATA_NULA And Len(Trim(objRegApuracaoItem.sDescricao)) > 0 Then
        
            'Lê o Item de apuração ICMS a partir do Tipo, Descrição e Data
            lErro = CF("RegApuracaoICMSItens_Le",objRegApuracaoItem)
            If lErro <> SUCESSO And lErro <> 67942 Then gError 67963
            
            'Se não encontrou, erro
            If lErro = 67942 Then gError 67964
                        
            'Traz os dados do Item de apuração para a tela
            Call Traz_ApuracaoICMSItens_Tela(objRegApuracaoItem)
        
        End If
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 67963
        
        Case 67964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOICMSITEM_NAO_CADASTRADA", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143028)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143029)

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

    'Libera variáveis globais
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

    'Informa tabela associada à Tela
    sTabela = "RegApuracaoICMSItem"

    'Le os dados da tela
    Call Move_Tela_Memoria(objRegApuracaoItem)

    'Preenche a coleção colCampoValor, com nome do campo,
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143030)

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
    
    'Verifica se o Código está preenchido
    If objRegApuracaoItem.iTipoReg <> 0 Then

        'Traz os dados dos itens de apuração ICMS para a tela tela
        Call Traz_ApuracaoICMSItens_Tela(objRegApuracaoItem)

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143031)

    End Select

    Exit Sub

End Sub

Sub Traz_ApuracaoICMSItens_Tela(objRegApuracaoItem As ClassRegApuracaoItem)
'Traz os dados dos itens de apuração ICMS para a tela tela

    Tipo.Text = objRegApuracaoItem.iTipoReg
    Descricao.Text = objRegApuracaoItem.sDescricao
    
    Call DateParaMasked(Data, objRegApuracaoItem.dtData)
    
    Valor.Text = Format(objRegApuracaoItem.dValor, "Standard")
    
    iAlterado = 0
    
End Sub

Sub Move_Tela_Memoria(objRegApuracaoItem As ClassRegApuracaoItem)
'Move dados da tela para a memória

    objRegApuracaoItem.iTipoReg = StrParaInt(Tipo.Text)
    objRegApuracaoItem.sDescricao = Descricao.Text
    objRegApuracaoItem.dtData = StrParaDate(Data.Text)
    objRegApuracaoItem.dValor = StrParaDbl(Valor.Text)

End Sub

Private Sub BotaoItens_Click()

Dim colSelecao As New Collection
Dim objApuracaoICMSItem As ClassRegApuracaoItem

    Call Chama_Tela("ApuracaoICMSItensLista", colSelecao, objApuracaoICMSItem, objEventoBotaoItens)

End Sub

Private Sub objEventoBotaoItens_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRegApuracaoItem As ClassRegApuracaoItem

On Error GoTo Erro_objEventoBotaoItens_evSelecao

    Set objRegApuracaoItem = obj1

    Call Traz_ApuracaoICMSItens_Tela(objRegApuracaoItem)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoBotaoItens_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143032)

    End Select

    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Se a Data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 67950

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True
    
    Select Case gErr

        Case 67950

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143033)

    End Select

    Exit Sub

End Sub

Private Sub LabelTipo_Click()

Dim colSelecao As New Collection
Dim objTipoReg As ClassTiposRegApuracao

    Call Chama_Tela("TiposRegApuracaoICMSLista", colSelecao, objTipoReg, objEventoTipo)

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 67951

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 67951

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143034)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 67952

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 67952

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143035)

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
            
    'Se o tipo não foi preenchido, sai da rotina
    If Len(Trim(Tipo.ClipText)) = 0 Then Exit Sub
    
    'Guarda o código do Tipo
    objTipoRegApuracao.iCodigo = CInt(Tipo.Text)
    
    'Verifica se o Tipo de Apuração está cadastrado
    lErro = CF("TipoRegApuracaoICMS_Le",objTipoRegApuracao)
    If lErro <> SUCESSO And lErro <> 67893 Then gError 67948
    
    'Se o Tipo de Apuração não está cadastrado, pergunta se desja criar
    If lErro = 67893 Then gError 67949
    
    Exit Sub
    
Erro_Tipo_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 67948
        
        Case 67949
            'Pergunta se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOAPURACAOICMS", objTipoRegApuracao.iCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Tipo de Apuração ICMS
                Call Chama_Tela("TiposRegApuracaoICMS", objTipoRegApuracao)
            End If
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143036)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoRegApuracaoICMS As ClassTiposRegApuracao
 
On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoRegApuracaoICMS = obj1
    
    'Coloca código do Tipo na tela
    Tipo.Text = objTipoRegApuracaoICMS.iCodigo
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoTipo_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143037)

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
        If lErro <> SUCESSO Then gError 67947
    
    End If
    
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 67947
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143038)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava um Item de apuraçao ICMS
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 67930

    'Limpa a tela
    Call Limpa_Tela_ApuracaoICMSItens

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 67930

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143039)

    End Select

Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRegApuracaoItem As New ClassRegApuracaoItem

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o tipo esta preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 67931

    'Verifica se a descrição está preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 67932

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 67933
    
    'Verifica se o valor foi preenchido
    If Len(Trim(Valor.Text)) = 0 Then gError 67934
    
    'Recolhe os dados da tela
    Call Move_Tela_Memoria(objRegApuracaoItem)

    'Grava um tipo de apuração
    lErro = CF("RegApuracaoICMSItem_Grava",objRegApuracaoItem)
    If lErro <> SUCESSO Then gError 67935

    'Limpa a tela
    Call Limpa_Tela_ApuracaoICMSItens

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 67931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPURACAO_NAO_PREENCHIDA", gErr)

        Case 67932
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 67933
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 67934
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
            
        Case 67935

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143040)

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
    If Len(Trim(Tipo.Text)) = 0 Then gError 67936

    'Verifica se a descrição está preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 67937

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 67938

    'Move os dqado da tela para a memória
    Call Move_Tela_Memoria(objRegApuracaoItem)
    
    'Lê o tipo de registro para apuração ICMS
    lErro = CF("RegApuracaoICMSItens_Le",objRegApuracaoItem)
    If lErro <> SUCESSO And lErro <> 67942 Then gError 67943

    'Se não encontrou, erro
    If lErro = 67942 Then gError 67944
    
    'Pede a confirmação da exclusão do Item de apuração ICMS
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REGAPURACAOICMSITEM", objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui o tipo de apuração
    lErro = CF("RegApuracaoICMSItem_Exclui",objRegApuracaoItem)
    If lErro <> SUCESSO Then gError 67945

    'Limpa a tela
    Call Limpa_Tela_ApuracaoICMSItens

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 67936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPURACAO_NAO_PREENCHIDA", gErr)

        Case 67937
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 67938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 67943, 67945

        Case 67944
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOICMSITEM_NAO_CADASTRADA", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143041)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 67946

    'Limpa a tela
    Call Limpa_Tela_ApuracaoICMSItens

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 67946

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143042)

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

Private Sub Limpa_Tela_ApuracaoICMSItens()
'Limpa a Tela de Apuracao ICMS

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
    Caption = "Lançamentos para Apuração de ICMS"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ApuracaoICMSItens"

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

