VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ClassificacaoFiscalOcx 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   LockControls    =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   6090
   Begin VB.CommandButton BotaoNCM 
      Caption         =   "NCMs pré-cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      TabIndex        =   21
      Top             =   3495
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alíquotas"
      Height          =   765
      Left            =   240
      TabIndex        =   15
      Top             =   1845
      Width           =   5715
      Begin MSMask.MaskEdBox AliquotaII 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AliquotaIPI 
         Height          =   285
         Left            =   2805
         TabIndex        =   3
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IPI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2415
         TabIndex        =   17
         Top             =   330
         Width           =   315
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "II:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   795
         TabIndex        =   16
         Top             =   330
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Importação"
      Height          =   750
      Left            =   240
      TabIndex        =   14
      Top             =   2670
      Width           =   5715
      Begin MSMask.MaskEdBox AliquotaICMS 
         Height          =   285
         Left            =   1095
         TabIndex        =   4
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AliquotaPIS 
         Height          =   285
         Left            =   2820
         TabIndex        =   5
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AliquotaCOFINS 
         Height          =   285
         Left            =   4845
         TabIndex        =   6
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "COFINS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4035
         TabIndex        =   20
         Top             =   330
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PIS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2370
         TabIndex        =   19
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ICMS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   495
         TabIndex        =   18
         Top             =   330
         Width           =   525
      End
   End
   Begin VB.TextBox Descricao 
      Height          =   885
      Left            =   1305
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   855
      Width           =   4650
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3840
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ClassificacaoFiscalOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ClassificacaoFiscalOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ClassificacaoFiscalOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ClassificacaoFiscalOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   345
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   10
      Format          =   "0000\.00\.00"
      Mask            =   "##########"
      PromptChar      =   " "
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
      Left            =   240
      TabIndex        =   12
      Top             =   930
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
      Left            =   510
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   390
      Width           =   660
   End
End
Attribute VB_Name = "ClassificacaoFiscalOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos do browse
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoNCM As AdmEvento
Attribute objEventoNCM.VB_VarHelpID = -1

Dim iAlterado As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long
    
On Error GoTo Erro_Form_Load
    
    'Inicializa o Browse
    Set objEventoCodigo = New AdmEvento
    Set objEventoNCM = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    iAlterado = 0
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150811)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objClassificacaoFiscal As ClassClassificacaoFiscal) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros
    
    If Not (objClassificacaoFiscal Is Nothing) Then
        
        'Traz os dados de ClassificacaoFiscal para Tela
        lErro = Traz_ClassificacaoFiscal_Tela(objClassificacaoFiscal)
        If lErro <> SUCESSO And lErro <> 123486 Then gError 123475
        
        If lErro = 123486 Then Codigo.Text = objClassificacaoFiscal.sCodigo
        
    End If
    
    iAlterado = 0
        
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
        
        Case 123475
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150812)
        
    End Select
        
    iAlterado = 0
        
    Exit Function
        
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO
Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 123476

    'Limpa a tela
    lErro = LimpaClassificacaoFiscal
    If lErro <> SUCESSO Then gError 123477
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 123476, 123477
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150813)
        
    End Select
        
    Exit Sub
        
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se os campos obrigatorios estao preenchidos
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 123478
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CLASSIFICACAOFISCAL")
    
    If vbMsgRes = vbYes Then
    
        'Move os elementos da tela para memória
        lErro = Move_Tela_Memoria(objClassificacaoFiscal)
        If lErro <> SUCESSO Then gError 123480
        
        'Chama a função que irá excluir o elemento do BD
        lErro = CF("ClassificacaoFiscal_Exclui", objClassificacaoFiscal)
        If lErro <> SUCESSO And lErro <> 125007 Then gError 123481
        
        If lErro = 125007 Then gError 123017
        
        'Limpa a tela
        lErro = LimpaClassificacaoFiscal
        If lErro <> SUCESSO Then gError 123482
    
        iAlterado = 0
        
    End If
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case gErr
    
        Case 123478
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAOPREECHIDO", gErr)
        
        Case 123480 To 123482
            
        Case 123017
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAO_EXISTENTE", gErr, objClassificacaoFiscal.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150814)
        
    End Select
        
    Exit Sub
        
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se o usuário deseja salvar o elemento no BD
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 123483
    
    'Limpa a tela
    lErro = LimpaClassificacaoFiscal
    If lErro <> SUCESSO Then gError 123484
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
    
        Case 123483, 123484
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150815)
        
    End Select
        
    Exit Sub
        
End Sub

Private Sub LabelCodigo_Click()

Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim colSelecao As New Collection
    
    'Preenche na memória o Código passado
    If Len(Trim(Codigo.ClipText)) > 0 Then objClassificacaoFiscal.sCodigo = Codigo.ClipText

    Call Chama_Tela("ClassificacaoFiscalLista", colSelecao, objClassificacaoFiscal, objEventoCodigo)

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

'*** FUNÇÕES DE APOIO À TELA - INÍCIO ***
Private Function Traz_ClassificacaoFiscal_Tela(objClassificacaoFiscal As ClassClassificacaoFiscal) As Long
'Realiza o preencimento da tela com as informações do BD

Dim lErro As Long

On Error GoTo Erro_Traz_ClassificacaoFiscal_Tela

    'Busca as informações sobre o código passado
    lErro = CF("ClassificacaoFiscal_Le", objClassificacaoFiscal)
    If lErro <> SUCESSO And lErro <> 123494 Then gError 123485
    
    If lErro = 123494 Then gError 123486
    
    'Preenche o Código
    Codigo.PromptInclude = False
    Codigo.Text = objClassificacaoFiscal.sCodigo
    Codigo.PromptInclude = True
    
    'Preenche o campo descrição da tela
    Descricao.Text = objClassificacaoFiscal.sDescricao
    
    If objClassificacaoFiscal.dIIAliquota <> 0 Then
        AliquotaII.Text = objClassificacaoFiscal.dIIAliquota * 100
    Else
        AliquotaII.Text = ""
    End If
    
    If objClassificacaoFiscal.dIPIAliquota <> 0 Then
        AliquotaIPI.Text = objClassificacaoFiscal.dIPIAliquota * 100
    Else
        AliquotaIPI.Text = ""
    End If
    
    If objClassificacaoFiscal.dPISAliquota <> 0 Then
        AliquotaPIS.Text = objClassificacaoFiscal.dPISAliquota * 100
    Else
        AliquotaPIS.Text = ""
    End If
    
    If objClassificacaoFiscal.dCOFINSAliquota <> 0 Then
        AliquotaCOFINS.Text = objClassificacaoFiscal.dCOFINSAliquota * 100
    Else
        AliquotaCOFINS.Text = ""
    End If
    
    If objClassificacaoFiscal.dICMSAliquota <> 0 Then
        AliquotaICMS.Text = objClassificacaoFiscal.dICMSAliquota * 100
    Else
        AliquotaICMS.Text = ""
    End If
    
    Traz_ClassificacaoFiscal_Tela = SUCESSO

    Exit Function
    
Erro_Traz_ClassificacaoFiscal_Tela:

    Traz_ClassificacaoFiscal_Tela = gErr
    
    Select Case gErr
    
        Case 123485
        
        Case 123486
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150816)
        
    End Select
        
    Exit Function
        
End Function

Private Function Move_Tela_Memoria(objClassificacaoFiscal As ClassClassificacaoFiscal) As Long
'Move para a memória todas as informações que estão preenchidas na tela

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Preenche o objClassificaçãoFiscal
    Codigo.PromptInclude = False
    objClassificacaoFiscal.sCodigo = Codigo.Text
    Codigo.PromptInclude = True
    
    objClassificacaoFiscal.sDescricao = Descricao.Text
    
    If Len(Trim(AliquotaII.Text)) > 0 Then objClassificacaoFiscal.dIIAliquota = CDbl(AliquotaII / 100)
    If Len(Trim(AliquotaICMS.Text)) > 0 Then objClassificacaoFiscal.dICMSAliquota = CDbl(AliquotaICMS / 100)
    If Len(Trim(AliquotaIPI.Text)) > 0 Then objClassificacaoFiscal.dIPIAliquota = CDbl(AliquotaIPI / 100)
    If Len(Trim(AliquotaCOFINS.Text)) > 0 Then objClassificacaoFiscal.dCOFINSAliquota = CDbl(AliquotaCOFINS / 100)
    If Len(Trim(AliquotaPIS.Text)) > 0 Then objClassificacaoFiscal.dPISAliquota = CDbl(AliquotaPIS / 100)

    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150817)
        
    End Select
        
    Exit Function
        
End Function

Function Gravar_Registro() As Long
'Grava o registro no Banco de dados
 
Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal

On Error GoTo Erro_Gravar_Registro

    'Verifica se os campos obrigatórios estão prenchidos
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 123487
    If Len(Trim(Descricao.Text)) = 0 Then gError 123490
    
    'Move as informações da tela para memória
    lErro = Move_Tela_Memoria(objClassificacaoFiscal)
    If lErro <> SUCESSO Then gError 123488
    
    'Faz a chamada a função que irá realizar a gravação no BD
    lErro = CF("ClassificacaoFiscal_Grava", objClassificacaoFiscal)
    If lErro <> SUCESSO Then gError 123489
    
    Gravar_Registro = SUCESSO

    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 123487
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAOPREECHIDO", gErr)
            
        Case 123488, 123489
        
        Case 123490
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_CLASSIFICACAOFISCAL_NAOPREECHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150818)
        
    End Select
        
    Exit Function
        
End Function

Private Function LimpaClassificacaoFiscal() As Long
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_LimpaClassificacaoFiscal

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'Limpa os Campos
    lErro = Limpa_Tela(Me)
    If lErro <> SUCESSO Then gError 123490
    
    LimpaClassificacaoFiscal = SUCESSO
    
    Exit Function
    
Erro_LimpaClassificacaoFiscal:

    LimpaClassificacaoFiscal = gErr
    
    Select Case gErr
    
        Case 123490
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150819)
        
    End Select
        
    Exit Function
    
End Function
'*** FUNÇÕES DE APOIO À TELA - FIM ***

'*** FUNÇÕES DO SISTEMA DE SETAS - INÍCIO***
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal

On Error GoTo Erro_Tela_Extrai
    
    'Informa a Tabela
    sTabela = "ClassificacaoFiscal"
    
    'Move os elementos da Tela para a memória
    lErro = Move_Tela_Memoria(objClassificacaoFiscal)
    If lErro <> SUCESSO Then gError 125011
    
    'preenche a Coleção
    colCampoValor.Add "Codigo", objClassificacaoFiscal.sCodigo, STRING_PRODUTO_IPI_CODIGO, "Codigo"
    'colCampoValor.Add "Descricao", objClassificacaoFiscal.sDescricao, STRING_CLASSIFICACAOFISCAL_DESCRICAO, "Descricao"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
            
        Case 125011
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150820)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal

On Error GoTo Erro_TelaPreenche
    
    'Informa o Código
    objClassificacaoFiscal.sCodigo = colCampoValor.Item("Codigo").vValor
    
    If Len(Trim(objClassificacaoFiscal.sCodigo)) <> 0 Then

        'Mostra os dados do na tela
        lErro = Traz_ClassificacaoFiscal_Tela(objClassificacaoFiscal)
        If lErro <> SUCESSO And lErro <> 123486 Then gError 125012
        
        If lErro = 123486 Then gError 125016

    End If
    
    iAlterado = 0

    Exit Sub

Erro_TelaPreenche:

    Select Case gErr
    
        Case 125012
        
        Case 125016
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAO_EXISTENTE", gErr, objClassificacaoFiscal.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150821)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera as variáveis globais
    Set objEventoCodigo = Nothing
    Set objEventoNCM = Nothing

    Call ComandoSeta_Liberar(Me.Name)

End Sub
'*** FUNÇÕES DO SISTEMA DE SETA - FIM

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

'*** FUNÇÕES DO BROWSE - INÍCIO

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim bCancel As Boolean
    
On Error GoTo Erro_objEventoCodigo_evSelecao
    
    Set objClassificacaoFiscal = obj1

    lErro = Traz_ClassificacaoFiscal_Tela(objClassificacaoFiscal)
    If lErro <> SUCESSO And lErro <> 123494 Then gError 125013
    
    If lErro = 123494 Then gError 125014

    Me.Show

    iAlterado = 0
    
    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 125013
        
        Case 125014
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAO_EXISTENTE", gErr, objClassificacaoFiscal.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150822)

    End Select

    Exit Sub

End Sub
'*** FUNÇÕES DO BROWSE - FIM ***

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EMPENHO
    Set Form_Load_Ocx = Me
    Caption = "Classificacao Fiscal"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ClassificacaoFiscal"

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

Private Sub AliquotaII_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub AliquotaII_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaII_Validate

    'Verifica se esta preenchida
    If Len(Trim(AliquotaII.Text)) = 0 Then Exit Sub

    'Critica
    lErro = Porcentagem_Critica(AliquotaII.Text)
    If lErro <> SUCESSO Then gError 64342

    Exit Sub

Erro_AliquotaII_Validate:

    Cancel = True
    
    Select Case gErr

        Case 64342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165478)

    End Select

    Exit Sub
    
End Sub


Private Sub AliquotaIPI_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub AliquotaIPI_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaIPI_Validate

    'Verifica se esta preenchida
    If Len(Trim(AliquotaIPI.Text)) = 0 Then Exit Sub

    'Critica
    lErro = Porcentagem_Critica(AliquotaIPI.Text)
    If lErro <> SUCESSO Then gError 64342

    Exit Sub

Erro_AliquotaIPI_Validate:

    Cancel = True
    
    Select Case gErr

        Case 64342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165478)

    End Select

    Exit Sub
    
End Sub


Private Sub AliquotaPIS_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub AliquotaPIS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaPIS_Validate

    'Verifica se esta preenchida
    If Len(Trim(AliquotaPIS.Text)) = 0 Then Exit Sub

    'Critica
    lErro = Porcentagem_Critica(AliquotaPIS.Text)
    If lErro <> SUCESSO Then gError 64342

    Exit Sub

Erro_AliquotaPIS_Validate:

    Cancel = True
    
    Select Case gErr

        Case 64342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165478)

    End Select

    Exit Sub
    
End Sub


Private Sub AliquotaICMS_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub AliquotaICMS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaICMS_Validate

    'Verifica se esta preenchida
    If Len(Trim(AliquotaICMS.Text)) = 0 Then Exit Sub

    'Critica
    lErro = Porcentagem_Critica(AliquotaICMS.Text)
    If lErro <> SUCESSO Then gError 64342

    Exit Sub

Erro_AliquotaICMS_Validate:

    Cancel = True
    
    Select Case gErr

        Case 64342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165478)

    End Select

    Exit Sub
    
End Sub


Private Sub AliquotaCOFINS_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub AliquotaCOFINS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaCOFINS_Validate

    'Verifica se esta preenchida
    If Len(Trim(AliquotaCOFINS.Text)) = 0 Then Exit Sub

    'Critica
    lErro = Porcentagem_Critica(AliquotaCOFINS.Text)
    If lErro <> SUCESSO Then gError 64342

    Exit Sub

Erro_AliquotaCOFINS_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 64342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165478)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoNCM_Click()

Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim colSelecao As New Collection
    
    'Preenche na memória o Código passado
    If Len(Trim(Codigo.ClipText)) > 0 Then objClassificacaoFiscal.sCodigo = Codigo.ClipText

    Call Chama_Tela("NCMLista", colSelecao, objClassificacaoFiscal, objEventoNCM)

End Sub

Private Sub objEventoNCM_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim bCancel As Boolean
    
On Error GoTo Erro_objEventoNCM_evSelecao
    
    Set objClassificacaoFiscal = obj1

    'Preenche o Código
    Codigo.PromptInclude = False
    Codigo.Text = objClassificacaoFiscal.sCodigo
    Codigo.PromptInclude = True
    
    'Preenche o campo descrição da tela
    Descricao.Text = objClassificacaoFiscal.sDescricao
    
    Me.Show

    iAlterado = 0
    
    Exit Sub

Erro_objEventoNCM_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150822)

    End Select

    Exit Sub

End Sub
