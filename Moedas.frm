VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Moedas 
   Caption         =   "Cadastro de Moedas"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Percentual 
      Caption         =   "Valor em percentual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   195
      TabIndex        =   13
      Top             =   1815
      Width           =   2280
   End
   Begin VB.ListBox ListaMoedas 
      Height          =   1425
      ItemData        =   "Moedas.frx":0000
      Left            =   3390
      List            =   "Moedas.frx":0002
      TabIndex        =   12
      Top             =   690
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3390
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Moedas.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Moedas.frx":015E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Moedas.frx":02E8
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "Moedas.frx":081A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1620
      Picture         =   "Moedas.frx":0998
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Numeração Automática"
      Top             =   345
      Width           =   300
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1035
      TabIndex        =   1
      Top             =   330
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1005
      TabIndex        =   3
      Top             =   855
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Simbolo 
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   1350
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   10
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
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
      TabIndex        =   5
      Top             =   900
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Símbolo:"
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
      Left            =   195
      TabIndex        =   4
      Top             =   1380
      Width           =   765
   End
   Begin VB.Label Label1 
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
      Left            =   285
      TabIndex        =   2
      Top             =   390
      Width           =   660
   End
End
Attribute VB_Name = "Moedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iAlterado As Integer

Private Const STRING_NOME_MOEDA = 20
Private Const STRING_SIMBOLO_MOEDA = 10

Private Type typeMoedas
    iCodigo As Integer
    sNome As String
    sSimbolo As String
End Type

Public Sub Form_Load()
    
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    ListaMoedas.Clear
    
    lErro = Preenche_Tela_Moedas()
    If lErro <> SUCESSO Then gError 108822
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162791)
    
    End Select
    
    Exit Sub
    
End Sub

Private Function Preenche_Tela_Moedas() As Long
'Preenche a tela com as moedas já existentes no bd

Dim lErro As Long
Dim colMoedas As New Collection
Dim objMoeda As ClassMoedas
Dim sMoedas As String

On Error GoTo Erro_Preenche_Tela_Moedas

    ListaMoedas.Clear
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 108822
    
    For Each objMoeda In colMoedas
        
        sMoedas = objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        
        'Adiciona a List
        ListaMoedas.AddItem (sMoedas)
    
    Next
       
    Preenche_Tela_Moedas = SUCESSO
    
    Exit Function
    
Erro_Preenche_Tela_Moedas:

    Preenche_Tela_Moedas = gErr
    
    Select Case gErr
    
        Case 108822
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162792)
    
    End Select
    
    Exit Function

End Function

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Obtém o próximo código disponível para TipoBloqueioPC
    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_MOEDA", "Moedas", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 108800

    'Mostra o código na tela
    Codigo.Text = CStr(lCodigo)
            
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 108800
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162793)
    
    End Select

End Sub

Private Sub ListaMoedas_DblClick()

Dim lErro As Long
Dim objMoedas As New ClassMoedas

On Error GoTo Erro_ListaMoedas_DblClick

    'Guarda o Código da Moeda
    objMoedas.iCodigo = Codigo_Extrai(ListaMoedas.List(ListaMoedas.ListIndex))
    
    'Le para trazer para a tela
    lErro = CF("Moedas_Le", objMoedas)
    If lErro <> SUCESSO And lErro <> 108821 Then gError 108825
    
    'Se nao encontrou => Erro
    If lErro = 108821 Then gError 108826
    
    'Preenche a tela
    Codigo.Text = CStr(objMoedas.iCodigo)
    NomeReduzido.Text = objMoedas.sNome
    Simbolo.Text = objMoedas.sSimbolo
    
    '###################################
    'Inserido por Wagner
    If objMoedas.iPercentual = MARCADO Then
        Percentual.Value = vbChecked
    Else
        Percentual.Value = vbUnchecked
    End If
    '###################################
    
    Exit Sub

Erro_ListaMoedas_DblClick:

    Select Case gErr
    
        Case 108825
        
        Case 108826
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_INEXISTENTE", gErr, objMoedas.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162794)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava os registros na tabela
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 50069

    Call Limpa_Tela_Indexador
    
    lErro = Preenche_Tela_Moedas()
    If lErro <> SUCESSO Then gError 108823

    iAlterado = 0

    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err

        Case 50069, 108823

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162795)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objMoedas As New ClassMoedas

On Error GoTo Erro_Gravar_Registro

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 108802
    
    'Verifica se a Descricao foi preenchida
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 108803
    
    'Verifica se o Simbolo foi preenchido
    If Len(Trim(Simbolo.Text)) = 0 Then gError 108804
    
    'Verifica se a Moeda é Dolar ou Real ... se for dá erro
    If StrParaInt(Codigo.Text) = MOEDA_REAL Then gError 108843
    
    'Verifica se a Moeda é Dolar ou Real ... se for dá erro
    If StrParaInt(Codigo.Text) = MOEDA_DOLAR Then gError 108805
    
    'Carrega o obj com os valores a serem passados como parametro
    objMoedas.iCodigo = StrParaInt(Codigo.Text)
    objMoedas.sNome = NomeReduzido.Text
    objMoedas.sSimbolo = Simbolo.Text
    
    '##########################
    'Inserido por Wagner
    If Percentual.Value = vbChecked Then
        objMoedas.iPercentual = MARCADO
    Else
        objMoedas.iPercentual = DESMARCADO
    End If
    '##########################
    
    'Chama função de gravação com obj carregado
    lErro = CF("Moedas_Grava", objMoedas)
    If lErro <> SUCESSO Then gError 108806
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 108802
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus
        
        Case 108803
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", gErr)
            NomeReduzido.SetFocus

        Case 108804
            Call Rotina_Erro(vbOKOnly, "ERRO_SIMBOLO_NAO_PREENCHIDO", gErr)
            Simbolo.SetFocus

        Case 108805
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_DOLAR", gErr)
            
        Case 108806
        
        Case 108843
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_REAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162796)

    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objMoedas As New ClassMoedas

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se a data está preenchida
    If Len(Trim(Codigo.Text)) = 0 Then gError 108813
    
    'Carrega o objMoedas
    objMoedas.iCodigo = StrParaInt(Codigo.Text)
    
    'Verifica se a Moeda é Dolar ... se for dá erro
    If StrParaInt(Codigo.Text) = MOEDA_DOLAR Then gError 108841
    
    'Verifica se a Moeda é Real ... se for dá erro
    If StrParaInt(Codigo.Text) = MOEDA_REAL Then gError 108853
    
    'Verifica se existe a moeda informada
    lErro = CF("Moedas_Le", objMoedas)
    If lErro <> SUCESSO And lErro <> 108821 Then gError 108816
    
    'Se não encontrou => ERRO
    If lErro <> SUCESSO Then gError 108817
    
    'Envia mensagem de confirmação de exclusão para o usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_MOEDAS", objMoedas.iCodigo)
    If vbMsgRes = vbYes Then
    
        'Exclui a CotacaoMoeda informada
        lErro = CF("Moedas_Exclui", objMoedas)
        If lErro <> SUCESSO Then gError 108815
        
        Call Limpa_Tela_Indexador
        
        Call Preenche_Tela_Moedas
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        
        Case 108813
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
                
        Case 108816, 108815
        
        Case 108817
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_INEXISTENTE", gErr, objMoedas.iCodigo)
            
        Case 108841
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_DOLAR", gErr)
            
        Case 108853
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_REAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162797)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 108801

    Call Limpa_Tela_Indexador
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 108801

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162798)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub


Private Sub Limpa_Tela_Indexador()

    Call Limpa_Tela(Me)
        
    iAlterado = 0
    
    Percentual.Value = vbUnchecked 'Inserido por Wagner
    
    Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode)
      
End Sub

Private Sub Simbolo_Change()

On Error GoTo Erro_Simbolo_Change

    If Len(Trim(Simbolo.Text)) > 0 Then
    
        If Not IniciaLetra(Simbolo.Text) Then gError 86258
        
    End If
    
    Exit Sub
    
Erro_Simbolo_Change:

    Select Case gErr
    
        Case 86258
            Call Rotina_Erro(vbOKOnly, "ERRO_SIMBOLO_COMECA_NUMERO", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162799)
    
    End Select
    
    Exit Sub
    
End Sub
