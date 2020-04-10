VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpLancNatOperacaoOcx 
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   7275
   Begin VB.Frame Frame3 
      Caption         =   "Tributo"
      Height          =   615
      Left            =   300
      TabIndex        =   24
      Top             =   630
      Width           =   2670
      Begin VB.OptionButton OptionTributo 
         Caption         =   "IPI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1755
         TabIndex        =   26
         Top             =   270
         Width           =   675
      End
      Begin VB.OptionButton OptionTributo 
         Caption         =   "ICMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   450
         TabIndex        =   25
         Top             =   300
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   720
      Left            =   300
      TabIndex        =   19
      Top             =   1290
      Width           =   4800
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1845
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   840
         TabIndex        =   6
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   4335
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   3330
         TabIndex        =   7
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dFim 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   2925
         TabIndex        =   23
         Top             =   315
         Width           =   360
      End
      Begin VB.Label dIni 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Height          =   240
         Left            =   450
         TabIndex        =   22
         Top             =   285
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   795
      Left            =   3420
      TabIndex        =   18
      Top             =   2070
      Width           =   3645
      Begin VB.OptionButton OptionTipo 
         Caption         =   "Ambos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   375
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton OptionTipo 
         Caption         =   "Saída"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2535
         TabIndex        =   5
         Top             =   375
         Width           =   885
      End
      Begin VB.OptionButton OptionTipo 
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1350
         TabIndex        =   4
         Top             =   375
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Natureza de Operação"
      Height          =   795
      Left            =   300
      TabIndex        =   15
      Top             =   2070
      Width           =   2925
      Begin MSMask.MaskEdBox NatOperacaoDe 
         Height          =   300
         Left            =   660
         TabIndex        =   1
         Top             =   330
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NatOperacaoAte 
         Height          =   300
         Left            =   2085
         TabIndex        =   2
         Top             =   330
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNatOperacaoAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   1620
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   390
         Width           =   360
      End
      Begin VB.Label LabelNatOperacaoDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   270
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   390
         Width           =   315
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5550
      Picture         =   "RelOpLancNatOperacao.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   900
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5010
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpLancNatOperacao.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpLancNatOperacao.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpLancNatOperacao.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpLancNatOperacao.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpLancNatOperacao.ctx":0A96
      Left            =   1095
      List            =   "RelOpLancNatOperacao.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2916
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Height          =   255
      Left            =   390
      TabIndex        =   14
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpLancNatOperacaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Imprime os Lançamentos detalhados das Naturezas de Operacao

'Campos do Relátorio

'Quebra por: Natureza de Operação - Descricao

'Campos: Série, Número da Nota, Data, Valor Contabil, Base Cálculo
'Imposto, Isentas, Outras

'EXEMPLO DA FORMA DO RELATORIO

'CODIGO FISCAL: 1.12      - Compras Para Comercializacao
 
'TIPO SERIE  NÚMERO    DATA       VLR.CONTABIL     BASE CALC.     IMPOSTO       ISENTAS      OUTRAS
' E     U      1     12/12/200      15.892,40       7.622,22     1.120,00     2.500,00      300,00

'NO FIM O TOTAL

'                                     TOTAL           TOTAL         TOTAL         TOTAL       TOTAL
'                                   15.892,40       7.622,22     1.120,00     2.500,00      300,00

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'Eventos dos Browses
Private WithEvents objEventoNatOpDe As AdmEvento
Attribute objEventoNatOpDe.VB_VarHelpID = -1
Private WithEvents objEventoNatOpAte As AdmEvento
Attribute objEventoNatOpAte.VB_VarHelpID = -1

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
            
    Set objEventoNatOpDe = New AdmEvento
    Set objEventoNatOpAte = New AdmEvento
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169718)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_PreencherParametrosNaTela

    'Limpa a tela
    Call Limpar_Tela

    'Carrega Opções de Relatório
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 75196
    
    'pega Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro Then gError 75197
    
    Call DateParaMasked(DataInicial, CDate(sParam))
    
    'pega Data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 75198

    Call DateParaMasked(DataFinal, CDate(sParam))
        
    'Definitiva
    lErro = objRelOpcoes.ObterParametro("NTIPO", sParam)
    If lErro Then gError 75199
    
    OptionTipo(CInt(sParam)).Value = True
      
    lErro = objRelOpcoes.ObterParametro("NTRIBUTO", sParam)
    If lErro Then gError 75199
    
    OptionTributo(CInt(sParam)).Value = True
      
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNATOPDE", sParam)
    If lErro Then gError 75200
            
    NatOperacaoDe.Text = sParam
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("TNATOPATE", sParam)
    If lErro Then gError 75201
        
    NatOperacaoAte.Text = sParam
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 75196 To 75201
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169719)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoNatOpDe = Nothing
    Set objEventoNatOpAte = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 75202
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 75203

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 75202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 75203
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169720)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    ComboOpcoes.SetFocus
    OptionTipo(0).Value = True
    OptionTributo(0).Value = True
    
End Sub

Private Function Formata_E_Critica_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
        
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 78085
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 78086
        
    'Data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 75204
    End If
                   
    'Natureza de operação inicial não pode ser maior que o final
    If Trim(NatOperacaoDe.Text) <> "" And Trim(NatOperacaoAte.Text) <> "" Then
        If CLng(NatOperacaoDe.Text) > CLng(NatOperacaoAte.Text) Then gError 75205
    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 75204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
                       
        Case 75205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INICIAL_MAIOR", gErr)
            NatOperacaoDe.SetFocus
        
        Case 78085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus
            
        Case 78086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
            DataFinal.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169721)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub LabelNatOperacaoDe_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection

    'Adiciona os limites de Natureza
    'Ambas
    If OptionTipo(0).Value = True Then
            colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
            colSelecao.Add NATUREZA_SAIDA_COD_FINAL
    
    'Entrada
    ElseIf OptionTipo(1).Value = True Then
            colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
            colSelecao.Add NATUREZA_ENTRADA_COD_FINAL
    
    'Saida
    ElseIf OptionTipo(2).Value = True Then
            colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
            colSelecao.Add NATUREZA_SAIDA_COD_FINAL
    End If
    
    'Se NaturezaOPDe estiver preenchida coloca no Obj
    If Len(Trim(NatOperacaoDe.ClipText)) > 0 Then objNaturezaOp.sCodigo = NatOperacaoDe.Text
    
    'Chama a Tela de browse de NaturezaOp
    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNatOpDe)

End Sub

Private Sub objEventoNatOpDe_evSelecao(obj1 As Object)

Dim objNaturezaOp As ClassNaturezaOp

    Set objNaturezaOp = obj1

    'Coloca a natureza Operação na tela
    NatOperacaoDe.Text = objNaturezaOp.sCodigo

    Me.Show

End Sub

Private Sub LabelNatOperacaoAte_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection

    'Adiciona os limites de Natureza
    'Ambas
    If OptionTipo(0).Value = True Then
            colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
            colSelecao.Add NATUREZA_SAIDA_COD_FINAL
    
    'Entrada
    ElseIf OptionTipo(1).Value = True Then
            colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
            colSelecao.Add NATUREZA_ENTRADA_COD_FINAL
    
    'Saida
    ElseIf OptionTipo(2).Value = True Then
            colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
            colSelecao.Add NATUREZA_SAIDA_COD_FINAL
    End If
    
    'Se NaturezaOPAte estiver preenchida coloca no Obj
    If Len(Trim(NatOperacaoAte.ClipText)) > 0 Then objNaturezaOp.sCodigo = NatOperacaoAte.Text
    
    'Chama a Tela de browse de NaturezaOp
    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNatOpAte)

End Sub

Private Sub objEventoNatOpAte_evSelecao(obj1 As Object)

Dim objNaturezaOp As ClassNaturezaOp

    Set objNaturezaOp = obj1

    'Coloca a natureza Operação na tela
    NatOperacaoAte.Text = objNaturezaOp.sCodigo

    Me.Show

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub NatOperacaoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNaturezaOp As New ClassNaturezaOp

On Error GoTo Erro_NaturezaOp_Validate
    
    'Verifica se Natureza de Operação foi preenchida
    If Len(Trim(NatOperacaoDe.ClipText)) = 0 Then Exit Sub

    objNaturezaOp.sCodigo = NatOperacaoDe.Text
    
    'Lê a Natureza de Operação
    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
    If lErro <> SUCESSO And lErro <> 17958 Then gError 75206

    'Se não encontrou a Natureza de Operação --> erro
    If lErro = 17958 Then gError 75207
    
    Exit Sub

Erro_NaturezaOp_Validate:

    Cancel = True

    Select Case gErr

        Case 75206

        Case 75207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", gErr, NatOperacaoDe.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169722)

    End Select

    Exit Sub

End Sub

Private Sub NatOperacaoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNaturezaOp As New ClassNaturezaOp

On Error GoTo Erro_NaturezaOp_Validate
    
    'Verifica se Natureza de Operação foi preenchida
    If Len(Trim(NatOperacaoAte.ClipText)) = 0 Then Exit Sub

    objNaturezaOp.sCodigo = NatOperacaoAte.Text
    
    'Lê a Natureza de Operação
    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
    If lErro <> SUCESSO And lErro <> 17958 Then gError 75208

    'Se não encontrou a Natureza de Operação --> erro
    If lErro = 17958 Then gError 75209
    
    Exit Sub

Erro_NaturezaOp_Validate:

    Cancel = True

    Select Case gErr

        Case 75208

        Case 75209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", gErr, NatOperacaoAte.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169723)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim sTipo As String, sTributo As String

On Error GoTo Erro_PreencherRelOp

    'Critica as datas
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 75210
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 75211
               
    lErro = objRelOpcoes.IncluirParametro("TNATOPDE", CStr(NatOperacaoDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 75212
    
    lErro = objRelOpcoes.IncluirParametro("TNATOPATE", CStr(NatOperacaoAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 75213
        
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 75214
    
    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 75215
    
    'verifica opção de ordenação selecionada
    For iIndice = 0 To 2
        If OptionTipo(iIndice).Value = True Then sTipo = CStr(iIndice)
    Next
            
    lErro = objRelOpcoes.IncluirParametro("NTIPO", sTipo)
    If lErro <> AD_BOOL_TRUE Then gError 75216
            
    For iIndice = 0 To 1
        If OptionTributo(iIndice).Value = True Then sTributo = CStr(iIndice)
    Next
            
    lErro = objRelOpcoes.IncluirParametro("NTRIBUTO", sTributo)
    If lErro <> AD_BOOL_TRUE Then gError 75216
            
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sTipo)
    If lErro <> SUCESSO Then gError 75217
            
    If OptionTributo(0).Value = True Then
        gobjRelatorio.sNomeTsk = "RegESNAT"
    Else
        gobjRelatorio.sNomeTsk = "RegESNATI"
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 75210 To 75217

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169724)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sTipo As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If Trim(NatOperacaoDe.Text) <> "" Then sExpressao = "NatOperacao >= @TNATOPDE"

   If Trim(NatOperacaoAte.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NatOperacao <= @TNATOPATE"

    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169725)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 75218

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 75219

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call Limpar_Tela
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 75218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 75219

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169726)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 75220

    Call gobjRelatorio.Executar_Prossegue2(Me)
        
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 75220
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169727)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 75221

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 75222

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 75223

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 75221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 75222, 75223

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169728)

    End Select

    Exit Sub

End Sub

Private Sub NatOperacaoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NatOperacaoDe)
    
End Sub

Private Sub NatOperacaoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NatOperacaoAte)
    
End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 75224

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 75224

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169729)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 75225

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 75225

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169730)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 75226

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 75226
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169731)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 75227

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 75227
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169732)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 75228

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 75228
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169733)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 75229

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 75229
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169734)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NatOperacaoDe Then
            Call LabelNatOperacaoDe_Click
        ElseIf Me.ActiveControl Is NatOperacaoAte Then
            Call LabelNatOperacaoAte_Click
        End If
        
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Lista de Reg. de Entrada/Saída p/ Nat. de Operação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpLancNatOperacao"
    
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

Public Sub Unload(objme As Object)
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

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub





Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub LabelNatOperacaoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNatOperacaoAte, Source, X, Y)
End Sub

Private Sub LabelNatOperacaoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNatOperacaoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNatOperacaoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNatOperacaoDe, Source, X, Y)
End Sub

Private Sub LabelNatOperacaoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNatOperacaoDe, Button, Shift, X, Y)
End Sub

