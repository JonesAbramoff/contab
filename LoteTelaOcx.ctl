VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl LoteTelaOcx 
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   7830
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1665
      Picture         =   "LoteTelaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Numeração Automática"
      Top             =   735
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "LoteTelaOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "LoteTelaOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "LoteTelaOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "LoteTelaOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton LancamentoLote 
      Caption         =   "Lançamentos do Lote"
      Height          =   765
      Left            =   5370
      Picture         =   "LoteTelaOcx.ctx":0A7E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3765
      Width           =   2100
   End
   Begin VB.ComboBox Periodo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3570
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   255
      Width           =   1590
   End
   Begin VB.ComboBox Exercicio 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1035
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   255
      Width           =   1590
   End
   Begin MSMask.MaskEdBox IdLoteExterno 
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   4035
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin VB.Frame SSFrame1 
      Caption         =   "Valores Atuais"
      Height          =   1650
      Left            =   135
      TabIndex        =   14
      Top             =   1950
      Width           =   7515
      Begin VB.CommandButton Botao_Recalcular 
         Caption         =   "  Recalcular Totais do Lote"
         Height          =   810
         Left            =   5670
         Picture         =   "LoteTelaOcx.ctx":0EC0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total dos Créditos:"
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   420
         Width           =   2115
      End
      Begin VB.Label TotCredAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         Height          =   300
         Left            =   2475
         TabIndex        =   16
         Top             =   375
         Width           =   1575
      End
      Begin VB.Label TotDebAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         Height          =   300
         Left            =   2475
         TabIndex        =   17
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label NumDocAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   300
         Left            =   2460
         TabIndex        =   18
         Top             =   1215
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total dos Débitos:"
         Height          =   195
         Left            =   345
         TabIndex        =   19
         Top             =   840
         Width           =   2070
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Número de Documentos:"
         Height          =   195
         Left            =   315
         TabIndex        =   20
         Top             =   1260
         Width           =   2100
      End
   End
   Begin MSMask.MaskEdBox TotDocInf 
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   1335
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   315
      Left            =   1035
      TabIndex        =   2
      Top             =   720
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSMask.MaskEdBox NumDocInf 
      Height          =   315
      Left            =   6810
      TabIndex        =   5
      Top             =   1350
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
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
      Format          =   "#,##0"
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label10 
      Caption         =   "Identificador de Lote Externo:"
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
      Height          =   255
      Left            =   180
      TabIndex        =   21
      Top             =   4080
      Width           =   2595
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Número de Documentos:"
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
      Left            =   4650
      TabIndex        =   22
      Top             =   1380
      Width           =   2100
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Valor Total dos Documentos:"
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
      Left            =   225
      TabIndex        =   23
      Top             =   1380
      Width           =   2475
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
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
      Left            =   2760
      TabIndex        =   24
      Top             =   315
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Exercício:"
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
      Left            =   105
      TabIndex        =   25
      Top             =   300
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Origem:"
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
      Left            =   2850
      TabIndex        =   26
      Top             =   780
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lote:"
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
      Left            =   540
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   27
      Top             =   780
      Width           =   450
   End
   Begin VB.Label Origem 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contabilidade"
      Height          =   315
      Left            =   3570
      TabIndex        =   28
      Top             =   720
      Width           =   1590
   End
End
Attribute VB_Name = "LoteTelaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
Dim sSiglaModulo As String

Dim iAlterado As Integer
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim objLote As New ClassLote

On Error GoTo Erro_BotaoProxNum_Click

    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    
    'Mostra número do proximo lote disponível
    lErro = CF("Lote_Automatico", objLote)
    If lErro <> SUCESSO Then Error 57516
    
    Lote.Text = CStr(objLote.iLote)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57516
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162530)
    
    End Select

    Exit Sub

End Sub

Private Sub Botao_Recalcular_Click()

Dim lErro As Long
Dim objLote As New ClassLote
Dim iLote As Integer
Dim iTotaisIguais As Integer
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_Botao_Recalcular_Click

    'Carrega em memória os dados da tela
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    
    If Len(Trim(Lote.Text)) > 0 Then
        objLote.iLote = CInt(Lote.Text)
    End If
    
    objLote.dTotCre = CDbl(TotCredAtual.Caption)
    objLote.dTotDeb = CDbl(TotDebAtual.Caption)
    objLote.iNumDocAtual = CInt(NumDocAtual.Caption)
    
    'Se o número do lote não foi fornecido ==> erro
    If objLote.iLote = 0 Then Error 8033
    
    lErro = CF("LanPendente_Critica_TotaisLote", objLote, iTotaisIguais)
    If lErro <> SUCESSO Then Error 8034
    
    'Se os totais cadastrados diferentes dos totais calculados
    If iTotaisIguais = 1 Then
        
        'totais diferentes
        vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_ATUALIZACAO_TOTAIS", objLote.dTotCre, objLote.dTotDeb, objLote.iNumDocAtual)
        If vbMsgRet = vbYes Then
        
            'modifica os totais do lote no banco de dados
            lErro = CF("LotePendente_Grava_Totais", objLote)
            If lErro <> SUCESSO Then Error 8035
            
            'Atualiza tela
            TotCredAtual.Caption = Format(objLote.dTotCre, "Standard")
            TotDebAtual.Caption = Format(objLote.dTotDeb, "Standard")
            NumDocAtual.Caption = Format(objLote.iNumDocAtual, "##,##0")
            
        End If
    
    Else
        'totais iguais
        vbMsgRet = Rotina_Aviso(vbOKOnly, "AVISO_IGUALDADE_TOTAIS", objLote.dTotCre, objLote.dTotDeb, objLote.iNumDocAtual)
        
    End If
            
    Exit Sub
            
Erro_Botao_Recalcular_Click:

    Select Case Err
    
        Case 8033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)
            
        Case 8034, 8035
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162531)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim objLote As New ClassLote
Dim iLote As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click
 
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Origem diferente do modulo em questão
    If gobjColOrigem.Origem(Origem.Caption) <> sSiglaModulo Then Error 59510

    'carrega em memória os dados da tela
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    objLote.iLote = StrParaInt(Lote.Text)
    
    'se o número do lote não foi fornecido ==> erro
    If objLote.iLote = 0 Or Len(Trim(Lote.ClipText)) = 0 Then Error 5405
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_LOTE")
    
    If vbMsgRes = vbYes Then
    
        'exclui o lote do banco de dados
        lErro = CF("LotePendente_Exclui", objLote)
        If lErro <> SUCESSO Then Error 5406
    
        'limpa o conteudo dos campos da tela
        lErro = Limpa_Tela(Me)
    
        iAlterado = 0
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
            
        Case 59510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_DIFERENTE", Err)

        Case 5405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)
            Lote.SetFocus
        
        Case 5406
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162532)
        
    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objLote As New ClassLote
Dim iExercicioFechado As Integer

On Error GoTo Erro_BotaoGravar_Click
 
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 10171
    
    'limpa o conteudo dos campos da tela
    Call Limpa_LoteTela
    
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:
    
    Select Case Err
    
        Case 10171
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162533)
        
    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iExercicioFechado As Integer
Dim objLote As New ClassLote

On Error GoTo Erro_Gravar_Registro
 
    GL_objMDIForm.MousePointer = vbHourglass
    
    Call Move_Tela_Memoria(objLote)
 
    'se o número do lote não foi fornecido ==> erro
    If objLote.iLote = 0 Then Error 5401
    
    If objLote.iExercicio = 0 Then Error 11784
          
    'Origem diferente do modulo em questão
    If gobjColOrigem.Origem(Origem.Caption) <> sSiglaModulo Then Error 59509
    
    lErro = Trata_Alteracao(objLote, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    If lErro <> SUCESSO Then Error 32300

    'inclui o lote no banco de dados
    lErro = CF("LotePendente_Grava", objLote)
    If lErro <> SUCESSO Then Error 5653
    
    Gravar_Registro = SUCESSO
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:
    
    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
            
        Case 32300
            
        Case 59509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_DIFERENTE", Err)

        Case 5401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)
        
        Case 5653
        
        Case 11784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_SELECIONADO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162534)
        
    End Select

    Exit Function

End Function

Private Sub Move_Tela_Memoria(objLote As ClassLote)
    
    'carrega em memória os dados da tela
    If Exercicio.ListIndex = -1 Then
        objLote.iExercicio = 0
    Else
        objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    End If
    
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    
    If Len(Trim(TotDocInf.Text)) = 0 Then
        objLote.dTotInf = 0
    Else
        objLote.dTotInf = StrParaDbl(TotDocInf.Text)
    End If
        
    If Len(Trim(NumDocInf.Text)) = 0 Then
        objLote.iNumDocInf = 0
    Else
        objLote.iNumDocInf = CInt(NumDocInf.Text)
    End If
    
    objLote.sIdOriginal = IdLoteExterno.Text
    
    If Len(Trim(Lote.Text)) = 0 Then
        objLote.iLote = 0
    Else
        objLote.iLote = CInt(Lote.Text)
    End If
    
End Sub

Private Sub BotaoLimpar_Click()
    
Dim lErro As Long
Dim objLote As New ClassLote

On Error GoTo Erro_BotaoLimpar_Click
    
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 9638

    Call Limpa_LoteTela
    
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 9638
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162535)
        
    End Select

    Exit Sub
    
End Sub

Private Sub Exercicio_Click()
    
Dim iExercicio As Integer
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo
Dim lErro As Long
Dim iIndice As Integer
    
On Error GoTo Erro_Exercicio_Click
    
    If Exercicio.ListIndex = -1 Then Exit Sub
        
    iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    
    'inicializar os periodos do exercicio atual
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 10934
        
    Periodo.Clear
    
    For iIndice = 1 To colPeriodos.Count
        Set objPeriodo = colPeriodos.Item(iIndice)
        Periodo.AddItem objPeriodo.sNomeExterno
        Periodo.ItemData(Periodo.NewIndex) = objPeriodo.iPeriodo
    Next
    
    'seleciona o primeiro periodo da combobox
    Periodo.ListIndex = 0
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_Exercicio_Click:

    Select Case Err
    
        Case 10934
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162536)
        
    End Select
    
    Exit Sub
        
End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim iLote As Integer
Dim lErro As Long
Dim sDescricao As String
Dim objExercicio As ClassExercicio
Dim colExerciciosAbertos As New Collection
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_Form_Load
    
    Set objEventoLote = New AdmEvento
    
    'ler os exercicios não fechados
    lErro = CF("Exercicios_Nao_Fechados_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 5984
    
    For iIndice = 1 To colExerciciosAbertos.Count
        Set objExercicio = colExerciciosAbertos.Item(iIndice)
        Exercicio.AddItem objExercicio.sNomeExterno
        Exercicio.ItemData(Exercicio.NewIndex) = objExercicio.iExercicio
    Next
        
    'mostra o exercicio atual
    For iIndice = 0 To Exercicio.ListCount - 1
        If Exercicio.ItemData(iIndice) = giExercicioAtual Then
            Exercicio.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'mostra o periodo atual
    For iIndice = 0 To Periodo.ListCount - 1
        If Periodo.ItemData(iIndice) = giPeriodoAtual Then
            Periodo.ListIndex = iIndice
            Exit For
        End If
    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
        
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 5984
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162537)
        
    End Select
    
    iAlterado = 0
    
    Exit Sub
        
End Sub

Function Trata_Parametros(Optional objLote As ClassLote, Optional vSiglaModulo As Variant) As Long

Dim lErro As Long
Dim iLote As Integer
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not IsMissing(vSiglaModulo) Then
        sSiglaModulo = vSiglaModulo
    Else
        sSiglaModulo = MODULO_CONTABILIDADE
    End If
    
    Origem.Caption = gobjColOrigem.Descricao(sSiglaModulo)

    'Se foi passado um lote como parametro, exibir seus dados
    If Not (objLote Is Nothing) Then
    
        lErro = Traz_Lote_Tela(objLote)
        If lErro <> SUCESSO Then Error 5834
        
    Else
        
        'mostra o Exercicio
        For iIndice = 0 To Exercicio.ListCount - 1
            If Exercicio.ItemData(iIndice) = giExercicioAtual Then
                Exercicio.ListIndex = iIndice
                Exit For
            End If
        Next
            
        'mostra o periodo
        For iIndice = 0 To Periodo.ListCount - 1
            If Periodo.ItemData(iIndice) = giPeriodoAtual Then
                Periodo.ListIndex = iIndice
                Exit For
            End If
        Next
        
        iAlterado = 0
        
    End If
    'alterado por cyntia
    If Exercicio.ListIndex = -1 And Exercicio.ListCount <> 0 Then Exercicio.ListIndex = 0
        
    If Periodo.ListIndex = -1 And Periodo.ListCount <> 0 Then Periodo.ListIndex = 0
        
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 5834
            If Exercicio.ListIndex = -1 Then Exercicio.ListIndex = 0
            If Periodo.ListIndex = -1 Then Periodo.ListIndex = 0
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162538)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Traz_Lote_Tela(objLote As ClassLote) As Long

Dim sDescricao As String
Dim sExercicio As String
Dim lErro As Long
Dim iIndice As Integer
Dim iLoteAtualizado As Integer

On Error GoTo Erro_Traz_Lote_Tela

    Call Limpa_LoteTela

    Origem.Caption = gobjColOrigem.Descricao(objLote.sOrigem)

    'verifica se o lote  está atualizado
    lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 5999
    
    'Se é um lote que já foi contabilizado, não pode sofrer alteração
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 9000

    'le o lote contido em objLote
    lErro = CF("LotePendente_Le", objLote)
    If lErro <> SUCESSO And lErro <> 5435 Then Error 5436
    
    'move os dados para a tela
    sDescricao = gobjColOrigem.Descricao(objLote.sOrigem)
    
    'mostra o Exercicio
    For iIndice = 0 To Exercicio.ListCount - 1
        If Exercicio.ItemData(iIndice) = objLote.iExercicio Then
            Exercicio.ListIndex = iIndice
            Exit For
        End If
    Next
        
    'mostra o periodo
    For iIndice = 0 To Periodo.ListCount - 1
        If Periodo.ItemData(iIndice) = objLote.iPeriodo Then
            Periodo.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Lote.Text = CStr(objLote.iLote)
    
    'se o lote está cadastrado, coloca o restante das informações na tela
    If lErro = SUCESSO Then
            
        TotDocInf = Format(objLote.dTotInf, "Fixed")
        NumDocInf = CStr(objLote.iNumDocInf)
        TotCredAtual.Caption = Format(objLote.dTotCre, "Standard")
        TotDebAtual.Caption = Format(objLote.dTotDeb, "Standard")
        NumDocAtual.Caption = Format(objLote.iNumDocAtual, "##,##0")
        IdLoteExterno = objLote.sIdOriginal
        
    End If
            
    iAlterado = 0
    
    Traz_Lote_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Lote_Tela:

    Traz_Lote_Tela = Err

    Select Case Err
    
        Case 5436, 5999
        
        Case 9000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_M_LOTE_LOTE_ATUALIZADO", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162539)
        
    End Select
    
    iAlterado = 0
    
    Exit Function
        
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
            
    Set objEventoLote = Nothing
    
End Sub

Private Sub IdLoteExterno_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Label1_Click()

Dim objLote As New ClassLote
Dim colSelecao As New Collection

    Call Move_Tela_Memoria(objLote)
    
    If sSiglaModulo = MODULO_CONTABILIDADE Then
        
        colSelecao.Add giFilialEmpresa
        colSelecao.Add 0
            
        Call Chama_Tela("LotePendenteCTBLista", colSelecao, objLote, objEventoLote)
    
    Else
        
        colSelecao.Add objLote.sOrigem
        colSelecao.Add giFilialEmpresa
        colSelecao.Add 0
            
        Call Chama_Tela("LotePendenteLista", colSelecao, objLote, objEventoLote)
    
    End If
    
End Sub

Private Sub LancamentoLote_Click()
    
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_LancamentoLote_Click

    If Len(Lote.Text) = 0 Then Error 59628

    colSelecao.Add gobjColOrigem.Origem(Origem.Caption)
    colSelecao.Add Exercicio.ItemData(Exercicio.ListIndex)
    colSelecao.Add Periodo.ItemData(Periodo.ListIndex)
    colSelecao.Add CInt(Lote.Text)

    Call Chama_Tela("LanPendenteLista_Lote", colSelecao)
    
    Exit Sub
    
Erro_LancamentoLote_Click:

    Select Case Err
        
        Case 59628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_PREENCHIDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162540)

    End Select

    Exit Sub
    
End Sub

Private Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Lote_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Lote, iAlterado)

End Sub

Private Sub Lote_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim objLote As New ClassLote
Dim iLoteAtualizado As Integer

On Error GoTo Erro_Lote_Validate
 
    'se o número do lote não foi fornecido, nao critica
    If Len(Lote.Text) = 0 Then Exit Sub
    If CInt(Lote.Text) = 0 Then Exit Sub
    
    'carrega em memória os dados da tela
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    objLote.iLote = CInt(Lote.Text)
    
    'verifica se o lote  está atualizado
    lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 5991
    
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 5432
        
    Exit Sub
    
Erro_Lote_Validate:

    Cancel = True

    
    Select Case Err
    
        Case 5432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_M_LOTE_LOTE_ATUALIZADO", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            
        Case 5991
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162541)
        
    End Select

    Exit Sub
    
End Sub

Private Sub NumDocInf_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NumDocInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumDocInf, iAlterado)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
    
Dim objLote As ClassLote
Dim lErro As Long

On Error GoTo Erro_objEventoLote_evSelecao

    Set objLote = obj1
    
    lErro = Traz_Lote_Tela(objLote)
    If lErro <> SUCESSO Then Error 9134
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    Me.Show
    
    Exit Sub
    
Erro_objEventoLote_evSelecao:

    Select Case Err
    
        Case 9134
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162542)
        
    End Select

    Exit Sub
        
End Sub

Private Sub Periodo_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TotDocInf_Change()

   iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TotDocInf_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long

On Error GoTo Erro_TotDocInf_Validate

    If Len(TotDocInf.Text) > 0 Then
    
        lErro = Valor_Critica(TotDocInf.Text)
        If lErro <> SUCESSO Then Error 5937
        
    End If

    Exit Sub

Erro_TotDocInf_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 5937
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162543)
            
    End Select
        
    Exit Sub

End Sub

Sub Limpa_LoteTela()

Dim lErro As Long
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    Call Limpa_Tela(Me)
    
    TotCredAtual.Caption = "0,00"
    TotDebAtual.Caption = "0,00"
    NumDocAtual.Caption = "0"
    Lote.Text = ""
    Origem.Caption = gobjColOrigem.Descricao(sSiglaModulo)
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objLote As New ClassLote
Dim colLancamento_Detalhe As New Collection

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "LotePendente"
            
    'Le os dados da Tela de Lotes
    Call Move_Tela_Memoria(objLote)
  
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objLote.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Origem", objLote.sOrigem, STRING_ORIGEM, "Origem"
    colCampoValor.Add "Lote", objLote.iLote, 0, "Lote"
    colCampoValor.Add "Exercicio", objLote.iExercicio, 0, "Exercicio"
    colCampoValor.Add "Periodo", objLote.iPeriodo, 0, "Periodo"
    
    'Exemplo de Filtro para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
'    If sSiglaModulo = MODULO_CONTABILIDADE Then colSelecao.Add "Origem", OP_IGUAL, MODULO_CONTABILIDADE
    If sSiglaModulo = MODULO_CONTASAPAGAR Then colSelecao.Add "Origem", OP_IGUAL, MODULO_CONTASAPAGAR
    If sSiglaModulo = MODULO_CONTASARECEBER Then colSelecao.Add "Origem", OP_IGUAL, MODULO_CONTASARECEBER
    If sSiglaModulo = MODULO_TESOURARIA Then colSelecao.Add "Origem", OP_IGUAL, MODULO_TESOURARIA
    If sSiglaModulo = MODULO_FATURAMENTO Then colSelecao.Add "Origem", OP_IGUAL, MODULO_FATURAMENTO
    If sSiglaModulo = MODULO_ESTOQUE Then colSelecao.Add "Origem", OP_IGUAL, MODULO_ESTOQUE
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162544)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objLote As New ClassLote

On Error GoTo Erro_Tela_Preenche

    objLote.iLote = colCampoValor.Item("Lote").vValor

    If objLote.iLote <> 0 Then
    
        objLote.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objLote.sOrigem = colCampoValor.Item("Origem").vValor
        objLote.iPeriodo = colCampoValor.Item("Periodo").vValor
        objLote.iExercicio = colCampoValor.Item("Exercicio").vValor
       
        lErro = Traz_Lote_Tela(objLote)
        If lErro <> SUCESSO Then Error 14974

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 14974

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162545)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_MANUTENCAO_LOTES
    Set Form_Load_Ocx = Me
    Caption = "Manutenção de Lote"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LoteTela"
    
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
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Lote Then
            Call Label1_Click
        End If
    
    End If

End Sub





Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub TotCredAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotCredAtual, Source, X, Y)
End Sub

Private Sub TotCredAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotCredAtual, Button, Shift, X, Y)
End Sub

Private Sub TotDebAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotDebAtual, Source, X, Y)
End Sub

Private Sub TotDebAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotDebAtual, Button, Shift, X, Y)
End Sub

Private Sub NumDocAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumDocAtual, Source, X, Y)
End Sub

Private Sub NumDocAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumDocAtual, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Origem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Origem, Source, X, Y)
End Sub

Private Sub Origem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Origem, Button, Shift, X, Y)
End Sub

