VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{454464FA-6BBA-4224-B6CD-4A4CA1778A0F}#1.0#0"; "AdmCalendar.ocx"
Begin VB.Form CotacaoMoeda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cotações"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   Icon            =   "CotacaoMoeda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin AdmCalendar.Calendar Calendar1 
      Height          =   3072
      Left            =   192
      TabIndex        =   0
      Top             =   660
      Width           =   3192
      _ExtentX        =   5636
      _ExtentY        =   5424
      Day             =   1
      Month           =   1
      Year            =   1999
      BeginProperty DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cotação "
      Height          =   3132
      Left            =   3465
      TabIndex        =   5
      Top             =   588
      Width           =   3705
      Begin VB.ComboBox Moeda 
         Height          =   288
         Left            =   888
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   390
         Width           =   2652
      End
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Trazer"
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
         Left            =   2505
         TabIndex        =   2
         Top             =   996
         Width           =   1035
      End
      Begin VB.CommandButton BotaoRepete 
         Caption         =   "Repetir Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2475
         TabIndex        =   4
         Top             =   1620
         Width           =   1065
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   216
         TabIndex        =   3
         Top             =   1728
         Width           =   2028
         _ExtentX        =   3598
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   732
         TabIndex        =   1
         Top             =   1008
         Width           =   1548
         _ExtentX        =   2725
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moeda:"
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
         Index           =   1
         Left            =   210
         TabIndex        =   15
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label1 
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
         Height          =   228
         Index           =   0
         Left            =   216
         TabIndex        =   13
         Top             =   1044
         Width           =   528
      End
      Begin VB.Label LabelCotacaoAnterior 
         BorderStyle     =   1  'Fixed Single
         Height          =   336
         Left            =   216
         TabIndex        =   12
         Top             =   2532
         Width           =   2028
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cotação Anterior:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   216
         TabIndex        =   11
         Top             =   2304
         Width           =   1500
      End
      Begin VB.Label Label4 
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
         Height          =   228
         Left            =   216
         TabIndex        =   10
         Top             =   1488
         Width           =   528
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   555
      Left            =   5544
      ScaleHeight     =   495
      ScaleWidth      =   1545
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   24
      Width           =   1605
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "CotacaoMoeda.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "CotacaoMoeda.frx":02D4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "CotacaoMoeda.frx":0452
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
End
Attribute VB_Name = "CotacaoMoeda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iMoedaAnterior As Integer

'esta tela faz a manutencao da tabela de feriados
'todo feriado tem que ter uma descricao

'para setar ou obter a data corrente do calendario: Calendar1.Value
'para repintar: Calendar1.Refresh
'Calendar1.DayBold: para colocar ou retirar negrito


Const ERRO_DATA_SEM_PREENCHIMENTO = 0
'A Data não foi informada

Const ERRO_VALOR_NAO_PREENCHIDO = 0
'O valor não foi infomardo

Const ERRO_LEITURA_COTACOESMOEDA = 0
'Erro na leitura da tabela CotacoesMoeda

'????? não colocar o obj como parâmetro. Coloque um nome sugestivo p\ indicar o q a msg espera - Ok
Const ERRO_INSERCAO_COTACOESMOEDA = 0 'Data, Valor
'Erro na inclusão da cotação relativa a data %s de valor %s

Const ERRO_ATUALIZACAO_COTACOESMOEDA = 0 'Data, Valor
'Erro na Atualização da cotação relativa a data %s de valor %s

Const COTACAO_GRAVADA = 0
'Cotação gravada com sucesso.

'???? Acento de exclusão - Ok
Const AVISO_CONFIRMA_EXCLUSAO_COTACOESMOEDA = 0 'Data
'Confirma a exclusão da Cotação do dia %s

Const ERRO_COTACOESMOEDA_INEXISTENTE = 0 'Data
'Cotação inexistente para o dia %s

Const ERRO_LOCK_COTACOESMOEDA = 0
'Erro na tentativa de Lock na tabela CotacoesMoeda

Const ERRO_EXCLUSAO_COTACOESMOEDA = 0 'Data
'Erro na exclusão da Cotação do dia %s

Dim iAlterado As Integer

Public Sub Form_Load()
    
Dim lErro As Long
    
On Error GoTo Erro_Form_Load
    
    Calendar1.Value = gdtDataHoje
    Data.Text = Format(Calendar1.Value, "dd/mm/yy")
    
    Call Preenche_Combo_Moeda
    
    iMoedaAnterior = 0
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155221)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCotacaoMoeda As New ClassCotacaoMoeda

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 80249
    
    'Carrega o objCotacaoMoeda
    objCotacaoMoeda.dtData = StrParaDate(Data.Text) '???? Por que não está convertendo para data? - OK
    objCotacaoMoeda.iMoeda = Codigo_Extrai(Moeda.List(Moeda.ListIndex))
    
    'Verifica se existe um valor de cotação para a data informada
    lErro = CF("CotacaoMoeda_Le", objCotacaoMoeda)
    If lErro <> SUCESSO And lErro <> 80257 Then gError 80261
    
    'Erro, a cotação informada não existe
    If lErro <> SUCESSO Then gError 80262
    
    'Envia mensagem de confirmação de exclusão para o usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_COTACOESMOEDA", objCotacaoMoeda.dtData)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui a CotacaoMoeda informada
    lErro = CF("CotacaoMoeda_Exclui", objCotacaoMoeda)
    If lErro <> SUCESSO Then gError 80263
    
    'Carrega a tela com CotaçãoMoeda do dia
    Calendar1.Value = gdtDataHoje
    Data.Text = Format(Calendar1.Value, "dd/mm/yy")
    
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        
        Case 80249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
                
        Case 80261, 80263
        
        Case 80262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COTACOESMOEDA_INEXISTENTE", gErr, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155222)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Chama função gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 80224
    
    'Envia mensagem confirmando gravação
    Call Rotina_Aviso(vbOKOnly, "COTACAO_GRAVADA")
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 80224

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155223)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRepete_Click()

    If Len(Trim(LabelCotacaoAnterior.Caption)) > 0 Then
        Valor.Text = LabelCotacaoAnterior.Caption
    End If

End Sub

Private Sub BotaoTrazer_Click()

Dim dtData As Date

    'Verifica se a da esta preenchida
    If Len(Trim(Data.Text)) > 0 Then
        
        Calendar1.Value = DateValue(Data.Text)
        dtData = CDate(Data.Text)
        Call Traz_CotacaoMoeda_Tela(dtData)
        
    End If
    
End Sub

Private Sub Calendar1_DateChange(ByVal OldDate As Date, ByVal NewDate As Date)
    
    Data.Text = Format(NewDate, "dd/mm/yy")

    Call Traz_CotacaoMoeda_Tela(NewDate)
    
End Sub

Private Sub Calendar1_WillChangeDate(ByVal NewDate As Date, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Calendar1_WillChangeDate

    If iAlterado <> 0 Then
    
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then gError 80228

        Calendar1.Refresh
        
        iAlterado = 0
            
    End If
    
    Exit Sub
    
Erro_Calendar1_WillChangeDate:
    
    Select Case gErr
    
        Case 80228
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155224)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 80226

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr
        
        Case 80226

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155225)

    End Select

    Exit Sub

End Sub

Private Sub Form_Activate()

    Calendar1.Refresh
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCotacaoMoeda As New ClassCotacaoMoeda
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 80238
    
    'Verifica se o valor foi preenchido
    If Len(Trim(Valor.Text)) = 0 Then gError 80239
    
    'Carrega o obj com os valores a serem passados como parametro
    objCotacaoMoeda.dtData = CDate(Data.Text)
    objCotacaoMoeda.dValor = StrParaDbl(Valor.Text)
    objCotacaoMoeda.iMoeda = Codigo_Extrai(Moeda.List(Moeda.ListIndex))
    
    If objCotacaoMoeda.dValor < -DELTA_VALORMONETARIO Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_VALOR_NEGATIVO")
        If vbMsgRes <> vbYes Then gError 130399
    End If
    
    'Chama função de gravação com obj carregado
    lErro = CF("CotacaoMoeda_Grava", objCotacaoMoeda)
    If lErro <> SUCESSO Then gError 80240
        
    iAlterado = 0
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 80238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
            Data.SetFocus
        
        Case 80239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO", gErr)

        Case 80240, 130399
        
        Case Else
            
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155226)

    End Select

    Exit Function

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
'''' Rotinas comentadas pertencem a GSilva
'Dim lErro As Long
'
'On Error GoTo Erro_Form_QueryUnload

'    If Len(Trim(Valor.Text)) = 0 Then gError 84725
      
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode)
      
    Exit Sub
    
'Erro_Form_QueryUnload:
'
'     Select Case gErr
'
'        Case 84725
'            lErro = Rotina_Aviso(vbYesNo, "AVISO_MANTER_COTACAO_ANTERIOR")
'            If lErro = 6 Then 'yes
'                Valor.Text = LabelCotacaoAnterior.Caption
'                Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode)
'            End If
'            If lErro = 7 Then 'no
'                Valor.SetFocus
'                Cancel = True
'            End If
'
'     End Select

End Sub

Private Sub Moeda_Click()

Dim lErro As Long

On Error GoTo Erro_Moeda_Click

    'Se trocou a moeda => Carrega a nova cotacao
    If iMoedaAnterior <> Codigo_Extrai(Moeda.List(Moeda.ListIndex)) Then
    
        iAlterado = REGISTRO_ALTERADO
        
        'Limpa o valor
        Valor.PromptInclude = False
        Valor.Text = ""
        Valor.PromptInclude = True
        
        'Se a data estiver preenchida => Traz a cotacao
        If Len(Trim(Data.ClipText)) > 0 Then
                
            'Traz a nova cotacao
            lErro = Traz_CotacaoMoeda_Tela(Data.Text)
            If lErro <> SUCESSO Then gError 108850
            
        End If
        
        iMoedaAnterior = Codigo_Extrai(Moeda.List(Moeda.ListIndex))
        
    End If

    Exit Sub

Erro_Moeda_Click:

    Select Case gErr
    
        Case 108850
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155227)
    
    End Select

End Sub

Private Sub Moeda_GotFocus()
    iMoedaAnterior = Codigo_Extrai(Moeda.List(Moeda.ListIndex))
End Sub

'Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label1, Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
'End Sub
'

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se o valor foi preenchido
    If Len(Trim(Valor.Text)) = 0 Then Exit Sub
    
    'Verifica se o valor foi preenchido
    lErro = Valor_Positivo_Critica(Valor.Text)
    If lErro <> SUCESSO Then gError 80227
    
    Valor.Text = Format(Valor.Text, "#.0000")
    
    Exit Sub

Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 80227
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155228)
        
    End Select
    
    Exit Sub

End Sub

Function Traz_CotacaoMoeda_Tela(ByVal NewDate As Date) As Long
'Função responsável pelo preenchimento da tela

Dim lErro As Long
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_Traz_CotacaoMoeda_Tela

    'Carrega objCotacao
    objCotacao.dtData = NewDate
    '??? Coloquei isso pra nao quando a combo nao estiver preenchida
    objCotacao.iMoeda = IIf(Moeda.List(Moeda.ListIndex) = "", MOEDA_DOLAR, Codigo_Extrai(Moeda.List(Moeda.ListIndex)))

    'Chama função de leitura
    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
    If lErro <> SUCESSO Then gError 80229

    'Verifica se o objCotacao está preenchido
    If objCotacao.dValor = 0 Then
        Valor.Text = ""
    Else
        Valor.Text = Format(objCotacao.dValor, "#.0000")
    End If
    
    
    '?????? erro PENDENCIAS  - WILLIAM ID=4 - Ok
    'Verifica se o objCotacaoAnterior está preenchido
    If objCotacaoAnterior.dValor = 0 Then
        LabelCotacaoAnterior.Caption = ""
        Label3.Caption = "Cotação Anterior"
    Else
        LabelCotacaoAnterior.Caption = Format(objCotacaoAnterior.dValor, "#.0000")
        Label3.Caption = "Cotação dia " & Format(objCotacaoAnterior.dtData, "dd/mm/yy")
    End If
    
    iAlterado = 0
    
    Traz_CotacaoMoeda_Tela = SUCESSO

    Exit Function
    
Erro_Traz_CotacaoMoeda_Tela:

    Traz_CotacaoMoeda_Tela = gErr
    
    Select Case gErr
    
        Case 80229
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155229)
            
    End Select
    
    Exit Function
    
End Function

Function Limpa_CotacaoMoeda() As Long

    Call Limpa_Tela(CotacaoMoeda)

    LabelCotacaoAnterior.Caption = ""

End Function

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub


Private Sub Preenche_Combo_Moeda()

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Combo_Moeda

    'Preenche a combo de moedas com as moedas existentes no bd
    lErro = CF("Cod_Nomes_Le", "Moedas", "Codigo", "Nome", STRING_NOME_MOEDA, colCodigoNome)
    If lErro <> SUCESSO Then gError 108845
    
    For Each objCodigoNome In colCodigoNome

        'Insere na Combo de moedas, se nao for moeda = R$ (Real)
        If objCodigoNome.iCodigo <> MOEDA_REAL Then
            Moeda.AddItem objCodigoNome.iCodigo & SEPARADOR & objCodigoNome.sNome
        End If
        
    Next
    
    Moeda.ListIndex = 0
    
    Exit Sub
    
Erro_Preenche_Combo_Moeda:

    Select Case gErr
    
        Case 108845
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155230)
    
    End Select

End Sub
