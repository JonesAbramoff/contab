VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpLancTipoICMSOcx 
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ScaleHeight     =   3585
   ScaleWidth      =   7200
   Begin VB.Frame Frame2 
      Caption         =   "Aliquota"
      Height          =   1005
      Left            =   150
      TabIndex        =   15
      Top             =   2340
      Width           =   4785
      Begin MSMask.MaskEdBox AliquotaDe 
         Height          =   285
         Left            =   615
         TabIndex        =   3
         Top             =   375
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AliquotaAte 
         Height          =   285
         Left            =   2700
         TabIndex        =   4
         Top             =   375
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2295
         TabIndex        =   17
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   255
         TabIndex        =   16
         Top             =   435
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de ICMS"
      Height          =   1395
      Left            =   150
      TabIndex        =   12
      Top             =   810
      Width           =   4785
      Begin VB.ComboBox ComboICMSTipoAte 
         Height          =   315
         ItemData        =   "RelOpLancTipoICMS.ctx":0000
         Left            =   630
         List            =   "RelOpLancTipoICMS.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   810
         Width           =   3336
      End
      Begin VB.ComboBox ComboICMSTipoDe 
         Height          =   315
         ItemData        =   "RelOpLancTipoICMS.ctx":0004
         Left            =   630
         List            =   "RelOpLancTipoICMS.ctx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3336
      End
      Begin VB.Label Label3 
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
         Left            =   210
         TabIndex        =   14
         Top             =   870
         Width           =   360
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   13
         Top             =   420
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
      Left            =   5415
      Picture         =   "RelOpLancTipoICMS.ctx":0008
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   810
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4860
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpLancTipoICMS.ctx":010A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "RelOpLancTipoICMS.ctx":0288
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpLancTipoICMS.ctx":07BA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOpLancTipoICMS.ctx":0944
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpLancTipoICMS.ctx":0A9E
      Left            =   960
      List            =   "RelOpLancTipoICMS.ctx":0AA0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
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
      Left            =   240
      TabIndex        =   11
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOpLancTipoICMSOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const ERRO_TIPOICMS_INICIAL_MAIOR = 0 'Sem parâmetros
'O tipo de ICMS inicial não pode ser maior que o final.

Const ERRO_ALIQUOTA_INICIAL_MAIOR = 0 'Sem parâmetros
'A Alíquota inicial não pode ser maior que a final.

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
            
    'Carrega Tipos de ICMS
    lErro = Carrega_TipoICMS()
    If lErro <> SUCESSO Then gError 75345
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
                    
        Case 75345
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169735)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoICMS() As Long

Dim lErro As Long
Dim objTipoTribICMS As ClassTipoTribICMS
Dim colTiposICMS As New Collection

On Error GoTo Erro_Carrega_TipoICMS

    'Lê Tipos de ICMS
    lErro = CF("TiposTribICMS_Le_Todos",colTiposICMS)
    If lErro <> SUCESSO Then gError 75346

    'Preenche ComboICMSTipo
    For Each objTipoTribICMS In colTiposICMS

        ComboICMSTipoDe.AddItem objTipoTribICMS.iTipo & SEPARADOR & objTipoTribICMS.sDescricao
        ComboICMSTipoDe.ItemData(ComboICMSTipoDe.NewIndex) = objTipoTribICMS.iTipo

        ComboICMSTipoAte.AddItem objTipoTribICMS.iTipo & SEPARADOR & objTipoTribICMS.sDescricao
        ComboICMSTipoAte.ItemData(ComboICMSTipoAte.NewIndex) = objTipoTribICMS.iTipo

    Next

    Carrega_TipoICMS = SUCESSO

    Exit Function

Erro_Carrega_TipoICMS:

    Carrega_TipoICMS = gErr

    Select Case gErr

        Case 75346

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169736)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    'Limpa a tela
    Call Limpar_Tela

    'Carrega Opções de Relatório
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 75347
    
    'pega Tipo de ICMS inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NTIPOINIC", sParam)
    If lErro Then gError 75348
    
    For iIndice = 0 To ComboICMSTipoDe.ListCount - 1
        If ComboICMSTipoDe.ItemData(iIndice) = CInt(sParam) Then
            ComboICMSTipoDe.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'pega Tipo de ICMS inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NTIPOFIM", sParam)
    If lErro Then gError 75349
    
    For iIndice = 0 To ComboICMSTipoAte.ListCount - 1
        If ComboICMSTipoAte.ItemData(iIndice) = CInt(sParam) Then
            ComboICMSTipoAte.ListIndex = iIndice
            Exit For
        End If
    Next
      
    'pega Aliquota Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NALIQINIC", sParam)
    If lErro Then gError 75350
        
    AliquotaDe.Text = Format(sParam, "Standard")
    
    'pega Aliquota final e exibe
    lErro = objRelOpcoes.ObterParametro("NALIQFIM", sParam)
    If lErro Then gError 75351
        
    AliquotaAte.Text = Format(sParam, "Standard")
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 75347 To 75351
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169737)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
        
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 75350
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 75351

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 75350
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 75351
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169738)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    ComboICMSTipoDe.ListIndex = -1
    ComboICMSTipoAte.ListIndex = -1
        
    ComboOpcoes.SetFocus
    
End Sub

Private Function Formata_E_Critica_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
        
    'Tipo de ICMS inicial não pode ser maior que o final
    If ComboICMSTipoDe.ListIndex <> -1 And ComboICMSTipoAte.ListIndex <> -1 Then
         If CInt(ComboICMSTipoDe.ItemData(ComboICMSTipoDe.ListIndex)) > CInt(ComboICMSTipoAte.ItemData(ComboICMSTipoAte.ListIndex)) Then gError 75352
    End If
                   
    'Aliquota inicial não pode ser maior que a final
    If Trim(AliquotaDe.Text) <> "" And Trim(AliquotaAte.Text) <> "" Then
        If CDbl(AliquotaDe.Text) > CDbl(AliquotaAte.Text) Then gError 75353
    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 75352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOICMS_INICIAL_MAIOR", gErr)
            ComboICMSTipoDe.SetFocus
                       
        Case 75353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTA_INICIAL_MAIOR", gErr)
            AliquotaDe.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169739)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_PreencherRelOp

    'Critica as datas
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 75354
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 75355
               
    If ComboICMSTipoDe.ListIndex <> -1 Then
        lErro = objRelOpcoes.IncluirParametro("NTIPOINIC", CStr(ComboICMSTipoDe.ItemData(ComboICMSTipoDe.ListIndex)))
    Else
        lErro = objRelOpcoes.IncluirParametro("NTIPOINIC", "-1")
    End If
    If lErro <> AD_BOOL_TRUE Then gError 75356
    
    If ComboICMSTipoAte.ListIndex <> -1 Then
        lErro = objRelOpcoes.IncluirParametro("NTIPOFIM", CStr(ComboICMSTipoAte.ItemData(ComboICMSTipoAte.ListIndex)))
    Else
        lErro = objRelOpcoes.IncluirParametro("NTIPOFIM", "-1")
    End If
    If lErro <> AD_BOOL_TRUE Then gError 75357
        
    lErro = objRelOpcoes.IncluirParametro("NALIQINIC", CStr(AliquotaDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 75358
    
    lErro = objRelOpcoes.IncluirParametro("NALIQFIM", CStr(AliquotaAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 75359
                
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 75360
            
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 75354 To 75360

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169740)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If ComboICMSTipoDe.ListIndex <> -1 Then sExpressao = "TipoICMS >= " & Forprint_ConvInt(CInt(ComboICMSTipoDe.ItemData(ComboICMSTipoDe.ListIndex)))

   If ComboICMSTipoAte.ListIndex <> -1 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoICMS <= " & Forprint_ConvInt(CInt(ComboICMSTipoAte.ItemData(ComboICMSTipoAte.ListIndex)))

    End If
    
    If Trim(AliquotaDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Aliquota >= " & Forprint_ConvDouble(CDbl(AliquotaDe.Text))

    End If

    If Trim(AliquotaAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Aliquota <= " & Forprint_ConvDouble(CDbl(AliquotaAte.Text))

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169741)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 75361

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 75362

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call Limpar_Tela
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 75361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 75362

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169742)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 75363

    Call gobjRelatorio.Executar_Prossegue2(Me)
        
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 75363
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169743)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 75364

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 75365

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 75366

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 75364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 75365, 75366

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169744)

    End Select

    Exit Sub

End Sub

Private Sub AliquotaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaDe_Validate

    'Se AliquotaDe foi preenchida
    If Len(Trim(AliquotaDe.Text)) > 0 Then
        
        'Critica o Valor
        lErro = Valor_Positivo_Critica(AliquotaDe.Text)
        If lErro <> SUCESSO Then gError 75367
        
    
    End If
    
    Exit Sub

Erro_AliquotaDe_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 75367
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169745)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub AliquotaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaAte_Validate

    'Se AliquotaAte foi preenchida
    If Len(Trim(AliquotaAte.Text)) > 0 Then
        
        'Critica o Valor
        lErro = Valor_Positivo_Critica(AliquotaAte.Text)
        If lErro <> SUCESSO Then gError 75368
        
    
    End If
    
    Exit Sub

Erro_AliquotaAte_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 75368
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169746)
    
    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Lista de Reg. de Entrada/Saída por Tipo ICMS"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpLancTipoICMS"
    
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub






Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label8(Index), Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
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

