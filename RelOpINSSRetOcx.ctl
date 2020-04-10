VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpINSSRetOcx 
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LockControls    =   -1  'True
   ScaleHeight     =   1605
   ScaleWidth      =   6630
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpINSSRetOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpINSSRetOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpINSSRetOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpINSSRetOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
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
      Left            =   4470
      Picture         =   "RelOpINSSRetOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   795
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpINSSRetOcx.ctx":0A96
      Left            =   1005
      List            =   "RelOpINSSRetOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   165
      Width           =   2730
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período"
      Height          =   810
      Left            =   120
      TabIndex        =   9
      Top             =   585
      Width           =   3990
      Begin MSComCtl2.UpDown DeUpDown 
         Height          =   315
         Left            =   1665
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   525
         TabIndex        =   1
         Top             =   330
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown AteUpDown 
         Height          =   315
         Left            =   3570
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   315
         Left            =   2445
         TabIndex        =   2
         Top             =   330
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
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
         Left            =   150
         TabIndex        =   13
         Top             =   390
         Width           =   315
      End
      Begin VB.Label Label2 
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
         Left            =   2010
         TabIndex        =   12
         Top             =   390
         Width           =   360
      End
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
      Left            =   300
      TabIndex        =   14
      Top             =   210
      Width           =   615
   End
End
Attribute VB_Name = "RelOpINSSRetOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 49792

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 49792

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169406)

    End Select

    Exit Sub

End Sub

Function Define_Padrao() As Long

Dim lErro As Long
Dim dtDataAux As Date

On Error GoTo Erro_Define_Padrao

    dtDataAux = gdtDataAtual - 2

    'se DataAtual diminuida de dois for sabado ou domingo entao diminui ate ser sexta-feira
    If Weekday(dtDataAux) = vbSunday Then

        dtDataAux = dtDataAux - 2

    Else

        If Weekday(dtDataAux) = vbSaturday Then

            dtDataAux = dtDataAux - 1

        End If

    End If

    'preenche a tela
    DataInicial.PromptInclude = False
    DataInicial.Text = Format(dtDataAux, "dd/mm/yy")
    DataInicial.PromptInclude = True

    'preenche a tela
    DataFinal.PromptInclude = False
    DataFinal.Text = Format(dtDataAux, "dd/mm/yy")
    DataFinal.PromptInclude = True

    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169407)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 49793
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 49791
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 49791
        
        Case 49793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169408)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 49794
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 49795
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 49794, 49795
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169409)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 49797

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 49798

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 49799
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 49800
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 49797
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 49798, 49799, 49800
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169410)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 49802

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 49803

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 49802
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 49803

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169411)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 49806

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 49806

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169412)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp
    
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then Error 49807

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 49808
        
    'se a data não foi preenchida passa a data nula
    If Trim(DataInicial.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
        If lErro <> AD_BOOL_TRUE Then Error 49809
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then Error 49810
    End If
    
    'se a data não foi preenchida passa a data nula
    If Trim(DataFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
        If lErro <> AD_BOOL_TRUE Then Error 49811
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then Error 49812
    End If
    
    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 49813

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 49807, 49808, 49809, 49810, 49811, 49812, 49813
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169413)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros() As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o Tipocliente e Cobrador

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'data da Baixa inicial não pode ser maior que a Baixa final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 49814
    
    End If
            
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
        
        Case 49814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
              
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169414)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
        
    If Trim(DataInicial.ClipText) <> "" Then sExpressao = "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169415)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoCliente As String
Dim sCobrador As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 49815
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 49816
    
    Call DateParaMasked(DataInicial, CDate(sParam))
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 49817
        
    Call DateParaMasked(DataFinal, CDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 49815, 49816, 49817
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169416)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then
        
        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then Error 49820

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 49820

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169417)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then Error 49821

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 49821

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169418)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
 
 End Sub
    
Private Sub AteUpDown_DownClick()

Dim lErro As Long

On Error GoTo Erro_AteUpDown_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 49822

    Exit Sub

Erro_AteUpDown_DownClick:

    Select Case Err

        Case 49822
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169419)

    End Select

    Exit Sub

End Sub

Private Sub AteUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_AteUpDown_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 49823

    Exit Sub

Erro_AteUpDown_UpClick:

    Select Case Err

        Case 49823
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169420)

    End Select

    Exit Sub
    
End Sub

Private Sub DeUpDown_DownClick()

Dim lErro As Long

On Error GoTo Erro_DeUpDown_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 49824

    Exit Sub

Erro_DeUpDown_DownClick:

    Select Case Err

        Case 49824
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169421)

    End Select

    Exit Sub

End Sub

Private Sub DeUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_DeUpDown_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 49825

    Exit Sub

Erro_DeUpDown_UpClick:

    Select Case Err

        Case 49825
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169422)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_INSSRET
    Set Form_Load_Ocx = Me
    Caption = "Relatório Para Recolhimento de INSS Retido"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpINSSRet"
    
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


Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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

