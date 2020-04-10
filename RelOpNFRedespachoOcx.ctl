VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpNFRedespachoOcx 
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   KeyPreview      =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   8025
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5745
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpNFRedespachoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpNFRedespachoOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpNFRedespachoOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpNFRedespachoOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpNFRedespachoOcx.ctx":0994
      Left            =   945
      List            =   "RelOpNFRedespachoOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2670
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
      Left            =   6000
      Picture         =   "RelOpNFRedespachoOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame FrameNF 
      Caption         =   "Nota Fiscal"
      Height          =   1965
      Left            =   135
      TabIndex        =   8
      Top             =   750
      Width           =   5505
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   690
         TabIndex        =   1
         Top             =   300
         Width           =   765
      End
      Begin VB.Frame Frame1 
         Caption         =   "Número"
         Height          =   1065
         Left            =   135
         TabIndex        =   11
         Top             =   765
         Width           =   5205
         Begin MSMask.MaskEdBox NFiscalInicial 
            Height          =   300
            Left            =   525
            TabIndex        =   12
            Top             =   465
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NFiscalFinal 
            Height          =   300
            Left            =   2265
            TabIndex        =   13
            Top             =   465
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label Label14 
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
            Left            =   120
            TabIndex        =   15
            Top             =   540
            Width           =   315
         End
         Begin VB.Label Label6 
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
            Left            =   1770
            TabIndex        =   14
            Top             =   540
            Width           =   360
         End
      End
      Begin VB.Label LabelSerie 
         AutoSize        =   -1  'True
         Caption         =   "Série:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   360
         Width           =   510
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
      Left            =   255
      TabIndex        =   10
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpNFRedespachoOcx"
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

Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoSerie = New AdmEvento

    'Carrega a combo série
    lErro = Carrega_Serie()
    If lErro Then gError 37634
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 37634

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 37635
      
    'pega Nota Fiscal inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NNFISCALINIC", sParam)
    If lErro Then gError 37640
    
    NFiscalInicial.Text = sParam

    'pega Nota Fiscal final e exibe
    lErro = objRelOpcoes.ObterParametro("NNFISCALFIM", sParam)
    If lErro Then gError 37641
    
    NFiscalFinal.Text = sParam
     
    'pega série e exibe
    lErro = objRelOpcoes.ObterParametro("TSERIE", sParam)
    If lErro <> SUCESSO Then gError 37644

    Serie.Text = sParam
       
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 37635, 37640, 37641, 37644
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 37629

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 37629
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoSerie = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros() As Long
'Formata os produtos retornando em sProd_I e sProd_F

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'Verifica se o numero da Nota Fiscal inicial é maior que o da final
    If Len(Trim(NFiscalInicial.ClipText)) > 0 And Len(Trim(NFiscalFinal.ClipText)) > 0 Then
    
        If CLng(NFiscalInicial.Text) > CLng(NFiscalFinal.Text) Then gError 37649
    
    End If
               
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 37649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
            NFiscalInicial.SetFocus
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 47157
    
    Serie.Text = ""
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 47157
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 37652
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 37653
            
    lErro = objRelOpcoes.IncluirParametro("NNFISCALINIC", NFiscalInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37656

    lErro = objRelOpcoes.IncluirParametro("NNFISCALFIM", NFiscalFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37657
       
    lErro = objRelOpcoes.IncluirParametro("TSERIE", Serie.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37660
       
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 37661
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 37652 To 37661
        
        Case 71363 To 71366

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 37662

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 37663

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 47158
    
        ComboOpcoes.Text = ""
        Serie.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 37662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 37663, 47158

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 37685

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 37685

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 37664

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 37665

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 37666

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 47159
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 37664
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 37665, 37666, 47159

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If Serie.Text <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Serie = " & Forprint_ConvTexto(Serie.Text)
    
    End If
    
    If NFiscalInicial.Text <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NotaFiscal >= " & Forprint_ConvLong(NFiscalInicial.Text)
    
    End If

   If NFiscalFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NotaFiscal <= " & Forprint_ConvLong(NFiscalFinal.Text)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub LabelSerie_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieListaModal
    Call Chama_Tela("SerieListaModal", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série na Tela
    Serie.Text = objSerie.sSerie
    
    Call Serie_Validate(bSGECancelDummy)

    Exit Sub

End Sub

Private Sub NFiscalInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalInicial_Validate
            
    lErro = Critica_Numero(NFiscalInicial.Text)
    If lErro <> SUCESSO Then gError 37673
              
    Exit Sub

Erro_NFiscalInicial_Validate:

    Cancel = True

    Select Case gErr
    
        Case 37673
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Sub

End Sub

Private Sub NFiscalFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalFinal_Validate
     
    lErro = Critica_Numero(NFiscalFinal.Text)
    If lErro <> SUCESSO Then gError 37674
        
    Exit Sub

Erro_NFiscalFinal_Validate:

    Cancel = True


    Select Case gErr
    
        Case 37674
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Sub

End Sub

Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero
         
    If Len(Trim(sNumero)) > 0 Then
        
        lErro = Long_Critica(sNumero)
        If lErro <> SUCESSO Then gError 37675
 
        If CLng(sNumero) < 0 Then gError 37676
        
    End If
 
    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = gErr

    Select Case gErr
                  
        Case 37675
            
        Case 37676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 37677
    
    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next
    
    Carrega_Serie = SUCESSO
    
    Exit Function
    
Erro_Carrega_Serie:

    Carrega_Serie = gErr
    
    Select Case gErr
    
        Case 37677
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function

End Function

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Serie_Validate

    'Verifica se a Serie foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
        
    'Verifica se é uma Serie selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 37678
    
    If lErro = 12253 Then gError 37679
    
    Exit Sub
    
Erro_Serie_Validate:

    Cancel = True


    Select Case gErr
    
        Case 37678
       
        Case 37679
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, Serie.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Notas Fiscais de Redespacho"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpNFRedespacho"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Serie Then
            Call LabelSerie_Click
        End If
    
    End If

End Sub
Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub LabelSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerie, Source, X, Y)
End Sub

Private Sub LabelSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerie, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub



