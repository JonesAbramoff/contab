VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpRotasOcx 
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   LockControls    =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   6270
   Begin VB.ComboBox Chave 
      Height          =   315
      Left            =   825
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   2730
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rotas"
      Height          =   795
      Left            =   120
      TabIndex        =   11
      Top             =   1185
      Width           =   5955
      Begin MSMask.MaskEdBox RotaDe 
         Height          =   315
         Left            =   705
         TabIndex        =   2
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RotaAte 
         Height          =   315
         Left            =   3840
         TabIndex        =   3
         Top             =   330
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelRotaAte 
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
         Height          =   255
         Left            =   3405
         TabIndex        =   13
         Top             =   375
         Width           =   435
      End
      Begin VB.Label LabelRotaDe 
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
         Height          =   255
         Left            =   315
         TabIndex        =   12
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRotasOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRotasOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRotasOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRotasOcx.ctx":0816
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
      Left            =   2175
      Picture         =   "RelOpRotasOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2100
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRotasOcx.ctx":0A96
      Left            =   825
      List            =   "RelOpRotasOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2730
   End
   Begin VB.Label LabelChave 
      Alignment       =   1  'Right Justify
      Caption         =   "Chave:"
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
      Height          =   315
      Left            =   -30
      TabIndex        =   14
      Top             =   825
      Width           =   825
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
      Left            =   120
      TabIndex        =   9
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRotasOcx"
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

    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_CHAVE_ROTA, Chave, True, True, True)
    If lErro <> SUCESSO Then gError 205380
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 205380
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205381)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 205382
   
    lErro = objRelOpcoes.ObterParametro("TROTADE", sParam)
    If lErro <> SUCESSO Then gError 205383

    RotaDe.Text = sParam
    
    'pega Região de Venda Final e exibe
    lErro = objRelOpcoes.ObterParametro("TROTAATE", sParam)
    If lErro <> SUCESSO Then gError 205384

    RotaAte.Text = sParam
    
    'pega vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NDIADASEMANA", sParam)
    If lErro <> SUCESSO Then gError 205385
    
    Call Combo_Seleciona_ItemData(Chave, StrParaLong(sParam))
          
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 205382 To 205385

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205386)

    End Select

    Exit Function

End Function

Private Function PreencheComboOpcoes(sCodRel As String) As Long

Dim colRelParametros As New Collection
Dim lErro As Long
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_PreencheComboOpcoes

    'le os nomes das opcoes do relatório existentes no BD
    lErro = CF("RelOpcoes_Le_Todos", sCodRel, colRelParametros)
    If lErro <> SUCESSO Then gError 205387

    'preenche o ComboBox com os nomes das opções do relatório
    For Each objRelOpcoes In colRelParametros
        ComboOpcoes.AddItem objRelOpcoes.sNome
    Next

    PreencheComboOpcoes = SUCESSO

    Exit Function

Erro_PreencheComboOpcoes:

    PreencheComboOpcoes = gErr

    Select Case gErr

        Case 205387

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205388)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long
Dim iOpcao As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 205389
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = PreencheComboOpcoes(gobjRelatorio.sCodRel)
    If lErro <> SUCESSO Then gError 205390

    'verifica se o nome da opção passada está no ComboBox
    For iOpcao = 0 To ComboOpcoes.ListCount - 1

        If ComboOpcoes.List(iOpcao) = gobjRelOpcoes.sNome Then

            ComboOpcoes.Text = ComboOpcoes.List(iOpcao)

            lErro = PreencherParametrosNaTela(gobjRelOpcoes)
            If lErro <> SUCESSO Then gError 205391

            Exit For

        End If

    Next
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 205389
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 205390, 205391
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205392)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
  
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 205393
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
       
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 205393
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205394)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

Dim lErro As Long

On Error GoTo Erro_ComboOpcoes_Click
    
    If ComboOpcoes.ListIndex = -1 Then Exit Sub

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Le", gobjRelOpcoes)
    If (lErro <> SUCESSO) Then gError 205395

    lErro = PreencherParametrosNaTela(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 205396

    Exit Sub

Erro_ComboOpcoes_Click:

    Select Case gErr

        Case 205395, 205396

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205397)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lChave As Long

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros(lChave)
    If lErro <> SUCESSO Then gError 205398
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 205399
         
    lErro = objRelOpcoes.IncluirParametro("TROTADE", RotaDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 205400

    lErro = objRelOpcoes.IncluirParametro("TROTAATE", RotaAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 205401
    
    lErro = objRelOpcoes.IncluirParametro("NDIADASEMANA", CStr(lChave))
    If lErro <> AD_BOOL_TRUE Then gError 205402

    lErro = objRelOpcoes.IncluirParametro("TDIADASEMANADESC", Chave.Text)
    If lErro <> AD_BOOL_TRUE Then gError 205403
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, RotaDe.Text, RotaAte.Text, lChave)
    If lErro <> SUCESSO Then gError 205404
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 205398 To 205404

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205405)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 205406

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELVENDREG")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 205407

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click
    
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 205406
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 205407

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205408)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 205409

    Me.Enabled = False
    Call gobjRelatorio.Executar_Prossegue

    Unload Me

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 205409

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205410)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 205411

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 205412

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 205413

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 205414
    
    Call BotaoLimpar_Click
    
    ComboOpcoes.Text = ""

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 205411
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 205412 To 205414

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205415)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sRotaDe As String, sRotaAte As String, lChave As Long) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
    
   If lChave <> 0 Then sExpressao = "DiaDaSemana >= " & Forprint_ConvLong(lChave)
     
    If Trim(sRotaDe) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Rota >= " & Forprint_ConvTexto(sRotaDe)
    
    End If
    
    If Trim(sRotaAte) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Rota <= " & Forprint_ConvTexto(sRotaAte)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205416)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(lChave As Long) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'Se RegiãoInicial e RegiãoFinal estão preenchidos
    If Len(Trim(RotaDe.Text)) > 0 And Len(Trim(RotaAte.Text)) > 0 Then
    
        'Se Região inicial for maior que Região final, erro
        If RotaDe.Text > RotaAte.Text Then gError 205417
        
    End If
    
    If Chave.ListIndex = -1 Then
        lChave = 0
    Else
        lChave = Chave.ItemData(Chave.ListIndex)
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 205417
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTA_INICIAL_MAIOR", gErr)
            RotaDe.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205418)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Rotas de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRotas"
    
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

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelRotaDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRotaDe, Source, X, Y)
End Sub

Private Sub LabelRotaDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRotaDe, Button, Shift, X, Y)
End Sub

Private Sub LabelRotaAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRotaAte, Source, X, Y)
End Sub

Private Sub LabelRotaAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRotaAte, Button, Shift, X, Y)
End Sub
