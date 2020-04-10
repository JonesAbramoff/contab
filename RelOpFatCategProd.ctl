VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpFatCategProdOcx 
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   ScaleHeight     =   2775
   ScaleWidth      =   7245
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5970
      ScaleHeight     =   495
      ScaleWidth      =   1065
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   1125
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpFatCategProd.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOpFatCategProd.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
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
      Left            =   5970
      Picture         =   "RelOpFatCategProd.ctx":06B0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1020
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Seleção"
      Height          =   2220
      Left            =   210
      TabIndex        =   0
      Top             =   270
      Width           =   5535
      Begin VB.Frame Frame2 
         Caption         =   "Dados Produto"
         Height          =   795
         Left            =   195
         TabIndex        =   12
         Top             =   1245
         Width           =   5160
         Begin VB.ComboBox CategoriaProduto 
            Height          =   315
            Left            =   1740
            TabIndex        =   5
            Top             =   300
            Width           =   2820
         End
         Begin VB.Label Label5 
            Caption         =   "Categoria:"
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
            Height          =   210
            Left            =   750
            TabIndex        =   13
            Top             =   345
            Width           =   930
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Intervalo Período"
         Height          =   900
         Left            =   210
         TabIndex        =   9
         Top             =   270
         Width           =   5160
         Begin MSMask.MaskEdBox DataDe 
            Height          =   300
            Left            =   855
            TabIndex        =   1
            Top             =   375
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissaoDe 
            Height          =   300
            Left            =   2010
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   300
            Left            =   3360
            TabIndex        =   3
            Top             =   360
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissaoAte 
            Height          =   300
            Left            =   4500
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   495
            TabIndex        =   11
            Top             =   420
            Width           =   315
         End
         Begin VB.Label Label4 
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
            Left            =   2895
            TabIndex        =   10
            Top             =   420
            Width           =   360
         End
      End
   End
End
Attribute VB_Name = "RelOpFatCategProdOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Help
Const IDH_RELOP_PEDDATAENTREGA = 0

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 84626 '75496
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, CategoriaProduto, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 84627 '75495

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 84627
        
        Case 84626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168871)

    End Select

    Exit Function

End Function

Function Critica_Datas_RelOpLancData() As Long
'a data inicial não pode ser maior que a data final

Dim lErro As Long

On Error GoTo Erro_Critica_Datas_RelOpLancData
    
    'Critica data se  a Inicial e a Final estiverem Preenchida
    If Len(DataDe.ClipText) <> 0 And Len(DataAte.ClipText) <> 0 Then
    
        'data inicial não pode ser maior que a data final
        If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 84628 '75497
    
    End If
            
    Critica_Datas_RelOpLancData = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpLancData:

    Critica_Datas_RelOpLancData = gErr

    Select Case gErr
    
        Case 84628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168872)

    End Select
    
    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Datas_RelOpLancData()
    If lErro <> SUCESSO Then gError 84629 '75498

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 84630 '75499
    

    lErro = objRelOpcoes.IncluirParametro("DINIC", DataDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 84631 '75500
    
    lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 84632 '75501
    
    lErro = objRelOpcoes.IncluirParametro("TCATEGORIAPRODUTO", CategoriaProduto.Text)
    If lErro <> AD_BOOL_TRUE Then gError 84633 '75501
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr

        Case 84629, 84630, 84631, 84632, 84633

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168873)
            
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arqquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 84634 '75503

    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 84635 '75504

    Call DateParaMasked(DataDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 84636 '75505

    Call DateParaMasked(DataAte, CDate(sParam))

    PreencherParametrosNaTela = SUCESSO

    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr
    
    Select Case gErr

        Case 84634, 84635, 84636
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168874)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If CategoriaProduto.ListIndex = -1 Then gError 84637 '75506

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 84638 '75507

        'retira nome das opções do ComboBox
        CategoriaProduto.RemoveItem CategoriaProduto.ListIndex
    
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 84639 '75508
        
    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case gErr

        Case 84637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)

        Case 84638, 84639

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168875)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

     
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 84640 '75509

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 84640

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168876)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If CategoriaProduto.Text = "" Then gError 84641 '75510

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 84642 '75511

    gobjRelOpcoes.sNome = CategoriaProduto.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 84643 '75512
        
    lErro = RelOpcoes_Testa_Combo(CategoriaProduto, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 84644 '75513
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 84641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            CategoriaProduto.SetFocus

        Case 784642, 84643, 84644

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168877)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 84645 '75514
    
    CategoriaProduto.Text = ""
    CategoriaProduto.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 84645
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168878)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, CategoriaProduto, Me)
    
End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(CategoriaProduto, Cancel)

End Sub

Private Sub DataAte_GotFocus()
    
     Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a Data Final foi digitada
    If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 84646 '90058

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 90058

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168879)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Change()

iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a Data Inicial foi digitada
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 84647 '90059

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 84647

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168880)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load
    
    'Lê as Categorias de Produtos
    lErro = CF("CategoriasProduto_Le_Todas",colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 84648 '33464

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        CategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
    
        Case 84648
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168881)
  
    End Select

    Exit Sub

End Sub



Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 84649

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 84649
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168882)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 84650

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 84650
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168883)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 84651

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 84651
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168884)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 84652

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 84652
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168885)

    End Select

    Exit Sub

End Sub

Public Function Name() As String

    Name = "RelOpFatCategProd"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ALCADA_FAT
    Set Form_Load_Ocx = Me
    Caption = "Tela de Parâmetros para Itens de Categoria"
    Call Form_Load
    
End Function

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

Private Sub DataAte_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub
