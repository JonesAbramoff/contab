VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl CustoDirFabrCalculaOcx 
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ScaleHeight     =   5325
   ScaleWidth      =   6570
   Begin VB.Frame Frame3 
      Caption         =   "Venda Mensal baseada em"
      Height          =   810
      Left            =   120
      TabIndex        =   26
      Top             =   1470
      Width           =   6315
      Begin VB.ComboBox MesFim 
         Height          =   315
         ItemData        =   "CustoDirFabrCalcula.ctx":0000
         Left            =   4845
         List            =   "CustoDirFabrCalcula.ctx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   300
         Width           =   1335
      End
      Begin VB.ComboBox MesIni 
         Height          =   315
         ItemData        =   "CustoDirFabrCalcula.ctx":0094
         Left            =   3240
         List            =   "CustoDirFabrCalcula.ctx":00BF
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   300
         Width           =   1350
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   285
         Left            =   1020
         TabIndex        =   27
         Top             =   330
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4635
         TabIndex        =   30
         Top             =   345
         Width           =   195
      End
      Begin VB.Label Label3 
         Caption         =   "meses de:"
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
         Left            =   2340
         TabIndex        =   29
         Top             =   360
         Width           =   930
      End
      Begin VB.Label LabelCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Previsão:"
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
         Height          =   285
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   345
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture 
      Height          =   555
      Left            =   4755
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   120
      Width           =   1665
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "CustoDirFabrCalcula.ctx":0128
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "CustoDirFabrCalcula.ctx":065A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   120
         Picture         =   "CustoDirFabrCalcula.ctx":07D8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Executa a rotina"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valores Mensais"
      Height          =   2025
      Left            =   120
      TabIndex        =   14
      Top             =   2310
      Width           =   6315
      Begin MSMask.MaskEdBox ValorFator 
         Height          =   300
         Index           =   1
         Left            =   1995
         TabIndex        =   2
         Top             =   285
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorFator 
         Height          =   300
         Index           =   2
         Left            =   1995
         TabIndex        =   3
         Top             =   720
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorFator 
         Height          =   300
         Index           =   3
         Left            =   2010
         TabIndex        =   4
         Top             =   1125
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorFator 
         Height          =   300
         Index           =   4
         Left            =   2010
         TabIndex        =   5
         Top             =   1515
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorTotal 
         Height          =   300
         Left            =   4620
         TabIndex        =   6
         Top             =   1545
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorFator 
         Height          =   300
         Index           =   5
         Left            =   4605
         TabIndex        =   33
         Top             =   360
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorFator 
         Height          =   300
         Index           =   6
         Left            =   4620
         TabIndex        =   35
         Top             =   765
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fator 6:"
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
         Index           =   6
         Left            =   3885
         TabIndex        =   36
         Top             =   825
         Width           =   675
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fator 5:"
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
         Index           =   5
         Left            =   3870
         TabIndex        =   34
         Top             =   420
         Width           =   675
      End
      Begin VB.Label LabelOutros 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4620
         TabIndex        =   25
         Top             =   1155
         Width           =   1290
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mão de Obra Direta:"
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
         Index           =   1
         Left            =   195
         TabIndex        =   15
         Top             =   345
         Width           =   1740
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Gás/BPF:"
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
         Index           =   2
         Left            =   1095
         TabIndex        =   16
         Top             =   765
         Width           =   840
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Energia:"
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
         Index           =   3
         Left            =   1215
         TabIndex        =   17
         Top             =   1185
         Width           =   720
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Água:"
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
         Index           =   4
         Left            =   1440
         TabIndex        =   18
         Top             =   1575
         Width           =   510
      End
      Begin VB.Label LabelTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   4050
         TabIndex        =   20
         Top             =   1575
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Outros:"
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
         Left            =   3900
         TabIndex        =   19
         Top             =   1200
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Calcular apenas o produto abaixo"
      Enabled         =   0   'False
      Height          =   795
      Left            =   120
      TabIndex        =   21
      Top             =   4395
      Width           =   6315
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   885
         TabIndex        =   7
         Top             =   330
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescProd 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2235
         TabIndex        =   23
         Top             =   330
         Width           =   3810
      End
      Begin VB.Label LabelProduto 
         Caption         =   "Produto:"
         Enabled         =   0   'False
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame FrameIdentificacao 
      Caption         =   "Identificação"
      Height          =   735
      Left            =   98
      TabIndex        =   0
      Top             =   705
      Width           =   6315
      Begin MSMask.MaskEdBox Ano 
         Height          =   285
         Left            =   585
         TabIndex        =   1
         Top             =   315
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   4
         Format          =   "0"
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataExecucao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5040
         TabIndex        =   13
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Executada em:"
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
         Left            =   3645
         TabIndex        =   12
         Top             =   315
         Width           =   1350
      End
      Begin VB.Label LabelAno 
         Caption         =   "Ano:"
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
         Height          =   285
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   315
         Width           =   465
      End
   End
End
Attribute VB_Name = "CustoDirFabrCalculaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAnoAlterado As Integer

Private WithEvents objEventoAno As AdmEvento
Attribute objEventoAno.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoPrevVenda As AdmEvento
Attribute objEventoPrevVenda.VB_VarHelpID = -1

Event Unload()

'Property Variables:
Dim m_Caption As String

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID = IDH_MOVIMENTOS_ESTOQUE_MOVIMENTO
    Set Form_Load_Ocx = Me
    Caption = "Custo Direto de Fabricação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CustoDirFabrCalculaOcx"

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

Private Sub Ano_GotFocus()
    Call MaskEdBox_TrataGotFocus(Ano, iAnoAlterado)
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Se o código foi preenchido
    If Len(Trim(Codigo.Text)) > 0 Then
    
        'Verifica se existe uma Previsão de Vendas cadastrada com o código passado
        lErro = CF("PrevVendaMensal_Le_Codigo", Codigo.Text, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 90203 Then gError 117248
        
        'Se não encontro PrevVenda, erro
        If lErro = 90203 Then gError 117249
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 117248
        
        Case 117249
            Call Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_CADASTRADA", gErr, Codigo.Text)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158567)
    
    End Select
    
    Exit Sub
    
End Sub

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

Private Sub LabelAno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAno, Source, X, Y)
End Sub

Private Sub LabelAno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAno, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub LabelTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotal, Button, Shift, X, Y)
End Sub

Private Sub LabelTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotal, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto, Source, X, Y)
End Sub

Public Sub Form_Load()

Dim lErro As Long, objCamposGenericos As New ClassCamposGenericos

On Error GoTo Erro_Form_Load

    'Seta os ObjEventos
    Set objEventoProduto = New AdmEvento
    Set objEventoPrevVenda = New AdmEvento
    Set objEventoAno = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 117246
    
    objCamposGenericos.lCodigo = CAMPOSGENERICOS_KIT_FATOR
    lErro = CF("CamposGenericosValores_Le_CodCampo", objCamposGenericos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    LabelFator(1) = objCamposGenericos.colCamposGenericosValores.Item(1).sValor & ":"
    LabelFator(2) = objCamposGenericos.colCamposGenericosValores.Item(2).sValor & ":"
    LabelFator(3) = objCamposGenericos.colCamposGenericosValores.Item(3).sValor & ":"
    LabelFator(4) = objCamposGenericos.colCamposGenericosValores.Item(4).sValor & ":"
    LabelFator(5) = objCamposGenericos.colCamposGenericosValores.Item(5).sValor & ":"
    LabelFator(6) = objCamposGenericos.colCamposGenericosValores.Item(6).sValor & ":"
    
    MesIni.ListIndex = 0
    MesFim.ListIndex = 11
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 117246, ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158568)

    End Select

    Exit Sub
   
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoAno = Nothing
    Set objEventoProduto = Nothing
    Set objEventoPrevVenda = Nothing
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Ano_Change()
    
    iAnoAlterado = REGISTRO_ALTERADO

End Sub

Private Function Limpa_Tela_CustoDirFabrCalcula() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CustoDirFabrCalcula
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'Limpa os labeis
    LabelDataExecucao.Caption = ""
    DescProd.Caption = ""
    LabelOutros.Caption = ""
    
    'Limpa os campos da tela
    Call Limpa_Tela(Me)
    
    'Desabilita frame de produto
    Frame2.Enabled = False
    LabelProduto.Enabled = False
 
    MesIni.ListIndex = 0
    MesFim.ListIndex = 11
    
    'Zera alterações
    iAnoAlterado = 0
    
    Exit Function
    
Erro_Limpa_Tela_CustoDirFabrCalcula:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158569)

    End Select

    Exit Function
    
End Function

Private Sub Ano_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Ano_Validate
    
    'se o ano não foi alterado => sai da função
    If iAnoAlterado = 0 Then Exit Sub
        
    'Chama a Função traz_custo_tela
    lErro = Traz_CustoDirFab_Tela()
    If lErro <> SUCESSO Then gError 117241
    
    Exit Sub

Erro_Ano_Validate:

    Cancel = True

    Select Case gErr
    
        Case 117241
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158570)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
'Está sendo declarado internamente pois não é utilizado em mais nenhum lugar
Dim iAlterado As Integer

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se algum campo foi alterado
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 117216

    'Limpa a tela CustoDirFabrCalcula
    Call Limpa_Tela_CustoDirFabrCalcula
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 117216

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158571)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub objEventoAno_evSelecao(obj1 As Object)

Dim objCustoDirFabr As ClassCustoDirFabr

On Error GoTo Erro_objEventoAno_evSelecao
    
    Set objCustoDirFabr = obj1

    'Preenche campo CustoDirFabr
    Ano.Text = CStr(objCustoDirFabr.iAno)
    
    Ano_Validate (bSGECancelDummy)
        
    Me.Show

    Exit Sub

Erro_objEventoAno_evSelecao:
   
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158572)

    End Select

    Exit Sub

End Sub

Private Sub LabelAno_Click()

Dim lErro As Long
Dim objCustoDirFabr As New ClassCustoDirFabr
Dim colSelecao As New Collection

On Error GoTo Erro_LabelAno_Click

    'Se o Ano está preenchido...
    If (Len(Trim(Ano.ClipText)) > 0) Then
        
        'Formata o ANO para o BD
        lErro = Long_Critica(Ano.Text)
        If lErro <> SUCESSO Then gError 117217
        
        'Guarda o Ano já criticado em objCustoDirFabr
        objCustoDirFabr.iAno = Ano.ClipText
        
    End If
    
    'chama a tela de browser
    Call Chama_Tela("AnoCustoDirFabrLista", colSelecao, objCustoDirFabr, objEventoAno)
    
    Exit Sub
    
Erro_LabelAno_Click:

    Select Case gErr
    
        Case 117217
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTODIRFABRCALCULA_ANO_INVALIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158573)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Validate
        
    'se o codigo estiver vazio  => sai da rotina
    If Len(Trim(Produto.ClipText)) = 0 Then Exit Sub

    lErro = CF("Produto_Critica_Filial", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 117223
            
    If lErro = 51381 Then gError 117224
    
    DescProd.Caption = objProduto.sDescricao

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 117223

        Case 117224
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158574)

    End Select

    Exit Sub

End Sub

Private Sub ValorFator_Validate(Index As Integer, Cancel As Boolean)
 
Dim lErro As Long

On Error GoTo Erro_ValorFator_Validate
 
    'Verifica se o valor foi digitado
    If Len(Trim(ValorFator(Index).ClipText)) = 0 Then Exit Sub
    
    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorFator(Index).Text)
    If lErro <> SUCESSO Then gError 117225
    
    lErro = Calcula_Valor_Outros
    If lErro <> SUCESSO Then gError 117245
    
    Exit Sub

Erro_ValorFator_Validate:
  
    Cancel = True
    
    Select Case gErr
        
        Case 117225, 117245
        'Tratado nas Rotinas Chamadas
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158575)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long, sNomeArqParam As String
Dim objCustoDirFab As New ClassCustoDirFabr

On Error GoTo Erro_BotaoGerar_Click
    
    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Chama o Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objCustoDirFab)
    If lErro <> SUCESSO Then gError 106630
        
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 106631
        
    lErro = CF("Rotina_Rateio_CustoDirFabr", sNomeArqParam, objCustoDirFab)
    If lErro <> SUCESSO Then gError 117230

    Unload Me
    
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
        
    Exit Sub

Erro_BotaoGerar_Click:
      
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
                                            
        Case 117230, 106630
        
        Case 106631
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158576)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros()
    Trata_Parametros = SUCESSO
End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor, lErro As Long
Dim objCustoDirFab As New ClassCustoDirFabr

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CustoDirFabr"

    lErro = Move_Tela_Memoria(objCustoDirFab)
    If lErro <> SUCESSO Then gError 117232
    
    'Preenche a coleção colCampoValor, com nome do campo,
    colCampoValor.Add "Ano", objCustoDirFab.iAno, 0, "Ano"
    colCampoValor.Add "Codigo", objCustoDirFab.sCodigoPrevVenda, STRING_PREVVENDA_CODIGO, "CodigoPrevVenda"
    colCampoValor.Add "LabelDataExecucao", objCustoDirFab.dtData, 0, "Data"
    colCampoValor.Add "ValorTotal", objCustoDirFab.dCustoTotal, 0, "CustoTotal"
    colCampoValor.Add "ValorFator(1)", objCustoDirFab.dCustoFator1, 0, "CustoFator1"
    colCampoValor.Add "ValorFator(2)", objCustoDirFab.dCustoFator2, 0, "CustoFator2"
    colCampoValor.Add "ValorFator(3)", objCustoDirFab.dCustoFator3, 0, "CustoFator3"
    colCampoValor.Add "ValorFator(4)", objCustoDirFab.dCustoFator4, 0, "CustoFator4"
    colCampoValor.Add "ValorFator(5)", objCustoDirFab.dCustoFator5, 0, "CustoFator5"
    colCampoValor.Add "ValorFator(6)", objCustoDirFab.dCustoFator6, 0, "CustoFator6"
        
    'Filtro para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 117232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158577)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCustoDirFab As New ClassCustoDirFabr

On Error GoTo Erro_Tela_Preenche

    Ano.Text = colCampoValor.Item("Ano").vValor
    
    'Traz os dados para tela
    lErro = Traz_CustoDirFab_Tela()
    If lErro <> SUCESSO Then gError 117233
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 117233

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158578)

    End Select

    Exit Sub

End Sub

Private Function Traz_CustoDirFab_Tela() As Long

Dim lErro As Long
Dim objCustoDirFabr As New ClassCustoDirFabr

On Error GoTo Erro_Traz_CustoDirFab_Tela
    
    Produto.PromptInclude = False
    Produto.Text = ""
    Produto.PromptInclude = True
    
    DescProd.Caption = ""
        
    Frame2.Enabled = False
    LabelProduto.Enabled = False
 
    'trazer dados de custodirfabr para a tela, se houver.
    'Verifica se o Ano está preeenchido
    If Len(Trim(Ano.ClipText)) > 0 Then
        
        'Se o ano tiver menos de 4 dígitos => erro
        If Len(Trim((Ano.Text))) < 4 Then gError 117213
        
        objCustoDirFabr.iAno = Ano.Text
        objCustoDirFabr.iFilialEmpresa = giFilialEmpresa
    
        lErro = CF("CustoDirFabr_Le", objCustoDirFabr)
        If lErro <> SUCESSO And lErro <> 117237 Then gError 117214
        
        If lErro = SUCESSO Then
        
            'Traz todos os campos para Tela
            Ano.Text = objCustoDirFabr.iAno
        
            'Preenche o código
            Codigo.PromptInclude = False
            Codigo.Text = objCustoDirFabr.sCodigoPrevVenda
            Codigo.PromptInclude = True
        
            'Preenche labelDataExecução
            If (objCustoDirFabr.dtData <> 0) Then
                LabelDataExecucao.Caption = objCustoDirFabr.dtData
            Else
                LabelDataExecucao.Caption = ""
            End If
        
            'Preenche o campo Valor Total
            ValorTotal.Text = Format(objCustoDirFabr.dCustoTotal, "standard")
    
            'Preenche os Valores Fatrores
            ValorFator(1).Text = Format(objCustoDirFabr.dCustoFator1, "standard")
            ValorFator(2).Text = Format(objCustoDirFabr.dCustoFator2, "standard")
            ValorFator(3).Text = Format(objCustoDirFabr.dCustoFator3, "standard")
            ValorFator(4).Text = Format(objCustoDirFabr.dCustoFator4, "standard")
            ValorFator(5).Text = Format(objCustoDirFabr.dCustoFator5, "standard")
            ValorFator(6).Text = Format(objCustoDirFabr.dCustoFator6, "standard")
            
            lErro = Calcula_Valor_Outros
            If lErro <> SUCESSO Then gError 117247
            
            Frame2.Enabled = True
            LabelProduto.Enabled = True
        
            Call Combo_Seleciona_ItemData(MesIni, IIf(objCustoDirFabr.iMesIni = 0, 1, objCustoDirFabr.iMesIni))
            Call Combo_Seleciona_ItemData(MesFim, IIf(objCustoDirFabr.iMesFim = 0, 12, objCustoDirFabr.iMesFim))
        
        End If
        
    End If
    
    Traz_CustoDirFab_Tela = SUCESSO
    
    Exit Function

Erro_Traz_CustoDirFab_Tela:

    Traz_CustoDirFab_Tela = gErr
    
    Select Case gErr
    
        Case 117213
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_ANO_INVALIDO", gErr)
        
        Case 117214, 117247
            'Erro tratado na Rotina chamada
            
        Case 117215
            'Call Rotina_Erro(vbOKOnly, "ERRO_CUSTODIRFABRCALCULA_SEM_DADOS", gErr, Ano.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158579)
    
    End Select
    
    Exit Function
    
End Function

Public Function Calcula_Valor_Outros() As Long

Dim lErro As Long
Dim iIndex As Integer
Dim dValorOutros As Double

On Error GoTo Erro_Calcula_Valor_Outros

    'Atribui o valor Total à label outros
    dValorOutros = StrParaDbl(ValorTotal.Text)
    
    'Recalcular outros: total - Valorfator(s)
    For iIndex = 1 To 6
         dValorOutros = dValorOutros - StrParaDbl(ValorFator(iIndex))
    Next
           
    'Se o valor outros for negativo => erro
    If dValorOutros < 0 Then
        LabelOutros.Caption = ""
    Else
        'Exibe o valor outros na tela
        LabelOutros.Caption = Format(dValorOutros, "#,##0.00")
    End If
    
    Calcula_Valor_Outros = SUCESSO
    
    Exit Function

Erro_Calcula_Valor_Outros:

    Calcula_Valor_Outros = gErr
    
    Select Case gErr
            
        Case 117243
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_OUTROS_NEGATIVO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158580)
    
    End Select
    
    Exit Function
    
End Function

Private Function Move_Tela_Memoria(objCustoDirFab As ClassCustoDirFabr) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se os campos Obrigatórios estão preenchidos
    If Len(Trim(Ano.ClipText)) = 0 Then gError 117226
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 117227
    If StrParaDbl(ValorTotal.Text) = 0 Then gError 117250
    
    'Preenche os campos do Objeto
    objCustoDirFab.iFilialEmpresa = giFilialEmpresa
    objCustoDirFab.iAno = StrParaInt(Ano.ClipText)
    objCustoDirFab.sCodigoPrevVenda = Codigo.Text
    objCustoDirFab.dtData = gdtDataHoje
    objCustoDirFab.dCustoTotal = StrParaDbl(ValorTotal.Text)
    objCustoDirFab.dCustoFator1 = StrParaDbl(ValorFator(1).ClipText)
    objCustoDirFab.dCustoFator2 = StrParaDbl(ValorFator(2).ClipText)
    objCustoDirFab.dCustoFator3 = StrParaDbl(ValorFator(3).ClipText)
    objCustoDirFab.dCustoFator4 = StrParaDbl(ValorFator(4).ClipText)
    objCustoDirFab.dCustoFator5 = StrParaDbl(ValorFator(5).ClipText)
    objCustoDirFab.dCustoFator6 = StrParaDbl(ValorFator(6).ClipText)
       
    If objCustoDirFab.dCustoTotal - (objCustoDirFab.dCustoFator1 + objCustoDirFab.dCustoFator2 + objCustoDirFab.dCustoFator3 + objCustoDirFab.dCustoFator4 + objCustoDirFab.dCustoFator5 + objCustoDirFab.dCustoFator6) < DELTA_VALORMONETARIO Then gError 106629
    
    'Formata o produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 117231
    
    objCustoDirFab.sProduto = sProdutoFormatado
    
    objCustoDirFab.iMesIni = MesIni.ItemData(MesIni.ListIndex)
    objCustoDirFab.iMesFim = MesFim.ItemData(MesFim.ListIndex)
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
        
        Case 117226
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
        
        Case 117227
            Call Rotina_Erro(vbOKOnly, "ERRO_COD_PREVISAO_NAO_PREENCHIDO", gErr)
                        
        Case 106629
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TOTAL_MENOR_PARCIAIS", gErr)
        
        Case 117250
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TOTAL_NAO_INFORMADO", gErr)
                        
        Case 117231
            'Tratado na Rotina Chamada
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158581)
            
    End Select

    Exit Function

End Function

Private Sub LabelCodigo_Click()

Dim objPrevVendaMensal As New ClassPrevVendaMensal
Dim colSelecao As Collection

On Error GoTo Erro_LabelCodigo_Click

    If Len(Trim(Codigo.Text)) > 0 Then
        
        objPrevVendaMensal.sCodigo = Codigo.Text
    End If
    
    Call Chama_Tela("PrevVMensalCodLista", colSelecao, objPrevVendaMensal, objEventoPrevVenda)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158582)

    End Select

    Exit Sub

End Sub
 
Private Sub objEventoPrevVenda_evSelecao(obj1 As Object)

Dim objPrevVendaMensal As ClassPrevVendaMensal

    Set objPrevVendaMensal = obj1
    
    Codigo.Text = objPrevVendaMensal.sCodigo
    
    Me.Show

    Exit Sub

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 117238

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 117238

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158583)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 117239

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 117240
    
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True
        
    Me.Show

    Call Produto_Validate(bSGECancelDummy)

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 117239

        Case 117240
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158584)

    End Select

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Ano Then
           Call LabelAno_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        End If
    
    End If

End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Validate
 
    'Verifica se o valor foi digitado
    If Len(Trim(ValorTotal.ClipText)) = 0 Then Exit Sub
    
    'Critica o valor
    lErro = Valor_Positivo_Critica(ValorTotal.Text)
    If lErro <> SUCESSO Then gError 117242
    
    lErro = Calcula_Valor_Outros
    If lErro <> SUCESSO Then gError 117244
     
    Exit Sub

Erro_ValorTotal_Validate:
  
    Cancel = True
    
    Select Case gErr
        
        Case 117242, 117244
        'Tratado nas Rotinas Chamadas
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158585)

    End Select

    Exit Sub

End Sub

Function PrevVendaMensal_Le_Codigo(sCodigo As String, iFilialEmpresa As Integer) As Long
'Verifica se a previsão de Vendas Mensal de códio e FilialEmpresa passados existem

Dim lErro As Long
Dim iFilial As Integer
Dim lComando As Long

On Error GoTo Erro_PrevVendaMensal_Le_Codigo

    'Abertura de comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 90200
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        'Pesquisa no BD se existe a Previsão de Vendas Mensais com o código passado, para a Empresa toda
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM PrevVendaMensal WHERE Codigo = ? ", iFilial, sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 90201
    Else
        'Pesquisa no BD se existe a Previsão de Vendas Mensais com o código passado, para uma FilialEmpresa
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM PrevVendaMensal WHERE Codigo = ? AND FilialEmpresa = ?", iFilial, sCodigo, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 90201
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90202
    
    'PrevVendas não encontradas
    If lErro = AD_SQL_SEM_DADOS Then gError 90203
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    PrevVendaMensal_Le_Codigo = SUCESSO
    
    Exit Function
    
Erro_PrevVendaMensal_Le_Codigo:
    
    PrevVendaMensal_Le_Codigo = gErr
    
    Select Case gErr
        
        Case 90200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90201, 90202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr, sCodigo)
        
        Case 90203 'PrevVendas não cadastrada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158586)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

