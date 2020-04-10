VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ProjetosCR 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame FramePRJ 
      Caption         =   "Nova associação"
      Height          =   630
      Left            =   120
      TabIndex        =   17
      Top             =   75
      Width           =   9240
      Begin VB.ComboBox PRJEtapa 
         Height          =   315
         Left            =   3345
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   210
         Width           =   2670
      End
      Begin VB.CommandButton BotaoInserir 
         Caption         =   "Inserir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7815
         TabIndex        =   18
         Top             =   195
         Width           =   1230
      End
      Begin MSMask.MaskEdBox PRJ 
         Height          =   300
         Left            =   795
         TabIndex        =   20
         Top             =   225
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Percentual 
         Height          =   300
         Left            =   7080
         TabIndex        =   21
         Top             =   210
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Etapa:"
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
         Index           =   25
         Left            =   2775
         TabIndex        =   24
         Top             =   255
         Width           =   570
      End
      Begin VB.Label LabelProjeto 
         AutoSize        =   -1  'True
         Caption         =   "Projeto:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   255
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Percentual:"
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
         Left            =   6090
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   255
         Width           =   990
      End
   End
   Begin VB.CommandButton BotaoLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8130
      TabIndex        =   13
      Top             =   4905
      Width           =   1230
   End
   Begin VB.CommandButton BotaoProjeto 
      Caption         =   "Projetos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   8
      Top             =   4905
      Width           =   1230
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   4950
      Picture         =   "ProjetosCR.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5340
      Width           =   1005
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   3210
      Picture         =   "ProjetosCR.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5340
      Width           =   1005
   End
   Begin VB.Frame FrameGrid 
      Caption         =   "Associações"
      Height          =   4050
      Left            =   105
      TabIndex        =   2
      Top             =   765
      Width           =   9255
      Begin VB.ComboBox Observacao 
         Height          =   315
         Left            =   4470
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   2145
         Width           =   3555
      End
      Begin VB.CheckBox CalcAuto 
         Height          =   315
         Left            =   3945
         TabIndex        =   16
         Top             =   1050
         Width           =   675
      End
      Begin MSMask.MaskEdBox QuantOriginal 
         Height          =   315
         Left            =   6435
         TabIndex        =   15
         Top             =   870
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorOriginal 
         Height          =   315
         Left            =   6210
         TabIndex        =   14
         Top             =   1320
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Percent 
         Height          =   315
         Left            =   4635
         TabIndex        =   12
         Top             =   1005
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   315
         Left            =   2940
         TabIndex        =   11
         Top             =   840
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin VB.ComboBox Item 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ProjetosCR.ctx":025C
         Left            =   525
         List            =   "ProjetosCR.ctx":025E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   570
         Width           =   1965
      End
      Begin VB.ComboBox Etapa 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ProjetosCR.ctx":0260
         Left            =   2385
         List            =   "ProjetosCR.ctx":0262
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1470
         Width           =   2790
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Left            =   4140
         TabIndex        =   5
         Top             =   1545
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Projeto 
         Height          =   315
         Left            =   390
         TabIndex        =   0
         Top             =   1515
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   2460
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   4339
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total:"
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
         Left            =   6735
         TabIndex        =   7
         Top             =   3675
         Width           =   1005
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7785
         TabIndex        =   6
         Top             =   3630
         Width           =   1245
      End
   End
End
Attribute VB_Name = "ProjetosCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public objGridItens As AdmGrid

Dim gobjTela As Object
Dim gcolPRJCustoReceita As Collection
Dim gcolItensPRJCR As Collection
Dim gbConsulta As Boolean
Dim gbTemQtd As Boolean

Private WithEvents objEventoProjeto As AdmEvento
Attribute objEventoProjeto.VB_VarHelpID = -1
Private WithEvents objEventoPRJ As AdmEvento
Attribute objEventoPRJ.VB_VarHelpID = -1

Dim iGrid_CalcAuto_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_Projeto_Col As Integer
Dim iGrid_Etapa_Col As Integer
Dim iGrid_Percent_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_QuantOriginal_Col As Integer
Dim iGrid_ValorOriginal_Col As Integer
Dim iGrid_Observacao_Col As Integer

Const ORD_PROJ = 0
Const ORD_ITEM = 1
Const ORD_VALOR = 2
Const ORD_QTD = 3

Public iAlterado As Integer
Dim iOrdenacaoAnt As Integer

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Projeto - Despesas/Receitas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ProjetosCR"

End Function

Public Sub Show()
'    Me.Show
'    Parent.SetFocus
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

Private Sub BotaoLimpar_Click()
    ValorTotal.Caption = ""
    Call Grid_Limpa(objGridItens)
End Sub

Private Sub LabelProjeto_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProjeto_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(PRJ.Text)) <> 0 Then

        lErro = Projeto_Formata(PRJ.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189103

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela_Modal("ProjetosLista", colSelecao, objProjeto, objEventoPRJ, , "Código")

    Exit Sub

Erro_LabelProjeto_Click:

    Select Case gErr

        Case 189103

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Sub
    
End Sub

Private Sub PRJ_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_PRJ_Validate

    If Len(Trim(PRJ.ClipText)) > 0 Then

        lErro = Projeto_Formata(PRJ.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189104

        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa
        
        'Le o almoxarifado pelo código ou pelo nome reduzido e joga o nome reduzido em Almoxarifado.Text
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181669
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 181670
        
    End If
    
    Call Trata_Etapa(objProjeto.lNumIntDoc, PRJEtapa)
    
    Exit Sub

Erro_PRJ_Validate:

    Cancel = True

    Select Case gErr
    
        Case 181669, 189104
        
        Case 181670
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 181671)

    End Select

    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Projeto Then
            Call BotaoProjeto_Click
        '#########################################
        'Inserido por Wagner 10/08/2006
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        '#########################################
        End If
        
    End If

End Sub

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

Public Sub Form_Load()

    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort
    
    Set objGridItens = New AdmGrid
    
    'Seta as Variáveis das Telas de browse
    Set objEventoProjeto = New AdmEvento
    Set objEventoPRJ = New AdmEvento
    
    iOrdenacaoAnt = 0
'
'    Call Inicializa_Grid_Itens(objGridItens)
    
    Call Inicializa_Mascara_Projeto(PRJ)
    
    Call Inicializa_Mascara_Projeto(Projeto)
    
    Observacao.Clear
    Call CF("Carrega_Combo_Texto", Observacao, "PRJCustoReceitaReal", STRING_PRJ_CR_OBSERVACAO, "Observacao")
   
    'Sinaliza que o Form_Loas ocorreu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

End Sub

Function Trata_Parametros(ByVal objTela As Object, ByVal colPRJCR As Collection, ByVal colItensPRJCR As Collection, Optional ByVal bConsulta As Boolean = False, Optional ByVal bTemQtd As Boolean = True) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Faz a variável global a tela apontar para a variável passada
    Set gobjTela = objTela
    Set gcolItensPRJCR = colItensPRJCR
    Set gcolPRJCustoReceita = colPRJCR
    
    gbConsulta = bConsulta
    gbTemQtd = bTemQtd
    
    Call Inicializa_Grid_Itens(objGridItens)
    
    Call Carrega_Itens(colItensPRJCR)
        
    lErro = Traz_Projeto_Tela(colPRJCR)
    If lErro <> SUCESSO Then gError 181695
       
    If bConsulta Then
        Call Desabilita_Controles
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 181695
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181695)
    
    End Select
    
    Exit Function
        
End Function

Function Saida_Celula(objGridItens As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridItens)

    If lErro = SUCESSO Then

         'Verifica qual a coluna do Grid em questão
        Select Case objGridItens.objGrid.Col

            Case iGrid_CalcAuto_Col
            
                lErro = Saida_Celula_CalcAuto(objGridItens)
                If lErro <> SUCESSO Then gError 181696
            
            Case iGrid_Item_Col
            
                lErro = Saida_Celula_Item(objGridItens)
                If lErro <> SUCESSO Then gError 181696

            Case iGrid_Projeto_Col
            
                lErro = Saida_Celula_Projeto(objGridItens)
                If lErro <> SUCESSO Then gError 181696

            Case iGrid_Etapa_Col
                
                lErro = Saida_Celula_Etapa(objGridItens)
                If lErro <> SUCESSO Then gError 181697
        
            Case iGrid_Quantidade_Col
                
                lErro = Saida_Celula_Quantidade(objGridItens)
                If lErro <> SUCESSO Then gError 181698
                
            Case iGrid_Percent_Col
                
                lErro = Saida_Celula_Percent(objGridItens)
                If lErro <> SUCESSO Then gError 181698
                
            Case iGrid_Valor_Col
                
                lErro = Saida_Celula_Valor(objGridItens)
                If lErro <> SUCESSO Then gError 181698
                
            Case iGrid_Observacao_Col
                
                lErro = Saida_Celula_Observacao(objGridItens)
                If lErro <> SUCESSO Then gError 181699
                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridItens)
        If lErro <> SUCESSO Then gError 181700

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 181696 To 181700

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181694)

    End Select

    Exit Function

End Function

Private Sub BotaoCancela_Click()
    
    'Nao mexer no obj da tela
    giRetornoTela = vbOK
    
    Unload Me
    
    Exit Sub

End Sub

Private Sub BotaoOK_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoOK_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 181701
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    iAlterado = 0
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 181701
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181693)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    lErro = Move_Tela_Memoria(gcolPRJCustoReceita)
    If lErro <> SUCESSO Then gError 181702
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 181702
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181692)

    End Select

    Exit Function

End Function

Private Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Valor()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CalcAuto_Click()

Dim dPercent As Double

    iAlterado = REGISTRO_ALTERADO
    
    dPercent = StrParaDbl(Replace(GridItens.TextMatrix(GridItens.Row, iGrid_Percent_Col), "%", "")) / 100

    'Se não estava marcado
    If GridItens.TextMatrix(GridItens.Row, iGrid_CalcAuto_Col) = S_MARCADO Then

        If gbTemQtd Then
            If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col)) <> 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dPercent * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col)))
            Else
                GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = ""
            End If
        End If

        If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col)) <> 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = Format(dPercent * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col)), "STANDARD")
        Else
            GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = ""
        End If
    
    End If
    
End Sub

Private Sub CalcAuto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub CalcAuto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub CalcAuto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CalcAuto()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Observacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Observacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Observacao()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Etapa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Etapa_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Etapa_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Etapa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Etapa
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Item_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Item_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Item_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Item_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Item
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Percent_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Percent_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Percent_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Percent_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Percent
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Projeto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Projeto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Projeto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Projeto()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridItens = Nothing
    
    Set objEventoProjeto = Nothing
    Set objEventoPRJ = Nothing

End Sub

Private Function Saida_Celula_Projeto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_Projeto

    Set objGridInt.objControle = Projeto
    
    If Len(Trim(Projeto.ClipText)) <> 0 Then

        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189089

        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa
        
        'Le o almoxarifado pelo código ou pelo nome reduzido e joga o nome reduzido em Almoxarifado.Text
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181678
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 181679
        
        GridItens.TextMatrix(GridItens.Row, iGrid_CalcAuto_Col) = S_MARCADO
        
        Call Trata_Etapa(objProjeto.lNumIntDoc, Etapa)

        If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If
    
    End If
    
    Call Trata_Etapa(objProjeto.lNumIntDoc, Etapa)
    
    Call Grid_Refresh_Checkbox(objGridInt)

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 181680

    Saida_Celula_Projeto = SUCESSO

    Exit Function

Erro_Saida_Celula_Projeto:

    Saida_Celula_Projeto = gErr

    Select Case gErr
    
        Case 181678, 181680, 189089
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 181679
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181681)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Etapa(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Etapa

    Set objGridInt.objControle = Etapa

'====================>
'TEM QUE IMPLEMENTAR QUANDO A PARTE DE ETAPAS FICAR PRONTA

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 181683

    Saida_Celula_Etapa = SUCESSO

    Exit Function

Erro_Saida_Celula_Etapa:

    Saida_Celula_Etapa = gErr

    Select Case gErr
    
        Case 181683
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181682)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao


    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 181684

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr
    
        Case 181684
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181685)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CalcAuto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CalcAuto

    Set objGridInt.objControle = CalcAuto


    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 181684

    Saida_Celula_CalcAuto = SUCESSO

    Exit Function

Erro_Saida_Celula_CalcAuto:

    Saida_Celula_CalcAuto = gErr

    Select Case gErr
    
        Case 181684
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181685)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Item(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objItemPRJCR As ClassItensPRJCR

On Error GoTo Erro_Saida_Celula_Item

    Set objGridInt.objControle = Item
    
    If Len(Trim(Item.Text)) <> 0 Then

        For Each objItemPRJCR In gcolItensPRJCR
        
            If SCodigo_Extrai(Item) = objItemPRJCR.sItem Then
                
                If gbTemQtd Then
                    If objItemPRJCR.dQuantidadeOriginal <> 0 Then
                        GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col) = Formata_Estoque(objItemPRJCR.dQuantidadeOriginal)
                    Else
                        GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col) = ""
                    End If
                End If
                
                If objItemPRJCR.dValorOriginal <> 0 Then
                    GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col) = Format(objItemPRJCR.dValorOriginal, "STANDARD")
                Else
                    GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col) = ""
                End If
           
                GridItens.TextMatrix(GridItens.Row, iGrid_Observacao_Col) = objItemPRJCR.sObservacao
           
                Exit For
            End If
        Next
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 181684

    Saida_Celula_Item = SUCESSO

    Exit Function

Erro_Saida_Celula_Item:

    Saida_Celula_Item = gErr

    Select Case gErr
    
        Case 181684
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181685)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercent As Double

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor

    If Len(Trim(Valor.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 181690
        
        If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col)) - StrParaDbl(Valor.Text) < -DELTA_VALORMONETARIO Then gError 181745
        
        If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col)) <> 0 Then
            dPercent = StrParaDbl(Valor.Text) / StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col))
        Else
            dPercent = 1
        End If
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Percent_Col) = Format(dPercent, "Percent")
        
        If StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_CalcAuto_Col)) = MARCADO Then
        
            If gbTemQtd Then
                If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col)) <> 0 Then
                    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dPercent * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col)))
                Else
                    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = ""
                End If
            End If
        
        End If

        Valor.Text = Format(Valor.Text, "STANDARD")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 181691

    Call Calcula_Totais

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr
    
        Case 181691, 181690
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 181745
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PROJETO_MAIOR_VALOR_ITEM2", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181689)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercent As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 181690

        If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col)) - StrParaDbl(Quantidade.Text) < -QTDE_ESTOQUE_DELTA Then gError 181745
                       
        If StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_CalcAuto_Col)) = MARCADO Then

            If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col)) <> 0 Then
                dPercent = StrParaDbl(Quantidade.Text) / StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col))
            Else
                dPercent = 1
            End If
        
            GridItens.TextMatrix(GridItens.Row, iGrid_Percent_Col) = Format(dPercent, "Percent")

            If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col)) <> 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = Format(dPercent * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col)), "STANDARD")
            Else
                GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = ""
            End If
        
        End If
        
        Quantidade.Text = Formata_Estoque(Quantidade.Text)
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 181691

    Call Calcula_Totais

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr
    
        Case 181691, 181690
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 181745
            Call Rotina_Erro(vbOKOnly, "ERRO_QTD_PROJETO_MAIOR_QTD_ITEM2", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181689)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Percent(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercent As Double

On Error GoTo Erro_Saida_Celula_Percent

    Set objGridInt.objControle = Percent

    If Len(Trim(Percent.Text)) > 0 Then
    
        lErro = Porcentagem_Critica(Val(Percent.Text))
        If lErro <> SUCESSO Then gError 181690

        dPercent = StrParaDbl(Replace(Percent.Text, "%", "")) / 100

        If StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_CalcAuto_Col)) = MARCADO Then

            If gbTemQtd Then
                If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col)) <> 0 Then
                    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dPercent * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantOriginal_Col)))
                Else
                    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = ""
                End If
            End If

            If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col)) <> 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = Format(dPercent * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorOriginal_Col)), "STANDARD")
            Else
                GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = ""
            End If
        
        End If
        
        Percent.Text = Format(dPercent, "Percent")
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 181691

    Call Calcula_Totais

    Saida_Celula_Percent = SUCESSO

    Exit Function

Erro_Saida_Celula_Percent:

    Saida_Celula_Percent = gErr

    Select Case gErr
    
        Case 181691, 181690
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181689)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub objEventoProjeto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoProjeto_evSelecao

    Set objProjeto = obj1
    
    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189127
    
    If Not (Me.ActiveControl Is Projeto) Then
        GridItens.TextMatrix(GridItens.Row, iGrid_Projeto_Col) = Projeto.Text
    End If

    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoProjeto_evSelecao:

    Select Case gErr
    
        Case 189127

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181688)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoProjeto_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjetoTela As String
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_BotaoProjeto_Click

    If GridItens.Row = 0 Then gError 181681

    'Verifica se o Codigo foi preenchido
    If Me.ActiveControl Is Projeto Then
        sProjetoTela = Projeto.Text
    Else
        sProjetoTela = GridItens.TextMatrix(GridItens.Row, iGrid_Projeto_Col)
    End If
    
    lErro = Projeto_Formata(sProjetoTela, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189090
    
    objProjeto.sCodigo = sProjeto

    Call Chama_Tela_Modal("ProjetosLista", colSelecao, objProjeto, objEventoProjeto, , "Código")

    Exit Sub
    
Erro_BotaoProjeto_Click:
    
    Select Case gErr
    
        Case 189090

        Case 181681
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181682)

    End Select
    
    Exit Sub

End Sub

Function Traz_Projeto_Tela(ByVal colPRJCR As Collection) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objPRJCR As ClassPRJCR
Dim objProjeto As ClassProjetos
Dim objItemPRJCR As ClassItensPRJCR
Dim objEtapa As ClassPRJEtapas

On Error GoTo Erro_Traz_Projeto_Tela

    iLinha = 0
    For Each objPRJCR In colPRJCR
    
        Set objProjeto = New ClassProjetos
        Set objEtapa = New ClassPRJEtapas
        
        iLinha = iLinha + 1
        
        objProjeto.lNumIntDoc = objPRJCR.lNumIntDocPRJ
        
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181683
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 181684
        
        If objPRJCR.lNumIntDocEtapa <> 0 Then
            
            objEtapa.lNumIntDoc = objPRJCR.lNumIntDocEtapa
            
            lErro = CF("PRJEtapas_Le_NumIntDoc", objEtapa)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185829
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 185830

            GridItens.TextMatrix(iLinha, iGrid_Etapa_Col) = objEtapa.sCodigo & SEPARADOR & objEtapa.sNomeReduzido
            
        End If
        
        
        Call CF("sCombo_Seleciona2", Item, objPRJCR.sItem)
        
        GridItens.TextMatrix(iLinha, iGrid_Item_Col) = Item.Text
        
        If objPRJCR.dValor <> 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objPRJCR.dValor, "STANDARD")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = ""
        End If
        
        If gbTemQtd Then
            If objPRJCR.dQuantidade <> 0 Then
                GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(objPRJCR.dQuantidade)
            Else
                GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = ""
            End If
        End If
        
        GridItens.TextMatrix(iLinha, iGrid_Observacao_Col) = objPRJCR.sObservacao
               
        GridItens.TextMatrix(iLinha, iGrid_CalcAuto_Col) = objPRJCR.iCalcAuto
        
        For Each objItemPRJCR In gcolItensPRJCR
            If objItemPRJCR.sItem = objPRJCR.sItem Then
                Exit For
            End If
        Next
        
        If gbTemQtd Then
            If objItemPRJCR.dQuantidadeOriginal <> 0 Then
                GridItens.TextMatrix(iLinha, iGrid_QuantOriginal_Col) = Formata_Estoque(objItemPRJCR.dQuantidadeOriginal)
            Else
                GridItens.TextMatrix(iLinha, iGrid_QuantOriginal_Col) = ""
            End If
        End If
        
        If objItemPRJCR.dValorOriginal <> 0 Then
            GridItens.TextMatrix(iLinha, iGrid_ValorOriginal_Col) = Format(objItemPRJCR.dValorOriginal, "STANDARD")
        Else
            GridItens.TextMatrix(iLinha, iGrid_ValorOriginal_Col) = ""
        End If

        If objItemPRJCR.dValorOriginal <> 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Percent_Col) = Format(objPRJCR.dValor / objItemPRJCR.dValorOriginal, "PERCENT")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Percent_Col) = Format(1, "PERCENT")
        End If
        
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189400
        
        GridItens.TextMatrix(iLinha, iGrid_Projeto_Col) = Projeto.Text 'objProjeto.sCodigo

    Next
    
    Call Grid_Refresh_Checkbox(objGridItens)
    
    objGridItens.iLinhasExistentes = iLinha
    
    Call Calcula_Totais
           
    Traz_Projeto_Tela = SUCESSO

    Exit Function

Erro_Traz_Projeto_Tela:

    Traz_Projeto_Tela = gErr
    
    Select Case gErr
    
        Case 181683, 185829, 189400
        
        Case 181684
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO", gErr, objProjeto.lNumIntDoc)
        
        Case 185830
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO", gErr, objEtapa.lNumIntDoc)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181685)
    
    End Select
    
    Exit Function
    
End Function

Function Move_Tela_Memoria(ByVal colPRJCR As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim objPRJCR As ClassPRJCR
Dim colPRJCustoReceitaAux As New Collection
Dim objItemPRJCR As ClassItensPRJCR
Dim dQuantidade As Double
Dim dValor As Double
Dim vbResult As VbMsgBoxResult
Dim objEtapa As ClassPRJEtapas
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria
   
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objPRJCR = New ClassPRJCR
        Set objProjeto = New ClassProjetos
        Set objEtapa = New ClassPRJEtapas
        
        lErro = Projeto_Formata(GridItens.TextMatrix(iIndice, iGrid_Projeto_Col), sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189310
        
        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181697
        
        objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
        objEtapa.sCodigo = SCodigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Etapa_Col))
        
        lErro = CF("PrjEtapas_Le", objEtapa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185831
    
        objPRJCR.sItem = SCodigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Item_Col))
        objPRJCR.iCalcAuto = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_CalcAuto_Col))
        objPRJCR.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Observacao_Col)
        objPRJCR.dValor = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Valor_Col))
        
        If gbTemQtd Then
            objPRJCR.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        End If
        
        objPRJCR.dPercent = StrParaDbl(Replace(GridItens.TextMatrix(iIndice, iGrid_Percent_Col), "%", "")) / 100
        objPRJCR.lNumIntDocPRJ = objProjeto.lNumIntDoc
        objPRJCR.sProjeto = objProjeto.sCodigo
        objPRJCR.lNumIntDocEtapa = objEtapa.lNumIntDoc
        
        colPRJCustoReceitaAux.Add objPRJCR
    
    Next
    
    '##############################################################
    'Validações
    'Se o total de cada item não é maior que o valor e quantidade original
    
    For Each objItemPRJCR In gcolItensPRJCR
    
        dValor = 0
        dQuantidade = 0
        For Each objPRJCR In colPRJCustoReceitaAux
    
            If objItemPRJCR.sItem = objPRJCR.sItem Then
                dValor = dValor + objPRJCR.dValor
                dQuantidade = dQuantidade + objPRJCR.dQuantidade
            End If
    
        Next
        
        If objItemPRJCR.dValorOriginal - dValor < -DELTA_VALORMONETARIO Then gError 181748
        If objItemPRJCR.dQuantidadeOriginal - dQuantidade < -DELTA_VALORMONETARIO Then gError 181749
        
        If gobjFAT.iTipoValidacaoPRJ = PRJ_TIPO_VALID_VLR_MENOR_TELA Or gobjFAT.iTipoValidacaoPRJ = PRJ_TIPO_VALID_VLR_MENOR_AMBOS Then
        
            If Abs(objItemPRJCR.dValorOriginal - dValor) > DELTA_VALORMONETARIO Then
            
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_VALOR_DIFERE_ITEMPRJ", objItemPRJCR.sItem, Format(objItemPRJCR.dValorOriginal, "STANDARD"), Format(dValor, "STANDARD"))
                If vbResult = vbNo Then gError 181750
            
            End If
            
            If Abs(objItemPRJCR.dQuantidadeOriginal - dQuantidade) > DELTA_VALORMONETARIO Then
            
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_QTD_DIFERE_ITEMPRJ", objItemPRJCR.sItem, Formata_Estoque(objItemPRJCR.dQuantidadeOriginal), Formata_Estoque(dQuantidade))
                If vbResult = vbNo Then gError 181751
            
            End If
            
        End If
    
    Next
    '##############################################################
    
    'Transfere de uma coleção para outra
    For iIndice = colPRJCR.Count To 1 Step -1
        colPRJCR.Remove iIndice
    Next
    
    For Each objPRJCR In colPRJCustoReceitaAux
        colPRJCR.Add objPRJCR
    Next
      
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 181697
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
    
        Case 181704
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DISTRIBUIDO_PRJ_MAIOR", gErr)
            
        Case 181748
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PROJETO_MAIOR_VALOR_ITEM", gErr, objItemPRJCR.sItem, Format(objItemPRJCR.dValorOriginal, "STANDARD"), Format(dValor, "STANDARD"))
        
        Case 181749
            Call Rotina_Erro(vbOKOnly, "ERRO_QTD_PROJETO_MAIOR_QTD_ITEM", gErr, objItemPRJCR.sItem, Formata_Estoque(objItemPRJCR.dQuantidadeOriginal), Formata_Estoque(dQuantidade))
        
        Case 181750, 181751, 185831, 189310
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181687)
    
    End Select
    
    Exit Function
    
End Function

Function Carrega_Itens(ByVal colItensPRJCR As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Carrega_Itens

    Item.Clear
    For Each objItensPRJCR In colItensPRJCR
        Item.AddItem objItensPRJCR.sItem & SEPARADOR & objItensPRJCR.sDescricao
    Next
      
    Carrega_Itens = SUCESSO

    Exit Function

Erro_Carrega_Itens:

    Carrega_Itens = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181687)
    
    End Select
    
    Exit Function
    
End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Projeto")
    objGridInt.colColuna.Add ("Auto")
    objGridInt.colColuna.Add ("Etapa")
    objGridInt.colColuna.Add ("Percentual")
    If gbTemQtd Then
        objGridInt.colColuna.Add ("Quantidade")
    End If
    objGridInt.colColuna.Add ("Valor")
    If gbTemQtd Then
        objGridInt.colColuna.Add ("Quant Orig")
    End If
    objGridInt.colColuna.Add ("Valor Orig")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (Projeto.Name)
    objGridInt.colCampo.Add (CalcAuto.Name)
    objGridInt.colCampo.Add (Etapa.Name)
    objGridInt.colCampo.Add (Percent.Name)
    If gbTemQtd Then
        objGridInt.colCampo.Add (Quantidade.Name)
    End If
    objGridInt.colCampo.Add (Valor.Name)
    If gbTemQtd Then
        objGridInt.colCampo.Add (QuantOriginal.Name)
    End If
    objGridInt.colCampo.Add (ValorOriginal.Name)
    objGridInt.colCampo.Add (Observacao.Name)

    iIndice = 0
    iGrid_Item_Col = 1 + iIndice
    iGrid_Projeto_Col = 2 + iIndice
    iGrid_CalcAuto_Col = 3 + iIndice
    iGrid_Etapa_Col = 4 + iIndice
    iGrid_Percent_Col = 5 + iIndice
    If gbTemQtd Then
        iGrid_Quantidade_Col = 6 + iIndice
        iIndice = iIndice + 1
    End If
    iGrid_Valor_Col = 6 + iIndice
    If gbTemQtd Then
        iGrid_QuantOriginal_Col = 7 + iIndice
        iIndice = iIndice + 1
    End If
    iGrid_ValorOriginal_Col = 7 + iIndice
    iGrid_Observacao_Col = 8 + iIndice
    
    'Etapa.Left = POSICAO_FORA_TELA
    
    If Not gbTemQtd Then
        Quantidade.Left = POSICAO_FORA_TELA
        QuantOriginal.Left = POSICAO_FORA_TELA
    End If

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
    
    Call Ordenacao_ClickGrid(objGridItens)

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    
    Call Calcula_Totais

End Sub

Private Sub Calcula_Totais()

Dim iIndice As Integer
Dim dValor As Double
Dim objItemPRJCR As ClassItensPRJCR

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        For Each objItemPRJCR In gcolItensPRJCR
            If SCodigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Item_Col)) = objItemPRJCR.sItem Then
                Exit For
            End If
        Next
        
        If Not (objItemPRJCR Is Nothing) Then
    
            If objItemPRJCR.iNegativo <> MARCADO Then
                dValor = dValor + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Valor_Col))
            Else
                dValor = dValor - StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Valor_Col))
            End If
            
        End If
    
    Next
    
    ValorTotal.Caption = Format(dValor, "STANDARD")

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sEtapaAnt As String
Dim objProjeto As New ClassProjetos
Dim iIndice As Integer
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable
   
    If Not gbConsulta Then
        
        Select Case objControl.Name
    
            Case Item.Name
                If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Item_Col))) > 0 Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
    
            Case Projeto.Name
                If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Item_Col))) > 0 Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
                
            Case Percent.Name
                'Se o projeto estiver preenchido, habilita o controle
                If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Projeto_Col))) > 0 And GridItens.TextMatrix(iLinha, iGrid_CalcAuto_Col) = S_MARCADO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
                
            Case QuantOriginal.Name, ValorOriginal.Name
                objControl.Enabled = False
                
            Case Else
                
                'Se o projeto estiver preenchido, habilita o controle
                If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Projeto_Col))) > 0 Then
                    objControl.Enabled = True
                
                    If objControl.Name = Etapa.Name Then
                    
                        lErro = Projeto_Formata(GridItens.TextMatrix(iLinha, iGrid_Projeto_Col), sProjeto, iProjetoPreenchido)
                        If lErro <> SUCESSO Then gError 189311
                    
                        objProjeto.sCodigo = sProjeto
                        objProjeto.iFilialEmpresa = giFilialEmpresa
                        
                        lErro = CF("Projetos_Le", objProjeto)
                        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185828
                    
                        sEtapaAnt = objControl.Text
                        
                        Call Trata_Etapa(objProjeto.lNumIntDoc, Etapa)
                        
                        For iIndice = 0 To Etapa.ListCount - 1
                            If Etapa.List(iIndice) = sEtapaAnt Then
                                Etapa.ListIndex = iIndice
                                Exit For
                            End If
                        Next
                    
                    End If
                
                Else
                    objControl.Enabled = False
                End If
                

                           
        End Select

    Else
        objControl.Enabled = False
    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 185828, 189311

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181686)

    End Select

    Exit Sub

End Sub


Private Sub objEventoPRJ_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoPRJ_evSelecao

    Set objProjeto = obj1
    
    lErro = Retorno_Projeto_Tela(PRJ, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189107

    Call PRJ_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoPRJ_evSelecao:

    Select Case gErr
    
        Case 189107

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181629)

    End Select

    Exit Sub

End Sub

Private Sub Percentual_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Percentual_Validate

    'Veifica se Percentual está preenchida
    If Len(Trim(Percentual.Text)) <> 0 Then

       'Critica a Percentual
       lErro = Porcentagem_Critica(Percentual.Text)
       If lErro <> SUCESSO Then gError 181746

    End If

    Exit Sub

Erro_Percentual_Validate:

    Cancel = True

    Select Case gErr

        Case 181746

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181747)

    End Select

    Exit Sub

End Sub

Private Sub BotaoInserir_Click()
    
Dim lErro As Long
Dim dPercent As Double
Dim iLinha As Integer
Dim objItemPRJCR As ClassItensPRJCR
    
On Error GoTo Erro_BotaoInserir_Click
    
    For Each objItemPRJCR In gcolItensPRJCR
    
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        iLinha = objGridItens.iLinhasExistentes
        
        dPercent = StrParaDbl(Val(Percentual.Text)) / 100
        
        GridItens.TextMatrix(iLinha, iGrid_Percent_Col) = Format(dPercent, "PERCENT")
        
        Call CF("sCombo_Seleciona2", Item, objItemPRJCR.sItem)
        
        GridItens.TextMatrix(iLinha, iGrid_Item_Col) = Item.Text
        GridItens.TextMatrix(iLinha, iGrid_CalcAuto_Col) = S_MARCADO
        
        If gbTemQtd Then
            If objItemPRJCR.dQuantidadeOriginal <> 0 Then
                GridItens.TextMatrix(iLinha, iGrid_QuantOriginal_Col) = Formata_Estoque(objItemPRJCR.dQuantidadeOriginal)
                GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(dPercent * objItemPRJCR.dQuantidadeOriginal)
            Else
                GridItens.TextMatrix(iLinha, iGrid_QuantOriginal_Col) = ""
                GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = ""
            End If
        End If
        
        If objItemPRJCR.dValorOriginal <> 0 Then
            GridItens.TextMatrix(iLinha, iGrid_ValorOriginal_Col) = Format(objItemPRJCR.dValorOriginal, "STANDARD")
            GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = Format(dPercent * objItemPRJCR.dValorOriginal, "STANDARD")
        Else
            GridItens.TextMatrix(iLinha, iGrid_ValorOriginal_Col) = ""
            GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = ""
        End If
        
        GridItens.TextMatrix(iLinha, iGrid_Observacao_Col) = objItemPRJCR.sObservacao
        
        GridItens.TextMatrix(iLinha, iGrid_Projeto_Col) = PRJ.Text
        GridItens.TextMatrix(iLinha, iGrid_Etapa_Col) = PRJEtapa.Text
        
    Next
    
    Call Grid_Refresh_Checkbox(objGridItens)
    
    Call Calcula_Totais

    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_BotaoInserir_Click:

    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181693)

    End Select

    Exit Sub
    
End Sub

Public Function Desabilita_Controles() As Long

Dim lErro As Long

On Error GoTo Erro_Desabilita_Controles

    FramePRJ.Visible = False
    
    BotaoLimpar.Enabled = False
    BotaoProjeto.Enabled = False

    Desabilita_Controles = SUCESSO

    Exit Function

Erro_Desabilita_Controles:

    Desabilita_Controles = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Private Function Trata_Etapa(ByVal lNumIntDocPRJ As Long, ByVal objCombo As Object) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos

On Error GoTo Erro_Trata_Etapa
    
    If lNumIntDocPRJ <> 0 Then

        objProjeto.lNumIntDoc = lNumIntDocPRJ
        
        objCombo.AddItem ""
    
        lErro = CF("CarregaCombo_Etapas", objProjeto, objCombo)
        If lErro <> SUCESSO Then gError 185234
    
    Else
    
        objCombo.Clear
        
    End If

    Trata_Etapa = SUCESSO

    Exit Function

Erro_Trata_Etapa:

    Trata_Etapa = gErr

    Select Case gErr
    
        Case 185234

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185235)

    End Select

    Exit Function

End Function
