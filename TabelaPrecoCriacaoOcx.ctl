VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl TabelaPrecoCriacaoOcx 
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   6210
   Begin VB.Frame Frame2 
      Caption         =   "Desconto sobre pre�o de tabela"
      Height          =   930
      Left            =   165
      TabIndex        =   20
      Top             =   3705
      Width           =   5715
      Begin VB.OptionButton LimitarDesconto 
         Caption         =   "No m�ximo de:"
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
         Left            =   1995
         TabIndex        =   23
         Top             =   435
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.OptionButton LimitarDesconto 
         Caption         =   "N�o Limitar"
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
         Index           =   0
         Left            =   210
         TabIndex        =   22
         Top             =   420
         Width           =   1425
      End
      Begin MSMask.MaskEdBox DescontoMaximo 
         Height          =   315
         Left            =   3630
         TabIndex        =   21
         Top             =   375
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      ItemData        =   "TabelaPrecoCriacaoOcx.ctx":0000
      Left            =   1125
      List            =   "TabelaPrecoCriacaoOcx.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1320
      Width           =   2310
   End
   Begin VB.ComboBox Moeda 
      Height          =   315
      Left            =   1125
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1830
      Width           =   1935
   End
   Begin VB.ComboBox CargoMinimo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3915
      TabIndex        =   12
      Top             =   3255
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ajuste autom�tico de pre�o"
      Height          =   735
      Left            =   180
      TabIndex        =   10
      Top             =   2280
      Width           =   5685
      Begin MSMask.MaskEdBox VlrCompCoef 
         Height          =   315
         Left            =   4635
         TabIndex        =   5
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin VB.CheckBox AjustaTabVlrCompCoefMaior 
         Caption         =   "Ajusta o valor quando o valor da compra vezes               for maior que o valor de tabela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   4
         Top             =   270
         Width           =   4665
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1680
      Picture         =   "TabelaPrecoCriacaoOcx.ctx":0029
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numera��o Autom�tica"
      Top             =   330
      Width           =   300
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   4065
      Picture         =   "TabelaPrecoCriacaoOcx.ctx":0113
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   5025
      Picture         =   "TabelaPrecoCriacaoOcx.ctx":026D
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   885
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Top             =   795
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   315
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Markup 
      Height          =   315
      Left            =   4020
      TabIndex        =   1
      Top             =   1770
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "#,##0.00##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Comissao 
      Height          =   315
      Left            =   1485
      TabIndex        =   16
      Top             =   3255
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "#,##0.00##"
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      Caption         =   "Tipo:"
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
      Height          =   270
      Left            =   555
      TabIndex        =   18
      Top             =   1365
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Comiss�o:"
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
      Left            =   540
      TabIndex        =   17
      Top             =   3300
      Width           =   855
   End
   Begin VB.Label Label7 
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
      Height          =   210
      Left            =   375
      TabIndex        =   15
      Top             =   1875
      Width           =   615
   End
   Begin VB.Label LabelCargoMinimo 
      AutoSize        =   -1  'True
      Caption         =   "Cargo M�nimo:"
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
      Left            =   2625
      TabIndex        =   13
      Top             =   3330
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mark-up:"
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
      Left            =   3240
      TabIndex        =   11
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descri��o:"
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
      Height          =   165
      Left            =   90
      TabIndex        =   9
      Top             =   840
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "C�digo:"
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
      Left            =   375
      TabIndex        =   8
      Top             =   360
      Width           =   660
   End
End
Attribute VB_Name = "TabelaPrecoCriacaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjTabelaPreco As ClassTabelaPreco

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lNumProxCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Pega o n�mero do pr�ximo c�digo
    lErro = CF("TabelaPreco_Automatico", lNumProxCodigo)
    If lErro <> SUCESSO Then Error 57539

    'Mostra o c�digo na tela
    Codigo.Text = CStr(lNumProxCodigo)
        
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57539
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174468)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_BotaoOK_Click

    'Verifica preenchimento do C�digo e Descri��o
    If Len(Trim(Codigo.Text)) = 0 Then Error 28116
    If Len(Trim(Descricao.Text)) = 0 Then Error 28117

    'Preenche objTabelaPreco
    objTabelaPreco.iCodigo = Codigo.Text
    objTabelaPreco.sDescricao = Descricao.Text
    
    If AjustaTabVlrCompCoefMaior.Value = vbChecked Then
        objTabelaPreco.iAjustaTabVlrCompCoefMaior = MARCADO
    Else
        objTabelaPreco.iAjustaTabVlrCompCoefMaior = DESMARCADO
    End If
    
    objTabelaPreco.dVlrCompCoef = StrParaDbl(VlrCompCoef.Text)
    objTabelaPreco.dMarkUp = StrParaDbl(Markup.Text)
    If Len(Trim(Comissao.Text)) <> 0 Then
        objTabelaPreco.dComissao = StrParaDbl(Comissao.Text) / 100
    Else
        objTabelaPreco.dComissao = -1
    End If
    
    objTabelaPreco.iCargoMinimo = Codigo_Extrai(CargoMinimo.Text)
    objTabelaPreco.iMoeda = Codigo_Extrai(Moeda.Text)
    
    If LimitarDesconto.Item(0).Value = True Then
    
        objTabelaPreco.iDescontoLimitado = 0
    
    Else
    
        objTabelaPreco.iDescontoLimitado = 1
        objTabelaPreco.dDescontoMaximo = StrParaDbl(DescontoMaximo.Text) / 100
    
    End If
    
    If Tipo.ListIndex <> -1 Then
        objTabelaPreco.iTipo = Tipo.ItemData(Tipo.ListIndex)
    End If
    
    'Chama TabelaPreco_Cria
    lErro = CF("TabelaPreco_Cria", objTabelaPreco)
    If lErro <> SUCESSO Then Error 28145

    'Iguala gobjTabelaPreco
    gobjTabelaPreco.iCodigo = objTabelaPreco.iCodigo
    gobjTabelaPreco.sDescricao = objTabelaPreco.sDescricao
    gobjTabelaPreco.dVlrCompCoef = objTabelaPreco.dVlrCompCoef
    gobjTabelaPreco.dMarkUp = objTabelaPreco.dMarkUp
    gobjTabelaPreco.dComissao = objTabelaPreco.dComissao
    gobjTabelaPreco.iAjustaTabVlrCompCoefMaior = objTabelaPreco.iAjustaTabVlrCompCoefMaior
    gobjTabelaPreco.iCargoMinimo = objTabelaPreco.iCargoMinimo
    gobjTabelaPreco.iMoeda = objTabelaPreco.iMoeda
    gobjTabelaPreco.iTipo = objTabelaPreco.iTipo
    gobjTabelaPreco.iDescontoLimitado = objTabelaPreco.iDescontoLimitado
    gobjTabelaPreco.dDescontoMaximo = objTabelaPreco.dDescontoMaximo

    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case Err

        Case 28116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 28117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)
        
        Case 28145
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174469)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se foi preenchido o campo C�digo
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub
    
    'Verifica se � do tipo Inteiro e positivo
    lErro = Inteiro_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 28115
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 28115

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174470)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega a combo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_CARGO_VENDEDOR, CargoMinimo)
    If lErro <> SUCESSO Then gError 124021
    
    Call Carrega_Moeda
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 124021
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174471)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjTabelaPreco = Nothing

End Sub

Function Trata_Parametros(objTabelaPreco As ClassTabelaPreco) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Faz gobjTabelaPreco referenciar objTabelaPreco
    Set gobjTabelaPreco = objTabelaPreco
    
    If objTabelaPreco.iCodigo <> 0 Then
    
        'Mostra o c�digo na tela
        Codigo.Text = CStr(objTabelaPreco.iCodigo)
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174472)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TABELA_PRECOS_CRIACAO
    Set Form_Load_Ocx = Me
    Caption = "Tabela de Pre�os"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TabelaPrecoCriacao"
    
End Function

Public Sub Show()
'    Parent.Show
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

Private Sub OptAjustaValorMenor_Click()

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

Private Sub AjustaTabVlrCompCoefMaior_Click()
    If AjustaTabVlrCompCoefMaior.Value = vbChecked Then
        VlrCompCoef.Enabled = True
    Else
        VlrCompCoef.Enabled = False
    End If
End Sub

Private Sub VlrCompCoef_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VlrCompCoef_Validate

    If Len(Trim(VlrCompCoef.Text)) = 0 Then Exit Sub

    'Critica quantidade
    lErro = Valor_Positivo_Critica(VlrCompCoef.Text)
    If lErro <> SUCESSO Then gError 200245
    
    VlrCompCoef.Text = Formata_Estoque(CDbl(VlrCompCoef.Text))

    Exit Sub

Erro_VlrCompCoef_Validate:

    Cancel = True

    Select Case gErr

        Case 200245
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200246)

    End Select

    Exit Sub

End Sub

Private Sub MarkUp_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MarkUp_Validate

    If Len(Trim(Markup.Text)) = 0 Then Exit Sub

    'Critica quantidade
    lErro = Valor_Positivo_Critica(Markup.Text)
    If lErro <> SUCESSO Then gError 200245
    
    Markup.Text = Formata_Estoque(CDbl(Markup.Text))

    Exit Sub

Erro_MarkUp_Validate:

    Cancel = True

    Select Case gErr

        Case 200245
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200246)

    End Select

    Exit Sub

End Sub

Private Sub CargoMinimo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CargoMinimo_Validate

    If CargoMinimo.Text <> "" Then
    
        'Valida o tipo de relacionamento selecionado pelo cliente
        lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_CARGO_VENDEDOR, CargoMinimo, "AVISO_CRIAR_CARGO_VENDEDOR")
        If lErro <> SUCESSO Then gError 195867
    
    End If
    
    Exit Sub

Erro_CargoMinimo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 195867
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195868)

    End Select

End Sub

Private Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas) 'leo colocar CF
    If lErro <> SUCESSO Then gError 103371
    
    'se n�o existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
    
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 103371
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164661)
    
    End Select

End Function

Private Sub Comissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Comissao_Validate

    If Len(Trim(Comissao.Text)) = 0 Then Exit Sub

    'Critica quantidade
    lErro = Valor_Positivo_Critica(Comissao.Text)
    If lErro <> SUCESSO Then gError 200245
    
    Comissao.Text = Formata_Estoque(CDbl(Comissao.Text))

    Exit Sub

Erro_Comissao_Validate:

    Cancel = True

    Select Case gErr

        Case 200245
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200246)

    End Select

    Exit Sub

End Sub

