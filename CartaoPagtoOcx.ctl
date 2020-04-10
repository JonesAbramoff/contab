VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl CartaoPagtoOcx 
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ScaleHeight     =   4395
   ScaleWidth      =   4770
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   1170
      Picture         =   "CartaoPagtoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   1035
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   555
      Left            =   2475
      Picture         =   "CartaoPagtoOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1035
   End
   Begin VB.TextBox Aprovacao 
      Height          =   315
      Left            =   1905
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2535
      Width           =   2760
   End
   Begin VB.TextBox NumeroCartao 
      Height          =   315
      Left            =   1905
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1334
      Width           =   2760
   End
   Begin VB.ComboBox Parcelamento 
      Height          =   315
      ItemData        =   "CartaoPagtoOcx.ctx":025C
      Left            =   1905
      List            =   "CartaoPagtoOcx.ctx":025E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   727
      Width           =   2760
   End
   Begin VB.ComboBox Adm 
      Height          =   315
      ItemData        =   "CartaoPagtoOcx.ctx":0260
      Left            =   1905
      List            =   "CartaoPagtoOcx.ctx":0262
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2760
   End
   Begin MSMask.MaskEdBox Validade 
      Height          =   300
      Left            =   1905
      TabIndex        =   6
      Top             =   1941
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   5
      Format          =   "mm/yyyy"
      Mask            =   "##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownEmissao 
      Height          =   300
      Left            =   3000
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataCartao 
      Height          =   300
      Left            =   1905
      TabIndex        =   13
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data da Transação:"
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
      Left            =   135
      TabIndex        =   14
      Top             =   3150
      Width           =   1710
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Aprovação:"
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
      Left            =   855
      TabIndex        =   9
      Top             =   2595
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Válido Até:"
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
      Left            =   900
      TabIndex        =   7
      Top             =   1965
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número:"
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
      Left            =   1125
      TabIndex        =   5
      Top             =   1395
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Parcelamento:"
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
      Left            =   615
      TabIndex        =   3
      Top             =   765
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cartão:"
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
      Left            =   1215
      TabIndex        =   2
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "CartaoPagtoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gobjParcPV As ClassParcelaPedidoVenda

Private Sub Adm_Click()
    
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

    
On Error GoTo Erro_Adm_Click
    
    iAlterado = REGISTRO_ALTERADO
    
    Parcelamento.Clear
    
    If Adm.ListIndex <> -1 Then
    
        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
        objAdmMeioPagto.iCodigo = Adm.ItemData(Adm.ListIndex)
    
        'Lê para cada admnistradoras os Pacelamentos Vinculados
        lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO Then gError 183045

        For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                
            Parcelamento.AddItem objAdmMeioPagtoCondPagto.sNomeParcelamento
            Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento
                        
        Next


    End If

    Exit Sub
    
Erro_Adm_Click:

    Select Case gErr
          
        Case 183045
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183046)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 178966

    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
          
        Case 178966
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178967)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colAdmMeioPagto As New Collection
Dim objAdmMeioPagto As ClassAdmMeioPagto

On Error GoTo Erro_Form_Load

    Validade.Format = ""

    'Le os meios de pagamento
    lErro = CF("AdmMeioPagto_Le_Todas", colAdmMeioPagto)
    If lErro <> SUCESSO Then gError 104033
    
    'Adcionar todos os Meios de Pagto na ListBox
    For Each objAdmMeioPagto In colAdmMeioPagto
               
        If objAdmMeioPagto.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Or objAdmMeioPagto.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO Then
               
            Adm.AddItem objAdmMeioPagto.sNome
            Adm.ItemData(Adm.NewIndex) = objAdmMeioPagto.iCodigo
        
        End If
        
    Next

    Adm.AddItem ""
    Adm.ItemData(Adm.NewIndex) = 0

    'preecher a data emissão com a data atual
    DataCartao.PromptInclude = False
    DataCartao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCartao.PromptInclude = True

    lErro_Chama_Tela = SUCESSO

    iAlterado = 0

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183044)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set gobjParcPV = Nothing
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim sValidade As String
Dim dtData As Date

On Error GoTo Erro_Gravar_Registro

    'Verifica se os campos essenciais da tela foram preenchidos
    If Len(Trim(Adm.Text)) = 0 Then gError 183036
    If Len(Trim(Parcelamento.Text)) = 0 Then gError 183037
    If Len(Trim(NumeroCartao.Text)) = 0 Then gError 183038
    If Len(Trim(Validade.ClipText)) = 0 Then gError 183039
    If Len(Trim(Aprovacao.Text)) = 0 Then gError 183040
    If Len(Trim(DataCartao.ClipText)) = 0 Then gError 183041

    gobjParcPV.iAdmMeioPagto = Adm.ItemData(Adm.ListIndex)
    gobjParcPV.iParcelamento = Parcelamento.ItemData(Parcelamento.ListIndex)
    gobjParcPV.sNumeroCartao = NumeroCartao.Text
    
    sValidade = "01/" & Validade.Text
    
    dtData = DateAdd("m", 1, StrParaDate(sValidade))
    
    dtData = DateAdd("d", -1, dtData)
    
    gobjParcPV.dtValidadeCartao = dtData
    gobjParcPV.sAprovacaoCartao = Aprovacao.Text
    gobjParcPV.dtDataTransacaoCartao = StrParaDate(DataCartao.Text)

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 183036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ADMINISTRADORA_NAO_PREENCHIDA", gErr)

        Case 183037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_SELECIONADO", gErr)

        Case 183038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_CARTAO_NAO_PREENCHIDO", gErr)

        Case 183039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALIDADE_CARTAO_NAO_PREENCHIDA", gErr)

        Case 183040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APROVACAO_CARTAO_NAO_PREENCHIDO", gErr)

        Case 183041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_TRANSACAO_CARTAO_NAO_PREENCHIDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183042)

     End Select

     Exit Function

End Function

Function Trata_Parametros(objParcPV As ClassParcelaPedidoVenda) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gobjParcPV = objParcPV
    
    If objParcPV.iAdmMeioPagto <> 0 Then
    
        For iIndice = 0 To Adm.ListCount - 1
            If Adm.ItemData(iIndice) = objParcPV.iAdmMeioPagto Then
                Adm.ListIndex = iIndice
                Exit For
            End If
        Next
                
        For iIndice = 0 To Parcelamento.ListCount - 1
            If Parcelamento.ItemData(iIndice) = objParcPV.iParcelamento Then
                Parcelamento.ListIndex = iIndice
                Exit For
            End If
        Next
        
        NumeroCartao.Text = objParcPV.sNumeroCartao
        
        Validade.Text = Format(objParcPV.dtValidadeCartao, "mm/yy")
        
        Aprovacao.Text = objParcPV.sAprovacaoCartao
    
        DataCartao.Text = Format(objParcPV.dtDataTransacaoCartao, "dd/mm/yy")
    
        iAlterado = 0
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183043)

    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Pagamento em Cartão"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CartaoPagto"
    
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

Private Sub Parcelamento_Click()
    iAlterado = REGISTRO_ALTERADO
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub NumeroCartao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Aprovacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Validade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Validade_GotFocus()

    Call MaskEdBox_TrataGotFocus(Validade)

End Sub


Private Sub Validade_Validate(Cancel As Boolean)

Dim sValidade As String
Dim iMes As Integer
Dim lErro As Long

On Error GoTo Erro_Validade_Validate

    'Verifica se a data de validade está preenchida
    If Len(Trim(Validade.ClipText)) = 0 Then Exit Sub

    iMes = StrParaInt(Left(Validade.Text, 2))
    
    If iMes < 1 Or iMes > 12 Then gError 178969

    sValidade = "01/" & Validade.Text

    'Critica a data digitada
    lErro = Data_Critica(sValidade)
    If lErro <> SUCESSO Then gError 178971
    
'    dtData = DateAdd("m", 1, StrParaDate(sValidade))
'
'    dtData = DateAdd("d", -1, dtData)
'
'    If dtData < gdtDataHoje Then gError 178972



    Exit Sub
    
Erro_Validade_Validate:

    Select Case gErr
          
        Case 178969
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_INVALIDO", gErr)
        
'        Case 178972
'            Call Rotina_Erro(vbOKOnly, "ERRO_CARTAO_FORA_VALIDADE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178970)
     
    End Select
     
    Exit Sub

End Sub

Private Sub DataCartao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Private Sub DataCartao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCartao_Validate

    'Verifica se a data de emissao está preenchida
    If Len(Trim(DataCartao.ClipText)) = 0 Then Exit Sub

    'Verifica se a data emissao é válida
    lErro = Data_Critica(DataCartao.Text)
    If lErro <> SUCESSO Then gError 183047

    Exit Sub

Erro_DataCartao_Validate:

    Cancel = True

    Select Case gErr

        Case 183046

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183048)

    End Select

    Exit Sub

End Sub

Private Sub DataCartao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataCartao)

End Sub

