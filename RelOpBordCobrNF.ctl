VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpBordCobrNFOcx 
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   ScaleHeight     =   1500
   ScaleWidth      =   5175
   Begin VB.ComboBox Serie 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1095
      Width           =   975
   End
   Begin VB.ComboBox Cobrador 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Caption         =   "Atributos"
      Height          =   225
      Left            =   135
      TabIndex        =   4
      Top             =   2010
      Visible         =   0   'False
      Width           =   4890
      Begin VB.Label LabelValor 
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
         Left            =   2850
         TabIndex        =   10
         Top             =   315
         Width           =   1350
      End
      Begin VB.Label LabelCarteira 
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
         Left            =   1095
         TabIndex        =   9
         Top             =   720
         Width           =   3150
      End
      Begin VB.Label LabelEmissao 
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
         Left            =   1095
         TabIndex        =   8
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   2250
         TabIndex        =   7
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Carteira:"
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
         Left            =   285
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emiss�o:"
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
         Left            =   255
         TabIndex        =   5
         Top             =   300
         Width           =   765
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
      Height          =   615
      Left            =   3450
      Picture         =   "RelOpBordCobrNF.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   1575
   End
   Begin MSMask.MaskEdBox NumBordero 
      Height          =   285
      Left            =   1245
      TabIndex        =   2
      Top             =   660
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NFiscalInicial 
      Height          =   300
      Left            =   1245
      TabIndex        =   13
      Top             =   1560
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NFiscalFinal 
      Height          =   300
      Left            =   2700
      TabIndex        =   14
      Top             =   1560
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "At�:"
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
      Left            =   2325
      TabIndex        =   17
      Top             =   1620
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   870
      TabIndex        =   16
      Top             =   1620
      Width           =   300
   End
   Begin VB.Label LabelSerie 
      AutoSize        =   -1  'True
      Caption         =   "S�rie:"
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
      Left            =   690
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   1155
      Width           =   510
   End
   Begin VB.Label Label9 
      Caption         =   "Cobrador:"
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
      Height          =   255
      Left            =   345
      TabIndex        =   11
      Top             =   225
      Width           =   855
   End
   Begin VB.Label LabelNumBordero 
      AutoSize        =   -1  'True
      Caption         =   "No. Border�:"
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
      Left            =   105
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "RelOpBordCobrNFOcx"
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

Private WithEvents objEventoBorderoCobranca As AdmEvento
Attribute objEventoBorderoCobranca.VB_VarHelpID = -1
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1


Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

   If Not (gobjRelatorio Is Nothing) Then gError 66577
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 66577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167360)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoBorderoCobranca = New AdmEvento
    Set objEventoSerie = New AdmEvento
    
    lErro = Carrega_Cobradores()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167361)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objSerie As New ClassSerie

On Error GoTo Erro_BotaoExecutar_Click

    If Len(Trim(NumBordero.Text)) = 0 Then gError 66570
    If Len(Trim(Cobrador.Text)) = 0 Then gError 66583
    'Verifica se a S�rie est� cadastrada
    If Len(Trim(Serie.Text)) = 0 Then gError 60375
    
    objBorderoCobranca.lNumBordero = CLng(NumBordero.Text)
    
    lErro = CF("BorderoCobranca_Le", objBorderoCobranca)
    If lErro <> SUCESSO And lErro <> 46366 Then gError 66584
    
    If lErro = 46366 Then gError 66585
    
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 66578
    
    objSerie.sSerie = Serie.List(Serie.ListIndex)

    'Tenta ler a s�rie no BD
    lErro = CF("Serie_Le", objSerie)
    If lErro <> SUCESSO And lErro <> 22202 Then gError ERRO_SEM_MENSAGEM
    
    gobjRelatorio.sNomeTsk = objSerie.sNomeTsk

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 60375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 66578, 66584

        Case 66570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_NAO_INFORMADO", gErr)
        
        Case 66583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
            
        Case 66585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDERO_COBRANCA_NAO_CADASTRADO", gErr, objBorderoCobranca.lNumBordero)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167362)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim lNumIni As Long, lNumFim As Long

On Error GoTo Erro_PreencherRelOp
    
    If Len(Trim(NumBordero.Text)) = 0 Then gError 66579
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 66580
         
    lErro = objRelOpcoes.IncluirParametro("NBORDERO", NumBordero.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
          
    lErro = objRelOpcoes.IncluirParametro("NCOBRADOR", Codigo_Extrai(Cobrador.Text))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TSERIE", Serie.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("Bordero_Le_NumNF_Serie", StrParaLong(NumBordero.Text), Serie.Text, lNumIni, lNumFim)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NNFISCALINIC", CStr(lNumIni))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NNFISCALFIM", CStr(lNumFim))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    objRelOpcoes.sSelecao = "FiltroBordero = 1"
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
                
        Case 66579
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCHA_CAMPOS_OBRIGATORIOS", gErr)
        
        Case 66585
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167363)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoBorderoCobranca = Nothing
    Set objEventoSerie = Nothing
 
 End Sub

Private Sub LabelNUmBordero_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objBorderoCobranca As New ClassBorderoCobranca

On Error GoTo Erro_LabelNUmBordero_Click
    
    If Len(Trim(Cobrador.Text)) = 0 Then gError 66586
    
    If Len(Trim(NumBordero.Text)) > 0 Then objBorderoCobranca.lNumBordero = CLng(NumBordero.Text)
    
    colSelecao.Add Codigo_Extrai(Cobrador.Text)
    
    'Chama Tela BorderoCobrancaLista
    Call Chama_Tela("BorderoDeCobrancaLista", colSelecao, objBorderoCobranca, objEventoBorderoCobranca)

    Exit Sub

Erro_LabelNUmBordero_Click:

    Select Case gErr

        Case 66586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167364)

    End Select

    Exit Sub
    
End Sub

Private Sub NumBordero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumBordero)

End Sub

Private Sub objEventoBorderoCobranca_evSelecao(obj1 As Object)

Dim objBorderoCobranca As ClassBorderoCobranca
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim lErro As Long

On Error GoTo Erro_objEventoBorderoCobranca_evSelecao

    Set objBorderoCobranca = obj1
    
    NumBordero.PromptInclude = False
    NumBordero.Text = objBorderoCobranca.lNumBordero
    NumBordero.PromptInclude = True
    LabelEmissao.Caption = Format(objBorderoCobranca.dtDataEmissao, "dd/mm/yy")
    LabelValor = Format(objBorderoCobranca.dValor, "Standard")
    
    objCarteiraCobranca.iCodigo = objBorderoCobranca.iCodCarteiraCobranca
    
    'L� a Carteira de Cobran�a
    lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
    If lErro <> SUCESSO And lErro <> 23413 Then gError 66571
        
    'Se n�o achou, erro
    If lErro = 23413 Then gError 66572
    
    'Coloca na a Carteira de Cobran�a
    LabelCarteira.Caption = objCarteiraCobranca.sDescricao
    
    Me.Show

    Exit Sub

Erro_objEventoBorderoCobranca_evSelecao:

    Select Case gErr
        
        Case 66571
        
        Case 66572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", gErr, objCarteiraCobranca.iCodigo)
                    
        Case Else
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167365)

    End Select
    
    Exit Sub
    
End Sub

Private Function Carrega_Cobradores() As Long

Dim lErro As Long
Dim objCobrador As ClassCobrador
Dim colCobrador As New Collection

On Error GoTo Erro_Carrega_Cobradores

    'Carrega a Cole��o de Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", colCobrador)
    If lErro <> SUCESSO Then gError 66582
    
    'Preenche a ComboBox Cobrador com os objetos da cole��o de Cobradores
    For Each objCobrador In colCobrador

        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        End If

    Next

    Carrega_Cobradores = SUCESSO
    
    Exit Function
    
Erro_Carrega_Cobradores:

    Carrega_Cobradores = gErr
    
    Select Case gErr
    
        Case 66582
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167366)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_BORDERO_COBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Emiss�o de NF por Border� de Cobran�a"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBordCobrNF"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NumBordero Then
            Call LabelNUmBordero_Click
        ElseIf Me.ActiveControl Is Serie Then
            Call LabelSerie_Click
        End If
    
    End If

End Sub

Private Sub LabelSerie_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a S�rie da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieListaModal
    Call Chama_Tela("SerieListaModal", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Private Function Carrega_Serie() As Long
'Carrega a combo de S�ries com as s�ries lidas do BD

Dim lErro As Long
Dim objSerie As ClassSerie
Dim colSerie As New colSerie
Dim sSeriePadrao As String

On Error GoTo Erro_Carrega_Serie

    'L� as s�ries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next
    
    'L� s�rie Padr�o
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If Len(Trim(sSeriePadrao)) > 0 Then
        Serie.Text = sSeriePadrao
        Call Serie_Validate(bSGECancelDummy)
    End If

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157166)

    End Select

    Exit Function

End Function

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie, iIndice As Integer

    Set objSerie = obj1

    'Coloca a S�rie na Tela
    For iIndice = 0 To Serie.ListCount - 1
        
        If Trim(Serie.List(iIndice)) = Trim(objSerie.sSerie) Then
            
            Serie.ListIndex = iIndice
            Exit For
        
        End If
    
    Next
        
    Call Serie_Validate(bSGECancelDummy)

    Exit Sub

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim lNumNotaUltima As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Serie_Validate

    'Verifica se a Serie foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
        
    'Verifica se � uma Serie selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then Error 38179
       
    Exit Sub
    
Erro_Serie_Validate:

    Cancel = True

    Select Case Err
    
        Case 38179, 60464 'Tratado na Rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168485)
    
    End Select
    
    Exit Sub

End Sub


Private Sub LabelNumBordero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumBordero, Source, X, Y)
End Sub

Private Sub LabelNumBordero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumBordero, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub LabelEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelCarteira_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub LabelCarteira_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub LabelValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub


Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

