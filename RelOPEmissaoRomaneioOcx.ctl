VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpEmissaoRomaneioOcx 
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   ScaleHeight     =   2520
   ScaleWidth      =   5010
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
      Height          =   555
      Left            =   840
      Picture         =   "RelOPEmissaoRomaneioOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1845
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3600
      ScaleHeight     =   495
      ScaleMode       =   0  'User
      ScaleWidth      =   1065
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1125
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOPEmissaoRomaneioOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOPEmissaoRomaneioOcx.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameNF 
      Caption         =   "Nota Fiscal"
      Height          =   1500
      Left            =   105
      TabIndex        =   0
      Top             =   930
      Width           =   4695
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin MSMask.MaskEdBox NFiscalInicial 
         Height          =   300
         Left            =   1290
         TabIndex        =   2
         Top             =   945
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFiscalFinal 
         Height          =   300
         Left            =   2745
         TabIndex        =   3
         Top             =   945
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2370
         TabIndex        =   6
         Top             =   1005
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
         Left            =   915
         TabIndex        =   5
         Top             =   1005
         Width           =   300
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   735
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   4
         Top             =   420
         Width           =   510
      End
   End
End
Attribute VB_Name = "RelOpEmissaoRomaneioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoSerie = New AdmEvento
    
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 90097
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 90097 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168508)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
  
    Set objEventoSerie = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes, Optional vParam As Variant) As Long

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 90074
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Caption = gobjRelatorio.sCodRel
    
    If Not IsMissing(vParam) Then
        
        'Se passou a Série
        If vParam <> "" Then
                    
            objSerie.iFilialEmpresa = giFilialEmpresa
            objSerie.sSerie = CStr(vParam)
            
            'Lê a Serie no BD
            lErro = CF("Serie_Le",objSerie)
            If lErro <> SUCESSO And lErro <> 22202 Then gError 90075
            
            'Se não encontrou Erro
            If lErro = 22202 Then gError 90076
        
        End If
            
        For iIndice = 0 To Serie.ListCount - 1
            
            If Trim(Serie.List(iIndice)) = Trim(objSerie.sSerie) Then
                Serie.ListIndex = iIndice
                Exit For
            End If
        
        Next
        
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 90074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 90075
        
        Case 90076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168509)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Critica_Parametros() As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long

On Error GoTo Erro_Critica_Parametros
      
    'Verifica se a Série está preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 90077
    
    'Verifica se a Nota Inicial está preenchida
    If Len(Trim(NFiscalInicial.Text)) = 0 Then gError 90078
    
    'Verifica se a Nota Final está preenchida
    If Len(Trim(NFiscalFinal.Text)) = 0 Then gError 90079
      
    'Verifica se o numero da Nota Fiscal inicial é maior que o da final
    If Len(Trim(NFiscalInicial.ClipText)) > 0 And Len(Trim(NFiscalFinal.ClipText)) > 0 Then
    
        If CLng(NFiscalInicial.Text) > CLng(NFiscalFinal.Text) Then gError 90080
    
    End If
        
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = gErr

    Select Case gErr

        Case 90080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
        
        Case 90077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 90078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DE_NAO_PREENCHIDO", gErr)
        
        Case 90079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168510)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 90081
      
    Serie.ListIndex = -1
    Serie.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 90081 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168511)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Parametros()
    If lErro <> SUCESSO Then gError 90082
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 90083
    
    lErro = objRelOpcoes.IncluirParametro("NNFISCALINIC", NFiscalInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90084

    lErro = objRelOpcoes.IncluirParametro("NNFISCALFIM", NFiscalFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90085
   
    lErro = objRelOpcoes.IncluirParametro("TSERIE", Serie.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90086
   
            
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 90082 To 90086 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168512)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim vbMsgRes As VbMsgBoxResult
Dim lFaixaFinal As Long


On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 90087
    
    objSerie.sSerie = Serie.Text
    objSerie.iFilialEmpresa = giFilialEmpresa
          
    lErro = gobjRelatorio.Executar_Prossegue
    If lErro <> SUCESSO And lErro <> 7072 Then gError 90092
    
    'Cancelou o relatório
    If lErro = 7072 Then gError 90093
     
    objSerie.lProxNumRomaneio = CLng(NFiscalFinal.Text) + 1

    'Atualiza a Tabela de Série
    lErro = CF("Serie_Atualiza_ImpressaoRomaneio",objSerie)
    If lErro <> SUCESSO And lErro <> 90116 Then gError 90095

    'Não encontrou a Série
    If lErro = 90116 Then gError 90096
       
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 90087, 90088, 90090, 90091, 90092, 90095 'Tratado na Rotina chamada
        
        Case 90089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
        
        Case 90094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_IMPRESSAO_NAO_PREENCHIDO", gErr)
        
        Case 90096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
       
        
        Case 90093
            Unload Me
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168513)

    End Select

    Exit Sub

End Sub

Private Sub LabelSerie_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieListaModal
    Call Chama_Tela("SerieListaModal", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Private Sub NFiscalInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalInicial_Validate
    
    lErro = Critica_Numero(NFiscalInicial.Text)
    If lErro <> SUCESSO Then gError 90098
              
    Exit Sub

Erro_NFiscalInicial_Validate:

    Cancel = True


    Select Case gErr
    
        Case 90098
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168514)
            
    End Select
    
    Exit Sub

End Sub

Private Sub NFiscalFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalFinal_Validate
     
    lErro = Critica_Numero(NFiscalFinal.Text)
    If lErro <> SUCESSO Then gError 90099
        
    Exit Sub

Erro_NFiscalFinal_Validate:

    Cancel = True


    Select Case gErr
    
        Case 90099
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168515)
            
    End Select
    
    Exit Sub

End Sub

Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero
         
    If Len(Trim(sNumero)) > 0 Then
        
        lErro = Long_Critica(sNumero)
        If lErro <> SUCESSO Then gError 90100
 
    End If
 
    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = gErr

    Select Case gErr
                  
        Case 90100 'Tratado na Rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168516)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le",colSerie)
    If lErro <> SUCESSO Then gError 90101
    
    Serie.Clear
        
    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next
        
    Carrega_Serie = SUCESSO
    
    Exit Function
    
Erro_Carrega_Serie:

    Carrega_Serie = gErr
    
    Select Case gErr
    
        Case 90101 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168517)
            
    End Select
    
    Exit Function

End Function

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie, iIndice As Integer

    Set objSerie = obj1

    'Coloca a Série na Tela
    For iIndice = 0 To Serie.ListCount - 1
        
        If Trim(Serie.List(iIndice)) = Trim(objSerie.sSerie) Then
            
            Serie.ListIndex = iIndice
            Exit For
        
        End If
    
    Next
        
    Call Serie_Validate(bSGECancelDummy)

    Exit Sub

End Sub

Private Sub Serie_Click()

Dim lErro As Long

On Error GoTo Erro_Serie_Click
    
    'Traz os números default
    lErro = Traz_Numeros_Default()
    If lErro <> SUCESSO Then gError 90102
        
    Exit Sub
    
Erro_Serie_Click:

    Select Case gErr
    
        Case 90102 'Tratado na Rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168518)
    
    End Select
    
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
        
    'Verifica se é uma Serie selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 90103
    
    'Se não encontrou na lista da ComboBox
    If lErro <> SUCESSO Then
        
        'Traz os números default
        lErro = Traz_Numeros_Default()
        If lErro <> SUCESSO Then gError 90104
    
    End If
    
    Exit Sub
    
Erro_Serie_Validate:

    Cancel = True


    Select Case gErr
    
        Case 90103, 90104 'Tratado na Rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168519)
    
    End Select
    
    Exit Sub

End Sub

Private Function Traz_Numeros_Default() As Long

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Traz_Numeros_Default

    If Serie.ListIndex = -1 Then Exit Function
    
    objSerie.sSerie = Serie.List(Serie.ListIndex)

    'Tenta ler a série no BD
    lErro = CF("Serie_Le",objSerie)
    If lErro <> SUCESSO And lErro <> 90110 Then gError 90105
    
    If lErro = 90110 Then gError 90106
        
    'Coloca número default de NFiscalInicial na tela
    If objSerie.lProxNumRomaneio > 0 Then
        NFiscalInicial.Text = objSerie.lProxNumRomaneio
        NFiscalFinal.Text = objSerie.lProxNumNFiscal - 1
    End If
    
    Traz_Numeros_Default = SUCESSO
    
    Exit Function
    
Erro_Traz_Numeros_Default:
    
    Traz_Numeros_Default = gErr
    
    Select Case gErr
    
        Case 90105  'Tratado na Rotina chamada
            Serie.SetFocus
       
        Case 90106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, Serie.Text)
            Serie.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168520)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_EMISSAO_NF
    Set Form_Load_Ocx = Me
    Caption = "Emissão de Romaneio de Nota Fiscal"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpEmissaoRomaneio"
    
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

Private Sub LabelSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerie, Source, X, Y)
End Sub

Private Sub LabelSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerie, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

'Função retirada de: RotinasFat2-ClassFatSelect
Function Serie_Le(objSerie As ClassSerie) As Long
'Lê a Serie a partir do código em objSerie.
'Devolve os dados em objSerie.

Dim lComando As Long
Dim lErro As Long
Dim tSerie As typeSerie

On Error GoTo Erro_Serie_Le

    tSerie.sSerie = String(STRING_SERIE, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90107

    'Verifica se a Serie existe, e se existir carrega seus dados em objSerie
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Serie, ProxNumNFiscal, ProxNumNFiscalEntrada, ProxNumNFiscalImpressa, Imprimindo, TipoFormulario, ProxNumRomaneio FROM Serie WHERE FilialEmpresa = ?  AND Serie = ?", tSerie.iFilialEmpresa, tSerie.sSerie, tSerie.lProxNumNFiscal, tSerie.lProxNumNFiscalEntrada, tSerie.lProxNumNFiscalImpressa, tSerie.iImprimindo, tSerie.iTipoFormulario, tSerie.lProxNumRomaneio, giFilialEmpresa, objSerie.sSerie)
    If lErro <> AD_SQL_SUCESSO Then gError 90108

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90109

    'Serie não está cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 90110

    'Carrega os dados lidos no objSerie
    objSerie.iFilialEmpresa = tSerie.iFilialEmpresa
    objSerie.sSerie = tSerie.sSerie
    objSerie.lProxNumNFiscal = tSerie.lProxNumNFiscal
    objSerie.lProxNumNFiscalEntrada = tSerie.lProxNumNFiscalEntrada
    objSerie.lProxNumNFiscalImpressa = tSerie.lProxNumNFiscalImpressa
    objSerie.iImprimindo = tSerie.iImprimindo
    objSerie.iTipoFormulario = tSerie.iTipoFormulario
    objSerie.lProxNumRomaneio = tSerie.lProxNumRomaneio
    
    Call Comando_Fechar(lComando)

    Serie_Le = SUCESSO

Exit Function

Erro_Serie_Le:

    Serie_Le = gErr

    Select Case gErr

        Case 90107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90108, 90109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE", gErr)

        Case 90110 'Serie não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168521)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Serie_Atualiza_ImpressaoRomaneio(objSerie As ClassSerie) As Long
'Atualiza o Número do último Romaneio de Nota impressa.

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim lProxNumRomaneio As Long

On Error GoTo Erro_Serie_Atualiza_ImpressaoRomaneio
    
    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 90111
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90112
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 90113
    
    'Lê a Série para poder atualizar o Numero de Romaneio
    lErro = Comando_ExecutarPos(lComando, "SELECT ProxNumRomaneio FROM Serie WHERE Serie = ? AND FilialEmpresa = ?", 0, lProxNumRomaneio, objSerie.sSerie, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 90114

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90115

    'Série não está cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 90116
   
    'Altera próximo Romaneio de Nota a imprimir
    lErro = Comando_ExecutarPos(lComando1, "UPDATE Serie SET ProxNumRomaneio = ?", lComando, objSerie.lProxNumRomaneio)
    If lErro <> AD_SQL_SUCESSO Then gError 90117

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 90118

    Serie_Atualiza_ImpressaoRomaneio = SUCESSO
    
    Exit Function
    
Erro_Serie_Atualiza_ImpressaoRomaneio:

    Serie_Atualiza_ImpressaoRomaneio = gErr
    
    Select Case gErr
        
        Case 90112, 90113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90114, 90115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE1", gErr, objSerie.sSerie)
        
        Case 90116 'Não encontrou , a ser Tratado na rotina chamadora
        
        Case 90117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SERIE", gErr, objSerie.sSerie)

        Case 90111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 90118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168522)


    End Select

    'Fechamento transação
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function


