VERSION 5.00
Begin VB.Form FormGeraCodigoLight 
   Caption         =   "Gera Codigo para Telas da Versao Light"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Diretorio 
      Height          =   375
      Left            =   2700
      TabIndex        =   8
      Top             =   2040
      Width           =   3465
   End
   Begin VB.TextBox ArqClasse 
      Height          =   375
      Left            =   2700
      TabIndex        =   6
      Top             =   1470
      Width           =   3465
   End
   Begin VB.TextBox ArqTelaModif 
      Height          =   375
      Left            =   2700
      TabIndex        =   5
      Top             =   795
      Width           =   3465
   End
   Begin VB.TextBox ArqTelaOrig 
      Height          =   375
      Left            =   2700
      TabIndex        =   4
      Top             =   195
      Width           =   3465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1965
      TabIndex        =   3
      Top             =   2865
      Width           =   2205
   End
   Begin VB.Label Label2 
      Caption         =   "Diretorio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1590
      TabIndex        =   7
      Top             =   2115
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo da Classe p/Tela:"
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
      Left            =   210
      TabIndex        =   2
      Top             =   1620
      Width           =   2235
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo da Tela Modificada:"
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
      Left            =   180
      TabIndex        =   1
      Top             =   915
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo da Tela Original:"
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
      TabIndex        =   0
      Top             =   315
      Width           =   2130
   End
End
Attribute VB_Name = "FormGeraCodigoLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gobjTela As Object
Private gsNomeClasse As String
Private mvarColCtls As Object

'flags de estado:
Private bJaAchouOptionExplicit As Boolean    'se já processou a linha com o "Option Explicit"
Private bJaProcessouSubOuFunc As Boolean    'se já processou 1a linha de sub ou function
Private bEmSubDeEvento As Boolean '    se está dentro de sub de evento
Private bEmTrechoNaoCopiar As Boolean

Private bInicSub As Boolean '    se é a 1a. linha de uma sub
Private bInicFunc As Boolean '    se é a 1a. linha de uma function
Private bInicSubEvento As Boolean '    se é a 1a. linha de uma sub de evento
Private bInicTrechoNaoCopiar As Boolean
    
Private bPularLinhaEmBranco As Boolean

Private bEdicaoTela As Boolean 'Inserido por Wagner

Public Sub TrataModeloCust()

Dim szLineIn As String

    bJaAchouOptionExplicit = False
    bJaProcessouSubOuFunc = False
    bEmTrechoNaoCopiar = False
    bPularLinhaEmBranco = False
    
    bEdicaoTela = False 'Inserido por Wagner
    
    'abrir arquivo original
    Open Diretorio.Text & "\" & ArqTelaOrig.Text For Input As #1
    
    'criar arquivos "novos"
    Open Diretorio.Text & "\" & ArqTelaModif.Text For Output As #2
    Open Diretorio.Text & "\" & ArqClasse.Text For Output As #3
    
    Call GeraCodigoInicialClasse

    Do While Not EOF(1)
        
        Line Input #1, szLineIn

        If bPularLinhaEmBranco = False Or szLineIn <> "" Then
        
            If bPularLinhaEmBranco = True Then bPularLinhaEmBranco = False
            
            'copiar a parte toda até o "option explicit" do arq original p/o modificado
            If bJaAchouOptionExplicit = False Then
                        
                Print #2, szLineIn
                    
                If szLineIn = "Option Explicit" Then
                    
                    Call Processa_Linha_Option_Explicit(szLineIn)
                    
                End If
                
            Else
            
                'analisa linha
                Call Analisa_linha(szLineIn)
                        
                'se está dentro de sub de evento
                If bEmSubDeEvento Then
                
                    '##################################
                    'Alterado por Wagner
                    'copiar linha p/nova classe
                    If bEdicaoTela Then
                        Print #2, szLineIn
                    Else
                        Print #3, szLineIn
                    End If
                    '##################################
                        
                    If szLineIn = "End Sub" Then
                        bEdicaoTela = False 'Inserido por Wagner
                        'desmarca flag de estado
                        bEmSubDeEvento = False
                        
                    End If
                    
                Else
        
                    'se ainda nao achei 1a sub ou function
                    If bJaProcessouSubOuFunc = False Then
                    
                        'se achei inicio de sub ou function
                        If bInicSub Or bInicFunc Then
                        
                            'gera property gets dos controles na classe que está sendo criada
                            Call GeraPGsDosControles
                            
                            'marca flag de 1a sub ou function processada
                            bJaProcessouSubOuFunc = True
                    
                        End If
                        
                    End If
        
                    'se for inicio de sub de evento
                    If bInicSubEvento Then
                    
                        Call Processa_Inicio_Sub_Evento(szLineIn)
                        
                    Else
                        
                        If InStr(1, szLineIn, "Function Trata_Parametros(") <> 0 Then
                        
                            Call GeraTrataParamCtlModif(szLineIn)
                        
                        End If
                        
                        If bEmTrechoNaoCopiar = False Then
                        
                            If bInicTrechoNaoCopiar Then
                            
                                bEmTrechoNaoCopiar = True
                                
                            Else
                            
                                If Left(szLineIn, Len("Attribute ")) <> "Attribute " Then
                                
                                    If Left(szLineIn, Len("Private Sub UserControl_")) <> "Private Sub UserControl_" Then
                                    
                                        'copiar linha p/classe nova
                                        Print #3, szLineIn
                                                                            
                                    Else
                                    
                                        'copiar linha p/classe nova
                                        Print #3, Replace(szLineIn, "Private", "Public")
                                    
                                    End If
                                
                                End If
                                
                            End If
                            
                        Else
                        
                            'se chegou ao fim do codigo que está sendo pulado
                            If Left$(szLineIn, Len("End ")) = "End " Then
                                bEmTrechoNaoCopiar = False
                                bPularLinhaEmBranco = True
                            End If
                            
                        End If
                        
                    End If
        
                End If
            
            End If
        
        Else
        
            If bPularLinhaEmBranco = True Then bPularLinhaEmBranco = False
            
        End If
        
    Loop

    Call GeraCodigoFixoCtlModif
    Call GeraCodigoFixoClasse
    
    Close #1
    Close #2
    Close #3
    
End Sub

Private Sub Command1_Click()
    Call TrataModeloCust
    Unload Me
End Sub

Private Sub Form_Load()
Dim obj As Object, obj1 As Object, sNomeAux As String
Dim iPos As Integer

    Set obj = GL_objMDIForm.ActiveForm
    If obj Is Nothing Then MsgBox ("nao encontrei form ativo")
    
    sNomeAux = Mid$(obj.sNomeTelaOcx, InStr(1, obj.sNomeTelaOcx, ".") + 1)
    gsNomeClasse = "CT" & sNomeAux
    If Right(gsNomeClasse, 3) = "Ocx" Then gsNomeClasse = Left(gsNomeClasse, Len(gsNomeClasse) - 3)
    
    ArqTelaOrig.Text = sNomeAux & ".ctl"
    ArqTelaModif.Text = sNomeAux & ".ctm"
    ArqClasse.Text = gsNomeClasse & ".cls"
    Diretorio.Text = "c:\contab"
    
    Set gobjTela = obj.objFormOcx
    
End Sub

Private Sub GeraPGsDosControles()
'gera property gets dos controles na classe que está sendo criada

Dim obj2 As Object, sTipo As String, iIndice As Integer
Dim vaObjs As Variant 'array de objetos
Dim szLineOut As String, bLabelInutil As Boolean

    Print #3, "'--- inicio dos properties get dos controles da tela"
    Print #3, ""
    
    vaObjs = mvarColCtls.Items
    
    For iIndice = 0 To mvarColCtls.Count - 1
 
        Set obj2 = vaObjs(iIndice)
        
        sTipo = TypeName(obj2)
        
        bLabelInutil = False
        If Left(obj2.Name, 5) = "Label" Then
            If IsNumeric(Mid$(obj2.Name, 6)) Then bLabelInutil = True
        End If
        
        If sTipo <> "SSFrame" And sTipo <> "PictureBox" And bLabelInutil = False Then
        
            Print #3, "Public Property Get " & obj2.Name & "() As Object"
            Print #3, "     Set " & obj2.Name & " = objUserControl.Controls(" & Chr$(34) & obj2.Name & Chr$(34) & ")"
            Print #3, "End Property"
            Print #3, ""
        
        End If
        
    Next
    
    Print #3, "'--- fim dos properties get dos controles da tela"
    Print #3, ""

End Sub

Private Sub Analisa_linha(sLinha As String)

Dim iLixo As Integer

    bInicSub = False
    bInicFunc = False
    bInicSubEvento = False
    bInicTrechoNaoCopiar = False
    
    'se é inicio ou fim de sub
    If InStr(1, sLinha, " Sub ") <> 0 Then
    
        If InStr(1, sLinha, " End Sub ") <> 0 Then
        
            iLixo = 1
            
        Else
        
            If InStr(1, sLinha, " Exit Sub ") = 0 Then
                
                bInicSub = True
                
                If Testa_Sub_Evento(sLinha) Then bInicSubEvento = True
                
            End If
            
        End If
        
    End If
    
    'se é inicio ou fim de function
    If InStr(1, sLinha, " Function ") <> 0 Then
    
        If InStr(1, sLinha, " End Function ") <> 0 Then
        
            iLixo = 1
        
        Else
        
            If InStr(1, sLinha, " Exit Function ") = 0 Then bInicFunc = True
            
        End If
            
    End If

    If InStr(1, sLinha, "Property Get Controls()") <> 0 Or _
        InStr(1, sLinha, "Property Get hWnd()") <> 0 Or _
        InStr(1, sLinha, "Property Get Height()") <> 0 Or _
        InStr(1, sLinha, "Property Get Width()") <> 0 Or _
        InStr(1, sLinha, "Property Get ActiveControl()") <> 0 Or _
        InStr(1, sLinha, "Property Get Enabled()") <> 0 Or _
        InStr(1, sLinha, "Property Let Enabled(") <> 0 Or _
        InStr(1, sLinha, "'WARNING!") Or _
        InStr(1, sLinha, "'Load property values from storage") Or _
        InStr(1, sLinha, "'Write property values to storage") Then
        
        bInicTrechoNaoCopiar = True
        
    End If

End Sub

Private Sub Carregar_Ctls_Tela()
Dim obj2 As Object

    Set mvarColCtls = CreateObject("Scripting.Dictionary")
    
    For Each obj2 In gobjTela.Controls

        If mvarColCtls.Exists(obj2.Name) = False Then Call mvarColCtls.Add(obj2.Name, obj2)
        
    Next

End Sub

Sub Processa_Linha_Option_Explicit(szLineIn As String)

    Print #3, szLineIn
    
    'inserir no ctl modificado a declaracao da variavel da classe com codigo
    Print #2, ""
    Print #2, "Event Unload()"
    Print #2, ""
    Print #2, "Private WithEvents objCT as " & gsNomeClasse
    Print #2, ""
    Print #2, "Private Sub UserControl_Initialize()"
    Print #2, "    Set objCT = New " & gsNomeClasse
    Print #2, "    Set objCT.objUserControl = Me"
    Print #2, "End Sub"
    Print #2, ""
    
    'inserir na classe a declaracao da variavel da tela
    Print #3, ""
    Print #3, "Dim m_objUserControl as Object"
    
    'obter os controles da tela
    Call Carregar_Ctls_Tela

    bJaAchouOptionExplicit = True
            
End Sub

Sub Parse_Sub_Evento(szEntrada As String, szSaida As String)
'coloca em szSaida szEntrada retirando os tipos dos parametros e o "Private Sub "

Dim szTemp As String, iPosIni As Integer, iPosFim As Integer
Dim sTipoParam As String

    szTemp = szEntrada
    szTemp = Replace(szTemp, "Private Sub ", "") 'p/eventos "normais" como click de botoes
    szTemp = Replace(szTemp, "Public Sub ", "") 'p/form_*
    
    'retira os tipos dos parametros
    iPosIni = InStr(1, szTemp, " As ")
    Do While iPosIni <> 0
    
        iPosFim = InStr(iPosIni, szTemp, ",")
        If iPosFim = 0 Then iPosFim = InStr(iPosIni, szTemp, ")")
            
        sTipoParam = Mid$(szTemp, iPosIni, iPosFim - iPosIni)
        
        szTemp = Replace(szTemp, sTipoParam, "")
        
        iPosIni = InStr(1, szTemp, " As ")
        
    Loop
    
    'retira os byVal dos parametros
    iPosIni = InStr(1, szTemp, "ByVal ")
    Do While iPosIni <> 0
        
        szTemp = Replace(szTemp, "ByVal ", "")
    
        iPosIni = InStr(1, szTemp, "ByVal ")
    Loop
    
    szSaida = szTemp
    
End Sub

Sub Processa_Inicio_Sub_Evento(szLineIn As String)
Dim szAux As String

    'marca flag de estado
    bEmSubDeEvento = True
    
    '##############################################
    'Alterado por wagner
    'copia linha p/classe que está sendo criada trocando Private por Public
    If InStr(1, szLineIn, "DragDrop") = 0 And InStr(1, szLineIn, "MouseDown") = 0 Then
    
        'copia linha p/ctl modif
        Print #2, szLineIn
        
        'obter parametros p/chamada da classe retirando os  " As..."
        Call Parse_Sub_Evento(szLineIn, szAux)
        
        'inclui no ctl modif chamada p/classe
        Print #2, "     Call objCT." & szAux
        
        'incluir end sub e pular linha
        Print #2, "End Sub"
        Print #2, ""
    
        bEdicaoTela = False
        Print #3, Replace(szLineIn, "Private", "Public")
    Else
        bEdicaoTela = True
        Print #2, szLineIn
    End If
    '##############################################
    
End Sub

Private Function Testa_Sub_Evento(sLinha As String) As Boolean
'retorna True se a linha corresponde a uma sub de tratamento de evento de controle

Dim sAux As String, iPosFim As Integer

    Testa_Sub_Evento = False
    
    If Left(sLinha, 12) = "Private Sub " Then
        
        iPosFim = InStr(13, sLinha, "_")
        If iPosFim > 1 Then
        
            sAux = Mid$(sLinha, 13, iPosFim - 13)
    
            'neste ponto tenho em sAux o nome da sub até o underscore, que deve ser o nome de um controle
            If mvarColCtls.Exists(sAux) Then Testa_Sub_Evento = True
            
        End If
    
    Else
    
        If InStr(1, sLinha, "Public Sub Form_") <> 0 And InStr(1, sLinha, "Public Sub Form_Unload") = 0 Then Testa_Sub_Evento = True
        
    End If
    
End Function

Sub GeraCodigoFixoCtlModif()

    Print #2, "Public Function Form_Load_Ocx() As Object"
    Print #2, ""
    Print #2, "    Call objCT.Form_Load_Ocx"
    Print #2, "    Set Form_Load_Ocx = Me"
    Print #2, ""
    Print #2, "End Function"
    Print #2, ""
    Print #2, "Public Sub Form_Unload(Cancel As Integer)"
    Print #2, "    If Not (objCT Is Nothing) Then"
    Print #2, "        Call objCT.Form_Unload(Cancel)"
    Print #2, "        If Cancel = False Then"
    Print #2, "             Set objCT.objUserControl = Nothing"
    Print #2, "             Set objCT = Nothing"
    Print #2, "        End If"
    Print #2, "    End If"
    Print #2, "End Sub"
    Print #2, ""
    Print #2, "Private Sub objCT_Unload()"
    Print #2, "   RaiseEvent Unload"
    Print #2, "End Sub"
    Print #2, ""
    Print #2, "Public Function Name() As String"
    Print #2, "    Call objCT.Name"
    Print #2, "End Function"
    Print #2, ""
    Print #2, "Public Sub Show()"
    Print #2, "    Call objCT.Show"
    Print #2, "End Sub"
    Print #2, ""
    Print #2, "'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!"
    Print #2, "'MappingInfo=UserControl,UserControl,-1,Controls"
    Print #2, "Public Property Get Controls() As Object"
    Print #2, "    Set Controls = UserControl.Controls"
    Print #2, "End Property"
    Print #2, ""
    Print #2, "Public Property Get hWnd() As Long"
    Print #2, "    hWnd = UserControl.hWnd"
    Print #2, "End Property"
    Print #2, ""
    Print #2, "Public Property Get Height() As Long"
    Print #2, "    Height = UserControl.Height"
    Print #2, "End Property"
    Print #2, ""
    Print #2, "Public Property Get Width() As Long"
    Print #2, "    Width = UserControl.Width"
    Print #2, "End Property"
    Print #2, ""
    Print #2, "'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!"
    Print #2, "'MappingInfo=UserControl,UserControl,-1,ActiveControl"
    Print #2, "Public Property Get ActiveControl() As Object"
    Print #2, "    Set ActiveControl = UserControl.ActiveControl"
    Print #2, "End Property"
    Print #2, ""
    Print #2, "'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!"
    Print #2, "'MappingInfo=UserControl,UserControl,-1,Enabled"
    Print #2, "Public Property Get Enabled() As Boolean"
    Print #2, "    Enabled = UserControl.Enabled"
    Print #2, "End Property"
    Print #2, ""
    Print #2, "Public Property Let Enabled(ByVal New_Enabled As Boolean)"
    Print #2, "    UserControl.Enabled() = New_Enabled"
    Print #2, "    PropertyChanged " & Chr$(34) & "Enabled" & Chr$(34)
    Print #2, "End Property"
    Print #2, ""
    Print #2, "Public Property Get Parent() As Object"
    Print #2, "    Set Parent = UserControl.Parent"
    Print #2, "End Property"
    Print #2, ""
    Print #2, "'Load property values from storage"
    Print #2, "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)"
    Print #2, "    UserControl.Enabled = PropBag.ReadProperty(" & Chr$(34) & "Enabled" & Chr$(34) & ", True)"
    Print #2, "End Sub"
    Print #2, ""
    Print #2, "'Write property values to storage"
    Print #2, "Private Sub UserControl_WriteProperties(PropBag As PropertyBag)"
    Print #2, "    Call PropBag.WriteProperty(" & Chr$(34) & "Enabled" & Chr$(34) & ", UserControl.Enabled, True)"
    Print #2, "End Sub"
    Print #2, ""
    Print #2, "Public Property Get Caption() As String"
    Print #2, "    Caption = objCT.Caption"
    Print #2, "End Property"
    Print #2, ""
    Print #2, "Public Property Let Caption(ByVal New_Caption As String)"
    Print #2, "    objCT.Caption = New_Caption"
    Print #2, "End Property"
    Print #2, ""
    Print #2, "Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)"
    Print #2, "    Call objCT.UserControl_KeyDown(KeyCode, Shift)"
    Print #2, "End Sub"
    Print #2, ""

End Sub

Sub Parse_Trata_Param(szEntrada As String, szSaida As String)
'coloca em szSaida szEntrada retirando os tipos dos parametros e o tipo do retorno

Dim szTemp As String, iPosIni As Integer, iPosFim As Integer
Dim sTipoParam As String

    szTemp = Left(szEntrada, Len(szEntrada) - Len(" As Long"))
    iPosIni = InStr(1, szEntrada, "Trata_")
    szTemp = Right(szTemp, Len(szTemp) - iPosIni + 1)
    
    'retira os tipos dos parametros
    iPosIni = InStr(1, szTemp, " As ")
    Do While iPosIni <> 0
    
        iPosFim = InStr(iPosIni, szTemp, ",")
        If iPosFim = 0 Then iPosFim = InStr(iPosIni, szTemp, ")")
            
        sTipoParam = Mid$(szTemp, iPosIni, iPosFim - iPosIni)
        
        szTemp = Replace(szTemp, sTipoParam, "")
        
        iPosIni = InStr(1, szTemp, " As ")
        
    Loop
    
    'retira os byVal dos parametros
    iPosIni = InStr(1, szTemp, "ByVal ")
    Do While iPosIni <> 0
        
        szTemp = Replace(szTemp, "ByVal ", "")
    
        iPosIni = InStr(1, szTemp, "ByVal ")
    Loop
    
    'retira os "Optional" dos parametros
    iPosIni = InStr(1, szTemp, "Optional ")
    Do While iPosIni <> 0
        
        szTemp = Replace(szTemp, "Optional ", "")
    
        iPosIni = InStr(1, szTemp, "Optional ")
        
    Loop
    
    szSaida = szTemp
    
End Sub

Sub GeraTrataParamCtlModif(szLineIn As String)
Dim szAux As String

    'copia linha p/ctl modif
    Print #2, szLineIn
    
    'obter parametros p/chamada da classe retirando os  " As..."
    Call Parse_Trata_Param(szLineIn, szAux)
        
    'inclui no ctl modif chamada p/classe
    Print #2, "     Trata_Parametros = objCT." & szAux
    
    'incluir end function e pular linha
    Print #2, "End Function"
    Print #2, ""
    
End Sub

Sub GeraCodigoFixoClasse()

    Print #3, "Public Property Get objUserControl() As Object"
    Print #3, "    Set objUserControl = m_objUserControl"
    Print #3, "End Property"
    Print #3, ""
    Print #3, "Public Property Set objUserControl(ByVal vData As Object)"
    Print #3, "    Set m_objUserControl = vData"
    Print #3, "End Property"
    Print #3, ""
    Print #3, "'Devolve Parent do User Control"
    Print #3, "Public Property Get Parent() As Object"
    Print #3, "    Set Parent = objUserControl.Parent"
    Print #3, "End Property"
    Print #3, ""
    Print #3, "Public Property Get Controls() As Object"
    Print #3, "    Set Controls = objUserControl.Controls"
    Print #3, "End Property"
    Print #3, ""
    Print #3, "Public Property Get ActiveControl() As Object"
    Print #3, "    Set ActiveControl = objUserControl.ActiveControl"
    Print #3, "End Property"
    Print #3, ""
    Print #3, "Public Property Get Enabled() As Boolean"
    Print #3, "    Enabled = objUserControl.Enabled"
    Print #3, "End Property"
    Print #3, ""
    Print #3, "Public Property Let Enabled(ByVal New_Enabled As Boolean)"
    Print #3, "    objUserControl.Enabled = New_Enabled"
    Print #3, "End Property"
    Print #3, ""
    
End Sub

Sub GeraCodigoInicialClasse()

    Print #3, "VERSION 1.0 CLASS"
    Print #3, "BEGIN"
    Print #3, "  MultiUse = -1  'True"
    Print #3, "  Persistable = 0  'NotPersistable"
    Print #3, "  DataBindingBehavior = 0  'vbNone"
    Print #3, "  DataSourceBehavior = 0   'vbNone"
    Print #3, "  MTSTransactionMode = 0   'NotAnMTSObject"
    Print #3, "End"
    Print #3, "Attribute VB_Name = " & Chr$(34) & gsNomeClasse & Chr$(34)
    Print #3, "Attribute VB_GlobalNameSpace = False"
    Print #3, "Attribute VB_Creatable = True"
    Print #3, "Attribute VB_PredeclaredId = False"
    Print #3, "Attribute VB_Exposed = True"
    Print #3, "Attribute VB_Ext_KEY = " & Chr$(34) & "SavedWithClassBuilder6" & Chr$(34) & " ," & Chr$(34) & "Yes" & Chr$(34)
    Print #3, "Attribute VB_Ext_KEY = " & Chr$(34) & "Top_Level" & Chr$(34) & " ," & Chr$(34) & "Yes" & Chr$(34)

End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

