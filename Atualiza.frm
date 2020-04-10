VERSION 5.00
Begin VB.Form Atualiza 
   Caption         =   "Atualiza Programa"
   ClientHeight    =   4005
   ClientLeft      =   2010
   ClientTop       =   2415
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1590
      TabIndex        =   7
      Top             =   3405
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.FileListBox File2 
      Height          =   480
      Left            =   2265
      TabIndex        =   6
      Top             =   3405
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Historico 
      Height          =   3135
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton BotaoEncerra 
      Caption         =   "Encerra"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton BotaoAtualiza 
      Caption         =   "Atualiza"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Titulo 
      Alignment       =   2  'Center
      Caption         =   "ATUALIZAÇÃO SGE-CORPORATOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label Informacao 
      Alignment       =   2  'Center
      Caption         =   $"Atualiza.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6255
   End
End
Attribute VB_Name = "Atualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BotaoAtualiza_Click()

Dim sArqLote As String
Dim sDirDestino As String
Dim sDirOrigem As String
Dim lArqNum As Long
Dim sLinha As String
Dim sNomeArq As String
Dim iEhEXE As Integer
Dim dDataArqOrigem As Date
Dim dDataArqDestino As Date
Dim iCopiaArq As Integer
Dim iIndice As Integer
Dim lErro As Long
   
On Error GoTo Erro_BotaoAtualiza
      
    Titulo.Visible = False
    Informacao.Visible = False
    Historico.Visible = True
    BotaoCancela.Visible = True
    BotaoAtualiza.Enabled = False
    BotaoEncerra.Enabled = False
   
    sArqLote = LeArqINI("Geral", "RegFile", NOME_ARQUIVO_ADM)
    If Len(sArqLote) = 0 Then Error 5005
          
    sDirDestino = App.Path
    If Right(sDirDestino, 1) <> "\" Then
       sDirDestino = sDirDestino & "\"
    End If
    
    sDirOrigem = LeArqINI("Geral", "DirProgram", NOME_ARQUIVO_ADM)
    If Len(sDirOrigem) = 0 Then Error 5007
    
    If Right(sDirOrigem, 1) <> "\" Then
       sDirOrigem = sDirOrigem & "\"
    End If
       
    Historico.Text = "Abrindo arquivo de Atualização " & sArqLote
    Historico.Text = Historico.Text & vbCrLf & "Copiando arquivos de " & sDirOrigem
    
    lErro = Teste_Acesso_Full_Pasta(sDirDestino)
    If lErro <> SUCESSO Then Error 10002
    
    lArqNum = FreeFile
    
    Open sArqLote For Input As lArqNum
 
    Do While Not EOF(lArqNum)
    
Le_Outro:

        If iParaTudo Then
           Exit Do
        End If
   
        If EOF(lArqNum) Then Exit Do
   
        Line Input #1, sLinha
      
        iEhEXE = False
      
        If UCase(Left(sLinha, 12)) = "REGSVR32 /S " Then
          
            sNomeArq = Mid(sLinha, 13, (Len(sLinha) - 12))
       
        ElseIf UCase(Right(sLinha, 15)) = ".EXE /REGSERVER" Then
          
            sNomeArq = Mid(sLinha, 1, (Len(sLinha) - 11))
            iEhEXE = True
       
        Else
        
           GoTo Le_Outro
       
        End If
        
        lErro = Testa_Arquivo(sDirOrigem & sNomeArq)
        If lErro <> SUCESSO Then Error 9999
                    
        If Len(Dir(sDirOrigem & sNomeArq)) = 0 Then Error 5008
       
        iCopiaArq = False
        dDataArqOrigem = FileDateTime(sDirOrigem & sNomeArq)
        
        If Len(Dir(sDirDestino & sNomeArq)) = 0 Then
           
           iCopiaArq = True
        
        Else
        
            lErro = Testa_Arquivo(sDirDestino & sNomeArq)
            If lErro <> SUCESSO Then Error 9999
        
           dDataArqDestino = FileDateTime(sDirDestino & sNomeArq)
        
           If dDataArqOrigem > dDataArqDestino Then
              iCopiaArq = True
           End If
        
        End If
        
        If iCopiaArq Then
                
           Atualiza.Historico.Text = Historico.Text & vbCrLf & sLinha
           DoEvents
           
            lErro = Copia_Arquivo(sDirOrigem, sDirDestino, sNomeArq)
            If lErro <> SUCESSO Then Error 9999
           
           'FileCopy sDirOrigem & sNomeArq, sDirDestino & sNomeArq
           
           If iEhEXE Then
              Call Shell(sDirDestino & sNomeArq & " /REGSERVER", vbMinimizedFocus)
           Else
              Call Shell("REGSVR32.EXE /S " & sDirDestino & sNomeArq, vbMinimizedFocus)
           End If
           
           Atualiza.Historico.Text = Historico.Text & " ...OK!"
           Atualiza.Historico.SelStart = Len(Historico.Text)
           DoEvents
           
        End If
              
    Loop
    
    Close lArqNum
      
    If iParaTudo Then
    
        Beep
        MsgBox "Atualização INTERROMPIDA pelo usuário!", vbExclamation, "Não atualizou ..."
    
        iParaTudo = False
    
        BotaoAtualiza.Enabled = True
        BotaoEncerra.Enabled = True
        BotaoCancela.Visible = False
        Screen.MousePointer = vbDefault
        
        Unload Me
    
    Else
    
        For iIndice = 1 To colOutrosArquivos.Count
        
            If iParaTudo Then
               Exit For
            End If

            sNomeArq = colOutrosArquivos.Item(iIndice)
            
            If Len(Dir(sDirOrigem & sNomeArq)) = 0 Then Error 5008
            
            lErro = Testa_Arquivo(sDirOrigem & sNomeArq)
            If lErro <> SUCESSO Then Error 9999
                         
            iCopiaArq = False
            dDataArqOrigem = FileDateTime(sDirOrigem & sNomeArq)
             
            If Len(Dir(sDirDestino & sNomeArq)) = 0 Then
                
               iCopiaArq = True
             
            Else
            
                lErro = Testa_Arquivo(sDirDestino & sNomeArq)
                If lErro <> SUCESSO Then Error 9999
             
               dDataArqDestino = FileDateTime(sDirDestino & sNomeArq)
             
               If dDataArqOrigem > dDataArqDestino Then
                  iCopiaArq = True
               End If
             
            End If
             
            If iCopiaArq Then
             
               Atualiza.Historico.Text = Historico.Text & vbCrLf & sNomeArq
               DoEvents
                
                lErro = Copia_Arquivo(sDirOrigem, sDirDestino, sNomeArq)
                If lErro <> SUCESSO Then Error 9999
'
'               FileCopy sDirOrigem & sNomeArq, sDirDestino & sNomeArq
                
               If sNomeArq = DLL_A_REGISTRAR Then
                  Call Shell("REGSVR32.EXE /S " & sDirDestino & sNomeArq, vbMinimizedFocus)
               End If
                
               Atualiza.Historico.Text = Historico.Text & " ...OK!"
               Atualiza.Historico.SelStart = Len(Historico.Text)
               DoEvents
                
            End If

        Next iIndice
        
        Historico.Text = Historico.Text & vbCrLf & "OK, Concluído!!!"
        Historico.SelStart = Len(Historico.Text)

         '##############################################################
         'Inserido por Wagner 16/05/2006
         'Atualiza os tsks
         lErro = Atualiza_Tsks
         If lErro <> SUCESSO Then Error 6003
         '##############################################################
                
        If iParaTudo Then
        
            Beep
            MsgBox "Atualização INTERROMPIDA pelo usuário!", vbExclamation, "Não atualizou ..."
        
            iParaTudo = False
        
            BotaoAtualiza.Enabled = True
            BotaoEncerra.Enabled = True
            BotaoCancela.Visible = False
            Screen.MousePointer = vbDefault
                    
        Else
       
           Beep
           MsgBox "Atualização efetuada com Sucesso!", vbInformation, "OK!"
           
        End If
       
        Unload Me
    
    End If
   
    Exit Sub
   
Erro_BotaoAtualiza:

    Beep
    Select Case Err.Number
   
        Case Is = 5005
            MsgBox "Não há informação sobre o arquivo de lote no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
      
        Case Is = 5006
            MsgBox "Informação incorreta sobre o arquivo de lote no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
      
        Case Is = 5007
            MsgBox "Não há informação sobre os arquivos de atualização no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
      
        Case Is = 5008
         
            MsgBox "O arquivo " & Chr(34) & sDirOrigem & sNomeArq & Chr(34) & " não existe. ", vbCritical, "Atenção!"
         
'            If MsgBox("O arquivo " & Chr(34) & sDirOrigem & sNomeArq & Chr(34) & " não existe. Deseja continuar mesmo assim?", vbYesNo + vbQuestion + vbDefaultButton2, "Continua?") = vbYes Then
'                Resume Next
'            End If
          
        Case 6003, 9999
      
        Case 10002
            MsgBox "Erro: " & Err.Number & " - O diretório " & sDirDestino & " está com restrições ao acesso.", vbCritical, "Atenção!"
      
        Case Else
            MsgBox "Erro do VB! (" & Err.Number & " - " & Err.Description & ").", vbCritical, "Atenção!"
   
    End Select
   
    BotaoAtualiza.Enabled = True
    BotaoEncerra.Enabled = True
    BotaoCancela.Visible = False
   
    Exit Sub

End Sub

Private Sub BotaoCancela_Click()
   
   If MsgBox("Confirma o cancelamento da atualização dos programas do SGE-Corporator?", _
      vbQuestion + vbYesNo + vbDefaultButton2, "Atualização SGE-Corporator") = vbYes Then
      
      iParaTudo = True
   
      Screen.MousePointer = vbHourglass
   
   End If

End Sub


Private Sub BotaoEncerra_Click()

   Beep
   MsgBox "Atualização INTERROMPIDA pelo usuário!", vbExclamation, "Não atualizou ..."
   Unload Me

End Sub

Private Sub Form_Load()

   Titulo.Visible = True
   Informacao.Visible = True
   Historico.Visible = False
   BotaoCancela.Visible = False

End Sub

'##################################
'Inserido por Wagner 16/05/2006
Public Function Atualiza_Tsks() As Long

Dim lErro As Long
Dim sDirTskSRV As String
Dim sDirTskPRG As String
Dim objFileSRV As FileListBox
Dim objFilePRG As FileListBox
Dim iIndSRV As Integer
Dim iIndPRG As Integer
Dim bTemTsk As Boolean
Dim bTskDataAnt As Boolean

On Error GoTo Erro_Atualiza_Tsks

    'Obtém o diretório dos tsk do servidor
    sDirTskSRV = LeArqINI("Geral", "Dirtsk", NOME_ARQUIVO_ADM)
    If Len(sDirTskSRV) = 0 Then Error 6001
    
    'Obtém o diretório dos tsk das máquinas
    sDirTskPRG = LeArqINI("ForPrint", "Dirtsks", NOME_ARQUIVO_ADM)
    If Len(sDirTskPRG) = 0 Then Error 6002
    
    If UCase(sDirTskSRV) <> UCase(sDirTskPRG) Then

        Set objFileSRV = Atualiza.File1
        Set objFilePRG = Atualiza.File2
        
        objFileSRV.Path = sDirTskSRV
        objFilePRG.Path = sDirTskPRG
        
        Historico.Text = Historico.Text & vbCrLf & "Copiando Tsks de " & sDirTskSRV
        Historico.SelStart = Len(Historico.Text)
        
        'Para cada arquivo no servidor
        For iIndSRV = 0 To objFileSRV.ListCount - 1
        
            If iParaTudo Then
               Exit For
            End If
        
            bTemTsk = False
            bTskDataAnt = False
            
            'Para cada arquivo na máquina
            For iIndPRG = 0 To objFilePRG.ListCount - 1
            
                'Se o arquivo for o mesmo
                If UCase(objFileSRV.List(iIndSRV)) = UCase(objFilePRG.List(iIndPRG)) Then
                    bTemTsk = True
                    'Se as data do tsk da máquina é anterior a do servidor
                    If FileDateTime(objFileSRV.Path & "\" & objFileSRV.List(iIndSRV)) > FileDateTime(objFilePRG.Path & "\" & objFilePRG.List(iIndPRG)) Then
                        bTskDataAnt = True
                    End If
                    Exit For
                End If
            
            Next
            
            'Se não existe o arquivo na máquina ou o arquivo do servidor é mais recente
            If bTskDataAnt Or Not bTemTsk Then
                
                Historico.Text = Historico.Text & vbCrLf & objFileSRV.Path & "\" & objFileSRV.List(iIndSRV)
                Historico.SelStart = Len(Historico.Text)
                
                DoEvents
                    
                'FileCopy objFileSRV.Path & "\" & objFileSRV.List(iIndSRV), objFilePRG.Path & "\" & objFileSRV.List(iIndSRV)
                    
                lErro = Copia_Arquivo(objFileSRV.Path & "\", objFilePRG.Path & "\", objFileSRV.List(iIndSRV))
                If lErro <> SUCESSO Then Error 9999
                    
                Historico.Text = Historico.Text & " ...OK!"
                Historico.SelStart = Len(Historico.Text)
                
                DoEvents
    
            End If
            
        Next
        
        Historico.Text = Historico.Text & vbCrLf & "Concluído."
        Historico.SelStart = Len(Historico.Text)
    
    End If
    
    Atualiza_Tsks = SUCESSO
   
    Exit Function

Erro_Atualiza_Tsks:

    Atualiza_Tsks = Err
   
    Select Case Err
    
        Case 6001, 6002
            MsgBox "Não há informação sobre o arquivo de lote no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
   
        Case Else
            MsgBox "Erro do VB! (" & Err.Number & " - " & Err.Description & ").", vbCritical, "Atenção!"

    End Select
   
   Exit Function
   
End Function
'#################################
