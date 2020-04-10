VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEstTipo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const NUM_MAX_EMBALAGENS = 9999
Private Const NUM_MAX_TESTESQUALIDADE = 9999

Private Function Tipo_Embalagem(ByVal sCampo As String) As enumTipo

If Len(Trim(sCampo)) = 0 Then
    Tipo_Embalagem = TIPO_VAZIO
ElseIf Not IsNumeric(sCampo) Then
    Tipo_Embalagem = TIPO_STRING
ElseIf Int(CDbl(sCampo)) <> CDbl(sCampo) Then
    Tipo_Embalagem = TIPO_DECIMAL
ElseIf CDbl(sCampo) <= 0 Then
    Tipo_Embalagem = TIPO_NAO_POSITIVO
ElseIf CDbl(sCampo) > NUM_MAX_EMBALAGENS Then
    Tipo_Embalagem = TIPO_OVERFLOW
Else
    Tipo_Embalagem = TIPO_CODIGO
End If

End Function

'Rotina criada por Daniel
'Revisão: Luiz G.F.Nogueira
'Data da inclusão no ClassEstTipo: 17/09/01
'Pendências: não
Function TP_Embalagem_Le_Grid(objEmbalagemTextBox As Object, objEmbalagem As ClassEmbalagem) As Long
'Recebe um controle TextBox e a partir do conteúdo desse controle chama uma função para leitura da embalagem
'O objEmbalagem só será retornado preenchido caso o conteúdo do controle seja a sigla oiu o código da embalagem

Dim sEmbalagem As String
Dim Embalagem As Object
Dim eTipoEmbalagem As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TP_Embalagem_Le_Grid

    Set Embalagem = objEmbalagemTextBox
    sEmbalagem = Trim(Embalagem.Text)
    eTipoEmbalagem = Tipo_Embalagem(sEmbalagem)

    Select Case eTipoEmbalagem

        Case TIPO_STRING

            'Guarda a sigla da embalagem no obj
            objEmbalagem.sSigla = sEmbalagem
            'Le a Embalagem a partir da sigla passada no obj
            lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
            If lErro <> SUCESSO And lErro <> 95088 Then gError 95083
            
            'Se nao encontrou => erro
            If lErro = 95088 Then gError 95104

        Case TIPO_CODIGO
                
            'Guarda o código da embalagem no obj
            objEmbalagem.iCodigo = sEmbalagem
            'Le a Embalagem a partir do código passado no obj
            lErro = CF("Embalagem_Le", objEmbalagem)
            If lErro <> SUCESSO And lErro <> 82763 Then gError 95084
            
            'Se nao encontrou => erro
            If lErro = 82763 Then gError 95105

        Case TIPO_OVERFLOW

            gError 95085

        Case TIPO_DECIMAL

            gError 95086

        Case TIPO_NAO_POSITIVO

            gError 95087

    End Select

    TP_Embalagem_Le_Grid = SUCESSO

    Exit Function

Erro_TP_Embalagem_Le_Grid:

    TP_Embalagem_Le_Grid = gErr

    Select Case gErr

        Case 95083, 95084
            
        Case 95085
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", gErr, sEmbalagem)

        Case 95086
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sEmbalagem)

        Case 95087
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sEmbalagem)
        
        Case 95104, 95105
            'Não encontrou Produto no BD
            
            If gErr = 95105 Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_EMBALAGEM", objEmbalagem.iCodigo)
            Else
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_EMBALAGEM1", objEmbalagem.sSigla)
            End If

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Embalagem", objEmbalagem)

            Else
                Embalagem.Text = ""
                Embalagem.SetFocus
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149710)

    End Select

End Function

Private Function Tipo_TesteQualidade(ByVal sCampo As String) As enumTipo

If Len(Trim(sCampo)) = 0 Then
    Tipo_TesteQualidade = TIPO_VAZIO
ElseIf Not IsNumeric(sCampo) Then
    Tipo_TesteQualidade = TIPO_STRING
ElseIf Int(CDbl(sCampo)) <> CDbl(sCampo) Then
    Tipo_TesteQualidade = TIPO_DECIMAL
ElseIf CDbl(sCampo) <= 0 Then
    Tipo_TesteQualidade = TIPO_NAO_POSITIVO
ElseIf CDbl(sCampo) > NUM_MAX_TESTESQUALIDADE Then
    Tipo_TesteQualidade = TIPO_OVERFLOW
Else
    Tipo_TesteQualidade = TIPO_CODIGO
End If

End Function

Function TP_TesteQualidade_Le_Grid(objTesteQualidadeTextBox As Object, objTesteQualidade As ClassTestesQualidade) As Long
'Recebe um controle TextBox e a partir do conteúdo desse controle chama uma função para leitura da TesteQualidade
'O objTesteQualidade só será retornado preenchido caso o conteúdo do controle seja o nome reduzido ou o código do Teste

Dim sTesteQualidade As String
Dim TesteQualidade As Object
Dim eTipoTesteQualidade As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TP_TesteQualidade_Le_Grid

    Set TesteQualidade = objTesteQualidadeTextBox
    sTesteQualidade = Trim(TesteQualidade.Text)
    eTipoTesteQualidade = Tipo_TesteQualidade(sTesteQualidade)

    Select Case eTipoTesteQualidade

        Case TIPO_STRING

            'Guarda a sigla da TesteQualidade no obj
            objTesteQualidade.sNomeReduzido = sTesteQualidade
            'Le a TesteQualidade a partir do nome reduzido no obj
            lErro = CF("TesteQualidade_Le_NomeReduzido", objTesteQualidade)
            If lErro <> SUCESSO And lErro <> 130109 Then gError 95083
            
            'Se nao encontrou => erro
            If lErro <> SUCESSO Then gError 95104

        Case TIPO_CODIGO
                
            'Guarda o código da TesteQualidade no obj
            objTesteQualidade.iCodigo = sTesteQualidade
            'Le a TesteQualidade a partir do código passado no obj
            lErro = CF("TestesQualidade_Le", objTesteQualidade)
            If lErro <> SUCESSO And lErro <> 130109 Then gError 95084
            
            'Se nao encontrou => erro
            If lErro <> SUCESSO Then gError 95105

        Case TIPO_OVERFLOW

            gError 95085

        Case TIPO_DECIMAL

            gError 95086

        Case TIPO_NAO_POSITIVO

            gError 95087

    End Select

    TP_TesteQualidade_Le_Grid = SUCESSO

    Exit Function

Erro_TP_TesteQualidade_Le_Grid:

    TP_TesteQualidade_Le_Grid = gErr

    Select Case gErr

        Case 95083, 95084
            
        Case 95085
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", gErr, sTesteQualidade)

        Case 95086
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sTesteQualidade)

        Case 95087
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sTesteQualidade)
        
        Case 95104, 95105
            'Não encontrou Produto no BD
            
            If gErr = 95105 Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TESTEQUALIDADE", objTesteQualidade.iCodigo)
            Else
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TESTEQUALIDADE1", objTesteQualidade.sNomeReduzido)
            End If

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("TestesQualidade", objTesteQualidade)

            Else
                TesteQualidade.Text = ""
                TesteQualidade.SetFocus
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149711)

    End Select

End Function


