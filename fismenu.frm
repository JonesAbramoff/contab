VERSION 5.00
Begin VB.Form fismenu 
   Caption         =   "Menu de Livros Fiscais"
   ClientHeight    =   3480
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   8
      Begin VB.Menu mnuFISMov 
         Caption         =   "Registro de Entrada"
         Index           =   1
      End
      Begin VB.Menu mnuFISMov 
         Caption         =   "Registro de Saida"
         Index           =   2
      End
      Begin VB.Menu mnuFISMov 
         Caption         =   "ICMS"
         Index           =   3
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Apuração de ICMS"
            Index           =   1
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Registro de Inventário"
            Index           =   2
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Cadastro de Emitentes"
            Index           =   3
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Cadastro de Produtos"
            Index           =   4
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Lançamentos para Apuração"
            Index           =   5
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Dados de Recolhimento GNR"
            Index           =   6
         End
      End
      Begin VB.Menu mnuFISMov 
         Caption         =   "IPI"
         Index           =   4
         Begin VB.Menu mnuFISMovIPI 
            Caption         =   "Apuração de IPI"
            Index           =   1
         End
         Begin VB.Menu mnuFISMovIPI 
            Caption         =   "Lançamentos para Apuração"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   8
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   8
      Begin VB.Menu mnuFISRel 
         Caption         =   "Livros Fiscais ISS"
         Index           =   1
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "Registro de Entradas (mod 1)"
            Index           =   1
         End
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "???Reg. de Utiliz. de Docs Fiscais e Termos de Ocor (mod 2)"
            Index           =   2
         End
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "Registro de Apuracao do ISS (mod 3)"
            Index           =   3
         End
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "???Reg. Entradas Mat. Serv. Terc (REMAS) (mod 4)"
            Index           =   4
         End
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "???Reg. Apuracao Constr. Civil (mod 5)"
            Index           =   5
         End
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "???Reg. Aux. Incorp. Imobiliarias (RADI) (mod 6)"
            Index           =   6
         End
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "???Reg. Apuracao ISS Fixo (mod 7)"
            Index           =   7
         End
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "???Reg. Apuracao Instituicoes financeiras (mod 8)"
            Index           =   8
         End
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "???Reg. Impr. Doctos Fiscais (mod 9)"
            Index           =   9
         End
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Livros Fiscais ICMS / IPI"
         Index           =   3
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Entradas (mod 1 e 1a)"
            Index           =   1
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Saídas (mod 2 e 2a)"
            Index           =   2
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "???Reg. de Impr. de Docs Fiscais (mod 5)"
            Index           =   3
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "???Reg. de Utiliz. de Docs Fiscais e Termos de Ocor (mod 6)"
            Index           =   4
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Inventário (mod 7)"
            Index           =   5
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Apuração do ICMS (mod 9)"
            Index           =   6
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Lista de Códigos de Emitentes (mod 10)"
            Index           =   7
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Tabela de Códigos de Mercadorias (mod 11)"
            Index           =   8
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Listagem de Operações por UF (mod 12)"
            Index           =   9
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Listagem de Prestações por UF (mod 13)"
            Index           =   10
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Dados de Recolhimentos - GNR (mod 14)"
            Index           =   11
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "DIPI, DIPAM, CIAP ???"
            Index           =   12
         End
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Gerenciais ICMS / IPI"
         Index           =   4
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "IR"
         Index           =   6
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "PIS"
         Index           =   7
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "COFINS"
         Index           =   8
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "INSS"
         Index           =   9
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Listagem IN 068/95"
         Index           =   10
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   8
      Begin VB.Menu mnuFISCad 
         Caption         =   "Naturezas de Operação"
         Index           =   1
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tipos de Tributação"
         Index           =   2
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Exceções ICMS"
         Index           =   3
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Exceções IPI"
         Index           =   4
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tributação Fornecedores"
         Index           =   5
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tributação Clientes"
         Index           =   6
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tipos de Registro p/ Apuração de ICMS"
         Index           =   7
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tipos de Registro p/ Apuração de IPI"
         Index           =   8
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   8
      Begin VB.Menu mnuFISRot 
         Caption         =   "Fechamento de Livro"
         Index           =   2
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "Geração Arq. ICMS"
         Index           =   3
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "Geração Arq. IPI (DIPI)"
         Index           =   4
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "Geração IN 068/95"
         Index           =   5
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   8
      Begin VB.Menu mnuFISConfig 
         Caption         =   "Configuração Geral"
         Index           =   1
      End
      Begin VB.Menu mnuFISConfig 
         Caption         =   "Configuração por Tributo"
         Index           =   2
      End
      Begin VB.Menu mnuFISConfig 
         Caption         =   "Alíquotas ICMS"
         Index           =   3
      End
   End
End
Attribute VB_Name = "fismenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Cadastros para o Menu do FIS
Const MENU_FIS_CAD_NATUREZAOP = 1
Const MENU_FIS_CAD_TIPOTRIB = 2
Const MENU_FIS_CAD_EXCECOESICMS = 3
Const MENU_FIS_CAD_EXCECOESIPI = 4
Const MENU_FIS_CAD_TRIBUTACAOFORN = 5
Const MENU_FIS_CAD_TRIBUTACAOCLI = 6
Const MENU_FIS_CAD_TIPOREGAPURACAOICMS = 7
Const MENU_FIS_CAD_TIPOREGAPURACAOIPI = 8

'Movimentos para o Menu do FIS
Const MENU_FIS_MOV_REGENTRADA = 1
Const MENU_FIS_MOV_REGSAIDA = 2

Const MENU_FIS_MOV_ICMS_APURACAO = 1
Const MENU_FIS_MOV_ICMS_REGINVENTARIO = 2
Const MENU_FIS_MOV_ICMS_REGEMITENTES = 3
Const MENU_FIS_MOV_ICMS_REGCADPRODUTOS = 4
Const MENU_FIS_MOV_ICMS_LANCAPURACAO = 5
Const MENU_FIS_MOV_ICMS_GNRICMS = 6

Const MENU_FIS_MOV_IPI_APURACAO = 1
Const MENU_FIS_MOV_IPI_LANCAPURACAO = 2

'Rotinas para o Menu do FIS
Const MENU_FIS_ROT_FECHAMENTOLIVRO = 1

Private Sub mnuFISCad_Click(Index As Integer)

    Select Case Index

        Case MENU_FIS_CAD_NATUREZAOP
            Call Chama_Tela("NaturezaOperacao")
        
        Case MENU_FIS_CAD_TIPOTRIB
            Call Chama_Tela("TipoDeTributacao")
        
        Case MENU_FIS_CAD_EXCECOESICMS
            Call Chama_Tela("ExcecoesICMS")

        Case MENU_FIS_CAD_EXCECOESIPI
            Call Chama_Tela("ExcecoesIPI")

        Case MENU_FIS_CAD_TRIBUTACAOFORN
            Call Chama_Tela("PadraoTribEntrada")

        Case MENU_FIS_CAD_TRIBUTACAOCLI
            Call Chama_Tela("PadraoTribSaida")
        
        Case MENU_FIS_CAD_TIPOREGAPURACAOICMS
            Call Chama_Tela("TipoRegApuracaoICMS")

        Case MENU_FIS_CAD_TIPOREGAPURACAOIPI
            Call Chama_Tela("TipoRegApuracaoIPI")
        
    End Select
    
End Sub

Private Sub mnuFISMov_Click(Index As Integer)

    Select Case Index
    
        Case MENU_FIS_MOV_REGENTRADA
            Call Chama_Tela("EdicaoRegEntrada")
          
        Case MENU_FIS_MOV_REGSAIDA
            Call Chama_Tela("EdicaoRegSaida")
        
    End Select
    
End Sub

Private Sub mnuFISMovICMS_Click(Index As Integer)
    
    Select Case Index
    
        Case MENU_FIS_MOV_ICMS_APURACAO
            Call Chama_Tela("ApuracaoICMS")
        
        Case MENU_FIS_MOV_ICMS_REGINVENTARIO
            Call Chama_Tela("EdicaoRegInventario")
        
        Case MENU_FIS_MOV_ICMS_REGEMITENTES
            Call Chama_Tela("RegESEmitentes")
    
        Case MENU_FIS_MOV_ICMS_REGCADPRODUTOS
            Call Chama_Tela("RegESCadProd")
        
        Case MENU_FIS_MOV_ICMS_LANCAPURACAO
            Call Chama_Tela("ApuracaoICMSItens")
            
        Case MENU_FIS_MOV_ICMS_GNRICMS
            Call Chama_Tela("CadastrarGNRICMS")
            
    End Select
    
End Sub

Private Sub mnuFISMovIPI_Click(Index As Integer)

    Select Case Index
    
        Case MENU_FIS_MOV_IPI_APURACAO
            Call Chama_Tela("ApuracaoIPI")
        
        Case MENU_FIS_MOV_IPI_LANCAPURACAO
            Call Chama_Tela("ApuracaoIPIItens")
        
    End Select

End Sub

Private Sub mnuFISRot_Click(Index As Integer)

    Select Case Index

        Case MENU_FIS_ROT_FECHAMENTOLIVRO
            Call Chama_Tela("EscrituracaoFechamento")
    
    End Select
    
End Sub
