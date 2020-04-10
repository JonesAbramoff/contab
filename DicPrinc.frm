VERSION 5.00
Begin VB.MDIForm Principal 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "SGE - Dicionário de Dados"
   ClientHeight    =   3885
   ClientLeft      =   -15
   ClientTop       =   1665
   ClientWidth     =   9480
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu Liberar_Bloqueios 
         Caption         =   "Liberar Bloqueios"
      End
      Begin VB.Menu Sair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Begin VB.Menu mnuUsuLog 
         Caption         =   "Usuários Logados"
      End
   End
   Begin VB.Menu mnuTeste 
      Caption         =   "&Cadastros"
      Begin VB.Menu Empresa 
         Caption         =   "Empresas"
      End
      Begin VB.Menu Filial 
         Caption         =   "Filiais"
      End
      Begin VB.Menu Cadsep1 
         Caption         =   "-"
      End
      Begin VB.Menu Grupo 
         Caption         =   "Grupos"
      End
      Begin VB.Menu Usuario 
         Caption         =   "Usuários"
      End
      Begin VB.Menu CadSep2 
         Caption         =   "-"
      End
      Begin VB.Menu Rotina 
         Caption         =   "Rotinas"
      End
      Begin VB.Menu menuTela 
         Caption         =   "Telas"
      End
      Begin VB.Menu menuCadRelatorio 
         Caption         =   "Relatórios"
      End
   End
   Begin VB.Menu mnuAcesso 
      Caption         =   "Ac&esso"
      Begin VB.Menu mnuGrupoRotina 
         Caption         =   "Grupo x Rotina"
      End
      Begin VB.Menu mnuGrupoTela 
         Caption         =   "Grupo x Tela"
      End
      Begin VB.Menu mnuGrupoRel 
         Caption         =   "Grupo x Relatórios"
      End
      Begin VB.Menu mnuGrupoBrowseCampo 
         Caption         =   "Grupo x Tela de Browse x Campo"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRotinaGrupo 
         Caption         =   "Rotina x Grupo"
      End
      Begin VB.Menu menuTelaGrupo 
         Caption         =   "Tela x Grupo"
      End
      Begin VB.Menu mnuRelatorioGrupo 
         Caption         =   "Relatório x Grupo"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiberacaoAcesso 
         Caption         =   "&Liberação de Acesso"
      End
      Begin VB.Menu mnuAcessoDados 
         Caption         =   "Dados do &Acesso"
      End
   End
   Begin VB.Menu mnuJanela 
      Caption         =   "&Janela"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Aju&da"
      Begin VB.Menu mnuIndice 
         Caption         =   "&Índice"
      End
      Begin VB.Menu mnuObterAjuda 
         Caption         =   "&Obter Ajuda Sobre..."
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Empresa_Click()
    Load EmpresaTela
    If lErro_Chama_Tela = SUCESSO Then EmpresaTela.Show
End Sub

Private Sub Filial_Click()
    Load FilialEmpresa
    If lErro_Chama_Tela = SUCESSO Then FilialEmpresa.Show
End Sub

Private Sub Grupo_Click()
    
    Load GrupoForm
    If lErro_Chama_Tela = SUCESSO Then GrupoForm.Show
    
End Sub

Private Sub Inst_Click()
'    EmpresaInst1.Show
End Sub

Private Sub Liberar_Bloqueios_Click()
    Load ControleLocks
    If lErro_Chama_Tela = SUCESSO Then ControleLocks.Show
End Sub

Private Sub MDIForm_Load()
    Set GL_objMDIForm = Me
End Sub

Private Sub menuCadRelatorio_Click()
    Load RelCadastro
    If lErro_Chama_Tela = SUCESSO Then RelCadastro.Show
End Sub

Private Sub mnuAcessoDados_Click()
    
Dim lErro As Long

    lErro = AcessoDados.Trata_Parametros()
    If lErro <> SUCESSO Then Exit Sub
    
    AcessoDados.Show
    
End Sub

Private Sub mnuGrupoBrowseCampo_Click()
    Load GrupoBrowseCampo
    If lErro_Chama_Tela = SUCESSO Then GrupoBrowseCampo.Show
End Sub

Private Sub mnuGrupoRel_Click()
    Load GrupoRelatorio
    If lErro_Chama_Tela = SUCESSO Then GrupoRelatorio.Show
End Sub

Private Sub mnuGrupoRotina_Click()
    Load GrupoRotina
    If lErro_Chama_Tela = SUCESSO Then GrupoRotina.Show
End Sub

Private Sub mnuGrupoTela_Click()
    Load GrupoTela
    If lErro_Chama_Tela = SUCESSO Then GrupoTela.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    Dim lErro As Long

    lErro = Sistema_Fechar()

End Sub

Private Sub menuTelaGrupo_Click()
    Load TelaGrupo
    If lErro_Chama_Tela = SUCESSO Then TelaGrupo.Show
End Sub

Private Sub mnuLiberacaoAcesso_Click()
    Load AcessoModulos
    If lErro_Chama_Tela = SUCESSO Then AcessoModulos.Show
End Sub

Private Sub mnuObterAjuda_Click()
    frmAboutDic.Show vbModal
End Sub

Private Sub mnuRelatorioGrupo_Click()
    Load RelatorioGrupo
    If lErro_Chama_Tela = SUCESSO Then RelatorioGrupo.Show
End Sub

Private Sub mnuUsuLog_Click()
    Load UsuLog
    If lErro_Chama_Tela = SUCESSO Then UsuLog.Show
End Sub

Private Sub Rotina_Click()
    Load RotinaTela
    If lErro_Chama_Tela = SUCESSO Then RotinaTela.Show
End Sub

Private Sub mnuRotinaGrupo_Click()
    Load RotinaGrupo
    If lErro_Chama_Tela = SUCESSO Then RotinaGrupo.Show
End Sub

Private Sub Sair_Click()

Dim lErro As Long

    Unload Me

End Sub

Private Sub menuTela_Click()
    Load Tela
    If lErro_Chama_Tela = SUCESSO Then Tela.Show
End Sub

Private Sub Usuario_Click()
    Load UsuarioTela
    If lErro_Chama_Tela = SUCESSO Then UsuarioTela.Show
End Sub

