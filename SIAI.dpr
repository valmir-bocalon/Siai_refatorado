program SIAI;

uses
  Forms,
  Classes,
  RLReport,
  uRuntimeFields in 'uRuntimeFields.pas',
  Windows,
  principal in 'principal.pas' {Frm_principal},
  tabelas in 'tabelas.pas' {DM_Tabelas: TDataModule},
  Funcoes in 'Funcoes.pas' {Frm_Funcoes},
  DigSenha in 'DigSenha.pas' {Frm_DigSenha},
  NovaSenha in 'NovaSenha.pas' {Frm_NovaSenha},
  AchaIpca in 'AchaIpca.pas' {Frm_AchaIpca},
  ipca in 'ipca.pas' {Frm_ipca},
  Usuario in 'Usuario.pas' {Frm_Usuario},
  Loteamento in 'Loteamento.pas' {Frm_Loteamento},
  Empresa in 'Empresa.pas' {FrmCad_Empresa},
  ContraSenha in 'ContraSenha.pas' {FrmContraSenha},
  Estrutura in 'Estrutura.pas' {FrmEstrutura},
  Mapa_loteamento in 'Mapa_loteamento.pas' {Frm_Mapa_loteamento},
  Participante in 'Participante.pas' {Frm_Participante},
  AchaParticipante in 'AchaParticipante.pas' {Frm_AchaParticipante},
  Achaloteamento in 'Achaloteamento.pas' {Frm_Achaloteamento},
  Corretor in 'Corretor.pas' {Frm_Corretor},
  Cad_TipoDoc in 'Cad_TipoDoc.pas' {FrmCad_TipoDoc},
  RelPlanoDeContas in 'RelPlanoDeContas.pas' {FrmRelPlanoDeContas},
  UPlanoDeContas in 'UPlanoDeContas.pas' {FrmCad_PlanodeContas},
  UFrmContasBancarias in 'UFrmContasBancarias.pas' {FrmCad_ContasBancarias},
  AchaImovel in 'AchaImovel.pas' {Frm_AchaImovel},
  AchaCorretor in 'AchaCorretor.pas' {Frm_AchaCorretor},
  AchaPlanoDeContas in 'AchaPlanoDeContas.pas' {Frm_AchaPlanoDeContas},
  Configuracoes in 'Configuracoes.pas' {Frm_Configuracoes},
  Cad_Recebimento in 'Cad_Recebimento.pas' {FrmCad_Recebimento},
  Inc_Recebimento in 'Inc_Recebimento.pas' {Frm_Inc_Recebimento},
  ImpRecibo in 'ImpRecibo.pas' {FrmImpRecibo},
  RelRecebGerencial in 'RelRecebGerencial.pas' {FrmRelRecebGerencial},
  PesqRecebimento_bx in 'PesqRecebimento_bx.pas' {FrmPesqRecebimento_bx},
  Acha_Contabancaria in 'Acha_Contabancaria.pas' {Frm_Acha_Contabancaria},
  RelContaBancaria in 'RelContaBancaria.pas' {Frm_RelBancaria},
  CobrancaBancaria in 'CobrancaBancaria.pas' {FrmCobrancaBancaria},
  RelVenda in 'RelVenda.pas' {Frm_RelVenda},
  RelVenda01 in 'RelVenda01.pas' {Frm_RelVenda01},
  RelRecibodeQuita in 'RelRecibodeQuita.pas' {FrmRelRecibodeQuita},
  NumRemessa in 'NumRemessa.pas' {Frm_NumRemessa},
  Transferencia in 'Transferencia.pas' {Frm_Transferencia},
  TrocaCPF_RG in 'TrocaCPF_RG.pas' {Frm_TrocaCPF_RG},
  AchaLoteVenda in 'AchaLoteVenda.pas' {Frm_AchaLoteVenda},
  resposavel_conjuge in 'resposavel_conjuge.pas' {Frm_resposavel_conjuge},
  acerta_adiamento in 'acerta_adiamento.pas' {Frm_acertaaditamento},
  QuadroResumo2 in 'QuadroResumo2.pas' {Frm_QuadroResumo2},
  ProximoVectoParcelas in 'ProximoVectoParcelas.pas' {Frm_ProximoVctoParcelas},
  RelUCobranca in 'RelUCobranca.pas' {Frm_RelUCobranca},
  TabelasDeMensagens in 'TabelasDeMensagens.pas' {Frm_TabelasDeMensagens},
  BaixaAutomatica in 'BaixaAutomatica.pas' {Frm_BaixaAutomatica},
  AchaVenda in 'AchaVenda.pas' {Frm_AchaVenda},
  RelImovelREMESSA2 in 'RelImovelREMESSA2.pas' {RLReportREMESSA2},
  RelLOTEAMENTO in 'RelLOTEAMENTO.pas' {Frm_Relloteamento},
  RelIpca in 'RelIpca.pas' {Frm_Relipca},
  RelREMESSA in 'RelREMESSA.pas' {Frm_RelREMESSA},
  RelImovel in 'RelImovel.pas' {Frm_RelImovel},
  RelImovel02 in 'RelImovel02.pas' {Frm_RelImovel02},
  Recebitotpago in 'Recebitotpago.pas' {FRM_Recebitotpago},
  RelVenda02_totais_analitico in 'RelVenda02_totais_analitico.pas' {Frm_RelVenda02_totais_analitico},
  Apresentacao in 'Apresentacao.pas' {FrmApresentacao},
  trocar_empreendimento in 'trocar_empreendimento.pas' {Frm_Trocar_Empre},
  RecebBaixa_subst in 'RecebBaixa_subst.pas' {FrmRecebBaixa_subst},
  quitacao in 'quitacao.pas' {Frmquitacao},
  RelReceb02_mensal_ab in 'RelReceb02_mensal_ab.pas' {FrmRelReceb02_mensal_ab},
  Venda in 'Venda.pas' {Frm_Venda},
  RelImovelREMESSA in 'RelImovelREMESSA.pas' {RLReportREMESSA},
  Finanmsg in 'Finanmsg.pas' {FormMensagem},
  PerguntaSIMNAO in 'PerguntaSIMNAO.pas' {FrmPerguntaSIMNAO},
  Pergunta in 'Pergunta.pas' {FrmPergunta},
  Balao in 'Balao.pas' {Frmbalao},
  Achaimoveis in 'Achaimoveis.pas' {Frm_Achaimoveis},
  Achavendas in 'Achavendas.pas' {Frm_Achavendas},
  Relparticipante2 in 'Relparticipante2.pas' {FrmRelparticipante2},
  HORLABEL2 in 'HORLABEL2.PAS' {frmHorLabels2},
  UFrmVisualizar in 'UFrmVisualizar.pas' {FrmVisualizar},
  RelReceb02_carteira in 'RelReceb02_carteira.pas' {FrmRelReceb02_carteira},
  Cidade in 'Cidade.pas' {Frm_Cidade},
  Relloteamento01 in 'Relloteamento01.pas' {Frm_Relloteamento01},
  acertaproximo_reajuste_em_recebimento in 'acertaproximo_reajuste_em_recebimento.pas' {Frm_acertaproximo_reajuste_em_recebimento},
  Excluir_loteamento in 'Excluir_loteamento.pas' {Frm_Excluir},
  RelReceb01 in 'RelReceb01.pas' {FrmRelReceb01},
  Resumo_quita in 'Resumo_quita.pas' {FrmResumo_quita},
  Price in 'Price.pas' {Frm_Price},
  RelPrice in 'RelPrice.pas' {Frm_Relprice},
  AchaCidade in 'AchaCidade.pas' {Frm_AchaCidade},
  acertavalor_final_venda in 'acertavalor_final_venda.pas' {Frm_acertavalor_final_venda},
  zerarNossoNumero_em_recebimento in 'zerarNossoNumero_em_recebimento.pas' {Frm_Zerar_nosso_numero},
  acertaorigem_em_recebimento in 'acertaorigem_em_recebimento.pas' {Frm_acertaorigem_em_recebimento},
  acerta_remessa_recebida in 'acerta_remessa_recebida.pas' {frm_acerta_remessa},
  RelVenda02 in 'RelVenda02.pas' {Frm_RelVenda02},
  RelReceb03 in 'RelReceb03.pas' {FrmRelReceb03},
  RelCobranca2 in 'RelCobranca2.pas' {FrmRelCobranca2},
  Imoveis in 'Imoveis.pas' {Frm_Imoveis},
  PesqRecebimento2 in 'PesqRecebimento2.pas' {FrmPesqRecebimento2},
  RelRecebimento in 'RelRecebimento.pas' {FrmRelRecebimento},
  PesqEndereco in 'PesqEndereco.pas' {FrmPesqEndereco},
  RelReceb02_quitado2 in 'RelReceb02_quitado2.pas' {FrmRelReceb02_quitado2},
  PesqCobranca in 'PesqCobranca.pas' {FrmPesqCobranca},
  Acerto_parcelas_reajuste in 'Acerto_parcelas_reajuste.pas' {Frm_Acerto_parcelas},
  RelReceb02_total_ab_lt in 'RelReceb02_total_ab_lt.pas' {FrmRelReceb02_total_ab_lt},
  PesqRecebimento in 'PesqRecebimento.pas' {FrmPesqRecebimento},
  Relajuste2 in 'Relajuste2.pas' {FrmRelReajuste2},
  ReajusteDeParcelas2 in 'ReajusteDeParcelas2.pas' {Frm_ReajusteDeParcelas2},
  RelCobranca in 'RelCobranca.pas' {FrmRelCobranca},
  RelReceb05 in 'RelReceb05.pas' {FrmRelReceb05},
  RelBolBanco in 'RelBolBanco.pas' {Frm_RelBolBanco},
  RelReceb02_mensal in 'RelReceb02_mensal.pas' {FrmRelReceb02_mensal},
  rel_SimulacaoAditamento in 'rel_SimulacaoAditamento.pas' {frm_SimulacaoAditamento},
  Umapa in 'Umapa.pas' {frmgoogle},
  Participante2 in 'Participante2.pas' {frmparticipante2},
  RelReceb02_diaria in 'RelReceb02_diaria.pas' {FrmRelReceb02_diaria},
  acerta_adversa_recebimento in 'acerta_adversa_recebimento.pas' {Frm_acertaadversa_recebimento},
  rel_diario_lote in 'rel_diario_lote.pas' {frm_rel_diario},
  hormain2 in 'hormain2.pas' {frmMain2},
  hormain in 'hormain.pas' {frmMain},
  HORLABEL in 'HORLABEL.PAS' {frmHorLabels},
  Relparticipante in 'Relparticipante.pas' {FrmRelparticipante},
  ReajusteDeParcelas in 'ReajusteDeParcelas.pas' {Frm_ReajusteDeParcelas},
  RelVenda02_totais in 'RelVenda02_totais.pas' {Frm_RelVenda02_totais},
  Uobs_estorno in 'Uobs_estorno.pas' {FrmObs_estorno},
  Cessao in 'Cessao.pas' {Frm_Cessao},
  RelVenda03 in 'RelVenda03.pas' {Frm_RelVenda03},
  QuadroMemorialt in 'QuadroMemorialt.pas' {Frm_QuadroMemorial},
  RelReceb02_rescisao in 'RelReceb02_rescisao.pas' {FrmRelReceb02_rescisao},
  alerta in 'alerta.pas' {FrmAlerta},
  Relajuste in 'Relajuste.pas' {FrmRelReajuste},
  Vcl.Themes,
  Vcl.Styles,
  RelReceb02_quitado in 'RelReceb02_quitado.pas' {FrmRelReceb02_quitado},
  Uparcelanaopaga in 'Uparcelanaopaga.pas' {frmnaopaga},
  RelReceb02 in 'RelReceb02.pas' {FrmRelReceb02},
  RelReceb04 in 'RelReceb04.pas' {FrmRelReceb04},
  zerarNossoNumero in 'zerarNossoNumero.pas' {FrmzerarNossoNumero},
  RelReceb02_total in 'RelReceb02_total.pas' {FrmRelReceb02_total},
  Acerto_ano_reajuste in 'Acerto_ano_reajuste.pas' {Frm_Acerto_ano},
  ProximoReajusteDeParcelas in 'ProximoReajusteDeParcelas.pas' {Frm_ProximoReajusteDeParcelas},
  Ucobranca in 'Ucobranca.pas' {frmcobranca},
  RelRetorno in 'RelRetorno.pas' {Frm_RelRetorno},
  RelReceb02_cessao in 'RelReceb02_cessao.pas' {FrmRelReceb02_cessao},
  RecebBaixa in 'RecebBaixa.pas' {FrmRecebBaixa},
  QuadroResumo in 'QuadroResumo.pas' {Frm_QuadroResumo},
  Relparticipante_empre_completa in 'Relparticipante_empre_completa.pas' {FrmRelparticipante_empre_completa},
  UFrmPainel in 'UFrmPainel.pas' {frmPainel},
  UFrmDash in 'UFrmDash.pas' {FrmDash},
  Relparticipante1_composta in 'Relparticipante1_composta.pas' {FrmRelparticipante_composta},
  igpm in 'igpm.pas' {Frm_igpm},
  RelIgpm in 'RelIgpm.pas' {Frm_Religpm},
  AchaIgpm in 'AchaIgpm.pas' {Frm_AchaIgpm},
  acertavenda_em_recebimento in 'acertavenda_em_recebimento.pas' {Frm_acertavenda_em_recebimento},
  trocar_lote in 'trocar_lote.pas' {Frm_Trocar},
  Recebi03 in 'Recebi03.pas' {FRM_Recebi03},
  Relparticipante1 in 'Relparticipante1.pas' {FrmRelparticipante1},
  Relparticipante_empre in 'Relparticipante_empre.pas' {FrmRelparticipante_empre},
  UAgenda in 'UAgenda.pas' {FrmAgenda},
  UAgenda_mov in 'UAgenda_mov.pas' {FrmAgenda_Mov},
  usimulaAditamento in 'usimulaAditamento.pas' {frmsimulaaditamento},
  rel_financiamento in 'rel_financiamento.pas' {frm_financiamento},
  RelImovel01 in 'RelImovel01.pas' {Frm_RelImovel01},
  UnPdf in 'UnPdf.pas' {FrmVisualizarPdf},
  RelReceb02_total_ab in 'RelReceb02_total_ab.pas' {FrmRelReceb02_total_ab},
  Aditamento in 'Aditamento.pas' {Frmaditamento},
  Relparticipante_ficha in 'Relparticipante_ficha.pas' {FrmRelparticipante_ficha};

{$R *.res}
var
  Aplicativo : THandle;

begin

    // nao executar mais de uma vez


   //  (1) - Troque o título do seu aplicativo para Teste;
   //  (2) - Veja essa modificação na unit Project1:

    Aplicativo := FindWindow('TApplication', 'Gerenciamento de Loteamento');
//    Aplicativo := FindWindow('FrmPrincipal', nil);
    { "Se, não estiver aberto" }
    If (Aplicativo = 0) then
    begin
      Application.MainFormOnTaskbar := True;

      // O relatorio de memory leaks deve aparecer somente no Debug.
      {$IFDEF DEBUG}
      ReportMemoryLeaksOnShutdown := True;
      {$ELSE}
      ReportMemoryLeaksOnShutdown := False;
      {$ENDIF}
  Application.Initialize;
  if Classes.GetClass('TRLReport') = nil then
    Classes.RegisterClass(TRLReport);
  Application.Title := 'Gerenciamento de Loteamento';
  Application.CreateForm(TFrm_principal, Frm_principal);
  Application.CreateForm(TDM_Tabelas, DM_Tabelas);
  EnsureLoteamentoNomeCidadeField(DM_Tabelas);
  EnsureImovelLoteamentoField(DM_Tabelas);
  EnsureCorretorLookupFields(DM_Tabelas);
  EnsureContaBancariaLookupFields(DM_Tabelas);
  EnsureVendaLookupFields(DM_Tabelas);
  EnsureCompradorLookupFields(DM_Tabelas);
  EnsureRecebimentoCalculatedFields(DM_Tabelas);
  EnsureRecebimentoTempAggregateField(DM_Tabelas);
  EnsureRecebBxTempCalculatedFields(DM_Tabelas);
  EnsureDivididoLookupFields(DM_Tabelas);
  EnsureRecebimentoLookupFields(DM_Tabelas);
  EnsureEmpresaLookupFields(DM_Tabelas);
  EnsureParticipanteLookupFields(DM_Tabelas);
  EnsureResponsavelLookupFields(DM_Tabelas);
  EnsureIncorpLoteamentoLookupFields(DM_Tabelas);
  EnsureRelVendaPartiFields(DM_Tabelas);
  EnsureRuntimeFields(DM_Tabelas);

  Application.CreateForm(TFrm_Funcoes, Frm_Funcoes);
  Application.CreateForm(TFrm_DigSenha, Frm_DigSenha);
  Application.CreateForm(TFrm_NovaSenha, Frm_NovaSenha);
  Application.CreateForm(TFrm_AchaIpca, Frm_AchaIpca);
  Application.CreateForm(TFrmContraSenha, FrmContraSenha);
  Application.CreateForm(TFrm_AchaParticipante, Frm_AchaParticipante);
  Application.CreateForm(TFrm_Achaloteamento, Frm_Achaloteamento);
  Application.CreateForm(TFrm_AchaImovel, Frm_AchaImovel);
  Application.CreateForm(TFrm_AchaCorretor, Frm_AchaCorretor);
  Application.CreateForm(TFrm_AchaPlanoDeContas, Frm_AchaPlanoDeContas);
  Application.CreateForm(TFrm_Configuracoes, Frm_Configuracoes);
  Application.CreateForm(TFrmPesqRecebimento_bx, FrmPesqRecebimento_bx);
  Application.CreateForm(TFrm_Acha_Contabancaria, Frm_Acha_Contabancaria);
  Application.CreateForm(TFrm_NumRemessa, Frm_NumRemessa);
  Application.CreateForm(TFrm_AchaLoteVenda, Frm_AchaLoteVenda);
  Application.CreateForm(TFrm_AchaVenda, Frm_AchaVenda);
  Application.CreateForm(TFormMensagem, FormMensagem);
  Application.CreateForm(TFrmPerguntaSIMNAO, FrmPerguntaSIMNAO);
  Application.CreateForm(TFrmPergunta, FrmPergunta);
  Application.CreateForm(TFrm_Achaimoveis, Frm_Achaimoveis);
  Application.CreateForm(TFrm_Achavendas, Frm_Achavendas);
  Application.CreateForm(TFrm_AchaCidade, Frm_AchaCidade);
  Application.CreateForm(TFrmPesqRecebimento2, FrmPesqRecebimento2);
  Application.CreateForm(TFrmPesqEndereco, FrmPesqEndereco);
  Application.CreateForm(TFrmPesqCobranca, FrmPesqCobranca);
  Application.CreateForm(TFrmPesqRecebimento, FrmPesqRecebimento);
  Application.CreateForm(TFrm_AchaIgpm, Frm_AchaIgpm);
  Application.Run;
    end
    Else
    begin
    
     { "Se, estiver aberto ele restaura a janela" }
     ShowWindow(Aplicativo, SW_RESTORE);
    end;
end.
