unit tabelas;

interface

uses
  Windows, Messages, SysUtils, Classes, DB, ZConnection, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, DBClient, ImgList, Controls, IniFiles,
  Graphics, Forms, Dialogs, StdCtrls, Shellapi, ExtCtrls, Menus, Registry,
  ZSqlUpdate, Midas, MidasLib, MidasCon, ZSqlMonitor, ZAbstractConnection,
  System.ImageList;


type
  TDM_Tabelas = class(TDataModule)
    DS_Usuario: TDataSource;
    ZQUsuario: TZQuery;
    zconeccao: TZConnection;
    DS_Login: TDataSource;
    DS_Loc_Senha: TDataSource;
    DS_Permissoes: TDataSource;
    DS_Cidade: TDataSource;
    ZQLogin: TZQuery;
    ZQLoc_Senha: TZQuery;
    ZQPermissoes: TZQuery;
    ZQCidade: TZQuery;
    DSLocaisSenha: TDataSource;
    CDSLocaisSenha: TClientDataSet;
    CDSLocaisSenhaGrupo: TWideStringField;
    CDSLocaisSenhapedir: TWideStringField;
    CDSLocaisSenhapermitir: TWideStringField;
    DS_UsuTemp: TDataSource;
    ZQUsuTemp: TZQuery;
    DS_AgrupaGrupo: TDataSource;
    ZQAgrupaGrupo: TZQuery;
    CDSLocaisSenhaidlocais: TIntegerField;
    CDSLocaisSenhadescricao: TWideStringField;
    DS_AchaCidade: TDataSource;
    ZQAchaCidade: TZQuery;
    DS_Loteamento: TDataSource;
    ZQLoteamento: TZQuery;
    DS_Empresa: TDataSource;
    ZQEmpresa: TZQuery;
    DS_Participante: TDataSource;
    ZqParticipante: TZQuery;
    DS_Incorp: TDataSource;
    CDSIncorp: TClientDataSet;
    CDSIncorpcodparticipante: TIntegerField;
    CDSIncorpnomeparti: TWideStringField;
    DS_AchaParticip: TDataSource;
    ZQAchaParticip: TZQuery;
    DS_Responsavel: TDataSource;
    ZQresponsavel: TZQuery;
    DSincorp_loteame: TDataSource;
    ZQincorp_loteame: TZQuery;
    CDSIncorppercent: TFloatField;
    DS_Imovel: TDataSource;
    ZQImovel: TZQuery;
    DS_AchaLotea: TDataSource;
    ZQAchaLotea: TZQuery;
    DS_Corretor: TDataSource;
    ZQCorretor: TZQuery;
    DS_TipoDoc: TDataSource;
    ZQTipoDoc: TZQuery;
    DS_PlanoDeContas: TDataSource;
    ZQPlanoDeContas: TZQuery;
    DS_ContaBancaria: TDataSource;
    DS_Especial: TDataSource;
    ZQContaBancaria: TZQuery;
    ZQEspecial: TZQuery;
    DS_Quadras: TDataSource;
    ZQQuadras: TZQuery;
    DS_QuadrasTemp: TDataSource;
    CDSQuadrasTemp: TClientDataSet;
    CDSQuadrasTempquadra: TWideStringField;
    CDSQuadrasTempordem: TIntegerField;
    CDSQuadrasTempapelido: TWideStringField;
    DS_Venda: TDataSource;
    ZQVenda: TZQuery;
    DS_ImovelTemp: TDataSource;
    ZQImovelTemp: TZQuery;
    DS_CompradorTemp: TDataSource;
    CDSCompradorTemp: TClientDataSet;
    CDSCompradorTempcodparticipante: TIntegerField;
    CDSCompradorTempnomeparticipante: TWideStringField;
    DS_Comprador: TDataSource;
    ZQComprador: TZQuery;
    DS_Vendedor: TDataSource;
    ZQVendedor: TZQuery;
    DS_VendedorTemp: TDataSource;
    CDSVendedorTemp: TClientDataSet;
    CDSVendedorTempnomecorretor: TWideStringField;
    DS_AchaCorretor: TDataSource;
    ZQAchaCorretor: TZQuery;
    CDSVendedorTempcodcorretor: TIntegerField;
    CDSCompradorTempdoc1: TWideStringField;
    CDSCompradorTempdoc2: TWideStringField;
    CDSCompradorTemppessoa: TWideStringField;
    CDSCompradorTempfone: TWideStringField;
    DSPag_Entrada: TDataSource;
    CDSPagEntrada: TClientDataSet;
    CDSPagEntradacontabil: TIntegerField;
    CDSPagEntradavencimento: TDateField;
    CDSPagEntradaordem: TWideStringField;
    CDSPagEntradavalor: TFloatField;
    CDSPagEntradatipodoc: TWideStringField;
    CDSPagEntradacontabildescri: TWideStringField;
    DS_Recebimento: TDataSource;
    ZQRecebimento: TZQuery;
    DS_Receb_baixa: TDataSource;
    ZQReceb_Baixa: TZQuery;
    DS_CobaRe: TDataSource;
    ZQCobaRe: TZQuery;
    DS_ReBxHi: TDataSource;
    ZQReBxHi: TZQuery;
    DS_RecebBxTemp: TDataSource;
    ZQRecebBxTemp: TZQuery;
    DS_NumOrdem: TDataSource;
    ZQNumOrdem: TZQuery;
    DS_Cheque: TDataSource;
    ZQCheque: TZQuery;
    ZQConfiguracoes: TZQuery;
    DS_AchaPlanoDeContas: TDataSource;
    ZQAchaPlanoDeContas: TZQuery;
    CDSPagEntradabanco: TWideStringField;
    CDSPagEntradaagencia: TWideStringField;
    CDSPagEntradaconta: TWideStringField;
    CDSPagEntradadocumento: TWideStringField;
    CDSPagEntradanumero: TWideStringField;
    DS_Dividido: TDataSource;
    ZQDividido: TZQuery;
    DS_Caixa: TDataSource;
    DS_MovBancaria: TDataSource;
    ZQCaixa: TZQuery;
    ZQMovBancaria: TZQuery;
    DS_Recebimento_temp: TDataSource;
    ZQRecebimento_temp: TZQuery;
    DS_PagParcela: TDataSource;
    CDSPagEntradanomedono: TWideStringField;
    DS_AchaContaBanc: TDataSource;
    ZQAchaContaBanc: TZQuery;
    CDSIncorpcodcontabancaria: TIntegerField;
    CDSIncorpcontabancaria: TWideStringField;
    DS_Bol_men: TDataSource;
    ZQBol_men: TZQuery;
    DS_Remes_Receb: TDataSource;
    ZQRemes_Receb: TZQuery;
    DS_RemesRec: TDataSource;
    DS_Empreendimento: TDataSource;
    CDSEmpreendimento: TClientDataSet;
    CDSEmpreendimentoidloteamento: TIntegerField;
    CDSEmpreendimentomarca: TWideStringField;
    CDSEmpreendimentonomeloteamento: TWideStringField;
    DS_BancRemes: TDataSource;
    ZQBancRemes: TZQuery;
    CDSEmpreendimentocodcontabancaria: TIntegerField;
    CDSEmpreendimentoconta: TWideStringField;
    CDSEmpreendimentonomebanco: TWideStringField;
    CDSEmpreendimentoidparticipante: TIntegerField;
    CDSEmpreendimentonomeparticipante: TWideStringField;
    CDSEmpreendimentoidordem: TIntegerField;
    DS_RecebParticp: TDataSource;
    ZQRecebParticp: TZReadOnlyQuery;
    CDSEmpreendimentolayoutBoleto: TWideStringField;
    CDSEmpreendimentonomearquivo: TWideStringField;
    ZQRemesRec: TZQuery;
    CDSEmpreendimentocod_no_banco: TWideStringField;
    CDSEmpreendimentonometitular: TWideStringField;
    CDSEmpreendimenton_agencia: TWideStringField;
    CDSEmpreendimenton_agencia_v: TWideStringField;
    CDSEmpreendimentoconta_v: TWideStringField;
    CDSEmpreendimentojurosemboleto: TFloatField;
    CDSEmpreendimentocarteira: TWideStringField;
    DS_MarcaTit: TDataSource;
    CDS_MarcaTit: TClientDataSet;
    CDS_MarcaTitIdreceb: TIntegerField;
    CDS_MarcaTitidparti: TIntegerField;
    CDS_MarcaTitDT_Entrada: TDateField;
    CDS_MarcaTitmarca: TWideStringField;
    CDS_MarcaTitDt_Vencimento: TDateField;
    CDS_MarcaTitValor: TFloatField;
    CDS_MarcaTitordem: TWideStringField;
    CDS_MarcaTitnome_parte: TWideStringField;
    CDS_MarcaTitende_cob: TWideStringField;
    CDS_MarcaTitcep_cob: TWideStringField;
    CDS_MarcaTitdoc1: TWideStringField;
    CDS_MarcaTitTipDoc: TWideStringField;
    CDS_MarcaTitnumordem: TIntegerField;
    CDS_MarcaTitcliente: TIntegerField;
    DS_RelVenda: TDataSource;
    ZQRelVenda: TZQuery;
    DS_RelVendaParti: TDataSource;
    ZQRelVendaParti: TZQuery;
    DS_RelVendaCorr: TDataSource;
    ZQRelVendaCorr: TZQuery;
    DS_ResponsavelTemp: TDataSource;
    CDSResponsavelTemp: TClientDataSet;
    CDSResponsavelTempdocumen1: TWideStringField;
    CDSResponsavelTempdocumen2: TWideStringField;
    CDSResponsavelTempnasicmento: TDateField;
    CDSResponsavelTempcid_natural: TLargeintField;
    CDSResponsavelTempnomerepres: TWideStringField;
    CDSResponsavelTempemailrepres: TWideStringField;
    CDSResponsavelTempprofiss: TWideStringField;
    CDSResponsavelTempassina: TWideStringField;
    CDSResponsavelTempnomecidade: TWideStringField;
    CDSResponsavelTemprenda: TFloatField;
    DS_Profissao: TDataSource;
    ZQProfissao: TZQuery;
    CDSCompradorTempordem: TIntegerField;
    CDSVendedorTempordem: TIntegerField;
    DS_Conjuge: TDataSource;
    ZQConjuge: TZQuery;
    CDSResponsavelTempadmissao: TDateField;
    CDSResponsavelTemptrabalho: TWideStringField;
    CDSResponsavelTempestcidade: TWideStringField;
    CDS_MarcaTitidloteam: TIntegerField;
    CDSEmpreendimentoapelido: TWideStringField;
    CDSResponsavelTempconjugenome: TWideStringField;
    CDSResponsavelTempconjugedoc1: TWideStringField;
    CDSResponsavelTempconjugedoc2: TWideStringField;
    CDSResponsavelTempconjugenasc: TDateField;
    CDSResponsavelTempconjugecid: TIntegerField;
    CDSResponsavelTempconjugeass: TWideStringField;
    CDSResponsavelTempconjugeDT: TDateField;
    CDSResponsavelTempconjugecartorio: TWideStringField;
    CDSResponsavelTempconjugelivro: TWideStringField;
    CDSResponsavelTempconjugefolha: TWideStringField;
    CDSResponsavelTempconjugenatur: TIntegerField;
    CDSResponsavelTempconjugenaturcidnome: TWideStringField;
    CDSResponsavelTempconjugenaturcidUF: TWideStringField;
    CDSResponsavelTempconjuge_cidcartorio: TIntegerField;
    CDSResponsavelTempconjuge_cid_cart_UF: TWideStringField;
    CDSResponsavelTempconjuge_cid_cart_nome: TWideStringField;
    DS_Compr_Dados: TDataSource;
    ZQCompr_Dados: TZQuery;
    DS_Compr_Resp_Dados: TDataSource;
    ZQCompr_Resp_Dados: TZQuery;
    CS_MensMot: TDataSource;
    ZQMensMot: TZQuery;
    DS_Mens_Retorno: TDataSource;
    ZQMens_Retorno: TZQuery;
    DS_Mens_Motivo: TDataSource;
    ZQMens_Motivo: TZQuery;
    ImageList1: TImageList;
    DS_AchaVenda: TDataSource;
    ZQAchaVenda: TZQuery;
    DS_Compr_conjuge: TDataSource;
    ZQCompr_conjuge: TZQuery;
    DS_Ass: TDataSource;
    CDSAss: TClientDataSet;
    CDSAssassinante: TWideStringField;
    CDSAsscargo: TWideStringField;
    CDSAssresponsavel: TWideStringField;
    DS_Ass1: TDataSource;
    CDSAss1: TClientDataSet;
    CDSAss1assinante: TWideStringField;
    CDSAss1cargo: TWideStringField;
    CDSAss1responsavel: TWideStringField;
    DS_Recebimento_tempE: TDataSource;
    ZQRecebimento_tempE: TZQuery;
    DataZQBancRemes2: TDataSource;
    ZQBancRemes2: TZQuery;
    DataZQRemes_Receb2: TDataSource;
    ZQRemes_Receb2: TZQuery;
    Datacontador_bx: TDataSource;
    contador_bx: TZQuery;
    DataZQRemes_Receb_atualiza: TDataSource;
    ZQRemes_Receb_atualiza: TZQuery;
    CDSIncorpdigito_dif: TWideStringField;
    CDSEmpreendimentodigito_dif: TWideStringField;
    CDSCompradorTempnacionalidade: TWideStringField;
    CDSCompradorTempprofissao: TWideStringField;
    CDSCompradorTempestadocivil: TWideStringField;
    Dataacha_imovel: TDataSource;
    Acha_imovel: TZQuery;
    DataZQvnd_quadro: TDataSource;
    ZQvnd_quadro: TZQuery;
    DS_Ass2: TDataSource;
    CDSAss2: TClientDataSet;
    CDSAss2Assinante: TWideStringField;
    CDSAss2cargo: TWideStringField;
    CDSAss2responsavel: TWideStringField;
    CDSIncorpconjuge: TWideStringField;
    CDSAss1conjuge: TWideStringField;
    CDSAss1idconjuge: TIntegerField;
    CDSIncorpidconjuge: TIntegerField;
    CDSAss1assina: TWideStringField;
    CDSIncorpassina: TWideStringField;
    ZQProcuradores: TZQuery;
    Dataprocuradores: TDataSource;
    CDSEmpreendimentoconvenio: TWideStringField;
    DS_Price: TDataSource;
    ZQprice: TZQuery;
    ZQigpm: TZQuery;
    DSIgpm: TDataSource;
    CDSCompradorTemppercentual: TFloatField;
    ZQAchaIgpm: TZQuery;
    DS_AchaIgpm: TDataSource;
    DS_log_logradouro: TDataSource;
    ZQlog_logradouro: TZQuery;
    localizacao: TZQuery;
    Datalocalizacao: TDataSource;
    CDSCompradorTempMarcar: TBooleanField;
    ZQEmail: TZQuery;
    DS_Email: TDataSource;
    DS_Configuracoes: TDataSource;
    Datalog_tipo_logr: TDataSource;
    log_tipo_logr: TZQuery;
    Datalog_faixa_uf: TDataSource;
    log_faixa_uf: TZQuery;
    DataSource1: TDataSource;
    log_localidade: TZQuery;
    DS_zqbairro: TDataSource;
    ZQbairro: TZQuery;
    CDSPagEntradaMarcar: TIntegerField;
    DS_ZQsequencia_remessa: TDataSource;
    ZQsequencia_remessa: TZQuery;
    ZQmemorial: TZQuery;
    DS_ZQmemorial: TDataSource;
    DS_proposta: TDataSource;
    ZQProposta: TZQuery;
    ZQCheque_proposta: TZQuery;
    DS_ZQCheque_proposta: TDataSource;
    ZQVendedor_proposta: TZQuery;
    DS_ZQVendedor_proposta: TDataSource;
    ZQComprador_proposta: TZQuery;
    DS_ZQComprador_proposta: TDataSource;
    ZQRecebimento_proposta: TZQuery;
    DS_ZQRecebimento_proposta: TDataSource;
    ZqParticipante_proposta: TZQuery;
    DS_Participante_proposta: TDataSource;
    ZQresponsavel_proposta: TZQuery;
    DS_Responsavel_proposta: TDataSource;
    ZQConjuge_proposta: TZQuery;
    DS_Conjuge_proposta: TDataSource;
    ZQagcompromisso: TZQuery;
    DS_agcompromisso: TDataSource;
    ZQagcompromisso_consulta: TZQuery;
    DS_agcompromisso_consulta: TDataSource;
    DS_Rescisao: TDataSource;
    ZQRescisao: TZQuery;
    ZSQLMonitor1: TZSQLMonitor;
    CDSPagParcela: TClientDataSet;
    CDSPagParcelaordem: TStringField;
    CDSPagParcelacontabil: TIntegerField;
    CDSPagParcelatipodoc: TStringField;
    CDSPagParcelavencimento: TDateField;
    CDSPagParcelavalor: TFloatField;
    CDSPagParcelacontabildescri: TStringField;
    CDSPagParcelabanco: TStringField;
    CDSPagParcelaagencia: TStringField;
    CDSPagParcelaconta: TStringField;
    CDSPagParcelanumero: TStringField;
    CDSPagParceladocumento: TStringField;
    CDSPagParcelanomedono: TStringField;
    CDSPagParcelaMarcar: TIntegerField;
    CDSPagParcelaSaldo: TFloatField;
    CDSCompradorTempPComissao_corretor: TFloatField;
    CDSVendedorTempPcomissao_corretor: TFloatField;
    CDSEmpreendimentoPerc_comissao: TFloatField;
    ZQIPCA: TZQuery;
    Ds_ipca: TDataSource;
    ZQAchaIpca: TZQuery;
    DS_AchaIpca: TDataSource;
    ds_qryagenda: TDataSource;
    qryagenda: TZQuery;
    procedure zconeccaoBeforeConnect(Sender: TObject);
    procedure ZQVendaBeforeOpen(DataSet: TDataSet);
    procedure ZQCompradorBeforeOpen(DataSet: TDataSet);
    procedure ZQRecebimentoBeforeOpen(DataSet: TDataSet);
    procedure ZQRecebBxTempBeforeOpen(DataSet: TDataSet);
    procedure ZQDivididoBeforeOpen(DataSet: TDataSet);
    procedure CDSIncorpCalcFields(DataSet: TDataSet);
    procedure ZQCompr_conjugeCalcFields(DataSet: TDataSet);
    procedure ZQRecebBxTempCalcFields(DataSet: TDataSet);
    procedure DS_ImovelDataChange(Sender: TObject; Field: TField);
    procedure ZQCompr_DadosCalcFields(DataSet: TDataSet);
    procedure DS_ParticipanteDataChange(Sender: TObject; Field: TField);
    procedure DS_ChequeDataChange(Sender: TObject; Field: TField);
    procedure ZQCompr_Resp_DadosCalcFields(DataSet: TDataSet);
    procedure DS_RescisaoDataChange(Sender: TObject; Field: TField);

  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

procedure EnsureLoteamentoNomeCidadeField(ADataModule: TDM_Tabelas);
procedure EnsureImovelLoteamentoField(ADataModule: TDM_Tabelas);
procedure EnsureCorretorLookupFields(ADataModule: TDM_Tabelas);
procedure EnsureContaBancariaLookupFields(ADataModule: TDM_Tabelas);
procedure EnsureVendaLookupFields(ADataModule: TDM_Tabelas);
procedure EnsureCompradorLookupFields(ADataModule: TDM_Tabelas);
procedure EnsureRecebimentoCalculatedFields(ADataModule: TDM_Tabelas);
procedure EnsureRecebimentoTempAggregateField(ADataModule: TDM_Tabelas);
procedure EnsureRelVendaPartiFields(ADataModule: TDM_Tabelas);
procedure EnsureRecebBxTempCalculatedFields(ADataModule: TDM_Tabelas);
procedure EnsureDivididoLookupFields(ADataModule: TDM_Tabelas);
procedure EnsureRecebimentoLookupFields(ADataModule: TDM_Tabelas);
procedure EnsureEmpresaLookupFields(ADataModule: TDM_Tabelas);
procedure EnsureParticipanteLookupFields(ADataModule: TDM_Tabelas);
procedure EnsureResponsavelLookupFields(ADataModule: TDM_Tabelas);
procedure EnsureIncorpLoteamentoLookupFields(ADataModule: TDM_Tabelas);

var
  DM_Tabelas: TDM_Tabelas;

implementation

uses funcoes, Imoveis, RelREMESSA, Participante, Cad_Recebimento,
  principal, uRuntimeFields;

{$R *.dfm}

procedure TDM_Tabelas.zconeccaoBeforeConnect(Sender: TObject);
Var
  ArqIni : tIniFile;
  varhostname, varsenha, varschemata, VarUsuarioSQL, VarPedeSenha, VarDestino,
  VarNomeArquivo, VarPorta, VarPath, VarNomePasta : string;

begin
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  try
  Try
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );
  if empty(varhostname) then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','HOSTNAME','localhost');
    varhostname := 'localhost';
  end;
  Finally
  end;
  Try
  varschemata := ArqIni.ReadString('BANCO_DE_DADOS','SCHEMATA', varschemata );
  if empty(varschemata) then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','SCHEMATA','siai');
    varschemata := 'siai';
  end;
  Finally
  end;
  Try
  VarUsuarioSQL := ArqIni.ReadString('BANCO_DE_DADOS','USUARIO', VarUsuarioSQL );
  if empty(VarUsuarioSQL) then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','USUARIO','root');
    VarUsuarioSQL := 'root';
  end;
  Finally
  end;
  Try
  VarPorta := ArqIni.ReadString('BANCO_DE_DADOS','PORTA', VarPorta );
  if empty(VarPorta) then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','PORTA','3306');
    VarPorta := '3306';
  end;
  Finally
  end;
  Try
  varsenha := ArqIni.ReadString('BANCO_DE_DADOS','SENHASQL', varsenha );
  if empty(varsenha) then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','SENHASQL','000001C70385');
    varsenha := '000001C70385';
  end;
  Finally
  end;
  Try
    VarNomePasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
    if empty(VarNomePasta) then Begin
      ArqIni.WriteString('BANCO_DE_DADOS','pasta_servidor','master');
      VarNomePasta := 'master';
    end;
  Finally
  end;
  Try
  VarDestino := ArqIni.ReadString('BACKUP','DESTINHO', VarDestino );
  if empty(VarDestino) then Begin
    ArqIni.WriteString('BACKUP','DESTINHO',VarPath);
    VarDestino := VarPath;
  end;
  Finally
  end;
  Try
  VarPedeSenha := ArqIni.ReadString('BACKUP','PEDESENHA', VarPedeSenha );
  if empty(VarPedeSenha) then Begin
    ArqIni.WriteString('BACKUP','PEDESENHA','NºO');
    VarPedeSenha := 'NºO';
  end;
  Finally
  end;
  Try
    VarNomeArquivo := ArqIni.ReadString('BACKUP','NOME_ARQUIVO', VarNomeArquivo );
    if empty(VarNomeArquivo) then Begin
      ArqIni.WriteString('BACKUP','NOME_ARQUIVO','LOTES');
      VarNomeArquivo := 'LOTES';
    end;
  Finally
  end;
  zconeccao.HostName := varhostname;
  zconeccao.Password := decodifica(varsenha);
  zconeccao.Database := varschemata;
  zconeccao.User     := VarUsuarioSQL;
  zconeccao.Port     := strtoint(VarPorta);
  finally
    ArqIni.Free;
  end;
end;

procedure TDM_Tabelas.ZQVendaBeforeOpen(DataSet: TDataSet);
begin
  EnsureVendaLookupFields(Self);
end;

procedure TDM_Tabelas.ZQCompradorBeforeOpen(DataSet: TDataSet);
begin
  EnsureCompradorLookupFields(Self);
end;

procedure TDM_Tabelas.ZQRecebimentoBeforeOpen(DataSet: TDataSet);
begin
  EnsureRecebimentoCalculatedFields(Self);
  EnsureRecebimentoLookupFields(Self);
end;

procedure TDM_Tabelas.ZQRecebBxTempBeforeOpen(DataSet: TDataSet);
begin
  EnsureRecebBxTempCalculatedFields(Self);
end;

procedure TDM_Tabelas.ZQDivididoBeforeOpen(DataSet: TDataSet);
begin
  EnsureDivididoLookupFields(Self);
end;

procedure TDM_Tabelas.CDSIncorpCalcFields(DataSet: TDataSet);
begin
  if ZQContaBancaria.Locate('idconta_bancaria',CDSIncorpcodcontabancaria.Value,[]) Then
    CDSIncorpcontabancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString+'-'+DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString+' - '+DM_Tabelas.ZQContaBancaria.FieldByName('nomebanco').AsString+' - '+DM_Tabelas.ZQContaBancaria.FieldByName('nometitular').AsString;

end;

procedure TDM_Tabelas.ZQCompr_conjugeCalcFields(DataSet: TDataSet);
begin
  ZQCompr_conjuge.FieldByName('comar_cid_est').AsString := alltrim(ZQCompr_conjuge.FieldByName('comar_cidnome').AsString)+'/'+alltrim(ZQCompr_conjuge.FieldByName('comar_cidest').AsString);
  ZQCompr_conjuge.FieldByName('cart_cid_est').AsString  := alltrim(ZQCompr_conjuge.FieldByName('cart_cidnome').AsString)+'/'+alltrim(ZQCompr_conjuge.FieldByName('cart_cidest').AsString);
  ZQCompr_conjuge.FieldByName('cidestnatural').AsString := alltrim(ZQCompr_conjuge.FieldByName('cidnatural').AsString)+'/'+alltrim(ZQCompr_conjuge.FieldByName('estnatural').AsString);
end;

procedure TDM_Tabelas.ZQRecebBxTempCalcFields(DataSet: TDataSet);
begin
  if ZQRecebBxTemp.FieldByName('percent_usado').AsFloat >0 Then Begin
    ZQRecebBxTemp.FieldByName('vrusado').AsFloat := ZQRecebBxTemp.FieldByName('Vr_rec').AsFloat*ZQRecebBxTemp.FieldByName('percent_usado').AsFloat/100;
  end
  else
    ZQRecebBxTemp.FieldByName('vrusado').AsFloat := ZQRecebBxTemp.FieldByName('Vr_rec').AsFloat;
  ZQRecebBxTemp.FieldByName('descusado').AsFloat := ZQRecebBxTemp.FieldByName('Desc_Vr').AsFloat*ZQRecebBxTemp.FieldByName('percent_usado').AsFloat/100;
  if ZQRecebBxTemp.FieldByName('Juros_Vr').AsFloat>0 Then
    ZQRecebBxTemp.FieldByName('jurusado').AsFloat :=  ZQRecebBxTemp.FieldByName('juros').AsFloat;
end;

procedure TDM_Tabelas.DS_ImovelDataChange(Sender: TObject; Field: TField);
begin
  if Frm_Imoveis=nil then
     Frm_Imoveis:=TFrm_Imoveis.create(Application);
  if DM_Tabelas.ZQimovel.FieldByName('matri').AsString='SIM' then
  begin
    Frm_Imoveis.Label23.Visible:=True;
    Frm_Imoveis.DBEMATRICULA.Visible:=True;
  end
  else
  begin
    Frm_Imoveis.Label23.Visible:=False;
    Frm_Imoveis.DBEMATRICULA.Visible:=False;
  end;
end;

procedure TDM_Tabelas.ZQCompr_DadosCalcFields(DataSet: TDataSet);
begin
  ZQCompr_Dados.FieldByName('cidadeestadonatual').AsString := alltrim(ZQCompr_Dados.FieldByName('natural').AsString)+'/'+alltrim(ZQCompr_Dados.FieldByName('estado').AsString);
  ZQCompr_Dados.FieldByName('cid_uf_cob').AsString:=alltrim(ZQCompr_Dados.FieldByName('cidnome_cob').AsString)+'/'+alltrim(ZQCompr_Dados.FieldByName('cidUF_cob').AsString);
end;

procedure TDM_Tabelas.DS_ParticipanteDataChange(Sender: TObject;
  Field: TField);
begin
  if Frm_Participante<>nil then
  begin
     if Frm_Participante.Pagina.ActivePage=Frm_Participante.TabSheet3 then
     begin
        Frm_Participante.ZQContratos.close;
        Frm_Participante.ZQContratos.SQL.Clear;
        Frm_Participante.ZQContratos.SQL.Add('SELECT * FROM participante as p join comprador as c on p.idpaticipante=c.paticipante_idpaticipante join cidade as ci on p.cidade_cob=ci.idcidade join venda as v on c.venda_idvenda=v.idvenda join imovel as i on v.imovel=i.idimovel ');
        Frm_Participante.ZQContratos.SQL.Add(' join loteamento as l on l.idloteamento=i.loteamento_idloteamento  where p.idpaticipante='+quotedstr(DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text)+' order by p.idpaticipante');
        Frm_Participante.ZQContratos.Open;
      end;
  end;
end;

procedure TDM_Tabelas.DS_ChequeDataChange(Sender: TObject; Field: TField);
begin
  if FrmCad_Recebimento<>nil then
  begin
    if FrmCad_Recebimento.Pag_Receb.PageIndex=0 then
    begin
      if DM_tabelas.ZQTipodoc.active=false then
         DM_tabelas.ZQTipodoc.open;
      DM_tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.ZQRecebimento.FieldByName('TipDoc').AsString,[]);
      IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then
        FrmCad_Recebimento.GBCheque.Visible := True
      else
        FrmCad_Recebimento.GBCheque.Visible := False;
    end;
  end;
end;

procedure TDM_Tabelas.ZQCompr_Resp_DadosCalcFields(DataSet: TDataSet);
begin
  if frm_principal.lblvnd.Caption<>'S' then
  begin
    ZQCidade.Locate('idcidade',ZQCompr_Resp_Dados.FieldByName('cid_natural').AsLargeInt,[]);
    ZQCompr_Resp_Dados.FieldByName('cid_est_resp').AsString:=alltrim(DM_Tabelas.ZQCidade.FieldByName('nomecid').AsString)+'/'+alltrim(DM_Tabelas.ZQCidade.FieldByName('estado').AsString);

{    if Acha_Cidade(ZQCompr_Resp_Dadoscidade_natural_resp.Value) Then
    Begin
      ZQCompr_Resp_Dadoscid_est_resp.Value:=alltrim(DM_Tabelas.ZQAchaCidadenomecid.Value)+'/'+alltrim(DM_Tabelas.ZQAchaCidadeestado.Value);
    end;}
  end;
  frm_principal.lblvnd.Caption:='N';
end;

procedure TDM_Tabelas.DS_RescisaoDataChange(Sender: TObject;
  Field: TField);
begin
  if Frm_Imoveis<>nil then
  begin
    if ZQRescisao.RecordCount>0 then
    begin
       Frm_Imoveis.ZQRecebimento.close;
       Frm_Imoveis.ZQRecebimento.SQL.clear;
       Frm_Imoveis.ZQRecebimento.SQL.add('Select *, sum(valor) as VRRECEBIDO from Recebimento as rb join rescisao as re on re.idvenda=rb.venda_idvenda order by rb.DT_Vencimento');
       Frm_Imoveis.ZQRecebimento.open;
    end
    else
    begin
       Frm_Imoveis.ZQRecebimento.close;
    end;
  end;
end;


procedure EnsureLoteamentoNomeCidadeField(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureCalculatedField(const AComponentName, AFieldName: string;
    ASize: Integer);
  var
    LField: TWideStringField;
  begin
    if ADataModule.ZQLoteamento.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := ASize;
      LField.DataSet := ADataModule.ZQLoteamento;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQLoteamento = nil) then
    Exit;
  if (ADataModule.ZQLoteamento.FindField('nomecidade') <> nil) and
     (ADataModule.ZQLoteamento.FindField('estado') <> nil) then
    Exit;

  { O grid de Loteamento referencia este campo no DFM. A existência do
    campo não pode depender de a consulta de cidades estar aberta durante a
    criação do formulário; o hook de campos calcula o lookup depois. }
  LWasActive := ADataModule.ZQLoteamento.Active;
  if LWasActive then
    ADataModule.ZQLoteamento.Close;
  try
    EnsureCalculatedField('ZQLoteamentonomecidade', 'nomecidade', 100);
    EnsureCalculatedField('ZQLoteamentoestado', 'estado', 2);
  finally
    if LWasActive then
      ADataModule.ZQLoteamento.Open;
  end;
end;

procedure EnsureImovelLoteamentoField(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;
  LField: TWideStringField;

  procedure EnsureCalculatedField(const AComponentName, AFieldName: string);
  begin
    if ADataModule.ZQImovel.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := 100;
      LField.DataSet := ADataModule.ZQImovel;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQImovel = nil) then
    Exit;
  if (ADataModule.ZQImovel.FindField('nomeloteamento') <> nil) and
     (ADataModule.ZQImovel.FindField('descricao') <> nil) then
    Exit;

  { This lookup is used by the main property query even when its SQL does not
    return the loteamento name as a physical column. }
  LWasActive := ADataModule.ZQImovel.Active;
  if LWasActive then
    ADataModule.ZQImovel.Close;
  try
    EnsureCalculatedField('ZQImovelnomeloteamento', 'nomeloteamento');
    EnsureCalculatedField('ZQImoveldescricao', 'descricao');
  finally
    if LWasActive then
      ADataModule.ZQImovel.Open;
  end;
end;

procedure EnsureCorretorLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureCalculatedField(const AComponentName, AFieldName: string;
    ASize: Integer);
  var
    LField: TWideStringField;
  begin
    if ADataModule.ZQCorretor.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := ASize;
      LField.DataSet := ADataModule.ZQCorretor;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQCorretor = nil) then
    Exit;
  if (ADataModule.ZQCorretor.FindField('nomecidade') <> nil) and
     (ADataModule.ZQCorretor.FindField('estado') <> nil) then
    Exit;

  { Corretor usa cidade e estado como lookups, mas a consulta principal não
    precisa retornar esses nomes como colunas físicas. }
  LWasActive := ADataModule.ZQCorretor.Active;
  if LWasActive then
    ADataModule.ZQCorretor.Close;
  try
    EnsureCalculatedField('ZQCorretornomecidade', 'nomecidade', 100);
    EnsureCalculatedField('ZQCorretorestado', 'estado', 2);
  finally
    if LWasActive then
      ADataModule.ZQCorretor.Open;
  end;
end;

procedure EnsureContaBancariaLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;
  LField: TWideStringField;
begin
  if (ADataModule = nil) or (ADataModule.ZQContaBancaria = nil) then
    Exit;
  if ADataModule.ZQContaBancaria.FindField('nomeparticip') <> nil then
    Exit;

  { nomeparticip é somente a apresentação do participante relacionado; não é
    uma coluna física de conta_bancaria. }
  LWasActive := ADataModule.ZQContaBancaria.Active;
  if LWasActive then
    ADataModule.ZQContaBancaria.Close;
  try
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := 'ZQContaBancarianomeparticip';
      LField.FieldName := 'nomeparticip';
      LField.FieldKind := fkCalculated;
      LField.Size := 100;
      LField.DataSet := ADataModule.ZQContaBancaria;
    except
      LField.Free;
      raise;
    end;
  finally
    if LWasActive then
      ADataModule.ZQContaBancaria.Open;
  end;
end;

procedure EnsureVendaLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureCalculatedField(const AComponentName, AFieldName: string;
    AFieldClass: TFieldClass; ASize: Integer);
  var
    LField: TField;
  begin
    if ADataModule.ZQVenda.FindField(AFieldName) <> nil then
      Exit;
    LField := AFieldClass.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      if (ASize > 0) and (LField is TStringField) then
        TStringField(LField).Size := ASize;
      LField.DataSet := ADataModule.ZQVenda;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQVenda = nil) then
    Exit;
  if ADataModule.ZQVenda.FindField('quadra') <> nil then
    Exit;

  { Estes campos são lookups de imovel, loteamento e cidade. Eles não são
    colunas físicas da consulta principal de venda. }
  LWasActive := ADataModule.ZQVenda.Active;
  if LWasActive then
    ADataModule.ZQVenda.Close;
  try
    EnsureCalculatedField('ZQVendaquadra', 'quadra', TWideStringField, 25);
    EnsureCalculatedField('ZQVendalote', 'lote', TWideStringField, 4);
    EnsureCalculatedField('ZQVendacodloteamento', 'codloteamento', TIntegerField, 0);
    EnsureCalculatedField('ZQVendanometoeam', 'nometoeam', TWideStringField, 100);
    EnsureCalculatedField('ZQVendaarea', 'area', TFloatField, 0);
    EnsureCalculatedField('ZQVendaruafrente', 'ruafrente', TWideStringField, 100);
    EnsureCalculatedField('ZQVendadedfrente', 'medfrente', TWideStringField, 0);
    EnsureCalculatedField('ZQVendamedfundo', 'medfundo', TWideStringField, 0);
    EnsureCalculatedField('ZQVendamedesquerda', 'medesquerda', TWideStringField, 0);
    EnsureCalculatedField('ZQVendameddireita', 'meddireita', TWideStringField, 0);
    EnsureCalculatedField('ZQVendavalorvenal', 'valorvenal', TFloatField, 0);
    EnsureCalculatedField('ZQVendaconf_frente', 'conf_frente', TWideStringField, 100);
    EnsureCalculatedField('ZQVendaconf_fundo', 'conf_fundo', TWideStringField, 100);
    EnsureCalculatedField('ZQVendaconf_esquerda', 'conf_esquerda', TWideStringField, 100);
    EnsureCalculatedField('ZQVendaconf_direita', 'conf_direita', TWideStringField, 100);
    EnsureCalculatedField('ZQVendalado', 'lado', TWideStringField, 0);
    EnsureCalculatedField('ZQVendaesquina', 'esquina', TWideStringField, 100);
    EnsureCalculatedField('ZQVendalinha', 'linha', TWideStringField, 0);
    EnsureCalculatedField('ZQVendacurva', 'curva', TWideStringField, 0);
    EnsureCalculatedField('ZQVendadisponivel', 'disponivel', TWideStringField, 3);
    EnsureCalculatedField('ZQVendamatricula', 'matricula', TWideStringField, 10);
    EnsureCalculatedField('ZQVendaimovel_matr', 'imovel_matr', TWideStringField, 10);
    EnsureCalculatedField('ZQVendaBairro', 'Bairro', TWideStringField, 80);
    EnsureCalculatedField('ZQVendaCidadeIdLoteamento',
      'cidade_idcidade_loteamento', TLargeintField, 0);
    EnsureCalculatedField('ZQVendaCidade', 'Cidade', TWideStringField, 150);
    EnsureCalculatedField('ZQVendauf', 'uf', TWideStringField, 2);
    EnsureCalculatedField('ZQVendaNomeEmpreend', 'NomeEmpreend', TWideStringField, 50);
    EnsureCalculatedField('ZQVendaApelido', 'Apelido', TWideStringField, 50);
  finally
    if LWasActive then
      ADataModule.ZQVenda.Open;
  end;
end;

procedure EnsureCompradorLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;
  LField: TWideStringField;
begin
  if (ADataModule = nil) or (ADataModule.ZQComprador = nil) then
    Exit;
  if ADataModule.ZQComprador.FindField('nomecomprador') <> nil then
    Exit;

  { nomecomprador é o nome do participante relacionado, não uma coluna física
    da consulta de comprador. }
  LWasActive := ADataModule.ZQComprador.Active;
  if LWasActive then
    ADataModule.ZQComprador.Close;
  try
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := 'ZQCompradornomecomprador';
      LField.FieldName := 'nomecomprador';
      LField.FieldKind := fkCalculated;
      LField.Size := 140;
      LField.DataSet := ADataModule.ZQComprador;
    except
      LField.Free;
      raise;
    end;
  finally
    if LWasActive then
      ADataModule.ZQComprador.Open;
  end;
end;

procedure EnsureRecebimentoCalculatedFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;
  LField: TFloatField;
begin
  if (ADataModule = nil) or (ADataModule.ZQRecebimento = nil) then
    Exit;
  if ADataModule.ZQRecebimento.FindField('saldocalc') <> nil then
    Exit;

  { saldocalc é um valor calculado usado pela tela/relatórios; não existe na
    tabela recebimento. }
  LWasActive := ADataModule.ZQRecebimento.Active;
  if LWasActive then
    ADataModule.ZQRecebimento.Close;
  try
    LField := TFloatField.Create(ADataModule);
    try
      LField.Name := 'ZQRecebimentosaldocalc';
      LField.FieldName := 'saldocalc';
      LField.FieldKind := fkCalculated;
      LField.DisplayFormat := '#,###,##0.00';
      LField.DataSet := ADataModule.ZQRecebimento;
    except
      LField.Free;
      raise;
    end;
  finally
    if LWasActive then
      ADataModule.ZQRecebimento.Open;
  end;
end;

procedure EnsureRecebimentoTempAggregateField(ADataModule: TDM_Tabelas);
var
  LField: TFloatField;
begin
  if (ADataModule = nil) or (ADataModule.ZQRecebimento_temp = nil) then
    Exit;
  if ADataModule.ZQRecebimento_temp.FindField('somavalor') <> nil then
    Exit;

  { somavalor é o alias do SUM(valor) usado pela rotina de parcelas da venda.
    A query é reutilizada com SQL montado em tempo de execução, portanto o
    campo precisa existir antes de a consulta ser aberta. }
  LField := TFloatField.Create(ADataModule);
  try
    LField.Name := 'ZQRecebimento_tempsomavalor';
    LField.FieldName := 'somavalor';
    LField.FieldKind := fkData;
    LField.DisplayFormat := '#,###,##0.00';
    LField.DataSet := ADataModule.ZQRecebimento_temp;
  except
    LField.Free;
    raise;
  end;
end;

procedure EnsureRelVendaPartiFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;
  LField: TWideStringField;
begin
  if (ADataModule = nil) or (ADataModule.ZQRelVendaParti = nil) then
    Exit;

  { A base atual não possui participante.apelido. O relatório ainda espera
    esse nome, então o resultado precisa expô-lo como alias de nome_parte. }
  LWasActive := ADataModule.ZQRelVendaParti.Active;
  if LWasActive then
    ADataModule.ZQRelVendaParti.Close;
  try
    if Pos(' as apelido', LowerCase(ADataModule.ZQRelVendaParti.SQL.Text)) = 0 then
    begin
      ADataModule.ZQRelVendaParti.SQL.Text :=
        'select comprador.*, participante.*, participante.nome_parte as apelido ' +
        'from comprador join participante on ' +
        'paticipante_idpaticipante=idpaticipante';
    end;

    if ADataModule.ZQRelVendaParti.FindField('apelido') = nil then
    begin
      LField := TWideStringField.Create(ADataModule);
      try
        LField.Name := 'ZQRelVendaPartiapelido';
        LField.FieldName := 'apelido';
        LField.FieldKind := fkData;
        LField.Size := 100;
        LField.DataSet := ADataModule.ZQRelVendaParti;
      except
        LField.Free;
        raise;
      end;
    end;
  finally
    if LWasActive then
      ADataModule.ZQRelVendaParti.Open;
  end;
end;

procedure EnsureRecebimentoLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureLookupField(const AComponentName, AFieldName: string;
    ASize: Integer);
  var
    LField: TWideStringField;
  begin
    if ADataModule.ZQRecebimento.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := ASize;
      LField.DataSet := ADataModule.ZQRecebimento;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQRecebimento = nil) then
    Exit;
  if (ADataModule.ZQRecebimento.FindField('nomecli') <> nil) and
     (ADataModule.ZQRecebimento.FindField('adversanome') <> nil) and
     (ADataModule.ZQRecebimento.FindField('nome_loteamento') <> nil) then
    Exit;

  { Estes nomes são lookups de participante e loteamento; não são colunas
    físicas da consulta de recebimento. }
  LWasActive := ADataModule.ZQRecebimento.Active;
  if LWasActive then
    ADataModule.ZQRecebimento.Close;
  try
    EnsureLookupField('ZQRecebimentonomecli', 'nomecli', 100);
    EnsureLookupField('ZQRecebimentoadversanome', 'adversanome', 100);
    EnsureLookupField('ZQRecebimentonome_loteamento', 'nome_loteamento', 100);
  finally
    if LWasActive then
      ADataModule.ZQRecebimento.Open;
  end;
end;

procedure EnsureRecebBxTempCalculatedFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureCalculatedField(const AComponentName, AFieldName: string);
  var
    LField: TFloatField;
  begin
    if ADataModule.ZQRecebBxTemp.FindField(AFieldName) <> nil then
      Exit;
    LField := TFloatField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.DisplayFormat := '#,###,##0.00';
      LField.DataSet := ADataModule.ZQRecebBxTemp;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQRecebBxTemp = nil) then
    Exit;
  if (ADataModule.ZQRecebBxTemp.FindField('vrusado') <> nil) and
     (ADataModule.ZQRecebBxTemp.FindField('jurusado') <> nil) and
     (ADataModule.ZQRecebBxTemp.FindField('descusado') <> nil) then
    Exit;

  { Estes valores são calculados no evento OnCalcFields e não pertencem ao
    SELECT da consulta temporária de baixa. }
  LWasActive := ADataModule.ZQRecebBxTemp.Active;
  if LWasActive then
    ADataModule.ZQRecebBxTemp.Close;
  try
    EnsureCalculatedField('ZQRecebBxTempvrusado', 'vrusado');
    EnsureCalculatedField('ZQRecebBxTempjurusado', 'jurusado');
    EnsureCalculatedField('ZQRecebBxTempdescusado', 'descusado');
  finally
    if LWasActive then
      ADataModule.ZQRecebBxTemp.Open;
  end;
end;

procedure EnsureDivididoLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureLookupField(const AComponentName, AFieldName: string);
  var
    LField: TWideStringField;
  begin
    if ADataModule.ZQDividido.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := 50;
      LField.DataSet := ADataModule.ZQDividido;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQDividido = nil) then
    Exit;
  if (ADataModule.ZQDividido.FindField('descrPlano') <> nil) and
     (ADataModule.ZQDividido.FindField('nomeparticp') <> nil) then
    Exit;

  { São lookups usados na tela de divisão; não são colunas físicas do
    SELECT da tabela dividido. }
  LWasActive := ADataModule.ZQDividido.Active;
  if LWasActive then
    ADataModule.ZQDividido.Close;
  try
    EnsureLookupField('ZQDivididodescrPlano', 'descrPlano');
    EnsureLookupField('ZQDivididonomeparticp', 'nomeparticp');
  finally
    if LWasActive then
      ADataModule.ZQDividido.Open;
  end;
end;

procedure EnsureEmpresaLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureCalculatedField(const AComponentName, AFieldName: string;
    ASize: Integer);
  var
    LField: TWideStringField;
  begin
    if ADataModule.ZQEmpresa.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := ASize;
      LField.DataSet := ADataModule.ZQEmpresa;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQEmpresa = nil) then
    Exit;
  if (ADataModule.ZQEmpresa.FindField('nomecidade') <> nil) and
     (ADataModule.ZQEmpresa.FindField('estado') <> nil) then
    Exit;

  { Empresa usa estes dois campos de lookup em eventos e relatórios, mas a
    consulta física não os retorna. Garanta a existência antes dos forms. }
  LWasActive := ADataModule.ZQEmpresa.Active;
  if LWasActive then
    ADataModule.ZQEmpresa.Close;
  try
    EnsureCalculatedField('ZQEmpresanomecidade', 'nomecidade', 100);
    EnsureCalculatedField('ZQEmpresaestado', 'estado', 2);
  finally
    if LWasActive then
      ADataModule.ZQEmpresa.Open;
  end;
end;

procedure EnsureParticipanteLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureCalculatedField(const AComponentName, AFieldName: string;
    ASize: Integer);
  var
    LField: TWideStringField;
  begin
    if ADataModule.ZqParticipante.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := ASize;
      LField.DataSet := ADataModule.ZqParticipante;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZqParticipante = nil) then
    Exit;
  if (ADataModule.ZqParticipante.FindField('nomecidade') <> nil) and
     (ADataModule.ZqParticipante.FindField('nomeestado') <> nil) and
     (ADataModule.ZqParticipante.FindField('nomecidadecob') <> nil) and
     (ADataModule.ZqParticipante.FindField('nomeestadocob') <> nil) and
     (ADataModule.ZqParticipante.FindField('nomecidadenatural') <> nil) and
     (ADataModule.ZqParticipante.FindField('nomeestadonatural') <> nil) then
    Exit;

  { Os dados de cidade ficam em lookups e não fazem parte do SELECT físico
    de participante. Garanta os campos antes de qualquer form ou relatório. }
  LWasActive := ADataModule.ZqParticipante.Active;
  if LWasActive then
    ADataModule.ZqParticipante.Close;
  try
    EnsureCalculatedField('ZqParticipantenomecidade', 'nomecidade', 100);
    EnsureCalculatedField('ZqParticipantenomeestado', 'nomeestado', 2);
    EnsureCalculatedField('ZqParticipantenomecidadecob', 'nomecidadecob', 100);
    EnsureCalculatedField('ZqParticipantenomeestadocob', 'nomeestadocob', 2);
    EnsureCalculatedField('ZqParticipantenomecidadenatural', 'nomecidadenatural', 100);
    EnsureCalculatedField('ZqParticipantenomeestadonatural', 'nomeestadonatural', 2);
  finally
    if LWasActive then
      ADataModule.ZqParticipante.Open;
  end;
end;

procedure EnsureResponsavelLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureCalculatedField(const AComponentName, AFieldName: string;
    ASize: Integer);
  var
    LField: TWideStringField;
  begin
    if ADataModule.ZQresponsavel.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := ASize;
      LField.DataSet := ADataModule.ZQresponsavel;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQresponsavel = nil) then
    Exit;
  if (ADataModule.ZQresponsavel.FindField('nomecidadecart') <> nil) and
     (ADataModule.ZQresponsavel.FindField('estcidadecart') <> nil) and
     (ADataModule.ZQresponsavel.FindField('nomecomarca') <> nil) and
     (ADataModule.ZQresponsavel.FindField('estcomarca') <> nil) and
     (ADataModule.ZQresponsavel.FindField('nomecidade') <> nil) and
     (ADataModule.ZQresponsavel.FindField('nomeesta') <> nil) then
    Exit;

  { Os nomes das cidades e dos estados são lookups calculados; não fazem
    parte do SELECT físico de representa_participante. }
  LWasActive := ADataModule.ZQresponsavel.Active;
  if LWasActive then
    ADataModule.ZQresponsavel.Close;
  try
    EnsureCalculatedField('ZQresponsavelnomecidadecart', 'nomecidadecart', 100);
    EnsureCalculatedField('ZQresponsavelestcidadecart', 'estcidadecart', 2);
    EnsureCalculatedField('ZQresponsavelnomecomarca', 'nomecomarca', 100);
    EnsureCalculatedField('ZQresponsavelestcomarca', 'estcomarca', 2);
    EnsureCalculatedField('ZQresponsavelnomecidade', 'nomecidade', 100);
    EnsureCalculatedField('ZQresponsavelnomeesta', 'nomeesta', 2);
  finally
    if LWasActive then
      ADataModule.ZQresponsavel.Open;
  end;
end;

procedure EnsureIncorpLoteamentoLookupFields(ADataModule: TDM_Tabelas);
var
  LWasActive: Boolean;

  procedure EnsureCalculatedField(const AComponentName, AFieldName: string;
    ASize: Integer);
  var
    LField: TWideStringField;
  begin
    if ADataModule.ZQincorp_loteame.FindField(AFieldName) <> nil then
      Exit;
    LField := TWideStringField.Create(ADataModule);
    try
      LField.Name := AComponentName;
      LField.FieldName := AFieldName;
      LField.FieldKind := fkCalculated;
      LField.Size := ASize;
      LField.DataSet := ADataModule.ZQincorp_loteame;
    except
      LField.Free;
      raise;
    end;
  end;
begin
  if (ADataModule = nil) or (ADataModule.ZQincorp_loteame = nil) then
    Exit;
  if (ADataModule.ZQincorp_loteame.FindField('nconta') <> nil) and
     (ADataModule.ZQincorp_loteame.FindField('conta_v') <> nil) and
     (ADataModule.ZQincorp_loteame.FindField('agencia') <> nil) and
     (ADataModule.ZQincorp_loteame.FindField('agencia_v') <> nil) and
     (ADataModule.ZQincorp_loteame.FindField('n_banco') <> nil) then
    Exit;

  { Os dados da conta bancária são exibidos por lookup no cadastro de
    incorporadores, mas não fazem parte do SELECT físico desta consulta. }
  LWasActive := ADataModule.ZQincorp_loteame.Active;
  if LWasActive then
    ADataModule.ZQincorp_loteame.Close;
  try
    EnsureCalculatedField('ZQincorp_loteamenconta', 'nconta', 50);
    EnsureCalculatedField('ZQincorp_loteameconta_v', 'conta_v', 3);
    EnsureCalculatedField('ZQincorp_loteameagencia', 'agencia', 6);
    EnsureCalculatedField('ZQincorp_loteameagencia_v', 'agencia_v', 3);
    EnsureCalculatedField('ZQincorp_loteamen_banco', 'n_banco', 3);
  finally
    if LWasActive then
      ADataModule.ZQincorp_loteame.Open;
  end;
end;

procedure TDM_Tabelas.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureLoteamentoNomeCidadeField(Self);
  EnsureImovelLoteamentoField(Self);
  EnsureCorretorLookupFields(Self);
  EnsureContaBancariaLookupFields(Self);
  ZQVenda.BeforeOpen := ZQVendaBeforeOpen;
  EnsureVendaLookupFields(Self);
  ZQComprador.BeforeOpen := ZQCompradorBeforeOpen;
  EnsureCompradorLookupFields(Self);
  ZQRecebimento.BeforeOpen := ZQRecebimentoBeforeOpen;
  EnsureRecebimentoCalculatedFields(Self);
  EnsureRecebimentoTempAggregateField(Self);
  EnsureRecebimentoLookupFields(Self);
  ZQRecebBxTemp.BeforeOpen := ZQRecebBxTempBeforeOpen;
  EnsureRecebBxTempCalculatedFields(Self);
  ZQDividido.BeforeOpen := ZQDivididoBeforeOpen;
  EnsureDivididoLookupFields(Self);
  EnsureEmpresaLookupFields(Self);
  EnsureParticipanteLookupFields(Self);
  EnsureResponsavelLookupFields(Self);
  EnsureIncorpLoteamentoLookupFields(Self);
  EnsureRelVendaPartiFields(Self);
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQUsuario', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuarioidusuario', 'idusuario', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuarionome', 'nome', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuariosenha', 'senha', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuarioadimitido', 'adimitido', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuariodemitido', 'demitido', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuariocargo', 'cargo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuarioendereco', 'endereco', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuariobairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuariocidade', 'cidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuariofone1', 'fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '(99)9999-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuariofone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '(99)9999-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuario', 'ZQUsuarioapelido', 'apelido', TWideStringField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQLogin', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLogin', 'ZQLoginidlogin_historico', 'idlogin_historico', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLogin', 'ZQLoginlocal_senha', 'local_senha', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLogin', 'ZQLogindata_2', 'data_2', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLogin', 'ZQLoginhora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLogin', 'ZQLoginmaquina', 'maquina', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLogin', 'ZQLoginusuario', 'usuario', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLogin', 'ZQLoginpassou', 'passou', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLogin', 'ZQLogindetalhe', 'detalhe', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQLoc_Senha', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoc_Senha', 'ZQLoc_Senhaidlocais_com_senha', 'idlocais_com_senha', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoc_Senha', 'ZQLoc_Senhagrupo', 'grupo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoc_Senha', 'ZQLoc_Senhadescricao', 'descricao', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoc_Senha', 'ZQLoc_Senhapedesenha', 'pedesenha', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQPermissoes', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPermissoes', 'ZQPermissoesidpermissoes', 'idpermissoes', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPermissoes', 'ZQPermissoesusuario', 'usuario', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPermissoes', 'ZQPermissoeslocal_senha', 'local_senha', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQCidade', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCidade', 'ZQCidadeidcidade', 'idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCidade', 'ZQCidadenomecid', 'nomecid', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCidade', 'ZQCidadeddd', 'ddd', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCidade', 'ZQCidadeemancipacao', 'emancipacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCidade', 'ZQCidadeestado', 'estado', TWideStringField, fkData, 2, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCidade', 'ZQCidadecepgeral', 'cepgeral', TWideStringField, fkData, 10, 0, False, '', '', '', '99\.999\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQUsuTemp', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempidusuario', 'idusuario', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempnome', 'nome', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempsenha', 'senha', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempadimitido', 'adimitido', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempdemitido', 'demitido', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempcargo', 'cargo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempendereco', 'endereco', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempcidade', 'cidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempfone1', 'fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQUsuTemp', 'ZQUsuTempapelido', 'apelido', TWideStringField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAgrupaGrupo', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAgrupaGrupo', 'ZQAgrupaGrupoidlocais_com_senha', 'idlocais_com_senha', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAgrupaGrupo', 'ZQAgrupaGrupogrupo', 'grupo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAgrupaGrupo', 'ZQAgrupaGrupodescricao', 'descricao', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAgrupaGrupo', 'ZQAgrupaGrupopedesenha', 'pedesenha', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAchaCidade', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCidade', 'ZQAchaCidadeidcidade', 'idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCidade', 'ZQAchaCidadenomecid', 'nomecid', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCidade', 'ZQAchaCidadeddd', 'ddd', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCidade', 'ZQAchaCidadeemancipacao', 'emancipacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCidade', 'ZQAchaCidadeestado', 'estado', TWideStringField, fkData, 2, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCidade', 'ZQAchaCidadecepgeral', 'cepgeral', TWideStringField, fkData, 10, 0, False, '', '', '', '00\-000\.999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQLoteamento', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentoidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentocidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentomatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentodatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentodatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentodataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentologotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentomapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentopasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentotestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentotestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentotestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentotestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentoObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentobairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentonomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_idcidade', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentoestado', 'estado', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_idcidade', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentotipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentonomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentoapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQLoteamento', 'ZQLoteamentoPerc_comissao', 'Perc_comissao', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQEmpresa', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresacodigo', 'codigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresarazao', 'razao', TWideStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresafantasia', 'fantasia', TWideStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresaendereco', 'endereco', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresabairro', 'bairro', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresacidade', 'cidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresacep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresainalgurada', 'inalgurada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresaie', 'ie', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresacnpj', 'cnpj', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresafone1', 'fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresafone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresafone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresaemail', 'email', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresacad_empresa', 'cad_empresa', TDateTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresalogo', 'logo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresaativa', 'ativa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresatipo', 'tipo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresafilial', 'filial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresalimite', 'limite', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresadia', 'dia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresatentativa', 'tentativa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresanomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresaestado', 'estado', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmpresa', 'ZQEmpresadt_importacao', 'dt_importacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZqParticipante', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantenome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantedoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantedoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantebairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '99\.999\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantebairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '99\.999\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)99999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantefone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)99999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantefone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)99999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantetipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantenacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantenomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade', 'ZQCidade', 'idcidade', 'nomecid', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantenomeestado', 'nomeestado', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade', 'ZQCidade', 'idcidade', 'estado', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantenomecidadecob', 'nomecidadecob', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cob', 'ZQCidade', 'idcidade', 'nomecid', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantenomeestadocob', 'nomeestadocob', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cob', 'ZQCidade', 'idcidade', 'estado', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantenaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantelocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantenomecidadenatural', 'nomecidadenatural', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'naturalidade', 'ZQCidade', 'idcidade', 'nomecid', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantenomeestadonatural', 'nomeestadonatural', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'naturalidade', 'ZQCidade', 'idcidade', 'estado', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteESTADOCIVIL', 'ESTADOCIVIL', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanterenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantefone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)99999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantesexo', 'sexo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantecex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantejb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantejudicial', 'judicial', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantexquadra', 'xquadra', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantexlote', 'xlote', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipanteplantao', 'plantao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante', 'ZqParticipantehistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAchaParticip', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '00\-000\.999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '00\-000\.999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticiptipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticiplocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipESTADOCIVIL', 'ESTADOCIVIL', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticiprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaParticip', 'ZQAchaParticipcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQresponsavel', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelidrepresenta_participante', 'idrepresenta_participante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsaveldocumen1', 'documen1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsaveldocumen2', 'documen2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelnasicmento', 'nasicmento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelcid_natural', 'cid_natural', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelnomerepres', 'nomerepres', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelprofiss', 'profiss', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsaveltrabalho', 'trabalho', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsaveladmissao', 'admissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelnomeconjuge', 'nomeconjuge', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsaveldoc1conjuge', 'doc1conjuge', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsaveldoc2conjuge', 'doc2conjuge', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelnascconjuge', 'nascconjuge', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelassconjuge', 'assconjuge', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelcidconjuge', 'cidconjuge', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsaveldata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelcartorio', 'cartorio', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavellivro', 'livro', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelfolha', 'folha', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelcidade_cart', 'cidade_cart', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelcomarca', 'comarca', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelnomecidadecart', 'nomecidadecart', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelestcidadecart', 'estcidadecart', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelnomecomarca', 'nomecomarca', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelestcomarca', 'estcomarca', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelnomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cid_natural', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelnomeesta', 'nomeesta', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cid_natural', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel', 'ZQresponsavelemailrepres', 'emailrepres', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQincorp_loteame', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteameincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteameloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteamepercent_parte', 'percent_parte', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteamecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteamedigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteamenconta', 'nconta', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQContaBancaria', 'idconta_bancaria', 'n_conta', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteameconta_v', 'conta_v', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQContaBancaria', 'idconta_bancaria', 'n_conta_v', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteameagencia', 'agencia', TWideStringField, fkLookup, 6, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQContaBancaria', 'idconta_bancaria', 'n_agencia', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteameagencia_v', 'agencia_v', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQContaBancaria', 'idconta_bancaria', 'n_agencia_v', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQincorp_loteame', 'ZQincorp_loteamen_banco', 'n_banco', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'ZQContaBancaria', 'idconta_bancaria', 'n_banco', True);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQImovel', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovellote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovellado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelarea', 'area', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '##,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelcornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelnomeloteamento', 'nomeloteamento', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'loteamento_idloteamento', 'ZQLoteamento', 'idloteamento', 'apelido', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImoveldisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelmedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelmedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelmedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelmedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImoveldescricao', 'descricao', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelmatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelmatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelproposta', 'proposta', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovel', 'ZQImovelvalorCusto', 'valorCusto', TFloatField, fkData, 0, 0, False, '', '##,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAchaLotea', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteaidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteacidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteamatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteadatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteadatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteadataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLotealogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteamapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteatestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteatestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteatestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteatestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteaObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteabairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteapasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteatipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteanomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteacodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaLotea', 'ZQAchaLoteaapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQCorretor', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretoridcorretor', 'idcorretor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretornomecorretor', 'nomecorretor', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretordoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '###\.###\.###\-##;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretorcreci', 'creci', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretortelefone1', 'telefone1', TWideStringField, fkData, 14, 0, False, '', '', '', '(00)-0000.0000;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretortelefone2', 'telefone2', TWideStringField, fkData, 14, 0, False, '', '', '', '(00)-0000.0000;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretorendereco_corr', 'endereco_corr', TWideStringField, fkData, 70, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretorbairro_corr', 'bairro_corr', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretorcep_corr', 'cep_corr', TWideStringField, fkData, 10, 0, False, '', '', '', '00\.000\-000;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretorObser_corr', 'Obser_corr', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretorcodcidade', 'codcidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretornomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'codcidade', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretorestado', 'estado', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'codcidade', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCorretor', 'ZQCorretorPcomissao_corretor', 'Pcomissao_corretor', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQTipoDoc', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDocidtipodocumento', 'idtipodocumento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDoctipodoc', 'tipodoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDocdescricao', 'descricao', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDocvend_receb', 'vend_receb', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDocvend_caixa', 'vend_caixa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDocreceb_receb', 'receb_receb', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDocreceb_caixa', 'receb_caixa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDocdados_chequ', 'dados_chequ', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDocso_avista', 'so_avista', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDoclancabanco', 'lancabanco', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDocsomapaga', 'somapaga', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQTipoDoc', 'ZQTipoDoclancinclus', 'lancinclus', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQPlanoDeContas', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContascodigo', 'codigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContasdoccomum', 'doccomum', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContasmascara', 'mascara', TWideStringField, fkData, 16, 0, False, '', '', '', '9.99.999.9999.99999;0;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContasdescricao', 'descricao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContasclassificacao', 'classificacao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContasvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContascod_reduzido', 'cod_reduzido', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContascred_debi', 'cred_debi', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContasusoativo', 'usoativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContasCDPlano', 'CDPlano', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQPlanoDeContas', 'ZQPlanoDeContasDespFixVar', 'DespFixVar', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQContaBancaria', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariaidconta_bancaria', 'idconta_bancaria', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarian_banco', 'n_banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarian_agencia', 'n_agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarian_agencia_v', 'n_agencia_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarian_conta', 'n_conta', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarian_conta_v', 'n_conta_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarianomebanco', 'nomebanco', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarian_no_banco', 'n_no_banco', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarianomeresposavel', 'nomeresposavel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarianometitular', 'nometitular', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariaaberta', 'aberta', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariasenha_interna', 'senha_interna', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariasenha_conta', 'senha_conta', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariaativa', 'ativa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariajurosemboleto', 'jurosemboleto', TFloatField, fkData, 0, 0, False, '', '###,##0.000', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariataxadiaria', 'taxadiaria', TFloatField, fkData, 0, 0, False, '', '###,##0.000', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariavaloroupercent', 'valoroupercent', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariacarteira', 'carteira', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariaboleto', 'boleto', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariaidparticipante', 'idparticipante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarianomeparticip', 'nomeparticip', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idparticipante', 'ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarian_dif_empreed', 'n_dif_empreed', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariaconvenio', 'convenio', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariadoc_titular', 'doc_titular', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarian_sequencial', 'n_sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarianosso_numero', 'nosso_numero', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariaApelido', 'Apelido', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariacod_transmissao', 'cod_transmissao', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariacomplemento', 'complemento', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariamoramensal', 'moramensal', TFloatField, fkData, 0, 0, False, '', '#.##0.000', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariavariacao', 'variacao', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariadata_maxima_desconto', 'data_maxima_desconto', TDateField, fkData, 0, 0, False, '', '99/99/9999', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancariaPerc_descontos', 'Perc_descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarialayoutversaoArquivo', 'layoutversaoArquivo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQContaBancaria', 'ZQContaBancarialayoutversaoLote', 'layoutversaoLote', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQEspecial', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEspecial', 'ZQEspecialidespecial', 'idespecial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEspecial', 'ZQEspecialdata_valida', 'data_valida', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEspecial', 'ZQEspecialConta_Bancaria_cod_banco', 'Conta_Bancaria_cod_banco', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEspecial', 'ZQEspecialvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQQuadras', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQQuadras', 'ZQQuadrasidquadras', 'idquadras', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQQuadras', 'ZQQuadrasloteamento_idLoteamento', 'loteamento_idLoteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQQuadras', 'ZQQuadrasdescricao', 'descricao', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQQuadras', 'ZQQuadrasapedlido', 'apedlido', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQVenda', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaquadra', 'quadra', TWideStringField, fkLookup, 25, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'quadra', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendalote', 'lote', TWideStringField, fkLookup, 4, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'lote', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendacodloteamento', 'codloteamento', TIntegerField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'loteamento_idloteamento', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendanometoeam', 'nometoeam', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'apelido', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaarea', 'area', TFloatField, fkLookup, 0, 0, False, '', '###,###,##0.00', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'area', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaruafrente', 'ruafrente', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontofrente', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendadedfrente', 'medfrente', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'medidafrente', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendamedfundo', 'medfundo', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'medidafundo', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendamedesquerda', 'medesquerda', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'medidaesquerda', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendameddireita', 'meddireita', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'medidadedireita', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendavalorvenal', 'valorvenal', TFloatField, fkLookup, 0, 0, False, '', '###,###.##0.00', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'valorvenal', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaconf_frente', 'conf_frente', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontofrente', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaconf_fundo', 'conf_fundo', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontofundo', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaconf_esquerda', 'conf_esquerda', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontoesquerda', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaconf_direita', 'conf_direita', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontodireita', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendalado', 'lado', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'lado', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaesquina', 'esquina', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'esquinanome', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendalinha', 'linha', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'esquinalinha', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendacurva', 'curva', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'esquinacurva', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendadisponivel', 'disponivel', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'disponivel', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendamatricula', 'matricula', TWideStringField, fkLookup, 10, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'matriculaloteamento', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaimovel_matr', 'imovel_matr', TWideStringField, fkLookup, 10, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'matricula', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaBairro', 'Bairro', TWideStringField, fkLookup, 80, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'bairro', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaCidadeIdLoteamento', 'cidade_idcidade_loteamento', TLargeintField, fkLookup, 0, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'cidade_idcidade', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaCidade', 'Cidade', TWideStringField, fkLookup, 150, 0, False, '', '', '', '', 0, 'cidade_idcidade_loteamento', 'ZQCidade', 'idcidade', 'nomecid', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendauf', 'uf', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_idcidade_loteamento', 'ZQCidade', 'idcidade', 'estado', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendatabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '##0.0000', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaNomeEmpreend', 'NomeEmpreend', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'nomeloteamento', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaApelido', 'Apelido', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'apelido', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendamarcar', 'marca', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendacodigo_contrato_ref', 'codigo_contrato_ref', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaMulta', 'Multa', TFloatField, fkData, 0, 0, False, '', '###0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaMora', 'Mora', TFloatField, fkData, 0, 0, False, '', '###0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendaPerc_comissao', 'Perc_comissao', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendavlr_comissao', 'vlr_comissao', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVenda', 'ZQVendapath_pdf', 'path_pdf', TWideStringField, fkData, 255, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQImovelTemp', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemploteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemplote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemplado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempmedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempmedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempmedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempmedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemparea', 'area', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempcornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempdisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempcidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempmatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempdatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempdatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempdataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemplogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempmapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemptestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemptestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemptestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemptestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempbairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemppasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempnomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTemptipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempcodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQImovelTemp', 'ZQImovelTempPerc_comissao', 'Perc_comissao', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQComprador', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador', 'ZQCompradoridcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador', 'ZQCompradorpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador', 'ZQCompradorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador', 'ZQCompradorpromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador', 'ZQCompradornomecomprador', 'nomecomprador', TWideStringField, fkLookup, 140, 0, False, '', '', '', '', 0, 'paticipante_idpaticipante', 'ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador', 'ZQCompradorpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador', 'ZQCompradorMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQVendedor', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVendedor', 'ZQVendedoridvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVendedor', 'ZQVendedorcorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVendedor', 'ZQVendedorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVendedor', 'ZQVendedorPcomissao_Corretor', 'Pcomissao_Corretor', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAchaCorretor', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretoridcorretor', 'idcorretor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretornomecorretor', 'nomecorretor', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretordoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '###\.###\.###\-##;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretorcreci', 'creci', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretortelefone1', 'telefone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretortelefone2', 'telefone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretorendereco_corr', 'endereco_corr', TWideStringField, fkData, 70, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretorbairro_corr', 'bairro_corr', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretorcep_corr', 'cep_corr', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretorObser_corr', 'Obser_corr', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretorcodcidade', 'codcidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaCorretor', 'ZQAchaCorretorPComissao_Corretor', 'PComissao_Corretor', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRecebimento', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentosaldocalc', 'saldocalc', TFloatField, fkCalculated, 0, 0, False, '', '#,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'ZqParticipante', 'idpaticipante', 'nome_parte', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'ZqParticipante', 'idpaticipante', 'nome_parte', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'ZQLoteamento', 'idloteamento', 'apelido', True, True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoobserv_estorno', 'observ_estorno', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentotip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentosld_antes_reajuste', 'sld_antes_reajuste', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentoPercentual_reajuste', 'Percentual_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentojuridico', 'juridico', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentodata_juridico', 'data_juridico', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentodt_nao_pagou_no_mes', 'dt_nao_pagou_no_mes', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento', 'ZQRecebimentodescricao_juridico', 'descricao_juridico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQReceb_Baixa', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_Baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_BaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_BaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_BaixaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_BaixaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_BaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_BaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_BaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_Baixaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_BaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_Baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_Baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_Baixasubstituicao', 'substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_Baixavr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_Baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReceb_Baixa', 'ZQReceb_Baixavalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQCobaRe', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCobaRe', 'ZQCobaRecotagem', 'cotagem', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCobaRe', 'ZQCobaRehoje', 'hoje', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCobaRe', 'ZQCobaReentrada', 'entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCobaRe', 'ZQCobaRetipobaixa', 'tipobaixa', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQReBxHi', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHiidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHirefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHiidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHivalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHidescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHijuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHipercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHidata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQReBxHi', 'ZQReBxHivalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRecebBxTemp', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTemprefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTemppercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTemprefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTemporigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTemprecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempvrusado', 'vrusado', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempjurusado', 'jurusado', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempdescusado', 'descusado', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebBxTemp', 'ZQRecebBxTempvalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQNumOrdem', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQNumOrdem', 'ZQNumOrdemidnumordem', 'idnumordem', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQCheque', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeBanco', 'Banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeDono', 'Dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequevalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequedeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequealias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequenumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequeSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque', 'ZQChequesq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQConfiguracoes', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesidconfiguracoes', 'idconfiguracoes', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesvenda_planodecontas', 'venda_planodecontas', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesjurosmensal', 'jurosmensal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoespaticipantedefault', 'paticipantedefault', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesdiasdeatrazo', 'diasdeatrazo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesrecebe_planodecontas', 'recebe_planodecontas', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesbxautodoc', 'bxautodoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesbxautoplc', 'bxautoplc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesprice', 'price', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesBaixa', 'Baixa', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesMulta', 'Multa', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesMora', 'Mora', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesTipo_reajuste', 'Tipo_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesresolucao_tela', 'resolucao_tela', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConfiguracoes', 'ZQConfiguracoesmesmodiames', 'mesmodiames', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAchaPlanoDeContas', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContascodigo', 'codigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContasdoccomum', 'doccomum', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContasmascara', 'mascara', TWideStringField, fkData, 16, 0, False, '', '', '', '9.99.999.9999.99999;0;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContasdescricao', 'descricao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContasclassificacao', 'classificacao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContasvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContascod_reduzido', 'cod_reduzido', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContascred_debi', 'cred_debi', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContasusoativo', 'usoativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContasCDPlano', 'CDPlano', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaPlanoDeContas', 'ZQAchaPlanoDeContasDespFixVar', 'DespFixVar', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQDividido', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQDividido', 'ZQDivididoiddividido', 'iddividido', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQDividido', 'ZQDivididoparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQDividido', 'ZQDivididopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQDividido', 'ZQDivididoplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQDividido', 'ZQDivididotipodocumento', 'tipodocumento', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQDividido', 'ZQDivididoordem', 'ordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQDividido', 'ZQDivididodescrPlano', 'descrPlano', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'planodeconta', 'ZQPlanoDeContas', 'codigo', 'descricao', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQDividido', 'ZQDivididonomeparticp', 'nomeparticp', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'participante', '', 'codigo', 'nome', True);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQCaixa', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixacodigo', 'codigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixaplano_contas_codigo', 'plano_contas_codigo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixadata_lan', 'data_lan', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixavr_lan', 'vr_lan', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixaobs', 'obs', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixadocumento', 'documento', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixavinculo', 'vinculo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixacredeb', 'credeb', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixacontraquem', 'contraquem', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixacontracodigo', 'contracodigo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCaixa', 'ZQCaixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQMovBancaria', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancarialancamento', 'lancamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariaconta_bancaria_cod_banco', 'conta_bancaria_cod_banco', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariaplano_contas_codigo', 'plano_contas_codigo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariadocumento', 'documento', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariadt_lanc', 'dt_lanc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariahora_lanc', 'hora_lanc', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariadt_conciliado', 'dt_conciliado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariavr_lanc', 'vr_lanc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariacredeb', 'credeb', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariasaldo_lanc', 'saldo_lanc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariasaldo_conci', 'saldo_conci', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariaobs', 'obs', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariacontraquem', 'contraquem', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariacontracodigo', 'contracodigo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariavinculo', 'vinculo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMovBancaria', 'ZQMovBancariasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRecebimento_temp', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_temprefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_temporigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_temprecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempsomavalor', 'somavalor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_temp', 'ZQRecebimento_tempnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAchaContaBanc', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancidconta_bancaria', 'idconta_bancaria', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancn_banco', 'n_banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancn_agencia', 'n_agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancn_agencia_v', 'n_agencia_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancn_conta', 'n_conta', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancn_conta_v', 'n_conta_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancnomebanco', 'nomebanco', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancn_no_banco', 'n_no_banco', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancnomeresposavel', 'nomeresposavel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancnometitular', 'nometitular', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancaberta', 'aberta', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancsenha_interna', 'senha_interna', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancsenha_conta', 'senha_conta', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancativa', 'ativa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancjurosemboleto', 'jurosemboleto', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanctaxadiaria', 'taxadiaria', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancvaloroupercent', 'valoroupercent', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccarteira', 'carteira', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancboleto', 'boleto', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancidparticipante', 'idparticipante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanctipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanclocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancn_dif_empreed', 'n_dif_empreed', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBanccadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancApelido', 'Apelido', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancdata_maxima_desconto', 'data_maxima_desconto', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaContaBanc', 'ZQAchaContaBancPerc_descontos', 'Perc_descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQBol_men', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBol_men', 'ZQBol_menidboleto_Mens', 'idboleto_Mens', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBol_men', 'ZQBol_menlinha1', 'linha1', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBol_men', 'ZQBol_menlinha2', 'linha2', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBol_men', 'ZQBol_menlinha3', 'linha3', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBol_men', 'ZQBol_menlinha4', 'linha4', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBol_men', 'ZQBol_menidordem', 'idordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRemes_Receb', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebnossonumero', 'nossonumero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebdt_entrada', 'dt_entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebdt_vencimento', 'dt_vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb', 'ZQRemes_Recebidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQBancRemes', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes', 'ZQBancRemesidbanco_remessa', 'idbanco_remessa', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes', 'ZQBancRemesremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes', 'ZQBancRemesidbanco', 'idbanco', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes', 'ZQBancRemesgerado', 'gerado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes', 'ZQBancRemesnomearq', 'nomearq', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes', 'ZQBancRemesarquivo', 'arquivo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes', 'ZQBancRemesdigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes', 'ZQBancRemesremessa2', 'remessa2', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRecebParticp', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticprefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticporigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticprecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticptipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticplocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpidcidade', 'idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpnomecid', 'nomecid', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebParticp', 'ZQRecebParticpestado', 'estado', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRemesRec', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesRecidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesRecremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesRecidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesRecretorno_motivo', 'retorno_motivo', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesRecacao', 'acao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesRecnossonumero', 'nossonumero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesRecocorrencia', 'ocorrencia', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesRecdesc_motivo', 'desc_motivo', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesReccredito', 'credito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemesRec', 'ZQRemesRecdt_ocorrencia', 'dt_ocorrencia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRelVenda', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaidvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendacorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendavenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendapaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendavenda_idvenda_1', 'venda_idvenda_1', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendapromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendalote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendalado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendamedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendamedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendamedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendamedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendavalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendacornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendadisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendacidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendamatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendadatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendadatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendadataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendalogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendamapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendatestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendatestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendatestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendatestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendaObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendabairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendapasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendanomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendatipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVenda', 'ZQRelVendacodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRelVendaParti', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartipaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartivenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartipromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartinome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiapelido', 'apelido', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartidoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartidoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartibairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaParticidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaParticep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartibairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaParticidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaParticep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartifone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartifone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartitipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartianiversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartinacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartinaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartilocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartirenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaPartiobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaParticodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaParticodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaParticodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaParti', 'ZQRelVendaParticodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRelVendaCorr', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorridvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrcorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorridcorretor', 'idcorretor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrnomecorretor', 'nomecorretor', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrdoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrcreci', 'creci', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrtelefone1', 'telefone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrtelefone2', 'telefone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrendereco_corr', 'endereco_corr', TWideStringField, fkData, 70, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrbairro_corr', 'bairro_corr', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrcep_corr', 'cep_corr', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrObser_corr', 'Obser_corr', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRelVendaCorr', 'ZQRelVendaCorrcodcidade', 'codcidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQProfissao', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProfissao', 'ZQProfissaoidProfissao', 'idProfissao', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProfissao', 'ZQProfissaoprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProfissao', 'ZQProfissaotipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQConjuge', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugeidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugenome', 'nome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugedoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugedoc2', 'doc2', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugenasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugenatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugeemail', 'email', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugeassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugeempresa', 'empresa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugecargo', 'cargo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugeadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugerenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugenomecida', 'nomecida', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'natural', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugeestcida', 'estcida', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'natural', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugeidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugeregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugecomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugecidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugedata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugefolha', 'folha', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugelivro', 'livro', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugecomar_cidnome', 'comar_cidnome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugecomar_cidest', 'comar_cidest', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugecart_cidnome', 'cart_cidnome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugecart_cidest', 'cart_cidest', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge', 'ZQConjugecartorio', 'cartorio', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQCompr_Dados', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadospaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadospromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_DadosFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadostipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoslocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_DadoscodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_DadoscodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscidnome_cob', 'cidnome_cob', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cob', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_DadoscidUF_cob', 'cidUF_cob', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cob', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosnatural', 'natural', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'naturalidade', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_DadosNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosidcidade', 'idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosnomecid', 'nomecid', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosddd', 'ddd', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosemancipacao', 'emancipacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadosestado', 'estado', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscepgeral', 'cepgeral', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscidadeestadonatual', 'cidadeestadonatual', TWideStringField, fkCalculated, 200, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Dados', 'ZQCompr_Dadoscid_uf_cob', 'cid_uf_cob', TWideStringField, fkCalculated, 200, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQCompr_Resp_Dados', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadospaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadospromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_DadosFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadostipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoslocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_DadoscodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_DadoscodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosidrepresenta_participante', 'idrepresenta_participante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadospaticipante_idpaticipante_1', 'paticipante_idpaticipante_1', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosdocumen1', 'documen1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosdocumen2', 'documen2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosnasicmento', 'nasicmento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscid_natural', 'cid_natural', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosnomerepres', 'nomerepres', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosprofiss', 'profiss', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscartorio', 'cartorio', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoslivro', 'livro', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosfolha', 'folha', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscidade_cart', 'cidade_cart', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscomarca', 'comarca', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadostrabalho', 'trabalho', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosadmissao', 'admissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosrenda_1', 'renda_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosnomeconjuge', 'nomeconjuge', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosdoc1conjuge', 'doc1conjuge', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosdoc2conjuge', 'doc2conjuge', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosnascconjuge', 'nascconjuge', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosassconjuge', 'assconjuge', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscidconjuge', 'cidconjuge', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscidade_natural_resp', 'cidade_natural_resp', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cid_natural', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscid_est_resp', 'cid_est_resp', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadosemailrepres', 'emailrepres', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_Resp_Dados', 'ZQCompr_Resp_Dadoscidade_conjuge', 'cidade_conjuge', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cidconjuge', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQMensMot', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMensMot', 'ZQMensMotidmens_motivo', 'idmens_motivo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMensMot', 'ZQMensMotretorno', 'retorno', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMensMot', 'ZQMensMotcodmotivo', 'codmotivo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMensMot', 'ZQMensMotmensagem', 'mensagem', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQMens_Retorno', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Retorno', 'ZQMens_Retornoidmens_retorno', 'idmens_retorno', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Retorno', 'ZQMens_Retornobanco', 'banco', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Retorno', 'ZQMens_Retornobase', 'base', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Retorno', 'ZQMens_Retornomensagem', 'mensagem', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Retorno', 'ZQMens_Retornonomebanco', 'nomebanco', TWideStringField, fkLookup, 30, 0, False, '', '', '', '', 0, 'nomebanco', 'ZQContaBancaria', 'idconta_bancaria', 'idconta_bancaria', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Retorno', 'ZQMens_Retornocodigo', 'codigo', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Retorno', 'ZQMens_Retornoatitude', 'atitude', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQMens_Motivo', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Motivo', 'ZQMens_Motivoidmens_motivo', 'idmens_motivo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Motivo', 'ZQMens_Motivoretorno', 'retorno', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Motivo', 'ZQMens_Motivocodmotivo', 'codmotivo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQMens_Motivo', 'ZQMens_Motivomensagem', 'mensagem', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAchaVenda', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendalote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendalado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendamedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendamedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendamedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendamedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendavalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendadisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendapaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendavenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendapromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendanome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendadoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendadoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendabairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendabairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendafone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendafone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendatipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendanacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendanaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendalocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendarenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendaobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendafone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaVenda', 'ZQAchaVendacadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQCompr_conjuge', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugeidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugenome', 'nome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugedoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugedoc2', 'doc2', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugenasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugenatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugeidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugeemail', 'email', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugeassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugeempresa', 'empresa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecargo', 'cargo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugeadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugerenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugeregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugefolha', 'folha', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugelivro', 'livro', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecartorio', 'cartorio', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecomar_cidnome', 'comar_cidnome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecomar_cidest', 'comar_cidest', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecart_cidnome', 'cart_cidnome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecart_cidest', 'cart_cidest', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugedata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecomar_cid_est', 'comar_cid_est', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecart_cid_est', 'cart_cid_est', TWideStringField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecidnatural', 'cidnatural', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'natural', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugeestnatural', 'estnatural', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'natural', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCompr_conjuge', 'ZQCompr_conjugecidestnatural', 'cidestnatural', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRecebimento_tempE', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempETipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempERefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempErefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempErecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempESubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_tempE', 'ZQRecebimento_tempEsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQBancRemes2', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes2', 'ZQBancRemes2idbanco_remessa', 'idbanco_remessa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes2', 'ZQBancRemes2remessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes2', 'ZQBancRemes2idbanco', 'idbanco', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes2', 'ZQBancRemes2gerado', 'gerado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes2', 'ZQBancRemes2nomearq', 'nomearq', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQBancRemes2', 'ZQBancRemes2arquivo', 'arquivo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRemes_Receb2', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2nomeadversa', 'nomeadversa', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2idremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2remessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2idrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2retorno_motivo', 'retorno_motivo', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2acao', 'acao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2documento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2Observ', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2ordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2TipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2marcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2origem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2recpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2quadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2numboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2Substituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2Nossonumero', 'Nossonumero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2ocorrencia', 'ocorrencia', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2desc_motivo', 'desc_motivo', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2sq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2somar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2credito', 'credito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb2', 'ZQRemes_Receb2dt_ocorrencia', 'dt_ocorrencia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'contador_bx', False);
  RegisterRuntimeField(TDM_Tabelas, 'contador_bx', 'contador_bxidcontador_bx', 'idcontador_bx', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRemes_Receb_atualiza', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRemes_Receb_atualiza', 'ZQRemes_Receb_atualizasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'Acha_imovel', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovellote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovellado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelmedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelmedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelmedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelmedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelarea', 'area', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelcornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveldisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelmatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelmatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelcidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelmatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveldatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveldatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveldataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovellogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelmapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveltestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveltestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveltestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveltestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelbairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelpasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelnomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveltipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelcodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imovelapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'Acha_imovel', 'Acha_imoveldigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQvnd_quadro', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrolote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrolado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadromedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadromedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadromedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadromedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrovalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrocornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrodisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadromatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadromatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrocidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadromatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrodatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrodatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrodataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrologotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadromapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrotestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrotestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrotestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrotestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrobairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadropasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadronomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrotipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrodigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroidcidade', 'idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadronomecid', 'nomecid', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroddd', 'ddd', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroemancipacao', 'emancipacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadroestado', 'estado', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQvnd_quadro', 'ZQvnd_quadrocepgeral', 'cepgeral', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQProcuradores', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProcuradores', 'ZQProcuradoresidProcuradores', 'idProcuradores', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProcuradores', 'ZQProcuradoresidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProcuradores', 'ZQProcuradoresidparti', 'idparti', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProcuradores', 'ZQProcuradoresnome', 'nome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProcuradores', 'ZQProcuradoresassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQprice', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQprice', 'ZQpriceidPrice', 'idPrice', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQprice', 'ZQpricecoeficiente_multiplicador', 'coeficiente_multiplicador', TFloatField, fkData, 0, 0, False, '', '##,###,##0.00000', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQprice', 'ZQpriceCoeficiente_amortizador', 'Coeficiente_amortizador', TFloatField, fkData, 0, 0, False, '', '##,###,##0.00000', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQprice', 'ZQpriceTaxa', 'Taxa', TFloatField, fkData, 0, 0, False, '', '##0.000', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQprice', 'ZQpriceMes', 'Mes', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQigpm', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQigpm', 'ZQigpmidIGPM', 'idIGPM', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQigpm', 'ZQigpmMes_ano', 'Mes_ano', TWideStringField, fkData, 7, 0, True, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQigpm', 'ZQigpmpercentual', 'percentual', TFloatField, fkData, 0, 0, True, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAchaIgpm', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaIgpm', 'ZQAchaIgpmidIGPM', 'idIGPM', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaIgpm', 'ZQAchaIgpmMes_ano', 'Mes_ano', TWideStringField, fkData, 7, 0, True, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaIgpm', 'ZQAchaIgpmpercentual', 'percentual', TFloatField, fkData, 0, 0, True, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQlog_logradouro', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQlog_logradouro', 'ZQlog_logradourolog_nu_sequencial', 'log_nu_sequencial', TIntegerField, fkData, 0, 0, True, 'N''#186'' Sequencial', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQlog_logradouro', 'ZQlog_logradourolog_no', 'log_no', TWideStringField, fkData, 70, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQlog_logradouro', 'ZQlog_logradourolog_nome', 'log_nome', TWideStringField, fkData, 125, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQlog_logradouro', 'ZQlog_logradourocep', 'cep', TWideStringField, fkData, 16, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQlog_logradouro', 'ZQlog_logradourolog_complemento', 'log_complemento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQlog_logradouro', 'ZQlog_logradourolog_tipo_logradouro', 'log_tipo_logradouro', TWideStringField, fkData, 72, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQlog_logradouro', 'ZQlog_logradourolog_no_sem_acento', 'log_no_sem_acento', TWideStringField, fkData, 70, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'localizacao', False);
  RegisterRuntimeField(TDM_Tabelas, 'localizacao', 'localizacaoufe_sg', 'ufe_sg', TWideStringField, fkData, 2, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'localizacao', 'localizacaocep', 'cep', TWideStringField, fkData, 16, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'localizacao', 'localizacaoloc_nu_sequencial', 'loc_nu_sequencial', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'localizacao', 'localizacaolog_tipo_logradouro', 'log_tipo_logradouro', TWideStringField, fkData, 72, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'localizacao', 'localizacaolog_no', 'log_no', TWideStringField, fkData, 70, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'localizacao', 'localizacaolog_no_sem_acento', 'log_no_sem_acento', TWideStringField, fkData, 70, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'localizacao', 'localizacaobai_nu_sequencial_ini', 'bai_nu_sequencial_ini', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'localizacao', 'localizacaoloc_no', 'loc_no', TWideStringField, fkLookup, 60, 0, False, '', '', '', '', 0, 'loc_nu_sequencial', 'log_localidade', 'loc_nu_sequencial', 'loc_no', True);
  RegisterRuntimeField(TDM_Tabelas, 'localizacao', 'localizacaobai_no', 'bai_no', TWideStringField, fkLookup, 60, 0, False, '', '', '', '', 0, 'bai_nu_sequencial_ini', 'ZQbairro', 'bai_nu_sequencial', 'bai_no', True);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQEmail', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailmatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaildatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaildatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaildataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaillogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailmapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailtestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailtestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailtestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailtestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailbairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailpasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailnomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailtipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaildigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaildocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailidloteamento_1', 'idloteamento_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaildescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaildoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaildoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailbairro_1', 'bairro_1', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailtipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmaillocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQEmail', 'ZQEmailInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'log_tipo_logr', False);
  RegisterRuntimeField(TDM_Tabelas, 'log_tipo_logr', 'log_tipo_logrtipologradouro', 'tipologradouro', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'log_faixa_uf', False);
  RegisterRuntimeField(TDM_Tabelas, 'log_faixa_uf', 'log_faixa_ufufe_sg', 'ufe_sg', TWideStringField, fkData, 2, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'log_faixa_uf', 'log_faixa_ufufe_no', 'ufe_no', TWideStringField, fkData, 72, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'log_localidade', False);
  RegisterRuntimeField(TDM_Tabelas, 'log_localidade', 'log_localidadeloc_no', 'loc_no', TWideStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'log_localidade', 'log_localidadeufe_sg', 'ufe_sg', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'log_localidade', 'log_localidadeloc_nu_sequencial', 'loc_nu_sequencial', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQbairro', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQbairro', 'ZQbairrobai_no', 'bai_no', TWideStringField, fkData, 72, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQbairro', 'ZQbairrobai_nu_sequencial', 'bai_nu_sequencial', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQsequencia_remessa', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQsequencia_remessa', 'ZQsequencia_remessaidremessa_sequencia', 'idremessa_sequencia', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQsequencia_remessa', 'ZQsequencia_remessaidconta_bancaria', 'idconta_bancaria', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQsequencia_remessa', 'ZQsequencia_remessasequencia', 'sequencia', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQmemorial', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQmemorial', 'ZQmemorialidmemorial_descritivo', 'idmemorial_descritivo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQmemorial', 'ZQmemorialidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQmemorial', 'ZQmemorialdescricao', 'descricao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQProposta', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaquadra', 'quadra', TWideStringField, fkLookup, 25, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'quadra', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostalote', 'lote', TWideStringField, fkLookup, 4, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'lote', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostacodloteamento', 'codloteamento', TIntegerField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'loteamento_idloteamento', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostanometoeam', 'nometoeam', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaarea', 'area', TFloatField, fkLookup, 0, 0, False, '', '###,###,##0.00', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'area', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaruafrente', 'ruafrente', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontofrente', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostamedfrente', 'medfrente', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'medidafrente', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostamedfundo', 'medfundo', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'medidafundo', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostamedesquerda', 'medesquerda', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'medidaesquerda', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostameddireita', 'meddireita', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'medidadedireita', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostavalorvenal', 'valorvenal', TFloatField, fkLookup, 0, 0, False, '', '###,###.##0.00', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'valorvenal', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaconf_frente', 'conf_frente', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontofrente', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaconf_fundo', 'conf_fundo', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontofundo', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaconf_esquerda', 'conf_esquerda', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontoesquerda', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaconf_direita', 'conf_direita', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'confrontodireita', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostalado', 'lado', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'lado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaesquina', 'esquina', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'esquinanome', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostalinha', 'linha', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'esquinalinha', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostacurva', 'curva', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'esquinacurva', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostadisponivel', 'disponivel', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'disponivel', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostamatricula', 'matricula', TWideStringField, fkLookup, 10, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'matriculaloteamento', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaimovel_matr', 'imovel_matr', TWideStringField, fkLookup, 10, 0, False, '', '', '', '', 0, 'imovel', 'ZQImovel', 'idimovel', 'matricula', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaBairro', 'Bairro', TWideStringField, fkLookup, 80, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'bairro', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaCidadeIdLoteamento', 'cidade_idcidade_loteamento', TLargeintField, fkLookup, 0, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'cidade_idcidade', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaCidade', 'Cidade', TWideStringField, fkLookup, 150, 0, False, '', '', '', '', 0, 'cidade_idcidade_loteamento', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostauf', 'uf', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_idcidade_loteamento', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostatabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaNomeEmpreend', 'NomeEmpreend', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codloteamento', 'ZQLoteamento', 'idloteamento', 'nomeloteamento', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaidproposta', 'idproposta', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostadataproposta', 'dataproposta', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostavalorproposta', 'valorproposta', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostaPerc_comissao', 'Perc_comissao', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQProposta', 'ZQPropostavlr_comissao', 'vlr_comissao', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQCheque_proposta', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostabanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostadono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostadeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaalias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostanumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQCheque_proposta', 'ZQCheque_propostasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQVendedor_proposta', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVendedor_proposta', 'ZQVendedor_propostaidvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVendedor_proposta', 'ZQVendedor_propostacorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVendedor_proposta', 'ZQVendedor_propostavenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVendedor_proposta', 'ZQVendedor_propostaidproposta', 'idproposta', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQVendedor_proposta', 'ZQVendedor_propostaPcomissao_corretor', 'Pcomissao_corretor', TFloatField, fkData, 0, 0, False, '', '##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQComprador_proposta', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador_proposta', 'ZQComprador_propostaidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador_proposta', 'ZQComprador_propostapaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador_proposta', 'ZQComprador_propostavenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador_proposta', 'ZQComprador_propostapromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador_proposta', 'ZQComprador_propostapercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador_proposta', 'ZQComprador_propostaMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador_proposta', 'ZQComprador_propostaidproposta', 'idproposta', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQComprador_proposta', 'ZQComprador_propostaParticipante', 'Participante', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'paticipante_idpaticipante', 'ZqParticipante_proposta', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRecebimento_proposta', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaidrecebimento_proposta', 'idrecebimento_proposta', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostanumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostanomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostadescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZqParticipante_proposta', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostanome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostadoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostadoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostabairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '99\.999\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostabairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '99\.999\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostafone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostafone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostatipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostanacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostanomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostanomeestado', 'nomeestado', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostanomecidadecob', 'nomecidadecob', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cob', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostanomeestadocob', 'nomeestadocob', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cob', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostanaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostalocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostanomecidadenatural', 'nomecidadenatural', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'naturalidade', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostanomeestadonatural', 'nomeestadonatural', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'naturalidade', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaESTADOCIVIL', 'ESTADOCIVIL', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostarenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostafone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostasexo', 'sexo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostacex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostajb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostajudicial', 'judicial', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaxquadra', 'xquadra', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaxlote', 'xlote', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostaplantao', 'plantao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZqParticipante_proposta', 'ZqParticipante_propostahistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQresponsavel_proposta', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostaidrepresenta_participante', 'idrepresenta_participante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostapaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostadocumen1', 'documen1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostadocumen2', 'documen2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostanasicmento', 'nasicmento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostacid_natural', 'cid_natural', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostanomerepres', 'nomerepres', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostaprofiss', 'profiss', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostaassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostatrabalho', 'trabalho', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostaadmissao', 'admissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostarenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostanomeconjuge', 'nomeconjuge', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostadoc1conjuge', 'doc1conjuge', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostadoc2conjuge', 'doc2conjuge', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostanascconjuge', 'nascconjuge', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostaassconjuge', 'assconjuge', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostacidconjuge', 'cidconjuge', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostacartorio', 'cartorio', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostalivro', 'livro', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostafolha', 'folha', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostacidade_cart', 'cidade_cart', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostacomarca', 'comarca', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostaregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostanomecidadecart', 'nomecidadecart', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostaestcidadecart', 'estcidadecart', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostanomecomarca', 'nomecomarca', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostaestcomarca', 'estcomarca', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostanomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cid_natural', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostanomeesta', 'nomeesta', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cid_natural', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQresponsavel_proposta', 'ZQresponsavel_propostaemailrepres', 'emailrepres', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQConjuge_proposta', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostaidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostanome', 'nome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostadoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostadoc2', 'doc2', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostanasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostanatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostaemail', 'email', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostaassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostaempresa', 'empresa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostacargo', 'cargo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostaadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostarenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostanomecida', 'nomecida', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'natural', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostaestcida', 'estcida', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'natural', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostaidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostaregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostacomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostacidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostafolha', 'folha', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostalivro', 'livro', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostacomar_cidnome', 'comar_cidnome', TWideStringField, fkLookup, 10, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostacomar_cidest', 'comar_cidest', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'comarca', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostacart_cidnome', 'cart_cidnome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostacart_cidest', 'cart_cidest', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cart', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TDM_Tabelas, 'ZQConjuge_proposta', 'ZQConjuge_propostacartorio', 'cartorio', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQagcompromisso', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso', 'ZQagcompromissoREGISTRO', 'REGISTRO', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso', 'ZQagcompromissoDT_REGISTRO', 'DT_REGISTRO', TDateField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso', 'ZQagcompromissoASSUNTO', 'ASSUNTO', TWideStringField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso', 'ZQagcompromissoDSC_ASSUNTO', 'DSC_ASSUNTO', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso', 'ZQagcompromissoCONF_OK', 'CONF_OK', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso', 'ZQagcompromissoHora', 'Hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQagcompromisso_consulta', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso_consulta', 'ZQagcompromisso_consultaREGISTRO', 'REGISTRO', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso_consulta', 'ZQagcompromisso_consultaDT_REGISTRO', 'DT_REGISTRO', TDateField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso_consulta', 'ZQagcompromisso_consultaASSUNTO', 'ASSUNTO', TWideStringField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso_consulta', 'ZQagcompromisso_consultaDSC_ASSUNTO', 'DSC_ASSUNTO', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso_consulta', 'ZQagcompromisso_consultaCONF_OK', 'CONF_OK', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQagcompromisso_consulta', 'ZQagcompromisso_consultaHora', 'Hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQRescisao', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaodatarescisao', 'datarescisao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaohora_rescisao', 'hora_rescisao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaoautocodigo', 'autocodigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQRescisao', 'ZQRescisaoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQIPCA', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQIPCA', 'ZQIPCAidipca', 'idipca', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQIPCA', 'ZQIPCAMes_ano', 'Mes_ano', TWideStringField, fkData, 7, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQIPCA', 'ZQIPCApercentual', 'percentual', TFloatField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'ZQAchaIpca', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaIpca', 'ZQAchaIpcaidipca', 'idipca', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaIpca', 'ZQAchaIpcaMes_ano', 'Mes_ano', TWideStringField, fkData, 7, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'ZQAchaIpca', 'ZQAchaIpcapercentual', 'percentual', TFloatField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TDM_Tabelas, 'qryagenda', False);
  RegisterRuntimeField(TDM_Tabelas, 'qryagenda', 'qryagendaidAgenda', 'idAgenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'qryagenda', 'qryagendadata', 'data', TDateField, fkData, 0, 0, True, '', '', '', '!99/99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'qryagenda', 'qryagendahora', 'hora', TTimeField, fkData, 0, 0, True, '', '', '', '!99:99:99;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'qryagenda', 'qryagendaassunto', 'assunto', TWideStringField, fkData, 30, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TDM_Tabelas, 'qryagenda', 'qryagendadetalhes', 'detalhes', TWideMemoField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);

end.
