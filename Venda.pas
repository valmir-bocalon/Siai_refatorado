unit Venda;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XDBNum, XNum, StdCtrls, wwdbdatetimepicker, Mask, DBCtrls,
  Grids, DBGrids, ComCtrls, TabNotBk, ExtCtrls, XBanner,
  XDate, XDBEdit, DB, JvExControls, JvXPCore, JvXPBar, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, FnpNumericEdit, XDBDate, Gauges, ImgList,
  DBClient, Provider, AppEvnts, System.ImageList, dxCoreA,system.Threading;

type
  TFrm_Venda = class(TForm)
    XBanner14: TXBanner;
    Label19: TLabel;
    Panel2: TPanel;
    XBanner8: TXBanner;
    BTExcluir: TdxButtonArround;
    BtEditar: TdxButtonArround;
    BtIncluir: TdxButtonArround;
    BtCancelar: TdxButtonArround;
    BtGravar: TdxButtonArround;
    BTPesquisar: TdxButtonArround;
    BtRelatorio: TdxButtonArround;
    BtUltimo: TdxButtonArround;
    BtProximo: TdxButtonArround;
    BTAnterior: TdxButtonArround;
    BtPrimeiro: TdxButtonArround;
    BTFechar: TdxButtonArround;
    BtProcessar: TdxButtonArround;
    Pag_Venda: TTabbedNotebook;
    XBanner1: TXBanner;
    DBGVenda: TDBGrid;
    DBECod: TDBEdit;
    Label1: TLabel;
    Label5: TLabel;
    Comprador: TGroupBox;
    Corretor: TGroupBox;
    XBanner2: TXBanner;
    XBanner3: TXBanner;
    DBGrid2: TDBGrid;
    Label14: TLabel;
    DBEImovel: TDBEdit;
    GBEntrada: TGroupBox;
    XBanner7: TXBanner;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    XNEEntrada: TXNumEdit;
    XNQtdeEnt: TXNumEdit;
    DBEEntrada: TDBGrid;
    DBEValor: TXDBNumEdit;
    XBanner9: TXBanner;
    Label23: TLabel;
    Label31: TLabel;
    XVencimentoEnt: TXDateEdit;
    EContabilEntrada: TEdit;
    Label6: TLabel;
    Label8: TLabel;
    DXBAchaLote: TdxButtonArround;
    DBELote: TDBEdit;
    DBEdit6: TDBEdit;
    Label2: TLabel;
    Label4: TLabel;
    Label9: TLabel;
    DBEEmpreend: TDBEdit;
    EComprador: TEdit;
    Ecorretor: TEdit;
    DBGrid1: TDBGrid;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    Label3: TLabel;
    DBEdit3: TDBEdit;
    Label7: TLabel;
    Label10: TLabel;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit7: TDBEdit;
    DBEdit8: TDBEdit;
    DBEdit9: TDBEdit;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    DBEdit10: TDBEdit;
    Labelx: TLabel;
    DBEOrdem: TDBEdit;
    Label17: TLabel;
    XDBNumEdit1: TXDBNumEdit;
    Label18: TLabel;
    CBPeriodoEnt: TComboBox;
    Label20: TLabel;
    Label21: TLabel;
    Label27: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    DBEBanco: TDBEdit;
    DBEAgencia: TDBEdit;
    DBEConta: TDBEdit;
    DBEDocumento: TDBEdit;
    DBENumeroEntrada: TDBEdit;
    DBCBDocVenda: TDBComboBox;
    GBParcelas: TGroupBox;
    XBanner4: TXBanner;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label37: TLabel;
    Label38: TLabel;
    Label39: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Label43: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    Label49: TLabel;
    XNParcela: TXNumEdit;
    XNQtdeParc: TXNumEdit;
    DBGPartelas: TDBGrid;
    DBEValorParc: TXDBNumEdit;
    XVencimentoParc: TXDateEdit;
    EContabilEntradaParc: TEdit;
    DBEOrdemParc: TDBEdit;
    DBEBancoParc: TDBEdit;
    DBEAgendaParc: TDBEdit;
    DBEContaParc: TDBEdit;
    DBEDocumentoParc: TDBEdit;
    DBENumeroParc: TDBEdit;
    DBCBDocVendaParc: TDBComboBox;
    CBPeriodoParc: TComboBox;
    EloteTemp: TEdit;
    Label22: TLabel;
    DBEVrvenda: TXDBNumEdit;
    Label50: TLabel;
    XDBNumEdit2: TXDBNumEdit;
    DBEDonoEntrada: TDBEdit;
    Label51: TLabel;
    Label52: TLabel;
    DBEDonoParcela: TDBEdit;
    JvXPBar1: TJvXPBar;
    Vias: TMaskEdit;
    pgn: TMaskEdit;
    Label53: TLabel;
    DBEdit11: TDBEdit;
    Panel1: TPanel;
    Timer1: TTimer;
    Label54: TLabel;
    DBEPRice: TDBComboBox;
    XBanner5: TXBanner;
    XBanner6: TXBanner;
    XBanner10: TXBanner;
    GroupBox1: TGroupBox;
    XBanner11: TXBanner;
    Label55: TLabel;
    Label58: TLabel;
    Label59: TLabel;
    Label60: TLabel;
    Label65: TLabel;
    XDentrada: TXNumEdit;
    XDqtdeentrada: TXNumEdit;
    DBGrid3: TDBGrid;
    DBEdit12: TDBEdit;
    Label89: TLabel;
    XDFinal: TXDBNumEdit;
    GroupBox3: TGroupBox;
    XBanner13: TXBanner;
    Label56: TLabel;
    Label57: TLabel;
    Label61: TLabel;
    Label62: TLabel;
    Label63: TLabel;
    ultparc: TXNumEdit;
    ZQParcelas: TZQuery;
    DS_Entrada: TDataSource;
    ZQParcelatotal: TZQuery;
    DS_Parcela: TDataSource;















    Label90: TLabel;
    tprice: TComboBox;
    xdresta: TXNumEdit;
    iprice: TMaskEdit;
    Label69: TLabel;
    Label70: TLabel;
    Label71: TLabel;
    Label72: TLabel;
    Label73: TLabel;
    Label74: TLabel;
    Label75: TLabel;
    vr_price: TXNumEdit;
    GroupBox2: TGroupBox;
    XBanner12: TXBanner;
    Label76: TLabel;
    Label77: TLabel;
    Label82: TLabel;
    Label64: TLabel;
    Label66: TLabel;
    Label67: TLabel;
    xdparcelas: TXNumEdit;
    xdqtdeparcelas: TXNumEdit;
    DBGrid4: TDBGrid;
    DBEdit19: TDBEdit;
    xdparpagas: TXNumEdit;
    xdsldvicenda: TXNumEdit;
    xdsldremanescente: TXNumEdit;
    Ultimo_Igpm: TXNumEdit;
    xpcorrecao_igpm: TXNumEdit;
    xvcorrecao_igpm: TXNumEdit;
    xdacrescimo: TXNumEdit;
    xvracrescimo: TXNumEdit;
    xddesconto: TXNumEdit;
    xvrdesconto: TXNumEdit;
    xdliquidado: TXNumEdit;
    vrfinal: TXNumEdit;
    sldpago: TXNumEdit;
    dxButton1: TdxButtonArround;
    dxButton2: TdxButtonArround;
    ZQparcelas_pagas: TZQuery;
    DS_parcelas_pagas: TDataSource;














    Label68: TLabel;
    DBComboBox1: TDBComboBox;
    Label78: TLabel;
    Epercentual: TFnpNumericEdit;
    DBEDAta: TXDBDateEdit;
    DBEVencimento: TXDBDateEdit;
    DBEVencimentoParc: TXDBDateEdit;
    Bar1: TProgressBar;
    Bar2: TProgressBar;
    ZQRecebimento: TZQuery;

    DS_Recebimento: TDataSource;

    FP: TRadioGroup;
    ZQAditamento: TZQuery;
    DS_Aditamento: TDataSource;











    ZQQuitacao: TZQuery;
    DS_quitacao: TDataSource;











    Gauge1: TGauge;
    ZQcomprador_cessao: TZQuery;
    DS_comprador_cessao: TDataSource;











    ZQrecebimento_historico_excluidos: TZQuery;
    DS_ZQrecebimento_historico_excluidos: TDataSource;
































    ZQrecebimento_historico_cessao: TZQuery;
    DS_recebimento_historico_cessao: TDataSource;





































    ImageList1: TImageList;
    btndes: TButton;
    Button1: TButton;
    ZQparcelas12: TZQuery;
    DS_ZQparcelas12: TDataSource;




    DS_ZQresumo_parcelas: TDataSource;
    ZQresumo_parcelas: TZQuery;






    DS_ZQResumo_entrada: TDataSource;
    ZQResumo_entrada: TZQuery;






    DS_ZQtot_parcelas: TDataSource;
    ZQtot_parcelas: TZQuery;







    DS_ZQInandimplencia: TDataSource;
    ZQInandimplencia: TZQuery;






    dxButton3: TdxButtonArround;
    XBanner15: TXBanner;
    DBRichEdit1: TDBRichEdit;
    DataSetProvider1: TDataSetProvider;
    CdsImovel: TClientDataSet;
    DS_CdsImovel: TDataSource;
    ZQimovel: TZQuery;
























    DS_Zqimoves: TDataSource;
    DS_ZQProposta: TDataSource;
    ZQProposta: TZQuery;







    ds_cds_proposta: TDataSource;
    cds_proposta: TClientDataSet;
    DataSetProvider2: TDataSetProvider;
    DataSetProvider3: TDataSetProvider;
    cds_comprador: TClientDataSet;
    ds_cds_comprador: TDataSource;
    ZQcomprador: TZQuery;






    DS_ZQcomprador: TDataSource;
    DataSetProvider5: TDataSetProvider;
    cds_recebimento: TClientDataSet;
    DS_cds_recebimento: TDataSource;
    ZQRecebimento_proposta: TZQuery;





































    DS_ZQRecebimento_proposta: TDataSource;
    DS_ZQcheques: TDataSource;
    ZQcheques: TZQuery;

















    ds_cds_cheque: TDataSource;
    cds_cheque: TClientDataSet;
    DataSetProvider6: TDataSetProvider;
    XBanner16: TXBanner;
    DBGrid5: TDBGrid;
    DBGrid6: TDBGrid;
    DBGrid7: TDBGrid;
    DBGrid8: TDBGrid;
    CdsImovelidimovel: TLargeintField;
    CdsImovelloteamento_idloteamento: TLargeintField;
    CdsImovelquadra: TWideStringField;
    CdsImovellote: TWideStringField;
    CdsImovellado: TWideStringField;
    CdsImovelmedidafrente: TWideStringField;
    CdsImovelmedidafundo: TWideStringField;
    CdsImovelmedidaesquerda: TWideStringField;
    CdsImovelmedidadedireita: TWideStringField;
    CdsImovelarea: TFloatField;
    CdsImovelconfrontofrente: TWideStringField;
    CdsImovelconfrontofundo: TWideStringField;
    CdsImovelconfrontodireita: TWideStringField;
    CdsImovelconfrontoesquerda: TWideStringField;
    CdsImovelesquinanome: TWideStringField;
    CdsImovelesquinalinha: TWideStringField;
    CdsImovelesquinacurva: TWideStringField;
    CdsImovelvalorvenal: TFloatField;
    CdsImovelObervacaoloteamento: TMemoField;
    CdsImovelcornomapa: TWideStringField;
    CdsImoveldisponivel: TWideStringField;
    CdsImovelmatri: TWideStringField;
    CdsImovelmatricula: TWideStringField;
    CdsImovelproposta: TWidestringField;
    cds_propostaidproposta: TLargeintField;
    cds_propostadataproposta: TDateField;
    cds_propostaimovel: TLargeintField;
    cds_propostavalorproposta: TFloatField;
    cds_propostaforma_reajuste: TWidestringField;
    cds_propostatabela_Price: TFloatField;
    cds_propostaEscriturado: TWidestringField;
    cds_compradoridcomprador: TLargeintField;
    cds_compradorpaticipante_idpaticipante: TLargeintField;
    cds_compradorpromissario: TWidestringField;
    cds_compradorpercentual: TFloatField;
    cds_compradorMarcar: TSmallintField;
    cds_recebimentoidrecebimento_proposta: TLargeintField;
    cds_recebimentodocumento: TWidestringField;
    cds_recebimentocliente: TIntegerField;
    cds_recebimentousuario: TIntegerField;
    cds_recebimentoDt_Entrada: TDateField;
    cds_recebimentoDt_Vencimento: TDateField;
    cds_recebimentoValor: TFloatField;
    cds_recebimentoObserv: TMemoField;
    cds_recebimentoVrDoc: TFloatField;
    cds_recebimentoordem: TWidestringField;
    cds_recebimentoTipDoc: TWidestringField;
    cds_recebimentosaldo: TFloatField;
    cds_recebimentomarcar: TWidestringField;
    cds_recebimentoRefBaixa: TIntegerField;
    cds_recebimentorefvinda: TIntegerField;
    cds_recebimentocontabil: TIntegerField;
    cds_recebimentoempresa: TIntegerField;
    cds_recebimentocustodaparcela: TFloatField;
    cds_recebimentoorigem: TWidestringField;
    cds_recebimentoadversa: TIntegerField;
    cds_recebimentorecpag: TWidestringField;
    cds_recebimentonumordem: TIntegerField;
    cds_recebimentoidloteamento: TIntegerField;
    cds_recebimentovenda_idvenda: TIntegerField;
    cds_recebimentoquadralote: TWidestringField;
    cds_recebimentonumboleto: TWidestringField;
    cds_recebimentoSubstituicao: TWidestringField;
    cds_recebimentosq: TLargeintField;
    cds_recebimentosomar: TWidestringField;
    cds_recebimentonomeadversa: TWidestringField;
    cds_recebimentoReajustado: TWidestringField;
    cds_recebimentoData_reajuste: TDateField;
    cds_recebimentoProximo_Reajuste: TWidestringField;
    cds_recebimentoParcelas_fixas: TWidestringField;
    cds_recebimentojuros: TFloatField;
    cds_recebimentodescontos: TFloatField;
    cds_recebimentoData_Quitacao: TDateField;
    cds_chequeidcheque: TLargeintField;
    cds_chequebanco: TWidestringField;
    cds_chequeagencia: TWidestringField;
    cds_chequedono: TWidestringField;
    cds_chequeCPF_CNPJ: TWidestringField;
    cds_chequevalor: TFloatField;
    cds_chequeemissao: TDateField;
    cds_chequedeposito: TDateField;
    cds_chequealias: TWidestringField;
    cds_chequeprorrogado: TDateField;
    cds_chequeconta: TWidestringField;
    cds_chequeidrecebimento: TIntegerField;
    cds_chequeidpagamento: TIntegerField;
    cds_chequenumero: TWidestringField;
    cds_chequeidvinculo: TIntegerField;
    cds_chequeSubstituicao: TWidestringField;
    cds_chequesq: TLargeintField;
    DBGrid9: TDBGrid;
    Label79: TLabel;
    Label80: TLabel;
    Label81: TLabel;
    Label83: TLabel;
    Label84: TLabel;
    DataSetProvider8: TDataSetProvider;
    cds_conjuge: TClientDataSet;
    DS_cds_conjuge: TDataSource;
    ZQConjuge: TZQuery;
    DS_ZQConjuge: TDataSource;




















    cds_conjugeidconjuge: TLargeintField;
    cds_conjugenome: TWidestringField;
    cds_conjugedoc1: TWidestringField;
    cds_conjugedoc2: TWidestringField;
    cds_conjugenasc: TDateField;
    cds_conjugenatural: TIntegerField;
    cds_conjugeidparti: TIntegerField;
    cds_conjugeemail: TWidestringField;
    cds_conjugeassina: TWidestringField;
    cds_conjugeempresa: TWidestringField;
    cds_conjugecargo: TWidestringField;
    cds_conjugeadmissa: TDateField;
    cds_conjugerenda: TFloatField;
    cds_conjugeregime: TWidestringField;
    cds_conjugecomarca: TFloatField;
    cds_conjugecidade_cart: TFloatField;
    cds_conjugefolha: TWidestringField;
    cds_conjugelivro: TWidestringField;
    cds_conjugecartorio: TWidestringField;
    cds_conjugedata: TDateField;
    JvXPBar2: TJvXPBar;
    ApplicationEvents1: TApplicationEvents;
    Gauge2: TGauge;
    DataSetProvider4: TDataSetProvider;
    cds_vendedor: TClientDataSet;
    ds_cds_vendedor: TDataSource;
    ZQvendedor: TZQuery;
    DS_ZQvendedor: TDataSource;



    cds_vendedoridvendedor: TLargeintField;
    cds_vendedorcorretor_idcorretor: TLargeintField;
    OpenDialog1: TOpenDialog;
    DataSetProvider7: TDataSetProvider;
    cds_participante: TClientDataSet;
    ds_cds_participante: TDataSource;
    ZQparticipante: TZQuery;










































    DS_participante: TDataSource;
    cds_participanteidpaticipante: TLargeintField;
    cds_participantenome_parte: TWidestringField;
    cds_participantedoc1: TWidestringField;
    cds_participantedoc2: TWidestringField;
    cds_participanteendereco: TWidestringField;
    cds_participantebairro: TWidestringField;
    cds_participantecidade: TIntegerField;
    cds_participantecep: TWidestringField;
    cds_participanteende_cob: TWidestringField;
    cds_participantebairro_cob: TWidestringField;
    cds_participantecidade_cob: TLargeintField;
    cds_participantecep_cob: TWidestringField;
    cds_participanteFone1: TWidestringField;
    cds_participantefone2: TWidestringField;
    cds_participantefone3: TWidestringField;
    cds_participantetipopessoa: TWidestringField;
    cds_participanteaniversario: TDateField;
    cds_participantenacionalidade: TWidestringField;
    cds_participanteemail: TWidestringField;
    cds_participantenaturalidade: TIntegerField;
    cds_participantelocaldetrab: TWidestringField;
    cds_participanteprofissao: TWidestringField;
    cds_participanteestadocivil: TWidestringField;
    cds_participanterenda: TFloatField;
    cds_participanteobservacao: TMemoField;
    cds_participantecodrecinc: TIntegerField;
    cds_participantecodrecBx: TIntegerField;
    cds_participantecodpaginc: TIntegerField;
    cds_participantecodpagBx: TIntegerField;
    cds_participantefone4: TWidestringField;
    cds_participantecomplemento: TWidestringField;
    cds_participantecomplemento_cob: TWidestringField;
    cds_participantecadastrado: TDateField;
    cds_participanteNome_Firma: TWidestringField;
    cds_participanteInativo: TWidestringField;
    cds_participanteSexo: TWidestringField;
    cds_participantecc: TDateField;
    cds_participantecex: TDateField;
    cds_participantejb: TDateField;
    cds_participantejudicial: TWidestringField;
    cds_participantexquadra: TWidestringField;
    cds_participantexlote: TWidestringField;
    SaveDialog1: TSaveDialog;
    dxButton4: TdxButtonArround;
    cds_compradorvenda_idvenda: TLargeintField;
    cds_vendedorvenda_idvenda: TLargeintField;


    cds_vendedoridproposta: TLargeintField;
    cds_compradoridproposta: TLargeintField;
    Label85: TLabel;

    cds_participanteplantao: TWidestringField;
    cds_participante2: TClientDataSet;
    ds_cds_participante2: TDataSource;
    cds_conjuge2: TClientDataSet;
    DS_cds_conjuge2: TDataSource;
    cds_conjuge2idconjuge: TLargeintField;
    cds_conjuge2nome: TWidestringField;
    cds_conjuge2doc1: TWidestringField;
    cds_conjuge2doc2: TWidestringField;
    cds_conjuge2nasc: TDateField;
    cds_conjuge2natural: TIntegerField;
    cds_conjuge2idparti: TIntegerField;
    cds_conjuge2email: TWidestringField;
    cds_conjuge2assina: TWidestringField;
    cds_conjuge2empresa: TWidestringField;
    cds_conjuge2cargo: TWidestringField;
    cds_conjuge2admissa: TDateField;
    cds_conjuge2renda: TFloatField;
    cds_conjuge2regime: TWidestringField;
    cds_conjuge2comarca: TFloatField;
    cds_conjuge2cidade_cart: TFloatField;
    cds_conjuge2folha: TWidestringField;
    cds_conjuge2livro: TWidestringField;
    cds_conjuge2cartorio: TWidestringField;
    cds_conjuge2data: TDateField;
    cds_comprador2: TClientDataSet;
    ds_cds_comprador2: TDataSource;
    cds_comprador2idcomprador: TLargeintField;
    cds_comprador2percentual: TFloatField;
    cds_comprador2Marcar: TSmallintField;
    cds_comprador2venda_idvenda: TLargeintField;
    cds_comprador2idproposta: TLargeintField;
    cds_comprador2paticipante_idpaticipante: TIntegerField;
    DS_Recebimento_temp: TDataSource;
    ZQRecebimento_temp: TZQuery;
    ZQRescisao: TZQuery;
    DS_Rescisao: TDataSource;











    ZQLimpar_participante_proposta: TZQuery;
    DS_limpar: TDataSource;

    gbrescisao: TGroupBox;
    mdata: TMaskEdit;
    Button2: TButton;





    Lmora: TLabel;
    XCorrigido: TXNumEdit;
    Label86: TLabel;
    DS_price: TDataSource;
    ZQPrice: TZQuery;





    cds_comprador2promissario: TStringField;
    cds_participante2idpaticipante: TIntegerField;
    cds_participante2idnovo: TIntegerField;
    cds_participante2nome_parte: TStringField;
    cds_participante2doc1: TStringField;
    cds_participante2doc2: TStringField;
    cds_participante2endereco: TStringField;
    cds_participante2bairro: TStringField;
    cds_participante2cidade: TIntegerField;
    cds_participante2cep: TStringField;
    cds_participante2ende_cob: TStringField;
    cds_participante2bairro_cob: TStringField;
    cds_participante2cidade_cob: TLargeintField;
    cds_participante2cep_cob: TStringField;
    cds_participante2Fone1: TStringField;
    cds_participante2fone2: TStringField;
    cds_participante2fone3: TStringField;
    cds_participante2tipopessoa: TStringField;
    cds_participante2aniversario: TDateField;
    cds_participante2nacionalidade: TStringField;
    cds_participante2email: TStringField;
    cds_participante2naturalidade: TIntegerField;
    cds_participante2localdetrab: TStringField;
    cds_participante2profissao: TStringField;
    cds_participante2estadocivil: TStringField;
    cds_participante2renda: TFloatField;
    cds_participante2observacao: TMemoField;
    cds_participante2codrecinc: TIntegerField;
    cds_participante2codrecBx: TIntegerField;
    cds_participante2codpaginc: TIntegerField;
    cds_participante2codpagBx: TIntegerField;
    cds_participante2fone4: TStringField;
    cds_participante2complemento: TStringField;
    cds_participante2complemento_cob: TStringField;
    cds_participante2cadastrado: TDateField;
    cds_participante2Nome_Firma: TStringField;
    cds_participante2Inativo: TStringField;
    cds_participante2Sexo: TStringField;
    cds_participante2cc: TDateField;
    cds_participante2cex: TDateField;
    cds_participante2jb: TDateField;
    cds_participante2judicial: TStringField;
    cds_participante2xquadra: TStringField;
    cds_participante2xlote: TStringField;
    cds_participante2plantao: TStringField;
    ZQParcelas_rescisao: TZQuery;
    DS_ZQParcelas_rescisao: TDataSource;













































    Label87: TLabel;
    DBEdit13: TDBEdit;
    xdqtdeparcelasAtrz: TXNumEdit;
    XNQtdeParcAtrz: TXNumEdit;
    xdsldatrasadas: TXNumEdit;
    GroupBox4: TGroupBox;
    Label88: TLabel;
    Label91: TLabel;
    DBEdit14: TDBEdit;
    DBEdit15: TDBEdit;


    ProgressBar1: TProgressBar;


    cds_propostaPerc_comissao: TFloatField;
    cds_propostavlr_comissao: TFloatField;

    cds_vendedorPcomissao_corretor: TFloatField;
    DBEdit16: TDBEdit;
    DBEdit17: TDBEdit;
    Label92: TLabel;
    Label93: TLabel;
    Label94: TLabel;
    XDBNumEdit3: TXDBNumEdit;
    XBanner17: TXBanner;
    dxButton5: TdxButtonArround;
    OpenDialog2: TOpenDialog;
    procedure ApplicationEvents1Message(var Msg: TMsg; var Handled: Boolean);    
    procedure FormShow(Sender: TObject);
    procedure botoes;
    procedure Atualiza;
    procedure Atualiza_comprador;
    procedure DesativaCampos;
    procedure AtivaCampos;
    procedure BtPrimeiroClick(Sender: TObject);
    procedure BTAnteriorClick(Sender: TObject);
    procedure BtProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtCancelarClick(Sender: TObject);
    procedure BtProcessarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BTExcluirClick(Sender: TObject);
    procedure BtRelatorioClick(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DXBAchaLoteClick(Sender: TObject);
    procedure ECompradorExit(Sender: TObject);
    procedure CompradorEnter(Sender: TObject);
    procedure DBGrid2KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure CorretorEnter(Sender: TObject);
    procedure EcorretorExit(Sender: TObject);
    procedure CorretorExit(Sender: TObject);
    procedure DBGVendaDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGVendaKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGVendaMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGVendaEnter(Sender: TObject);
    procedure DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGrid2Enter(Sender: TObject);
    procedure DBEEntradaDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBEOrdemChange(Sender: TObject);
    procedure GBEntradaEnter(Sender: TObject);
    procedure CBPeriodoEntExit(Sender: TObject);
    procedure atualiza_tela_venda_entrada;
    procedure atualiza_tela_venda_parcela;
    procedure atualizatebelaspagamento;
    procedure DBEEntradaEnter(Sender: TObject);
    procedure DBEEntradaKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EContabilEntradaExit(Sender: TObject);
    procedure EContabilEntradaEnter(Sender: TObject);
    procedure DBEValorExit(Sender: TObject);
    procedure DBEValorEnter(Sender: TObject);
    procedure DBEDocumentoExit(Sender: TObject);
    procedure XVencimentoEntExit(Sender: TObject);
    procedure DBCBDocVendaExit(Sender: TObject);
    procedure XNEEntradaExit(Sender: TObject);
    procedure GBParcelasEnter(Sender: TObject);
    procedure CBPeriodoParcExit(Sender: TObject);
    procedure XVencimentoParcExit(Sender: TObject);
    procedure DBEDocumentoParcExit(Sender: TObject);
    procedure DBGPartelasDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBEVrvendaExit(Sender: TObject);
    procedure DBCBDocVendaParcExit(Sender: TObject);
    procedure EContabilEntradaParcEnter(Sender: TObject);
    procedure EContabilEntradaParcExit(Sender: TObject);
    procedure DBEOrdemParcChange(Sender: TObject);
    procedure DBEValorParcEnter(Sender: TObject);
    procedure DBEValorParcExit(Sender: TObject);
    procedure DBEEntradaMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGPartelasKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGPartelasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGPartelasEnter(Sender: TObject);
    procedure DBEDonoEntradaExit(Sender: TObject);
    procedure DBEDonoParcelaExit(Sender: TObject);
    procedure JvXPBar1Items0Click(Sender: TObject);
    procedure JvXPBar1Items1Click(Sender: TObject);
    procedure JvXPBar1Exit(Sender: TObject);
    procedure BTPesquisarClick(Sender: TObject);
    procedure DBEdit11Enter(Sender: TObject);
    procedure DBEdit11Exit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Pag_VendaClick(Sender: TObject);
    procedure JvXPBar1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure DBEPRiceEnter(Sender: TObject);
    procedure tpriceExit(Sender: TObject);
    procedure xvracrescimoExit(Sender: TObject);
    procedure xvrdescontoExit(Sender: TObject);
    procedure xdliquidadoEnter(Sender: TObject);
    procedure xdacrescimoExit(Sender: TObject);
    procedure xddescontoExit(Sender: TObject);
    procedure vrfinalEnter(Sender: TObject);
    procedure Ultimo_IgpmEnter(Sender: TObject);
    procedure Ultimo_IgpmExit(Sender: TObject);
    procedure dxButton2Click(Sender: TObject);
    procedure dxButton1Click(Sender: TObject);
    procedure DBGrid2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid2CellClick(Column: TColumn);
    procedure DBComboBox1Exit(Sender: TObject);
    procedure DBEPRiceExit(Sender: TObject);
    procedure DBEEntradaColEnter(Sender: TObject);
    procedure DBGPartelasColEnter(Sender: TObject);
    procedure btndesClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure dxButton3Click(Sender: TObject);
    procedure JvXPBar1Items2Click(Sender: TObject);
    procedure JvXPBar2Items0Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure JvXPBar2Items1Click(Sender: TObject);
    procedure dxButton4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ZQInandimplenciaCalcFields(DataSet: TDataSet);
    procedure tpriceEnter(Sender: TObject);
    procedure JvXPBar1Items3Click(Sender: TObject);
    procedure DBEVrvendaEnter(Sender: TObject);
    procedure XDBNumEdit3Enter(Sender: TObject);
    procedure dxButton5Click(Sender: TObject);
    procedure DBGVendaDblClick(Sender: TObject);

  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Venda: TFrm_Venda;

implementation

uses tabelas, Funcoes, AchaLoteVenda, RelVenda, QuadroResumo, AchaVenda,
  uSiaiPerformance,
  Loteamento, principal, quitacao, Resumo_quita, QuadroMemorialt, QuadroResumo2,
  usimulaAditamento, UnPdf, uRuntimeFields;

{$R *.dfm}

procedure PosicionarDadosParaQuadroResumo;
begin
  if DM_Tabelas = nil then
    Exit;

  { A montagem dos assinantes percorre os detalhes ate EOF. O FortesReport
    precisa recebe-los novamente no primeiro registro para imprimir os dados
    vinculados ao comprador atual. }
  if DM_Tabelas.ZQCompr_Dados.Active and
    not DM_Tabelas.ZQCompr_Dados.IsEmpty then
  begin
    DM_Tabelas.ZQCompr_Dados.First;

    if DM_Tabelas.ZQCompr_Resp_Dados.Active then
      DM_Tabelas.ZQCompr_Resp_Dados.First;

    if DM_Tabelas.ZQCompr_Conjuge.Active then
      DM_Tabelas.ZQCompr_Conjuge.First;
  end;
end;

function ListaVndE: TStrings;
begin
  Result := TStringList.Create;
  Result.Clear;
  Result.BeginUpdate;
  DM_Tabelas.ZQTipoDoc.open;
  DM_Tabelas.ZQTipoDoc.First;
  DM_Tabelas.ZQTipoDoc.DisableControls;
  while not DM_Tabelas.ZQTipoDoc.Eof do Begin
            Result.Add(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
            DM_Tabelas.ZQTipoDoc.Next;
  end;
  DM_Tabelas.ZQTipoDoc.EnableControls;
  Result.EndUpdate;
end;
function ListaVndP: TStrings;
begin
  Result := TStringList.Create;
  Result.Clear;
  Result.BeginUpdate;
  DM_Tabelas.ZQTipoDoc.open;
  DM_Tabelas.ZQTipoDoc.First;
  DM_Tabelas.ZQTipoDoc.DisableControls;
  while not DM_Tabelas.ZQTipoDoc.Eof do Begin
            Result.Add(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
            DM_Tabelas.ZQTipoDoc.Next;
  end;
  DM_Tabelas.ZQTipoDoc.EnableControls;
  Result.EndUpdate;
end;

// para rolar registro no dbgrid com o rolon do mouse
procedure TFrm_Venda.ApplicationEvents1Message(var Msg: TMsg; var Handled: Boolean);
var i: SmallInt;

begin
    if Msg.message = WM_MOUSEWHEEL then
    begin
      Msg.message := WM_KEYDOWN;
      Msg.lParam := 0;
      i := HiWord(Msg.wParam) ;
      if i > 0 then
         Msg.wParam := VK_UP
      else
         Msg.wParam := VK_DOWN;
      Handled := False;
    end;
end;


procedure TFrm_Venda.FormShow(Sender: TObject);
var
Task : Itask;
I:integer;
Lista: TStrings;
  LStarted, LStep: UInt64;
begin
  LStarted := PerformanceStart;
  LStep := PerformanceStart;
  Pag_Venda.PageIndex := 0;
  Frm_principal.Panel1.Visible:=true;
  //Application.ProcessMessages;
  DBCBDocVenda.Clear;
  DBCBDocVendaParc.Clear;
  DM_tabelas.ZQContaBancaria.Close;
  DM_tabelas.ZQContaBancaria.SQL.Clear;
  DM_tabelas.ZQContaBancaria.SQL.Add('select  idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,nometitular,aberta,senha_interna,senha_conta,ativa,');
  DM_tabelas.ZQContaBancaria.SQL.Add('        jurosemboleto,taxadiaria,valoroupercent,carteira,boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,Apelido,cod_transmissao,');
  DM_tabelas.ZQContaBancaria.SQL.Add('        complemento,moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote ');
  DM_tabelas.ZQContaBancaria.SQL.Add('  from conta_bancaria');
  if DM_tabelas.ZQContaBancaria.Active=false then
     DM_tabelas.ZQContaBancaria.open;
  PerformanceElapsed('Venda: abrir contas bancarias', LStep);
  LStep := PerformanceStart;

  DM_tabelas.ZQQuadras.Close;
  DM_tabelas.ZQQuadras.SQL.Clear;
  DM_tabelas.ZQQuadras.SQL.Add('select idquadras,loteamento_idLoteamento,descricao,apedlido from quadras order by descricao');
  if DM_tabelas.ZQQuadras.Active=false then
     DM_tabelas.ZQQuadras.open;
  PerformanceElapsed('Venda: abrir quadras', LStep);
  LStep := PerformanceStart;

  DM_Tabelas.ZQReBxHi.Close;
  DM_Tabelas.ZQReBxHi.SQL.Clear;
  DM_Tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from RecBxHist limit 1');
  if DM_Tabelas.ZQReBxHi.Active=false then
     DM_Tabelas.ZQReBxHi.open;
  PerformanceElapsed('Venda: abrir historico de baixa', LStep);
  LStep := PerformanceStart;

  DM_Tabelas.ZQincorp_loteame.Close;
  DM_Tabelas.ZQincorp_loteame.SQL.Clear;
  DM_Tabelas.ZQincorp_loteame.SQL.Add('Select incorporador_idincorporador,loteamento_idloteamento,percent_parte,codcontabancaria,digito_dif from incorporador_loteamento');
  if DM_Tabelas.ZQincorp_loteame.Active=false then
     DM_Tabelas.ZQincorp_loteame.Open;
  PerformanceElapsed('Venda: abrir incorporadores', LStep);
  LStep := PerformanceStart;


  DM_Tabelas.ZQLoteamento.Close;
  DM_Tabelas.ZQLoteamento.SQL.Clear;
  DM_Tabelas.ZQLoteamento.SQL.Add('Select idloteamento,cidade_idcidade,matriculaloteamento,datacadastro,datainalguracao,dataconclusao,logotipo,mapa,pasta_mapa,testemunha1,testemunha2,testemunha3,');
  DM_Tabelas.ZQLoteamento.SQL.Add('       testemunha4,Observincorp,bairro,tipo,nomeloteamento,codcontabancaria,apelido,Perc_comissao ');
  DM_Tabelas.ZQLoteamento.SQL.Add('  from loteamento');
  if DM_Tabelas.ZQLoteamento.Active=false then
     DM_Tabelas.ZQLoteamento.Open;
  PerformanceElapsed('Venda: abrir empreendimentos', LStep);
  LStep := PerformanceStart;


  DM_Tabelas.zqprocuradores.Close;
  DM_Tabelas.zqprocuradores.SQL.Clear;
  DM_Tabelas.zqprocuradores.SQL.Add('Select idProcuradores,idloteamento,idparti,nome,assina from  Procuradores');
  if DM_Tabelas.zqprocuradores.Active=false then
     DM_Tabelas.zqprocuradores.Open;
  PerformanceElapsed('Venda: abrir procuradores', LStep);
  LStep := PerformanceStart;


  DM_tabelas.ZQNumOrdem.Close;
  DM_tabelas.ZQNumOrdem.SQL.Clear;
  DM_tabelas.ZQNumOrdem.SQL.Add('select idnumordem from numordem');
  if DM_tabelas.ZQNumOrdem.Active=false then
     DM_tabelas.ZQNumOrdem.Open;
  PerformanceElapsed('Venda: abrir numeracao', LStep);
  LStep := PerformanceStart;

  DM_Tabelas.ZQImovel.Close;
  DM_Tabelas.ZQImovel.SQL.Clear;
  DM_Tabelas.ZQImovel.SQL.Add('Select  loteamento_idloteamento,quadra,lote,lado,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,esquinanome,valorvenal,Obervacaoloteamento,');
  DM_Tabelas.ZQImovel.SQL.Add('        cornomapa,idimovel,disponivel,medidafrente,medidafundo,medidadedireita,esquinalinha,esquinacurva,medidaesquerda,matri,matricula,proposta,valorCusto');
  DM_Tabelas.ZQImovel.SQL.Add(' from Imovel');
  if DM_Tabelas.ZQImovel.Active = false then
     DM_Tabelas.ZQImovel.open;
  PerformanceElapsed('Venda: abrir imoveis', LStep);
  LStep := PerformanceStart;

  DM_Tabelas.ZQVendedor.Close;
  DM_Tabelas.ZQVendedor.SQL.Clear;
  DM_Tabelas.ZQVendedor.SQL.Add('Select idvendedor,corretor_idcorretor,venda_idvenda,PComissao_corretor from vendedor');
  if DM_Tabelas.ZQVendedor.Active=false then
     DM_Tabelas.ZQVendedor.open;
  PerformanceElapsed('Venda: abrir vendedores', LStep);
  LStep := PerformanceStart;

  DM_Tabelas.ZQRecebimento.Close;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
  DM_Tabelas.ZQRecebimento.SQL.Add(' from Recebimento order by DT_Vencimento limit 200, 20');
  if DM_Tabelas.ZQRecebimento.Active=false then
     DM_Tabelas.ZQRecebimento.open;
  PerformanceElapsed('Venda: abrir recebimentos', LStep);
  LStep := PerformanceStart;

  DM_Tabelas.ZQVenda.Close;
  DM_Tabelas.ZQVenda.SQL.Clear;
  DM_Tabelas.ZQVenda.SQL.Add('Select idvenda,datavenda,imovel,valorvenda,forma_reajuste,tabela_Price,Escriturado,marca,codigo_contrato_ref,Multa,Mora,Perc_comissao,vlr_comissao,path_pdf from venda limit 50');
  if DM_Tabelas.ZQVenda.Active=false then
     DM_Tabelas.ZQVenda.open;
  PerformanceElapsed('Venda: abrir vendas', LStep);
  LStep := PerformanceStart;

  DM_Tabelas.ZQComprador.Close;
  DM_Tabelas.ZQComprador.SQL.Clear;
  DM_Tabelas.ZQComprador.SQL.Add('Select idcomprador,paticipante_idpaticipante,venda_idvenda,promissario,percentual,Marcar from comprador');
  if DM_Tabelas.ZQComprador.Active=false then
     DM_Tabelas.ZQComprador.open;
  PerformanceElapsed('Venda: abrir compradores', LStep);
  LStep := PerformanceStart;

  DM_tabelas.ZQCheque.Close;
  DM_tabelas.ZQCheque.SQL.Clear;
  DM_tabelas.ZQCheque.SQL.Add('select Banco,Dono,CPF_CNPJ,valor,emissao,deposito,alias,prorrogado,conta,idrecebimento,idpagamento,idcheque,idvinculo,numero,agencia,Substituicao,sq from cheque_rec');
  if DM_tabelas.ZQCheque.Active=false then
     DM_tabelas.ZQCheque.open;
  PerformanceElapsed('Venda: abrir cheques', LStep);
  LStep := PerformanceStart;

  DM_Tabelas.ZQTipoDoc.Close;
  DM_Tabelas.ZQTipoDoc.SQL.Clear;
  DM_Tabelas.ZQTipoDoc.SQL.Add('select idtipodocumento,tipodoc,descricao,vend_receb,vend_caixa,receb_receb,receb_caixa,dados_chequ,so_avista,lancabanco,somapaga,lancinclus from tipodocumento');
  if DM_Tabelas.ZQTipoDoc.Active=false then
     DM_Tabelas.ZQTipoDoc.open;
  PerformanceElapsed('Venda: abrir tipos de documento', LStep);
  LStep := PerformanceStart;

  DM_Tabelas.ZQTipoDoc.First;
  DBCBDocVenda.Clear;
  Lista := ListaVndE;
  try
    DBCBDocVenda.Items := Lista;
  finally
    Lista.Free;
  end;
  DBCBDocVenda.Sorted := True;
  i:=0;
{  for i := 0 to DBCBDocVenda.Items.Count -1 do
  begin
    DBCBDocVenda.Checked[I] := True;
  end;
   }
  DBCBDocVendaParc.Clear;
  Lista := ListaVndP;
  try
    DBCBDocVendaParc.Items := Lista;
  finally
    Lista.Free;
  end;
  DBCBDocVendaParc.Sorted:=True;

  {i:=0;
  for i := 0 to DBCBDocVendaParc.Items.Count -1 do
  begin
    DBCBDocVendaParc.Checked[I] := True;
  end;
 }

//  Task := TTask.create(procedure
//                       begin
//                          While not DM_Tabelas.ZQTipoDoc.Eof do Begin
//                            DBCBDocVenda.Items.Add(DM_Tabelas.ZQTipoDoctipodoc.Value);
//                            DBCBDocVendaParc.Items.Add(DM_Tabelas.ZQTipoDoctipodoc.Value);
//                            DM_Tabelas.ZQTipoDoc.Next;
//                          end;
//                       end);
//  Task.Start;


  DM_TAbelas.ZQmemorial.open;
  PerformanceElapsed('Venda: abrir memorial', LStep);
  LStep := PerformanceStart;
  DesativaCampos;
  DBGVenda.SetFocus;
  Frm_principal.Panel1.Visible:=false;
  Application.ProcessMessages;

  iprice.Clear;
  sldpago.Value:=0;
  vr_price.Value:=0;
  Ultimo_Igpm.Value:=0;
  xpcorrecao_igpm.Value:=0;
  xvcorrecao_igpm.Value:=0;
  xdacrescimo.Value:=0;
  xvracrescimo.Value:=0;
  xddesconto.Value:=0;
  xvrdesconto.Value:=0;
  xdliquidado.Value:=0;
  vrfinal.Value:=0;
  PerformanceElapsed('Venda: preparar controles', LStep);
  PerformanceElapsed('Venda: abertura completa', LStarted);
  FlushPerformanceLog;
end;

procedure TFrm_Venda.botoes;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQVenda.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQVenda.RecNo=DM_Tabelas.ZQVenda.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  Atualiza;
End;

procedure TFrm_Venda.Atualiza;
Begin
  DM_TAbelas.CDSCompradorTemp.Close;
  DM_Tabelas.CDSCompradorTemp.CreateDataSet;
  DM_Tabelas.ZQComprador.First;
  while not DM_Tabelas.ZQComprador.Eof do begin
    DM_Tabelas.CDSCompradorTemp.Insert;
    DM_Tabelas.CDSCompradorTempcodparticipante.Value := DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').AsLargeInt;
    DM_Tabelas.CDSCompradorTempnomeparticipante.Value := DM_Tabelas.ZQComprador.FieldByName('nomecomprador').AsString;
    DM_Tabelas.CDSCompradorTemppercentual.Value := DM_Tabelas.ZQcomprador.FieldByName('percentual').AsFloat;
    if DM_Tabelas.ZQcomprador.FieldByName('Marcar').AsInteger=0 then
       DM_Tabelas.CDSCompradorTempMarcar.Value := False
    else
      DM_Tabelas.CDSCompradorTempMarcar.Value := True;
      
    DM_tabelas.CDSCompradorTemp.Post;
    dm_tabelas.ZQComprador.Next;
  end;
  DM_Tabelas.CDSVendedorTemp.Close;
  DM_Tabelas.CDSVendedorTemp.CreateDataSet;
  DM_Tabelas.ZQVendedor.First;
  While not DM_Tabelas.ZQVendedor.Eof do Begin
     DM_Tabelas.CDSVendedorTemp.Insert;
     DM_Tabelas.CDSVendedorTempcodcorretor.Value := DM_Tabelas.ZQvendedor.FieldByName('corretor_idcorretor').AsLargeInt;
     DM_Tabelas.CDSVendedorTempPComissao_corretor.Value := DM_Tabelas.ZQVendedor.FieldByName('Pcomissao_Corretor').AsFloat;
     DM_Tabelas.CDSVendedorTemp.Post;
    DM_Tabelas.ZQVendedor.Next;
  end;
  Atualiza_comprador;
  if Pag_Venda.PageIndex = 1 Then
    Atualizatebelaspagamento;
End;

procedure TFrm_Venda.Atualiza_comprador;
Begin
  Label10.Visible := True;
  if DM_Tabelas.CDSCompradorTemppessoa.Value = 'F' Then Begin
    Label7.Caption := 'C.P.F.';
    Label10.Caption := 'R.G.';
  end
  else if DM_Tabelas.CDSCompradorTemppessoa.Value = 'J' Then Begin
    Label7.Caption := 'C.N.P.J.';
    Label10.Caption := 'Insc. Est.';
  end
  else if DM_Tabelas.CDSCompradorTemppessoa.Value = 'R' Then Begin
    Label7.Caption := 'C.N.P.J.';
    Label10.Caption := 'Insc. Rural.';
  end
  else if DM_Tabelas.CDSCompradorTemppessoa.Value = 'P' Then Begin
    Label7.Caption := 'C.N.P.J.';
    Label10.Visible := False;
  end;
end;

procedure TFrm_Venda.DesativaCampos;
Begin
  DBEDAta.ReadOnly := True;
  Ecomprador.ReadOnly := True;
  Ecorretor.ReadOnly := True;
  DXBAchaLote.Enabled := False;
  DBEVrvenda.ReadOnly := True;
  XNEEntrada.ReadOnly := True;
  XNQtdeEnt.ReadOnly := True;
  XVencimentoEnt.ReadOnly := True;
  CBPeriodoEnt.Enabled := False;
  EContabilEntrada.ReadOnly := True;
  DBCBDocVenda.ReadOnly := True;
  DBEVencimento.ReadOnly := True;
  DBEValor.ReadOnly := True;
  DBEBanco.ReadOnly := True;
  DBEAgencia.ReadOnly := True;
  DBEConta.ReadOnly := True;
  DBENumeroEntrada.ReadOnly := True;
  DBEDocumento.ReadOnly := True;
  XNParcela.ReadOnly := True;
  XNQtdeParc.ReadOnly := True;
  XVencimentoParc.ReadOnly := True;
  CBPeriodoParc.Enabled := False;
  EContabilEntradaParc.ReadOnly := True;
  DBCBDocVendaParc.ReadOnly := True;
  DBEVencimentoParc.ReadOnly := True;
  DBEValorParc.ReadOnly := True;
  DBEBancoParc.ReadOnly := True;
  DBEAgendaParc.ReadOnly := True;
  DBEContaParc.ReadOnly := True;
  DBENumeroParc.ReadOnly := True;
  DBEDocumentoParc.ReadOnly := True;
  BTPrimeiro.Enabled := true;
  BTAnterior.Enabled := true;
  BTProximo.Enabled := true;
  BtUltimo.Enabled := true;
  BtGravar.Enabled := False;
  BtCancelar.Enabled := False;
  BtProcessar.Enabled := true;
  BtIncluir.Enabled := true;
  BtEditar.Enabled := true;
  BtExcluir.Enabled := true;
  BtPesquisar.Enabled := true;
  BtRelatorio.Enabled := true;
  DBGVenda.Enabled := True;
end;

procedure TFrm_Venda.AtivaCampos;
Begin
  DBEDAta.ReadOnly := False;
  Ecomprador.ReadOnly := False;
  Ecorretor.ReadOnly := False;
  DXBAchaLote.Enabled := True;
  DBEVrvenda.ReadOnly := False;
  XNEEntrada.ReadOnly := False;
  XNQtdeEnt.ReadOnly := False;
  XVencimentoEnt.ReadOnly := False;
  CBPeriodoEnt.Enabled := True;
  EContabilEntrada.ReadOnly := False;
  DBCBDocVenda.ReadOnly := False;
  DBEVencimento.ReadOnly := False;
  DBEValor.ReadOnly := False;
  DBEBanco.ReadOnly := False;
  DBEAgencia.ReadOnly := False;
  DBEConta.ReadOnly := False;
  DBENumeroEntrada.ReadOnly := False;
  DBEDocumento.ReadOnly := False;
  DBEDonoEntrada.ReadOnly := False;
  DBEDonoParcela.ReadOnly := False;
  XNParcela.ReadOnly := False;
  XNQtdeParc.ReadOnly := False;
  XVencimentoParc.ReadOnly := False;
  CBPeriodoParc.Enabled := True;
  EContabilEntradaParc.ReadOnly := False;
  DBCBDocVendaParc.ReadOnly := False;
  DBEVencimentoParc.ReadOnly := False;
  DBEValorParc.ReadOnly := False;
  DBEBancoParc.ReadOnly := False;
  DBEAgendaParc.ReadOnly := False;
  DBEContaParc.ReadOnly := False;
  DBENumeroParc.ReadOnly := False;
  DBEDocumentoParc.ReadOnly := False;
  BTPrimeiro.Enabled := False;
  BTAnterior.Enabled := False;
  BTProximo.Enabled := False;
  BtUltimo.Enabled := False;
  BtGravar.Enabled := True;
  BtCancelar.Enabled := True;
  BtProcessar.Enabled := False;
  BtIncluir.Enabled := False;
  BtEditar.Enabled := False;
  BtExcluir.Enabled := False;
  BtPesquisar.Enabled := False;
  BtRelatorio.Enabled := False;
  DBGVenda.Enabled := False;
end;

procedure TFrm_Venda.BtPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.First;
  DBGVenda.SetFocus;
end;


procedure TFrm_Venda.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Prior;
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.BtProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Next;
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Last;
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.BtGravarClick(Sender: TObject);
var
  varidrecebimento,varnumordemE, varnumordemP, varvezes, varvenda : integer;
  qlote,fx,pr,edicao,passou,entrada:String;
  data:Tdatetime;
begin
    FP.Visible:=false;
    entrada:='N';
    edicao:='N';
    if (Pag_Venda.PageIndex=3) and (DM_TAbelas.ZQmemorial.State in [DsEdit,DsInsert]) then
    begin
      DM_TAbelas.ZQmemorial.post;
      DM_Tabelas.ZQVenda.Cancel;
      DesativaCampos;
      exit;
    end;

    if DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt=0 Then Begin
      mensagem('Deve ser indicado um lote para ser vendido...');
      Pag_Venda.PageIndex := 0;
      DXBAchaLote.SetFocus;
      exit;
    end;
//  if DM_Tabelas.ZQVendavalorvenda.Value<DM_Tabelas.ZQVendavalorvenal.Value Then Begin
//    Showmessage('O valor de VENDA não deve ser menor que o valor venal...');
//    Pag_Venda.PageIndex := 1;
//    DBEVrvenda.SetFocus;
//    exit;
//  end;
    if DM_Tabelas.CDSCompradorTemp.RecordCount=0 Then Begin
      Showmessage('Ao menos um COMPRADOR deve ser informado para que a venda possa ser concluida..');
      Pag_Venda.PageIndex := 0;
      EComprador.SetFocus;
      exit;
    end;
    if DM_Tabelas.CDSVendedorTemp.RecordCount=0 Then Begin
      Showmessage('Ao menos um VENDEDOR deve ser informado para que a venda possa ser concluida..');
      Pag_Venda.PageIndex := 0;
      Ecorretor.SetFocus;
      exit;
    end;

    if fp.ItemIndex=-1 then begin
       Showmessage('Deve indicar o tipo de reajuste das parcelas!');
       fp.SetFocus;
       exit;
    end;

    if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then
    begin
      DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat:=XNParcela.Value+XNEEntrada.Value;
    end;
    if DM_Tabelas.ZQVenda.State in [DsEdit] then
    begin
      edicao:='S';
    end
    else
    begin
      edicao:='N';
    end;
    DM_Tabelas.ZQVenda.Post;



    if not empty(EloteTemp.Text) Then Begin
      DM_Tabelas.ZQImovel.Locate('idimovel',strtoint(EloteTemp.Text),[]);
      DM_Tabelas.ZQImovel.Edit;
      DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString := 'SIM';
    end;
    DM_Tabelas.ZQImovel.Locate('idimovel',DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt,[]);
    DM_Tabelas.ZQImovel.Edit;
    DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString := 'NºO';
    DM_Tabelas.ZQImovel.Post;

    if DM_TAbelas.ZQmemorial.Active=false then
       DM_TAbelas.ZQmemorial.open;
    if edicao='N' then
       DM_TAbelas.ZQmemorial.Insert
    else
       DM_TAbelas.ZQmemorial.Edit;
    DM_TAbelas.ZQmemorial.FieldByName('idvenda').AsLargeInt:= DM_TAbelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
    DM_TAbelas.ZQmemorial.FieldByName('descricao').AsString:=DM_TAbelas.ZQimovel.FieldByName('Obervacaoloteamento').AsString;
    DM_TAbelas.ZQmemorial.Post;

    DM_Tabelas.ZQComprador.First;
    while not DM_Tabelas.ZQComprador.Eof do
      DM_Tabelas.ZQComprador.Delete;
    DM_Tabelas.CDSCompradorTemp.First;

    // aqui marcar
    while not DM_Tabelas.CDSCompradorTemp.Eof do
    begin
      DM_Tabelas.ZQComprador.Insert;
      DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').AsLargeInt := DM_Tabelas.CDSCompradorTempcodparticipante.Value;
      DM_Tabelas.ZQcomprador.FieldByName('percentual').AsFloat:=DM_Tabelas.CDSCompradorTemppercentual.Value;
      if DM_Tabelas.CDSCompradorTempmarcar.Value=False then
         DM_Tabelas.ZQcomprador.FieldByName('Marcar').AsInteger:=0
      else
         DM_Tabelas.ZQcomprador.FieldByName('Marcar').AsInteger:=1;
      DM_Tabelas.CDSCompradorTemp.Next;
    end;


    DM_Tabelas.ZQVendedor.First;
    while not DM_Tabelas.ZQVendedor.Eof do
      DM_Tabelas.ZQVendedor.Delete;
    DM_Tabelas.CDSVendedorTemp.First;
    while not DM_Tabelas.CDSVendedorTemp.Eof do
    begin
      DM_Tabelas.ZQVendedor.Insert;
      DM_Tabelas.ZQvendedor.FieldByName('corretor_idcorretor').AsLargeInt := DM_Tabelas.CDSVendedorTempcodcorretor.Value;
      DM_Tabelas.ZQVendedor.FieldByName('Pcomissao_Corretor').AsFloat := DM_Tabelas.CDSVendedorTempPComissao_corretor.Value;
      DM_Tabelas.CDSVendedorTemp.Next;
    end;
    if DM_Tabelas.ZQRecebimento.Locate('origem,documento',VarArrayOf(['V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text,DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-E-']),[]) Then
      varnumordemE := DM_Tabelas.ZQRecebimento.FieldByName('numordem').AsInteger
    else
    Begin
      DM_tabelas.ZQNumOrdem.Insert;
      DM_Tabelas.ZQNumOrdem.Post;
      varnumordemE := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').AsLargeInt;
    end;
    if DM_Tabelas.ZQRecebimento.Locate('origem,documento',VarArrayOf(['V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text,DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-P-']),[]) Then
      varnumordemp := DM_Tabelas.ZQRecebimento.FieldByName('numordem').AsInteger
    else Begin
      DM_tabelas.ZQNumOrdem.Insert;
      DM_Tabelas.ZQNumOrdem.Post;
      varnumordemp := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').AsLargeInt;
    end;


    if edicao<>'S' then
    begin
      DM_Tabelas.ZQRecebimento.SQL.Clear;
      DM_Tabelas.ZQRecebimento.SQL.Add('delete from recebimento');
      DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
      DM_Tabelas.ZQRecebimento.ExecSQL;
      DM_Tabelas.ZQRecebimento.close;
      DM_Tabelas.ZQRecebimento.SQL.Clear;
      DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
      DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
      DM_Tabelas.ZQRecebimento.SQL.Add(' from recebimento');
      DM_Tabelas.ZQRecebimento.Open;
      DM_Tabelas.ZQLoteamento.Locate('idloteamento',DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger,[]);

      DM_Tabelas.ZQincorp_loteame.first;

      // para inibir duplicar entrada qdo tem 2 donos o lote. ex. jardim canaa
      entrada:='N';
      passou:='N';
      while not DM_Tabelas.ZQincorp_loteame.eof do
      begin
        DM_Tabelas.CDSCompradorTemp.First;
        DM_Tabelas.CDSPagEntrada.First;

        if entrada='N' then
        begin

         DM_Tabelas.CDSPagEntrada.DisableControls;
         Bar1.Visible:=true;
         Bar1.Position:=0;
         Bar1.Max:=DM_Tabelas.CDSPagEntrada.RecordCount;
         while not DM_Tabelas.CDSPagEntrada.Eof do
         Begin
            Bar1.Position:=DM_Tabelas.CDSPagEntrada.RecNo;
            if DM_Tabelas.CDSCompradorTemp.recordcount>1 then
            begin
              DM_Tabelas.CDSCompradorTemp.Filtered:=false;
              DM_Tabelas.CDSCompradorTemp.Filter:='marcar=''true''';
              DM_Tabelas.CDSCompradorTemp.Filtered:=true;
            end;
            if fp.ItemIndex=0 then
               fx:='S'
            else
               fx:='N';


            DM_Tabelas.ZQRecebimento.close;
            DM_Tabelas.ZQRecebimento.SQL.Clear;
            DM_tabelas.ZQRecebimento.SQL.Add('INSERT INTO  Recebimento (idloteamento,documento,tipdoc,cliente,dt_entrada,dt_vencimento,ordem,saldo,Valor,VrDoc,contabil,origem,adversa,nomeadversa,'+
                                                'recpag,numordem,venda_idvenda,quadralote,Observ,Parcelas_fixas) VALUES ('+
                                                quotedstr(inttostr(DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger))+','+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-E- '+CBPeriodoEnt.Text)+','+quotedstr(DM_Tabelas.CDSPagEntradatipodoc.Value)+','+
                                                quotedstr(inttostr(DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').AsLargeInt))+','+quotedstr(FormatDateTime('YYYY-MM-DD',date))+','+
                                                quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.CDSPagEntradavencimento.Value))+','+quotedstr(inttostr(varnumordemE)+'-'+DM_Tabelas.CDSPagEntradaordem.Value)+','+
                                                quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_Tabelas.CDSPagEntradavalor.Value),'.')))+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_Tabelas.CDSPagEntradavalor.Value),'.')))+','+
                                                quotedstr(TrocaVirgPPto(TrimChar(floattostr(XNEEntrada.Value),'.')))+','+quotedstr(inttostr(DM_Tabelas.CDSPagEntradacontabil.Value))+','+
                                                quotedstr('V-'+inttostr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt))+','+quotedstr(inttostr(DM_Tabelas.CDSCompradorTempcodparticipante.Value))+','+
                                                quotedstr(DM_Tabelas.CDSCompradorTempnomeparticipante.Value)+',''R'','+quotedstr(inttostr(varnumordemE))+','+quotedstr(inttostr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt))+','+
                                                quotedstr(DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+','+quotedstr('Empreendimento: '+DM_Tabelas.ZQVenda.FieldByName('nometoeam').AsString+chr(13)+'Quadra: '+DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+chr(13)+'Lote: '+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+','+
                                                quotedstr(fx)+')');
  //         mensagem(DM_tabelas.ZQRecebimento.sql.text);

            DM_tabelas.ZQRecebimento.ExecSQL;



            qlote:=DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString;
            DM_Tabelas.ZQRecebimento.close;
            ZQRecebimento.close;
            ZQRecebimento.SQL.Clear;
            ZQRecebimento.SQL.Add('Select idrecebimento,quadralote from Recebimento where quadralote='+quotedstr(qlote)+' order by idrecebimento');
            ZQRecebimento.Open;
            ZQRecebimento.last;
            varidrecebimento:=ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
            ZQRecebimento.Locate('idrecebimento',varidrecebimento,[]);

            DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagEntradatipodoc.Value,[]);
            if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
              DM_tabelas.ZQCheque.Insert;
              DM_Tabelas.ZQCheque.FieldByName('idrecebimento').AsInteger := ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
              DM_Tabelas.ZQCheque.FieldByName('idvinculo').AsInteger := ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
              DM_Tabelas.ZQCheque.FieldByName('Banco').AsString := DM_Tabelas.CDSPagEntradabanco.Value;
              DM_Tabelas.ZQCheque.FieldByName('agencia').AsString := DM_Tabelas.CDSPagEntradaagencia.Value;
              DM_Tabelas.ZQCheque.FieldByName('numero').AsString := DM_Tabelas.CDSPagEntradanumero.Value;
              DM_Tabelas.ZQCheque.FieldByName('conta').AsString := DM_Tabelas.CDSPagEntradaconta.Value;
              DM_Tabelas.ZQCheque.FieldByName('Dono').AsString := DM_Tabelas.CDSPagEntradanomedono.Value;
              DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').AsString := DM_Tabelas.CDSPagEntradadocumento.Value;
              DM_Tabelas.ZQCheque.FieldByName('valor').AsFloat := DM_Tabelas.CDSPagEntradavalor.Value;
              DM_Tabelas.ZQCheque.FieldByName('emissao').AsDateTime := Date;
              DM_Tabelas.ZQCheque.FieldByName('deposito').AsDateTime := DM_Tabelas.CDSPagEntradavencimento.Value;
              DM_Tabelas.ZQCheque.Post;
            end;
            DM_Tabelas.CDSPagEntrada.Next;
         end;
         DM_Tabelas.CDSPagEntrada.EnableControls;
         Bar1.Visible:=false;
         Bar1.Position:=0;

         entrada:='S';
        // até aqui
      end;

      DM_Tabelas.CDSPagParcela.First;
      DM_Tabelas.CDSPagParcela.DisableControls;
      Bar2.Visible:=true;
      Bar2.Position:=0;
      Bar2.Max:=DM_Tabelas.cdSPagParcela.RecordCount;

      while not DM_Tabelas.cdSPagParcela.Eof do
      Begin
        bar2.Position:=DM_Tabelas.cdSPagParcela.RecNo;
        if DM_Tabelas.CDSCompradorTemp.recordcount>1 then
        begin
          DM_Tabelas.CDSCompradorTemp.Filtered:=false;
          DM_Tabelas.CDSCompradorTemp.Filter:='marcar=''true''';
          DM_Tabelas.CDSCompradorTemp.Filtered:=true;
        end;

        if fp.ItemIndex=1 then
           fx:='N'
        else
           fx:='S';


        if (fx='N') and (strtoInt(copy(DM_Tabelas.cdSPagParcelaordem.Value,1,3))>=13) and (passou='N') then
        begin
          // adiciona 12 meses
//          data:=IncMonth(DM_Tabelas.cdSPagParcelavencimento.Value,12);
          data:=DM_Tabelas.cdSPagParcelavencimento.Value;
          pr:=copy(datetostr(data),4,7);
          passou:='S';

        end;
        if (strtoInt(copy(DM_Tabelas.cdSPagParcelaordem.Value,1,3))>13) then
           pr:=copy(datetostr(data),4,7);


        DM_Tabelas.ZQRecebimento.close;
        DM_Tabelas.ZQRecebimento.SQL.Clear;

        DM_tabelas.ZQRecebimento.SQL.Add('INSERT INTO  Recebimento (idloteamento,documento,tipdoc,cliente,dt_entrada,dt_vencimento,ordem,'+
                                              'saldo,Valor,VrDoc,contabil,origem,adversa,nomeadversa,recpag,numordem,venda_idvenda,quadralote,Observ,'+
                                              'Parcelas_fixas,Proximo_Reajuste) VALUES ('+
                                              quotedstr(inttostr(DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger))+','+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-P- '+CBPeriodoParc.Text)+','+quotedstr(DM_Tabelas.cdSPagParcelatipodoc.Value)+','+
                                              quotedstr(inttostr(DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').AsLargeInt))+','+quotedstr(FormatDateTime('YYYY-MM-DD',date))+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.cdSPagParcelavencimento.Value))+','+
                                              quotedstr(inttostr(varnumordemP)+'-'+DM_Tabelas.cdSPagParcelaordem.Value)+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr((DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').AsFloat/100)),'.')))+','+
                                              quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').AsFloat/100),'.')))+','+
                                              quotedstr(TrocaVirgPPto(TrimChar(floattostr(XNEEntrada.Value),'.')))+','+quotedstr(inttostr(DM_Tabelas.cdSPagParcelacontabil.Value))+','+quotedstr('V-'+inttostr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt))+','+
                                              quotedstr(inttostr(DM_Tabelas.CDSCompradorTempcodparticipante.Value))+','+quotedstr(DM_Tabelas.CDSCompradorTempnomeparticipante.Value)+',''R'','+quotedstr(inttostr(varnumordemP))+','+quotedstr(inttostr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt))+','+
                                              quotedstr(DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+','+quotedstr('Empreendimento: '+DM_Tabelas.ZQVenda.FieldByName('nometoeam').AsString+chr(13)+'Quadra: '+DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+chr(13)+'Lote: '+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+','+
                                              quotedstr(fx)+','+quotedstr(pr)+')');

        DM_tabelas.ZQRecebimento.ExecSQL;



{        DM_Tabelas.ZQRecebimento.Insert;
        DM_Tabelas.ZQRecebimentoidloteamento.Value := DM_Tabelas.ZQVendacodloteamento.Value;
        DM_Tabelas.ZQRecebimentodocumento.Value := DM_Tabelas.ZQVendaidvenda.Text+'-P- '+CBPeriodoParc.Text;
        DM_Tabelas.ZQRecebimentoTipDoc.Value := DM_Tabelas.cdSPagParcelatipodoc.Value;
        DM_Tabelas.ZQRecebimentocliente.Value := DM_Tabelas.ZQincorp_loteameincorporador_idincorporador.Value;
        DM_Tabelas.ZQRecebimentoDt_Entrada.Value := date;
        DM_Tabelas.ZQRecebimentoDt_Vencimento.Value := DM_Tabelas.cdSPagParcelavencimento.Value;
        DM_Tabelas.ZQRecebimentoordem.Value := inttostr(varnumordemP)+'-'+DM_Tabelas.cdSPagParcelaordem.Value;

        DM_Tabelas.ZQRecebimentosaldo.Value := DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteamepercent_parte.Value/100;
        DM_Tabelas.ZQRecebimentoValor.Value := DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteamepercent_parte.Value/100;

        DM_Tabelas.ZQRecebimentoVrDoc.Value := XNEEntrada.Value;
        DM_Tabelas.ZQRecebimentocontabil.Value := DM_Tabelas.cdSPagParcelacontabil.Value;
        DM_Tabelas.ZQRecebimentoorigem.Value := 'V-'+DM_Tabelas.ZQVendaidvenda.Text;

        if DM_Tabelas.CDSCompradorTemp.recordcount>1 then
        begin
          DM_Tabelas.CDSCompradorTemp.Filtered:=false;
          DM_Tabelas.CDSCompradorTemp.Filter:='marcar=''true''';
          DM_Tabelas.CDSCompradorTemp.Filtered:=true;
        end;

        DM_Tabelas.CDSCompradorTemp.First;
        DM_Tabelas.ZQRecebimentoadversa.Value := DM_Tabelas.CDSCompradorTempcodparticipante.Value;
        DM_Tabelas.ZQRecebimentonomeadversa.Value := DM_Tabelas.CDSCompradorTempnomeparticipante.Value;


        DM_Tabelas.ZQRecebimentorecpag.Value := 'R';
        DM_Tabelas.ZQRecebimentonumordem.Value := varnumordemP;
        DM_Tabelas.ZQRecebimentoquadralote.Value := DM_Tabelas.ZQVendaquadra.Value+'-'+DM_Tabelas.ZQVendalote.Value;
        DM_Tabelas.ZQRecebimentovenda_idvenda.Value := DM_Tabelas.ZQVendaidvenda.Value;
        DM_Tabelas.ZQRecebimentoObserv.Value := 'Empreendimento: '+DM_Tabelas.ZQVendanometoeam.Value+chr(13);
        DM_Tabelas.ZQRecebimentoObserv.Value := DM_Tabelas.ZQRecebimentoObserv.Value+'Quadra: '+DM_Tabelas.ZQVendaquadra.Value+chr(13);
        DM_Tabelas.ZQRecebimentoObserv.Value := DM_Tabelas.ZQRecebimentoObserv.Value+'Lote: '+DM_Tabelas.ZQVendalote.Value;
        if DM_Tabelas.ZQVendaforma_reajuste.Value='IGPM - FGV' then
           DM_Tabelas.ZQRecebimentoParcelas_fixas.Value := 'N'
        else
           DM_Tabelas.ZQRecebimentoParcelas_fixas.Value := 'S';

        if (DM_Tabelas.ZQRecebimentoParcelas_fixas.Value='N') and (strtoInt(copy(DM_Tabelas.cdSPagParcelaordem.Value,1,3))>=13) and (passou='N') then
        begin
          // adiciona 12 meses
//          data:=IncMonth(DM_Tabelas.cdSPagParcelavencimento.Value,12);
          data:=DM_Tabelas.cdSPagParcelavencimento.Value;
          DM_Tabelas.ZQRecebimentoProximo_Reajuste.value:=copy(datetostr(data),4,7);
          passou:='S';
        end;
        if (strtoInt(copy(DM_Tabelas.cdSPagParcelaordem.Value,1,3))>13) then
           DM_Tabelas.ZQRecebimentoProximo_Reajuste.value:=copy(datetostr(data),4,7);

        DM_Tabelas.ZQRecebimento.Post;
 }

        qlote:=DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString;
        DM_Tabelas.ZQRecebimento.close;
        ZQRecebimento.close;
        ZQRecebimento.SQL.Clear;
        ZQRecebimento.SQL.Add('Select idrecebimento,quadralote from Recebimento where quadralote='+quotedstr(qlote)+' order by idrecebimento');
        ZQRecebimento.Open;
        ZQRecebimento.last;
        varidrecebimento:=ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
        ZQRecebimento.Locate('idrecebimento',varidrecebimento,[]);



        DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagParcelatipodoc.Value,[]);
        if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
          DM_tabelas.ZQCheque.Insert;
          DM_Tabelas.ZQCheque.FieldByName('idrecebimento').AsInteger := ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
          DM_Tabelas.ZQCheque.FieldByName('idvinculo').AsInteger := ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
          DM_Tabelas.ZQCheque.FieldByName('Banco').AsString := DM_Tabelas.cdSPagParcelabanco.Value;
          DM_Tabelas.ZQCheque.FieldByName('agencia').AsString := DM_Tabelas.cdSPagParcelaagencia.Value;
          DM_Tabelas.ZQCheque.FieldByName('numero').AsString := DM_Tabelas.CDSPagParcelanumero.Value;
          DM_Tabelas.ZQCheque.FieldByName('Dono').AsString := DM_Tabelas.cdSPagParcelanomedono.Value;
          DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').AsString := DM_Tabelas.cdSPagParceladocumento.Value;
          DM_Tabelas.ZQCheque.FieldByName('valor').AsFloat := DM_Tabelas.cdSPagParcelavalor.Value;
          DM_Tabelas.ZQCheque.FieldByName('emissao').AsDateTime := Date;
          DM_Tabelas.ZQCheque.FieldByName('deposito').AsDateTime := DM_Tabelas.cdSPagParcelavencimento.Value;
          DM_Tabelas.ZQCheque.Post;
        end;
        DM_Tabelas.cdSPagParcela.Next;
      end;
      Bar2.Visible:=false;
      Bar2.Position:=0;
      DM_Tabelas.CDSPagParcela.EnableControls;
      DM_Tabelas.ZQincorp_loteame.Next;
    end;
  end;  
  if edicao='S' then
  begin
    if simnao(' Deseja que o financeiro seja alterado ?','SIM') then
    begin
      DM_Tabelas.ZQRecebimento.SQL.Clear;
      DM_Tabelas.ZQRecebimento.SQL.Add('delete from recebimento');
      DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
      DM_Tabelas.ZQRecebimento.ExecSQL;
      DM_Tabelas.ZQRecebimento.SQL.Clear;
      DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
      DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
      DM_Tabelas.ZQRecebimento.SQL.Add(' from recebimento');
      DM_Tabelas.ZQRecebimento.Open;


      DM_Tabelas.ZQincorp_loteame.first;

    // para inibir duplicar entrada qdo tem 2 donos o lote. ex. jardim canaa
      entrada:='N';
      passou:='N';
      while not DM_Tabelas.ZQincorp_loteame.eof do
      begin
        DM_Tabelas.CDSCompradorTemp.Filtered:=false;
        DM_Tabelas.CDSCompradorTemp.First;
        DM_Tabelas.CDSPagEntrada.First;

        if entrada='N' then
        begin

         DM_Tabelas.CDSPagEntrada.DisableControls;
         Bar1.Visible:=true;
         Bar1.Position:=0;
         Bar1.Max:=DM_Tabelas.CDSPagEntrada.RecordCount;
         while not DM_Tabelas.CDSPagEntrada.Eof do
         Begin
           Bar1.Position:=DM_Tabelas.CDSPagEntrada.RecNo;
          if DM_Tabelas.CDSCompradorTemp.recordcount>1 then
          begin
            DM_Tabelas.CDSCompradorTemp.Filtered:=false;
            DM_Tabelas.CDSCompradorTemp.Filter:='marcar=''true''';
            DM_Tabelas.CDSCompradorTemp.Filtered:=true;
          end;

          if fp.ItemIndex=1 then
             fx:='N'
          else
             fx:='S';

          DM_Tabelas.ZQRecebimento.close;
          DM_Tabelas.ZQRecebimento.SQL.Clear;

          DM_tabelas.ZQRecebimento.SQL.Add('INSERT INTO  Recebimento (idloteamento,documento,tipdoc,cliente,dt_entrada,dt_vencimento,ordem,'+
                                                'saldo,Valor,VrDoc,contabil,origem,adversa,nomeadversa,recpag,numordem,venda_idvenda,quadralote,Observ,'+
                                                'Parcelas_fixas) VALUES ('+
                                                quotedstr(inttostr(DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger))+','+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-E- '+CBPeriodoEnt.Text)+','+quotedstr(DM_Tabelas.CDSPagEntradatipodoc.Value)+','+
                                                quotedstr(inttostr(DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').AsLargeInt))+','+quotedstr(FormatDateTime('YYYY-MM-DD',date))+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.CDSPagEntradavencimento.Value))+','+
                                                quotedstr(inttostr(varnumordemE)+'-'+DM_Tabelas.CDSPagEntradaordem.Value)+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_Tabelas.CDSPagEntradavalor.Value),'.')))+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_Tabelas.CDSPagEntradavalor.Value),'.')))+','+
                                                quotedstr(TrocaVirgPPto(TrimChar(floattostr(XNEEntrada.Value),'.')))+','+quotedstr(inttostr(DM_Tabelas.CDSPagEntradacontabil.Value))+','+quotedstr('V-'+inttostr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt))+','+
                                                quotedstr(inttostr(DM_Tabelas.CDSCompradorTempcodparticipante.Value))+','+quotedstr(DM_Tabelas.CDSCompradorTempnomeparticipante.Value)+',''R'','+quotedstr(inttostr(varnumordemE))+','+quotedstr(inttostr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt))+','+
                                                quotedstr(DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+','+quotedstr('Empreendimento: '+DM_Tabelas.ZQVenda.FieldByName('nometoeam').AsString+chr(13)+'Quadra: '+DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+chr(13)+'Lote: '+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+','+
                                                quotedstr(fx)+')');

          DM_tabelas.ZQRecebimento.ExecSQL;


         { DM_Tabelas.ZQRecebimento.Insert;
          DM_Tabelas.ZQRecebimentoidloteamento.Value := DM_Tabelas.ZQVendacodloteamento.Value;
          DM_Tabelas.ZQRecebimentodocumento.Value := DM_Tabelas.ZQVendaidvenda.Text+'-E- '+CBPeriodoEnt.Text;
          DM_Tabelas.ZQRecebimentoTipDoc.Value := DM_Tabelas.CDSPagEntradatipodoc.Value;
          DM_Tabelas.ZQRecebimentocliente.Value := DM_Tabelas.ZQincorp_loteameincorporador_idincorporador.Value;
          DM_Tabelas.ZQRecebimentoDt_Entrada.Value := date;
          DM_Tabelas.ZQRecebimentoDt_Vencimento.Value := DM_Tabelas.CDSPagEntradavencimento.Value;
          DM_Tabelas.ZQRecebimentoordem.Value := inttostr(varnumordemE)+'-'+DM_Tabelas.CDSPagEntradaordem.Value;

          //  DM_Tabelas.ZQRecebimentosaldo.Value := DM_Tabelas.CDSPagEntradavalor.Value*DM_Tabelas.ZQincorp_loteamepercent_parte.Value/100;
          //      DM_Tabelas.ZQRecebimentoValor.Value := DM_Tabelas.CDSPagEntradavalor.Value*DM_Tabelas.ZQincorp_loteamepercent_parte.Value/100;

          DM_Tabelas.ZQRecebimentosaldo.Value := DM_Tabelas.CDSPagEntradavalor.Value;
          DM_Tabelas.ZQRecebimentoValor.Value := DM_Tabelas.CDSPagEntradavalor.Value;

          DM_Tabelas.ZQRecebimentoVrDoc.Value := XNEEntrada.Value;
          DM_Tabelas.ZQRecebimentocontabil.Value := DM_Tabelas.CDSPagEntradacontabil.Value;
          DM_Tabelas.ZQRecebimentoorigem.Value := 'V-'+DM_Tabelas.ZQVendaidvenda.Text;

          if DM_Tabelas.CDSCompradorTemp.RecordCount>1 then
          begin
            DM_Tabelas.CDSCompradorTemp.Filtered:=false;
            DM_Tabelas.CDSCompradorTemp.Filter:='marcar=''true''';
            DM_Tabelas.CDSCompradorTemp.Filtered:=true;
          end;

          DM_Tabelas.CDSCompradorTemp.First;
          DM_Tabelas.ZQRecebimentoadversa.Value := DM_Tabelas.CDSCompradorTempcodparticipante.Value;
          DM_Tabelas.ZQRecebimentonomeadversa.Value := DM_Tabelas.CDSCompradorTempnomeparticipante.Value;


          DM_Tabelas.ZQRecebimentorecpag.Value := 'R';
          DM_Tabelas.ZQRecebimentonumordem.Value := varnumordemE;
          DM_Tabelas.ZQRecebimentovenda_idvenda.Value := DM_Tabelas.ZQVendaidvenda.Value;
          DM_Tabelas.ZQRecebimentoquadralote.Value := DM_Tabelas.ZQVendaquadra.Value+'-'+DM_Tabelas.ZQVendalote.Value;
          DM_Tabelas.ZQRecebimentoObserv.Value := 'Empreendimento: '+DM_Tabelas.ZQVendanometoeam.Value+chr(13);
          DM_Tabelas.ZQRecebimentoObserv.Value := DM_Tabelas.ZQRecebimentoObserv.Value+'Quadra: '+DM_Tabelas.ZQVendaquadra.Value+chr(13);
          DM_Tabelas.ZQRecebimentoObserv.Value := DM_Tabelas.ZQRecebimentoObserv.Value+'Lote: '+DM_Tabelas.ZQVendalote.Value;
          if DM_Tabelas.ZQVendaforma_reajuste.Value='IGPM - FGV' then
           DM_Tabelas.ZQRecebimentoParcelas_fixas.Value := 'N'
          else
             DM_Tabelas.ZQRecebimentoParcelas_fixas.Value := 'S';

          DM_Tabelas.ZQRecebimento.Post;
          }
          qlote:=DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString;
          DM_Tabelas.ZQRecebimento.close;
          ZQRecebimento.close;
          ZQRecebimento.SQL.Clear;
          ZQRecebimento.SQL.Add('Select idrecebimento,quadralote from Recebimento where quadralote='+quotedstr(qlote)+' order by idrecebimento');
          ZQRecebimento.Open;
          ZQRecebimento.last;
          varidrecebimento:=ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
          ZQRecebimento.Locate('idrecebimento',varidrecebimento,[]);



          DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagEntradatipodoc.Value,[]);
          if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
            DM_tabelas.ZQCheque.Insert;
            DM_Tabelas.ZQCheque.FieldByName('idrecebimento').AsInteger := ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
            DM_Tabelas.ZQCheque.FieldByName('idvinculo').AsInteger := ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
            DM_Tabelas.ZQCheque.FieldByName('Banco').AsString := DM_Tabelas.CDSPagEntradabanco.Value;
            DM_Tabelas.ZQCheque.FieldByName('agencia').AsString := DM_Tabelas.CDSPagEntradaagencia.Value;
            DM_Tabelas.ZQCheque.FieldByName('numero').AsString := DM_Tabelas.CDSPagEntradanumero.Value;
            DM_Tabelas.ZQCheque.FieldByName('conta').AsString := DM_Tabelas.CDSPagEntradaconta.Value;
            DM_Tabelas.ZQCheque.FieldByName('Dono').AsString := DM_Tabelas.CDSPagEntradanomedono.Value;
            DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').AsString := DM_Tabelas.CDSPagEntradadocumento.Value;
            DM_Tabelas.ZQCheque.FieldByName('valor').AsFloat := DM_Tabelas.CDSPagEntradavalor.Value;
            DM_Tabelas.ZQCheque.FieldByName('emissao').AsDateTime := Date;
            DM_Tabelas.ZQCheque.FieldByName('deposito').AsDateTime := DM_Tabelas.CDSPagEntradavencimento.Value;
            DM_Tabelas.ZQCheque.Post;
          end;
          DM_Tabelas.CDSPagEntrada.Next;
         end;
         DM_Tabelas.CDSPagEntrada.EnableControls;
         Bar1.Visible:=false;
         Bar1.Position:=0;

         entrada:='S';
         // até aqui
        end;
        DM_Tabelas.CDSPagParcela.First;


        DM_Tabelas.CDSPagParcela.DisableControls;
        Bar2.Visible:=true;
        Bar2.Position:=0;
        Bar2.Max:=DM_Tabelas.cdSPagParcela.RecordCount;

        while not DM_Tabelas.cdSPagParcela.Eof do
        Begin
          bar2.Position:=DM_Tabelas.cdSPagParcela.RecNo;
          if DM_Tabelas.CDSCompradorTemp.recordcount>1 then
          begin
            DM_Tabelas.CDSCompradorTemp.Filtered:=false;
            DM_Tabelas.CDSCompradorTemp.Filter:='marcar=''true''';
            DM_Tabelas.CDSCompradorTemp.Filtered:=true;
          end;

          if fp.ItemIndex=1 then
             fx:='N'
          else
             fx:='S';


          if (fx='N') and (strtoInt(copy(DM_Tabelas.cdSPagParcelaordem.Value,1,3))>=13) and (passou='N') then
          begin
            // adiciona 12 meses
  //          data:=IncMonth(DM_Tabelas.cdSPagParcelavencimento.Value,12);
            data:=DM_Tabelas.cdSPagParcelavencimento.Value;
            pr:=copy(datetostr(data),4,7);
            passou:='S';

          end;
          if (strtoInt(copy(DM_Tabelas.cdSPagParcelaordem.Value,1,3))>13) then
             pr:=copy(datetostr(data),4,7);


          DM_Tabelas.ZQRecebimento.close;
          DM_Tabelas.ZQRecebimento.SQL.Clear;

          DM_tabelas.ZQRecebimento.SQL.Add('INSERT INTO  Recebimento (idloteamento,documento,tipdoc,cliente,dt_entrada,dt_vencimento,ordem,'+
                                                'saldo,Valor,VrDoc,contabil,origem,adversa,nomeadversa,recpag,numordem,venda_idvenda,quadralote,Observ,'+
                                                'Parcelas_fixas,Proximo_Reajuste) VALUES ('+
                                                quotedstr(inttostr(DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger))+','+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-P- '+CBPeriodoParc.Text)+','+quotedstr(DM_Tabelas.cdSPagParcelatipodoc.Value)+','+
                                                quotedstr(inttostr(DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').AsLargeInt))+','+quotedstr(FormatDateTime('YYYY-MM-DD',date))+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.cdSPagParcelavencimento.Value))+','+
                                                quotedstr(inttostr(varnumordemP)+'-'+DM_Tabelas.cdSPagParcelaordem.Value)+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr((DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').AsFloat/100)),'.')))+','+
                                                quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').AsFloat/100),'.')))+','+
                                                quotedstr(TrocaVirgPPto(TrimChar(floattostr(XNEEntrada.Value),'.')))+','+quotedstr(inttostr(DM_Tabelas.cdSPagParcelacontabil.Value))+','+quotedstr('V-'+inttostr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt))+','+
                                                quotedstr(inttostr(DM_Tabelas.CDSCompradorTempcodparticipante.Value))+','+quotedstr(DM_Tabelas.CDSCompradorTempnomeparticipante.Value)+',''R'','+quotedstr(inttostr(varnumordemP))+','+quotedstr(inttostr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt))+','+
                                                quotedstr(DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+','+quotedstr('Empreendimento: '+DM_Tabelas.ZQVenda.FieldByName('nometoeam').AsString+chr(13)+'Quadra: '+DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+chr(13)+'Lote: '+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+','+
                                                quotedstr(fx)+','+quotedstr(pr)+')');

          DM_tabelas.ZQRecebimento.ExecSQL;


{          DM_Tabelas.ZQRecebimento.Insert;
          DM_Tabelas.ZQRecebimentoidloteamento.Value := DM_Tabelas.ZQVendacodloteamento.Value;
          DM_Tabelas.ZQRecebimentodocumento.Value := DM_Tabelas.ZQVendaidvenda.Text+'-P- '+CBPeriodoParc.Text;
          DM_Tabelas.ZQRecebimentoTipDoc.Value := DM_Tabelas.cdSPagParcelatipodoc.Value;
          DM_Tabelas.ZQRecebimentocliente.Value := DM_Tabelas.ZQincorp_loteameincorporador_idincorporador.Value;
          DM_Tabelas.ZQRecebimentoDt_Entrada.Value := date;
          DM_Tabelas.ZQRecebimentoDt_Vencimento.Value := DM_Tabelas.cdSPagParcelavencimento.Value;
          DM_Tabelas.ZQRecebimentoordem.Value := inttostr(varnumordemP)+'-'+DM_Tabelas.cdSPagParcelaordem.Value;

          DM_Tabelas.ZQRecebimentosaldo.Value := DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteamepercent_parte.Value/100;
          DM_Tabelas.ZQRecebimentoValor.Value := DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteamepercent_parte.Value/100;

          DM_Tabelas.ZQRecebimentoVrDoc.Value := XNEEntrada.Value;
          DM_Tabelas.ZQRecebimentocontabil.Value := DM_Tabelas.cdSPagParcelacontabil.Value;
          DM_Tabelas.ZQRecebimentoorigem.Value := 'V-'+DM_Tabelas.ZQVendaidvenda.Text;

          if DM_Tabelas.CDSCompradorTemp.RecordCount>1 then
          begin
            DM_Tabelas.CDSCompradorTemp.Filtered:=false;
            DM_Tabelas.CDSCompradorTemp.Filter:='marcar=''true''';
            DM_Tabelas.CDSCompradorTemp.Filtered:=true;
          end;

          DM_Tabelas.CDSCompradorTemp.First;
          DM_Tabelas.ZQRecebimentoadversa.Value := DM_Tabelas.CDSCompradorTempcodparticipante.Value;
          DM_Tabelas.ZQRecebimentonomeadversa.Value := DM_Tabelas.CDSCompradorTempnomeparticipante.Value;


          DM_Tabelas.ZQRecebimentorecpag.Value := 'R';
          DM_Tabelas.ZQRecebimentonumordem.Value := varnumordemP;
          DM_Tabelas.ZQRecebimentoquadralote.Value := DM_Tabelas.ZQVendaquadra.Value+'-'+DM_Tabelas.ZQVendalote.Value;
          DM_Tabelas.ZQRecebimentovenda_idvenda.Value := DM_Tabelas.ZQVendaidvenda.Value;
          DM_Tabelas.ZQRecebimentoObserv.Value := 'Empreendimento: '+DM_Tabelas.ZQVendanometoeam.Value+chr(13);
          DM_Tabelas.ZQRecebimentoObserv.Value := DM_Tabelas.ZQRecebimentoObserv.Value+'Quadra: '+DM_Tabelas.ZQVendaquadra.Value+chr(13);
          DM_Tabelas.ZQRecebimentoObserv.Value := DM_Tabelas.ZQRecebimentoObserv.Value+'Lote: '+DM_Tabelas.ZQVendalote.Value;
          if DM_Tabelas.ZQVendaforma_reajuste.Value='IGPM - FGV' then
             DM_Tabelas.ZQRecebimentoParcelas_fixas.Value := 'N'
          else
             DM_Tabelas.ZQRecebimentoParcelas_fixas.Value := 'S';

          if (DM_Tabelas.ZQRecebimentoParcelas_fixas.Value='N') and (strtoInt(copy(DM_Tabelas.cdSPagParcelaordem.Value,1,3))>=13) and (passou='N') then
          begin
            // adiciona 12 meses
  //          data:=IncMonth(DM_Tabelas.cdSPagParcelavencimento.Value,12);
            data:=DM_Tabelas.cdSPagParcelavencimento.Value;
            DM_Tabelas.ZQRecebimentoProximo_Reajuste.value:=copy(datetostr(data),4,7);
            passou:='S';
          end;
          if (strtoInt(copy(DM_Tabelas.cdSPagParcelaordem.Value,1,3))>13) then
             DM_Tabelas.ZQRecebimentoProximo_Reajuste.value:=copy(datetostr(data),4,7);

          DM_Tabelas.ZQRecebimento.Post;}



          qlote:=DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString;
          DM_Tabelas.ZQRecebimento.close;
          ZQRecebimento.close;
          ZQRecebimento.SQL.Clear;
          ZQRecebimento.SQL.Add('Select idrecebimento,quadralote from Recebimento where quadralote='+quotedstr(qlote)+' order by idrecebimento');
          ZQRecebimento.Open;
          ZQRecebimento.last;
          varidrecebimento:=ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
          ZQRecebimento.Locate('idrecebimento',varidrecebimento,[]);


          DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagParcelatipodoc.Value,[]);
          if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
            DM_tabelas.ZQCheque.Insert;
            DM_Tabelas.ZQCheque.FieldByName('idrecebimento').AsInteger := ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
            DM_Tabelas.ZQCheque.FieldByName('idvinculo').AsInteger := ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
            DM_Tabelas.ZQCheque.FieldByName('Banco').AsString := DM_Tabelas.cdSPagParcelabanco.Value;
            DM_Tabelas.ZQCheque.FieldByName('agencia').AsString := DM_Tabelas.cdSPagParcelaagencia.Value;
            DM_Tabelas.ZQCheque.FieldByName('numero').AsString := DM_Tabelas.CDSPagParcelanumero.Value;
            DM_Tabelas.ZQCheque.FieldByName('Dono').AsString := DM_Tabelas.cdSPagParcelanomedono.Value;
            DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').AsString := DM_Tabelas.cdSPagParceladocumento.Value;
            DM_Tabelas.ZQCheque.FieldByName('valor').AsFloat := DM_Tabelas.cdSPagParcelavalor.Value;
            DM_Tabelas.ZQCheque.FieldByName('emissao').AsDateTime := Date;
            DM_Tabelas.ZQCheque.FieldByName('deposito').AsDateTime := DM_Tabelas.cdSPagParcelavencimento.Value;
            DM_Tabelas.ZQCheque.Post;
          end;
          DM_Tabelas.cdSPagParcela.Next;
        end;
        DM_Tabelas.CDSPagParcela.DisableControls;
        Bar2.Visible:=false;
        Bar2.Position:=0;

        DM_Tabelas.ZQincorp_loteame.Next;
      end;
    end;
  end;

 {  DM_Tabelas.ZQRecebimento.close;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('select * from recebimento');
  DM_Tabelas.ZQRecebimento.Open;}
  DM_Tabelas.CDSCompradorTemp.Filtered:=false;
  DM_Tabelas.CDSCompradorTemp.First;

  DesativaCampos;

  DM_Tabelas.ZQRecebimento.close;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
  DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento order by DT_Vencimento limit 0,100');
  DM_Tabelas.ZQRecebimento.open;

  DBGVenda.SetFocus;
  varvenda := DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
  DM_TAbelas.ZQVenda.Filter := 'idvenda='+inttostr(varvenda);
  DM_Tabelas.ZQVenda.Filtered := True;
  DM_TAbelas.ZQCompr_Dados.Open;
  DM_TAbelas.ZQCompr_Dados.Refresh;
//  if Frm_QuadroResumo=nil then
//    Frm_QuadroResumo:=TFrm_QuadroResumo.Create(Application);
//  Frm_QuadroResumo.RLReport1.Preview;
  DM_Tabelas.ZQVenda.Filtered := False;
  edicao:='N';
END;

procedure TFrm_Venda.BtCancelarClick(Sender: TObject);
begin
  DM_TAbelas.ZQmemorial.cancel;
  FP.Visible:=false;
  DM_Tabelas.ZQVenda.Cancel;
  DesativaCampos;
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.BtProcessarClick(Sender: TObject);
begin
  JvXPBar2.Visible:=true;
end;

procedure TFrm_Venda.BtIncluirClick(Sender: TObject);
begin
    if not Verif_senha('Venda','Incluir ','') then exit;
    FP.Visible:=true;
    DM_Tabelas.ZQVenda.Insert;
    EloteTemp.Text := '';
    AtivaCampos;
    Pag_Venda.PageIndex := 0;
    DM_TAbelas.CDSCompradorTemp.Close;
    DM_Tabelas.CDSCompradorTemp.CreateDataSet;
    DM_Tabelas.CDSVendedorTemp.Close;
    DM_Tabelas.CDSVendedorTemp.CreateDataSet;
    DM_Tabelas.CDSPagEntrada.Close;
    DM_Tabelas.CDSPagEntrada.CreateDataSet;
    DM_Tabelas.CDSPagParcela.Close;
    DM_Tabelas.CDSPagParcela.CreateDataSet;
    DM_Tabelas.ZQVenda.FieldByName('datavenda').AsDateTime := date;
    XNEEntrada.Value := 0;
    XNQtdeEnt.Value := 0;
    XNQtdeParc.Value := 0;
    XNQtdeParcAtrz.Value := 0;
    XVencimentoEnt.DateValue := date;
    XVencimentoParc.DateValue := date;
    CBPeriodoEnt.ItemIndex := 3;
    CBPeriodoParc.ItemIndex := 3;
    DBEDAta.SetFocus;
end;

procedure TFrm_Venda.BtEditarClick(Sender: TObject);
var
tv:string;
begin
  if not Verif_senha('Venda','Editar ','Venda Nº: '+DM_tabelas.ZQVenda.FieldByName('idvenda').Text) then exit;
  if Pag_Venda.PageIndex=3 then
  begin
      DM_TAbelas.ZQmemorial.Edit;
      AtivaCampos;
  end
  else
  begin
    FP.Visible:=true;
    Panel1.Visible:=true;
    Application.ProcessMessages;
    Atualizatebelaspagamento;
    Panel1.Visible:=false;
    Application.ProcessMessages;


    tv:=copy(alltrim(DM_Tabelas.ZQVenda.FieldByName('forma_reajuste').AsString),1,4);
    if tv='IGPM' then
       fp.ItemIndex:=1
    else
       fp.ItemIndex:=0;

    DM_Tabelas.ZQVenda.Edit;
    EloteTemp.Text := DM_tabelas.ZQVenda.FieldByName('imovel').Text;
    AtivaCampos;
    if Pag_Venda.PageIndex = 0 Then
      DBEDAta.SetFocus;

  end;
end;

procedure TFrm_Venda.BTExcluirClick(Sender: TObject);
var
cd:integer;
begin
  if not Verif_senha('Venda','Excluir ','Venda Nº: '+DM_tabelas.ZQVenda.FieldByName('idvenda').Text) then exit;
  if SIMNAO('Deseja Realmente Excluir esse Registro ?','SIM') then
  begin
    if Pag_Venda.PageIndex<>4 then
    begin
      if SIMNAO('É Rescisão de Contrato ?','SIM') then
      begin
        gbrescisao.Visible:=true;
        mdata.SetFocus;
        exit;
      end
      else
      begin
        DM_Tabelas.ZQRecebimento.close;
        DM_Tabelas.ZQRecebimento.SQL.Clear;
        DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
        DM_Tabelas.ZQRecebimento.SQL.Add(' from recebimento where documento like '+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-%'));
        DM_Tabelas.ZQRecebimento.open;
        if DM_Tabelas.ZQRecebimento.RecordCount>0 then
        begin
          Gauge1.MaxValue:=DM_Tabelas.ZQRecebimento.RecordCount;
          Gauge1.Progress:=0;
          Gauge1.Visible:=true;
          while not DM_Tabelas.ZQRecebimento.Eof do
          begin
            Gauge1.Progress:=DM_Tabelas.ZQRecebimento.RecNo;
            DM_Tabelas.ZQReBxHi.SQL.clear;
            DM_Tabelas.ZQReBxHi.SQL.Add('delete from recbxhist where idrecib='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
            DM_Tabelas.ZQReBxHi.ExecSQL;

            DM_Tabelas.ZQReceb_Baixa.SQL.clear;
            DM_Tabelas.ZQReceb_Baixa.SQL.Add('delete from Receb_Baixa where refbaixa='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('RefBaixa').Text));
            DM_Tabelas.ZQReceb_Baixa.ExecSQL;

            ZQAditamento.SQL.clear;
            ZQAditamento.SQL.Add('delete from aditamento where idvenda='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').Text));
            ZQAditamento.ExecSQL;

            ZQQuitacao.SQL.clear;
            ZQQuitacao.SQL.Add('delete from quitacao where idvenda='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').Text));
            ZQQuitacao.ExecSQL;

            ZQcomprador_cessao.SQL.clear;
            ZQcomprador_cessao.SQL.Add('delete from comprador_cessao where idvenda='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').Text));
            ZQcomprador_cessao.ExecSQL;


            ZQrecebimento_historico_excluidos.SQL.clear;
            ZQrecebimento_historico_excluidos.SQL.Add('delete from recebimento_historico_excluidos where idrecebimento='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
            ZQrecebimento_historico_excluidos.ExecSQL;

            ZQrecebimento_historico_cessao.SQL.clear;
            ZQrecebimento_historico_cessao.SQL.Add('delete from recebimento_historico_cessao where idrecebimento='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
            ZQrecebimento_historico_cessao.ExecSQL;


            DM_Tabelas.ZQCheque.Close;
            DM_Tabelas.ZQCheque.SQL.Clear;
            DM_Tabelas.ZQCheque.SQL.Add('delete from cheque_rec where idvinculo='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
            DM_Tabelas.ZQCheque.ExecSQL;


            DM_Tabelas.ZQRecebimento.Next;
          end;
        end;
        Gauge1.Progress:=0;
        Gauge1.Visible:=false;

        DM_Tabelas.ZQRecebimento.close;
        DM_Tabelas.ZQRecebimento.SQL.Clear;
        DM_Tabelas.ZQRecebimento.SQL.Add('delete from recebimento where documento like '+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-%'));
        DM_Tabelas.ZQRecebimento.ExecSQL;
        DM_Tabelas.ZQRecebimento.SQL.Clear;
        DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
        DM_Tabelas.ZQRecebimento.SQL.Add(' from recebimento ');
        DM_Tabelas.ZQRecebimento.Open;

        DM_Tabelas.ZQReBxHi.SQL.Clear;
        DM_Tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from recbxhist ');
        DM_Tabelas.ZQReBxHi.Open;

        DM_Tabelas.ZQReceb_Baixa.SQL.Clear;
        DM_Tabelas.ZQReceb_Baixa.SQL.Add('select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento,substituicao,vr_ab,sq,valor_parcela from Receb_Baixa ');
        DM_Tabelas.ZQReceb_Baixa.Open;



        DM_Tabelas.ZQComprador.close;
        DM_Tabelas.ZQComprador.SQL.Clear;
        DM_Tabelas.ZQComprador.SQL.Add('delete from comprador where  venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text));
        DM_Tabelas.ZQComprador.ExecSQL;
        DM_Tabelas.ZQComprador.SQL.Clear;
        DM_Tabelas.ZQComprador.SQL.Add('select idcomprador,paticipante_idpaticipante,venda_idvenda,promissario,percentual,Marcar from comprador ');
        DM_Tabelas.ZQComprador.Open;

        DM_Tabelas.ZQVendedor.close;
        DM_Tabelas.ZQVendedor.SQL.Clear;
        DM_Tabelas.ZQVendedor.SQL.Add('delete from Vendedor where  venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text));
        DM_Tabelas.ZQVendedor.ExecSQL;
        DM_Tabelas.ZQVendedor.SQL.Clear;
        DM_Tabelas.ZQVendedor.SQL.Add('select idvendedor,corretor_idcorretor,venda_idvenda,PComissao_corretor from Vendedor ');
        DM_Tabelas.ZQVendedor.Open;

        DM_Tabelas.ZQImovel.Locate('idimovel',DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt,[]);
        DM_Tabelas.ZQImovel.Edit;
        DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString := 'SIM';
        DM_Tabelas.ZQImovel.FieldByName('proposta').AsString := 'NºO';
        DM_Tabelas.ZQImovel.Post;

        if DM_TAbelas.ZQmemorial.active= false then
           DM_TAbelas.ZQmemorial.open;
        if DM_TAbelas.ZQmemorial.recordcount>0 then
           DM_TAbelas.ZQmemorial.delete;

        DM_Tabelas.ZQVenda.Delete;

        DM_Tabelas.ZQCheque.Close;
        DM_Tabelas.ZQCheque.SQL.Clear;
        DM_Tabelas.ZQCheque.SQL.Add('select * from cheque_rec ');
        DM_Tabelas.ZQCheque.Open;


        mensagem('Exclusão Efetuada Com Sucesso!');


      end;
    end;
    if Pag_Venda.PageIndex=4 then
    begin

      cd:=DM_Tabelas.ZQProposta.FieldByName('idproposta').AsLargeInt;
      DM_Tabelas.ZQproposta.Close;
      DM_Tabelas.ZQproposta.SQL.Clear;
      DM_Tabelas.ZQproposta.SQL.Add('SELECT * FROM proposta');
      DM_Tabelas.ZQproposta.open;
      DM_Tabelas.ZQproposta.Locate('idproposta',cd,[]);


      DM_Tabelas.ZQRecebimento_proposta.close;
      DM_Tabelas.ZQRecebimento_proposta.SQL.Clear;
      DM_Tabelas.ZQRecebimento_proposta.SQL.Add('select * from recebimento_proposta where venda_idvenda='+quotedstr(DM_Tabelas.ZQProposta.FieldByName('idproposta').Text));
      DM_Tabelas.ZQRecebimento_proposta.Open;
      DM_Tabelas.ZQRecebimento_proposta.First;
      while not DM_Tabelas.ZQRecebimento_proposta.Eof do
      begin
        DM_Tabelas.ZQCheque_proposta.Close;
        DM_Tabelas.ZQCheque_proposta.SQL.Clear;
        DM_Tabelas.ZQCheque_proposta.SQL.Add('delete from cheque_rec_proposta where idvinculo='+quotedstr(DM_Tabelas.ZQRecebimento_proposta.FieldByName('idrecebimento_proposta').Text));
        DM_Tabelas.ZQCheque_proposta.ExecSQL;
        DM_Tabelas.ZQRecebimento_proposta.Next;
      end;
      DM_Tabelas.ZQRecebimento_proposta.close;
      DM_Tabelas.ZQRecebimento_proposta.SQL.Clear;
      DM_Tabelas.ZQRecebimento_proposta.SQL.Add('select * from recebimento_proposta where venda_idvenda='+quotedstr(DM_Tabelas.ZQProposta.FieldByName('idproposta').Text));
      DM_Tabelas.ZQRecebimento_proposta.Open;
      DM_Tabelas.ZQRecebimento_proposta.First;
      if DM_Tabelas.ZQRecebimento_proposta.RecordCount>0 then
      begin
        DM_Tabelas.ZQRecebimento_proposta.close;
        DM_Tabelas.ZQRecebimento_proposta.SQL.Clear;
        DM_Tabelas.ZQRecebimento_proposta.SQL.Add('delete from recebimento_proposta where  venda_idvenda='+quotedstr(DM_Tabelas.ZQProposta.FieldByName('idproposta').Text));
        DM_Tabelas.ZQRecebimento_proposta.ExecSQL;
        DM_Tabelas.ZQRecebimento_proposta.close;
        DM_Tabelas.ZQRecebimento_proposta.SQL.Clear;
        DM_Tabelas.ZQRecebimento_proposta.SQL.Add('select * from recebimento_proposta ');
        DM_Tabelas.ZQRecebimento_proposta.Open;

      end;
      DM_Tabelas.ZQReBxHi.SQL.Clear;
      DM_Tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from recbxhist ');
      DM_Tabelas.ZQReBxHi.Open;

      DM_Tabelas.ZQReceb_Baixa.SQL.Clear;
      DM_Tabelas.ZQReceb_Baixa.SQL.Add('select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento,substituicao,vr_ab,sq,valor_parcela from Receb_Baixa ');
      DM_Tabelas.ZQReceb_Baixa.Open;

      DM_Tabelas.ZQImovel.Locate('idimovel',  DM_Tabelas.ZQProposta.FieldByName('imovel').AsLargeInt,[]);
      DM_Tabelas.ZQImovel.Edit;
      DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString := 'SIM';
      DM_Tabelas.ZQImovel.FieldByName('proposta').AsString := 'NºO';
      DM_Tabelas.ZQImovel.Post;

      DM_Tabelas.ZQComprador_proposta.close;
      DM_Tabelas.ZQComprador_proposta.SQL.Clear;
      DM_Tabelas.ZQComprador_proposta.SQL.Add('select * from comprador_proposta where  venda_idvenda='+quotedstr(  DM_Tabelas.ZQProposta.FieldByName('idproposta').Text));
      DM_Tabelas.ZQComprador_proposta.Open;
      if DM_Tabelas.ZQComprador_proposta.RecordCount>0 then
      begin
        DM_Tabelas.ZQComprador_proposta.close;
        DM_Tabelas.ZQComprador_proposta.SQL.Clear;
        DM_Tabelas.ZQComprador_proposta.SQL.Add('delete from comprador_proposta where  venda_idvenda='+quotedstr(  DM_Tabelas.ZQProposta.FieldByName('idproposta').Text));
        DM_Tabelas.ZQComprador_proposta.ExecSQL;
        DM_Tabelas.ZQComprador_proposta.SQL.Clear;
        DM_Tabelas.ZQComprador_proposta.SQL.Add('select * from comprador_proposta ');
        DM_Tabelas.ZQComprador_proposta.Open;
      end;

      DM_Tabelas.ZQVendedor_proposta.close;
      DM_Tabelas.ZQVendedor_proposta.SQL.Clear;
      DM_Tabelas.ZQVendedor_proposta.SQL.Add('select * from Vendedor_proposta where  venda_idvenda='+quotedstr(  DM_Tabelas.ZQProposta.FieldByName('idproposta').Text));
      DM_Tabelas.ZQVendedor_proposta.open;

      if DM_Tabelas.ZQVendedor_proposta.RecordCount>0 then
      begin
        DM_Tabelas.ZQVendedor_proposta.close;
        DM_Tabelas.ZQVendedor_proposta.SQL.Clear;
        DM_Tabelas.ZQVendedor_proposta.SQL.Add('delete from Vendedor_proposta where  venda_idvenda='+quotedstr(  DM_Tabelas.ZQProposta.FieldByName('idproposta').Text));
        DM_Tabelas.ZQVendedor_proposta.ExecSQL;
        DM_Tabelas.ZQVendedor_proposta.SQL.Clear;
        DM_Tabelas.ZQVendedor_proposta.SQL.Add('select * from Vendedor_proposta ');
        DM_Tabelas.ZQVendedor_proposta.Open;

      end;
      if DM_Tabelas.ZQproposta.RecordCount>0 then
         DM_Tabelas.ZQproposta.Delete;

      cd :=   DM_Tabelas.ZQProposta.FieldByName('idproposta').AsLargeInt;
      DM_Tabelas.ZQproposta.Close;
      DM_Tabelas.ZQproposta.SQL.Clear;
      DM_Tabelas.ZQproposta.SQL.Add('SELECT * FROM proposta ');
      DM_Tabelas.ZQproposta.open;

      DM_Tabelas.ZQproposta.locate('idproposta',cd,[]);


      DM_Tabelas.ZQRecebimento_proposta.close;
      DM_Tabelas.ZQRecebimento_proposta.SQL.Clear;
      DM_Tabelas.ZQRecebimento_proposta.SQL.Add('select * from recebimento_proposta');
      DM_Tabelas.ZQRecebimento_proposta.Open;

      DM_Tabelas.ZQCheque_proposta.Close;
      DM_Tabelas.ZQCheque_proposta.SQL.Clear;
      DM_Tabelas.ZQCheque_proposta.SQL.Add('select * from cheque_rec_proposta ');
      DM_Tabelas.ZQCheque_proposta.open;


      DM_Tabelas.ZQComprador_proposta.Close;      
      DM_Tabelas.ZQComprador_proposta.SQL.Clear;
      DM_Tabelas.ZQComprador_proposta.SQL.Add('select * from comprador_proposta ');
      DM_Tabelas.ZQComprador_proposta.Open;


      DM_Tabelas.ZQVendedor_proposta.Close;
      DM_Tabelas.ZQVendedor_proposta.SQL.Clear;      
      DM_Tabelas.ZQVendedor_proposta.SQL.Add('select * from Vendedor_proposta ');
      DM_Tabelas.ZQVendedor_proposta.Open;
      
      mensagem('Exclusão Efetuada Com Sucesso!');

    end;
  end;
end;

procedure TFrm_Venda.BtRelatorioClick(Sender: TObject);
begin
  if not Verif_senha('Vendas','Relatórios','') Then Exit;
  JvXPBar1.Visible := True;
  JvXPBar1.SetFocus;
end;

procedure TFrm_Venda.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Venda.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Venda.DXBAchaLoteClick(Sender: TObject);
begin
  AbrirModal(Self, Frm_AchaLoteVenda);
  Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrm_Venda.ECompradorExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then
    if (DM_Tabelas.CDSCompradorTemp.RecordCount=0) or (not empty(Ecomprador.Text)) Then
    Begin
      if AchaCLiente(200, 10,Ecomprador.Text) Then
      Begin
        DM_Tabelas.CDSCompradorTemp.Insert;
        DM_Tabelas.CDSCompradorTempcodparticipante.Value := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
        DM_Tabelas.CDSCompradorTempnomeparticipante.Value := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
        DM_Tabelas.CDSCompradorTempordem.Value := DM_Tabelas.CDSCompradorTemp.RecordCount;
        DM_Tabelas.CDSCompradorTemppercentual.Value:=epercentual.Value;
        DM_Tabelas.CDSCompradorTemp.Post;
        Atualiza_comprador;
      end;
      Ecomprador.Text := '';
      Ecomprador.SetFocus;
    end;
    epercentual.Value:=100;
end;

procedure TFrm_Venda.CompradorEnter(Sender: TObject);
begin
  Epercentual.SetFocus;
end;

procedure TFrm_Venda.DBGrid2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (key=46) and (DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit]) Then
    DM_Tabelas.CDSCompradorTemp.Delete;
  Atualiza_comprador;
end;

procedure TFrm_Venda.CorretorEnter(Sender: TObject);
begin
  Ecorretor.SetFocus;
end;

procedure TFrm_Venda.EcorretorExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then
    if (DM_Tabelas.CDSVendedorTemp.RecordCount=0) or (not empty(Ecorretor.Text)) Then Begin
      if AchaCorretor(200, 10,Ecorretor.Text) Then Begin
        DM_Tabelas.CDSVendedorTemp.Insert;
        DM_Tabelas.CDSVendedorTempcodcorretor.Value := DM_TAbelas.ZQAchaCorretor.FieldByName('idcorretor').AsLargeInt;
        DM_Tabelas.CDSVendedorTempPComissao_corretor.Value := DM_TAbelas.ZQAchaCorretor.FieldByName('PComissao_Corretor').AsFloat;
        DM_TAbelas.CDSVendedorTempordem.Value := DM_Tabelas.CDSVendedorTemp.RecordCount;
        DM_Tabelas.CDSVendedorTemp.Post;
      end;
      Ecorretor.Text := '';
      Ecorretor.SetFocus;
    end;
end;

procedure TFrm_Venda.CorretorExit(Sender: TObject);
begin
  DBComboBox1.SetFocus;
end;

procedure TFrm_Venda.DBGVendaDblClick(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.FieldByName('path_pdf').AsString=emptystr then
  begin
    if not fileexists(ExtractFilePath(Application.ExeName)) then
    begin
       ForceDirectories(ExtractFilePath(Application.ExeName)+'Contratos');
       DM_Tabelas.ZQVenda.Edit;
       DM_Tabelas.ZQVenda.FieldByName('path_pdf').AsString := ExtractFilePath(Application.ExeName)+'Contratos';
       DM_Tabelas.ZQVenda.Post;
    end;
    if fileexists(ExtractFilePath(Application.ExeName)+'Contratos\'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'.pdf') then
    begin
      if FrmVisualizarPdf=nil then
         FrmVisualizarPdf:=TFrmVisualizarPdf.create(self);
      FrmVisualizarPdf.PathPdf.Caption  := ExtractFilePath(Application.ExeName)+'Contratos\'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'.pdf';
      AbrirModal(Self, FrmVisualizarPdf);
      FreeAndNil(FrmVisualizarPdf);
    end
    else
    begin
      mensagem('PDF do Contrato não encontrado');
    end;
  end
  else
  begin
    if fileexists(ExtractFilePath(Application.ExeName)+'Contratos\'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'.pdf') then
    begin
      if FrmVisualizarPdf=nil then
         FrmVisualizarPdf:=TFrmVisualizarPdf.create(self);
      FrmVisualizarPdf.PathPdf.Caption  := ExtractFilePath(Application.ExeName)+'Contratos\'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'.pdf';
      AbrirModal(Self, FrmVisualizarPdf);
      FreeAndNil(FrmVisualizarPdf);
    end
    else
    begin
      mensagem('PDF do Contrato não encontrado');
    end;
  end;
end;

procedure TFrm_Venda.DBGVendaDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQVenda.FieldByName('idvenda').Text Then Begin
    DBGVenda.Canvas.Brush.Color :=$006CFFFF;
    DBGVenda.Canvas.Font.Color := $00A80000;
    DBGVenda.Canvas.Font.Style := [FsBold];
  end;
  DBGVenda.DefaultDrawDataCell(Rect, DBGVenda.columns[datacol].field, State);

end;

procedure TFrm_Venda.DBGVendaKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  Panel1.Visible:=true;
  Application.ProcessMessages;
  Botoes;
  Panel1.Visible:=false;
  Application.ProcessMessages;
end;

procedure TFrm_Venda.DBGVendaMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Panel1.Visible:=true;
  Application.ProcessMessages;
  Botoes;
  Panel1.Visible:=false;
  Application.ProcessMessages;
end;

procedure TFrm_Venda.DBGVendaEnter(Sender: TObject);
begin
  Panel1.Visible:=true;
  Application.ProcessMessages;
  Botoes;
  Panel1.Visible:=false;
  Application.ProcessMessages;

end;

procedure TFrm_Venda.DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Atualiza_comprador;
end;

procedure TFrm_Venda.DBGrid2Enter(Sender: TObject);
begin
  Atualiza_comprador;
end;

procedure TFrm_Venda.DBEEntradaDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBEOrdem.Text = DM_Tabelas.CDSPagEntradaordem.Text Then Begin
    DBEEntrada.Canvas.Brush.Color :=$006CFFFF;
    DBEEntrada.Canvas.Font.Color := $00A80000;
    DBEEntrada.Canvas.Font.Style := [FsBold];
  end;
  DBEEntrada.DefaultDrawDataCell(Rect, DBEEntrada.columns[datacol].field, State);

  if Column.Field = DM_Tabelas.CDSPagEntradaMarcar then
  begin
    DBEEntrada.Canvas.FillRect(Rect);
    if DM_Tabelas.CDSPagEntradaMarcar.Value = 0 then
      ImageList1.Draw(DBEEntrada.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBEEntrada.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
  end;

end;

procedure TFrm_Venda.DBEOrdemChange(Sender: TObject);
begin
  Label17.Caption := copy(DM_Tabelas.CDSPagEntradaordem.AsString,1,3);
end;

procedure TFrm_Venda.GBEntradaEnter(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then
    XNEEntrada.SetFocus;
end;

procedure TFrm_Venda.CBPeriodoEntExit(Sender: TObject);
var
  varvezes, varfinal, vardias: integer;
  varsoma : double;
  varvenci : Tdatetime;
  vartipo : string;
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then Begin
    if empty(CBPeriodoEnt.Text) Then Begin
      CBPeriodoEnt.SetFocus;
      Exit;
    end;
    varvenci  :=  XVencimentoEnt.DateValue;
    DM_Tabelas.CDSPagEntrada.Close;
    DM_Tabelas.CDSPagEntrada.CreateDataSet;
    varfinal  := strtoint(floattostr(XNQtdeEnt.Value));
    if StrIsNumero(CBPeriodoEnt.Text) Then Begin
      vardias := strtoint(CBPeriodoEnt.Text);
      vartipo := 'N';
    end
    else Begin
      if CBPeriodoEnt.ItemIndex<0 Then Begin
        CBPeriodoEnt.SetFocus;
        Exit;
      end;
      vartipo := 'S';
      case CBPeriodoEnt.ItemIndex of
        0 : vardias := 1;
        1 : vardias := 7;
        2 : vardias := 15;
        3 : vardias := 1;
        4 : vardias := 2;
        5 : vardias := 3;
        6 : vardias := 4;
        7 : vardias := 6;
        8 : vardias := 12;
        9 : vardias := 24;
       10 : vardias := 0;
      end;
      if (CBPeriodoEnt.ItemIndex<=2) and (CBPeriodoEnt.ItemIndex=10) Then
        vartipo := 'N';
    end;
    DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger,[]);
    varsoma                                        := 0;
    for varvezes                                   := 1 to varfinal do Begin
      DM_Tabelas.CDSPagEntrada.Insert;
      DM_Tabelas.CDSPagEntradaordem.Value          := inttostrZero(varvezes,3)+'/'+inttostrZero(varfinal,3);
      DM_Tabelas.CDSPagEntradavalor.Value          := truncar(XNEEntrada.Value/varfinal,6);
      DM_Tabelas.CDSPagEntradacontabil.Value       := DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger;
      DM_Tabelas.CDSPagEntradatipodoc.Value        := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString;
      if varvezes = 1 Then
         DM_Tabelas.CDSPagEntradavencimento.Value  := XVencimentoEnt.DateValue
      else Begin
        if vartipo = 'S' then
          DM_Tabelas.CDSPagEntradavencimento.Value := IncMonth(XVencimentoEnt.DateValue,(varvezes*vardias)-(vardias))
        else
          DM_Tabelas.CDSPagEntradavencimento.Value := varvenci+vardias;
        varvenci                                   := DM_Tabelas.CDSPagEntradavencimento.Value;
      end;
      DM_Tabelas.CDSPagEntrada.Post;
      varsoma                                      := varsoma + DM_Tabelas.CDSPagEntradavalor.Value;
    end;
    DM_Tabelas.CDSPagEntrada.First;
    DM_Tabelas.CDSPagEntrada.Edit;
    DM_Tabelas.CDSPagEntradavalor.Value            := DM_Tabelas.CDSPagEntradavalor.Value+(XNEEntrada.Value-varsoma);
    DM_Tabelas.CDSPagEntrada.Post;
  end;
  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.atualiza_tela_venda_entrada;
Begin
  EContabilEntrada.Text := DM_Tabelas.CDSPagEntradacontabildescri.AsString;
  DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagEntradatipodoc.Value,[]);
  if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
    Label20.Visible := True;
    Label21.Visible := True;
    Label27.Visible := True;
    Label32.Visible := True;
    Label33.Visible := True;
    Label51.Visible := True;
    DBEBanco.Visible := True;
    DBEAgencia.Visible := True;
    DBEConta.Visible := True;
    DBEDocumento.Visible := True;
    DBENumeroEntrada.Visible := True;
    DBEDonoEntrada.Visible := True;
  end
  else Begin
    Label20.Visible := False;
    Label21.Visible := False;
    Label27.Visible := False;
    Label32.Visible := False;
    Label33.Visible := False;
    Label51.Visible := False;
    DBEBanco.Visible := False;
    DBEAgencia.Visible := False;
    DBEConta.Visible := False;
    DBEDocumento.Visible := False;
    DBENumeroEntrada.Visible := False;
    DBEDonoEntrada.Visible := False;
  end;
End;

procedure TFrm_Venda.atualiza_tela_venda_parcela;
Begin
  EContabilEntradaParc.Text := DM_Tabelas.CDSPagParcelacontabildescri.AsString;
  DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagParcelatipodoc.Value,[]);
  if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
    Label45.Visible := True;
    Label46.Visible := True;
    Label47.Visible := True;
    Label48.Visible := True;
    Label49.Visible := True;
    Label52.Visible := True;
    DBEBancoParc.Visible := True;
    DBEAgendaParc.Visible := True;
    DBEContaParc.Visible := True;
    DBEDocumentoParc.Visible := True;
    DBENumeroParc.Visible := True;
    DBEDonoParcela.Visible := True;
  end
  else Begin
    Label45.Visible := False;
    Label46.Visible := False;
    Label47.Visible := False;
    Label48.Visible := False;
    Label49.Visible := False;
    Label52.Visible := False;
    DBEBancoParc.Visible := False;
    DBEAgendaParc.Visible := False;
    DBEContaParc.Visible := False;
    DBEDocumentoParc.Visible := False;
    DBENumeroParc.Visible := False;
    DBEDonoParcela.Visible := False;
  end;
End;

procedure TFrm_Venda.atualizatebelaspagamento;
var
  VpAtrz,varsomaE, varsomaP, VE, VP : double;
Begin
  DM_Tabelas.ZQRecebimento_temp.SQL.Clear;
  DM_Tabelas.ZQRecebimento_temp.SQL.Add('Select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento_temp.SQL.Add('       custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto ');
  DM_Tabelas.ZQRecebimento_temp.SQL.Add(', sum(valor) as somavalor from recebimento as r ');
  DM_Tabelas.ZQRecebimento_temp.SQL.Add(' Where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text));
  DM_Tabelas.ZQRecebimento_temp.SQL.Add(' group by r.idrecebimento');
  DM_Tabelas.ZQRecebimento_temp.Open;
  DM_Tabelas.ZQRecebimento_temp.RecordCount;
  DM_Tabelas.CDSPagEntrada.Close;
  DM_Tabelas.CDSPagEntrada.CreateDataSet;
  DM_Tabelas.CDSPagParcela.Close;
  DM_Tabelas.CDSPagParcela.CreateDataSet;
  varsomaE := 0;
  varsomaP := 0;
  Ve := 0;
  Vp := 0;
  VpAtrz:=0;
  XVencimentoEnt.DateText := '';
  XVencimentoParc.DateText := '';
  DM_Tabelas.ZQRecebimento_temp.DisableControls;
  ProgressBar1.Visible:=true;
  ProgressBar1.Position:=0;
  ProgressBar1.Max:=DM_Tabelas.ZQRecebimento_temp.RecordCount;

  while not DM_Tabelas.ZQRecebimento_temp.Eof do begin
    ProgressBar1.Position:=DM_Tabelas.ZQRecebimento_temp.RecNo;
    IF pos('-E-',DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').AsString)>0 Then
    Begin
      if Copy(XVencimentoEnt.DateText,1,2)='00'  Then Begin
        XVencimentoEnt.DateValue := DM_Tabelas.ZQRecebimento_temp.FieldByName('Dt_Vencimento').AsDateTime;
        CBPeriodoEnt.Text := copy(DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').AsString,pos(' ',DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').AsString)+1,20);
      end;
      DM_Tabelas.CDSPagEntrada.Insert;
      DM_Tabelas.CDSPagEntradacontabil.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('contabil').AsLargeInt;
      DM_Tabelas.CDSPagEntradavencimento.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('Dt_Vencimento').AsDateTime;
      DM_Tabelas.CDSPagEntradaordem.Value := copy(DM_TAbelas.ZQRecebimento_temp.FieldByName('ordem').AsString,pos('-',DM_TAbelas.ZQRecebimento_temp.FieldByName('ordem').AsString)+1,7);
      DM_Tabelas.CDSPagEntradavalor.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('somavalor').AsFloat;
      DM_Tabelas.CDSPagEntradatipodoc.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('TipDoc').AsString;
      DM_tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.CDSPagEntradatipodoc.Value,[]);
      IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin

         DM_Tabelas.ZQRecebimento.Close;
         DM_Tabelas.ZQRecebimento.SQL.Clear;
         DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
         DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
         DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
         DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
         DM_Tabelas.ZQRecebimento.SQL.Add(' from Recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.idrecebimento='+quotedstr(DM_Tabelas.ZQRecebimento_temp.FieldByName('idrecebimento').Text)+' order by r.DT_Vencimento');
         DM_Tabelas.ZQRecebimento.open;


        DM_Tabelas.ZQRecebimento.Locate('idrecebimento',DM_Tabelas.ZQRecebimento_temp.FieldByName('idrecebimento').AsLargeInt,[]);


        DM_Tabelas.CDSPagEntradabanco.Value := DM_Tabelas.ZQCheque.FieldByName('Banco').AsString;
        DM_Tabelas.CDSPagEntradaagencia.Value := DM_Tabelas.ZQCheque.FieldByName('agencia').AsString;
        DM_Tabelas.CDSPagEntradaconta.Value := DM_Tabelas.ZQCheque.FieldByName('conta').AsString;
        DM_Tabelas.CDSPagEntradanumero.Value := DM_Tabelas.ZQCheque.FieldByName('numero').AsString;
        DM_Tabelas.CDSPagEntradadocumento.Value := DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').AsString;
        DM_Tabelas.CDSPagEntradanomedono.Value := DM_Tabelas.ZQCheque.FieldByName('Dono').AsString;
      end;
      DM_Tabelas.CDSPagEntrada.Post;
      varsomaE := varsomaE + DM_Tabelas.ZQRecebimento_temp.FieldByName('somavalor').AsFloat;
      ve := ve+1;
    end
    else Begin
      if Copy(XVencimentoParc.DateText,1,2)='00'  Then Begin
        XVencimentoParc.DateValue := DM_Tabelas.ZQRecebimento_temp.FieldByName('Dt_Vencimento').AsDateTime;
      CBPeriodoParc.Text := copy(DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').AsString,pos(' ',DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').AsString)+1,20);
      end;
      DM_Tabelas.CDSPagParcela.Insert;
      DM_Tabelas.CDSPagParcelacontabil.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('contabil').AsLargeInt;
      DM_Tabelas.CDSPagParcelavencimento.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('Dt_Vencimento').AsDateTime;
      DM_Tabelas.CDSPagParcelaordem.Value := copy(DM_TAbelas.ZQRecebimento_temp.FieldByName('ordem').AsString,pos('-',DM_TAbelas.ZQRecebimento_temp.FieldByName('ordem').AsString)+1,7);
      DM_Tabelas.CDSPagParcelavalor.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('somavalor').AsFloat;
      DM_Tabelas.CDSPagParcelatipodoc.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('TipDoc').AsString;
      DM_Tabelas.CDSPagParcelavencimento.Value:=DM_Tabelas.ZQRecebimento_temp.FieldByName('Dt_Vencimento').AsDateTime;
      DM_Tabelas.CDSPagParcelaSaldo.Value:=DM_Tabelas.ZQRecebimento_temp.FieldByName('saldo').AsFloat;
      DM_tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.CDSPagEntradatipodoc.Value,[]);
      IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin

        DM_Tabelas.ZQRecebimento.Close;
        DM_Tabelas.ZQRecebimento.SQL.Clear;
        DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
        DM_Tabelas.ZQRecebimento.SQL.Add(' from Recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.idrecebimento='+quotedstr(DM_Tabelas.ZQRecebimento_temp.FieldByName('idrecebimento').Text)+' order by r.DT_Vencimento');
        DM_Tabelas.ZQRecebimento.open;

        DM_Tabelas.ZQRecebimento.Locate('idrecebimento',DM_Tabelas.ZQRecebimento_temp.FieldByName('idrecebimento').AsLargeInt,[]);

        DM_Tabelas.cdSPagParcelabanco.Value := DM_Tabelas.ZQCheque.FieldByName('Banco').AsString;
        DM_Tabelas.cdSPagParcelaagencia.Value := DM_Tabelas.ZQCheque.FieldByName('agencia').AsString;
        DM_Tabelas.cdSPagParcelaconta.Value := DM_Tabelas.ZQCheque.FieldByName('conta').AsString;
        DM_Tabelas.cdSPagParcelanumero.Value := DM_Tabelas.ZQCheque.FieldByName('numero').AsString;
        DM_Tabelas.cdSPagParceladocumento.Value := DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').AsString;
        DM_Tabelas.cdSPagParcelanomedono.Value := DM_Tabelas.ZQCheque.FieldByName('Dono').AsString;
      end;
      DM_Tabelas.CDSPagParcela.Post;
      varsomaP := varsomaP + DM_Tabelas.ZQRecebimento_temp.FieldByName('somavalor').AsFloat;
      if (DM_Tabelas.CDSPagParcelavencimento.Value<date) and (DM_Tabelas.CDSPagParcelaSaldo.Value>0) then
         vpAtrz := vpAtrz +1
      else
         vp := vp +1;
    end;
    DM_Tabelas.ZQRecebimento_temp.Next;
  end;
  DM_Tabelas.ZQRecebimento_temp.EnableControls;
  ProgressBar1.Position:=0;
  ProgressBar1.Visible:=false;
  DM_Tabelas.CDSPagEntrada.First;
  DM_Tabelas.CDSPagParcela.First;
  XNEEntrada.Value := varsomaE;
  XNParcela.Value := varsomaP;
  XNQtdeEnt.Value :=  ve;
  XNQtdeParc.Value := vp+vpAtrz;
  XNQtdeParcAtrz.Value:= vpAtrz;
  XDentrada.Value:=XNEEntrada.Value;
  xdparcelas.Value:=XNParcela.Value;
  XDqtdeentrada.Value:=XNQtdeEnt.Value;
  xdqtdeparcelas.Value:=XNQtdeParc.Value;
  xdqtdeparcelasAtrz.Value:=XNQtdeParcAtrz.Value;

  zqparcelatotal.close;
  zqparcelatotal.SQL.Clear;
  zqparcelatotal.SQL.Add('Select valor,quadralote,idrecebimento,saldo,origem,adversa,venda_idvenda, sum(valor) as parcela from recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.quadralote='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+' and r.saldo >0 and r.documento like  ''%-P-%'' order by r.ordem');
  zqparcelatotal.open;

  zqparcelas.close;
  zqparcelas.SQL.Clear;
  zqparcelas.SQL.Add('Select valor,quadralote,idrecebimento,saldo,origem,adversa,venda_idvenda from recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.quadralote='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString)+' and r.saldo >0 and r.documento like  ''%-P-%'' order by r.ordem');
  zqparcelas.open;


  if ZQParcelatotal.FieldByName('valor').AsFloat>0 then
  begin
    xdsldremanescente.Value:=ZQParcelatotal.FieldByName('parcela').AsFloat;
  end
  else
  begin
    xdsldremanescente.Value:=ZQParcelatotal.FieldByName('parcela').AsFloat;
  end;

  if ZQParcelas.FieldByName('valor').AsFloat>0 then
  begin
    xdparpagas.Value:=(xdqtdeparcelas.Value-zqparcelas.RecordCount);
    xdsldvicenda.Value:=zqparcelas.RecordCount;
    xdsldatrasadas.Value:=xdqtdeparcelasAtrz.Value;
  end
  else
  begin
    xdparpagas.Value:=xdqtdeparcelas.Value;
    xdsldvicenda.Value:=0;
    xdsldatrasadas.Value:=0;
  end;
  zqparcelas.close;
  zqparcelatotal.close;
  xdresta.Value:=xdsldvicenda.Value-xdsldatrasadas.Value;
  DM_Tabelas.CDSPagParcela.IndexFieldNames:='Vencimento';

//  DM_Tabelas.CDSPagParcela.Last;
  DM_Tabelas.CDSPagParcela.First;
  while not DM_Tabelas.CDSPagParcela.Eof do
  begin
    if DM_Tabelas.CDSPagParcelaSaldo.Value>0 then
       ultparc.Value:=DM_Tabelas.CDSPagParcelavalor.Value;
    DM_Tabelas.CDSPagParcela.Next;
  end;

//  ultparc.Value:=DM_Tabelas.CDSPagParcelavalor.Value;
  DM_Tabelas.CDSPagParcela.first;
  

  atualiza_tela_venda_entrada;
  atualiza_tela_venda_parcela;

  DM_Tabelas.ZQRecebimento.Close;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
  DM_Tabelas.ZQRecebimento.SQL.Add(' from Recebimento order by DT_Vencimento limit 200, 30');
  DM_Tabelas.ZQRecebimento.open;


  if Pag_Venda.PageIndex=2 then
  begin
    if xdqtdeparcelas.Value=xdparpagas.Value then
       Showmessage('Não há Valores para Quitação.');
  end;
end;



procedure TFrm_Venda.DBEEntradaEnter(Sender: TObject);
begin
  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.DBEEntradaKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if DBEEntrada.SelectedField.fieldname = 'Marcar' Then
  Begin
    DM_Tabelas.CDSPagEntrada.Edit;
    if  DM_Tabelas.CDSPagEntradaMarcar.Value = 0 Then
    begin
      DM_Tabelas.CDSPagEntradaMarcar.Value := 1;
    end
    else
    begin
      DM_Tabelas.CDSPagEntradaMarcar.Value := 0;
    end;
    DM_Tabelas.CDSPagEntrada.post;
  end;
  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.EContabilEntradaExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsEdit,DsInsert] Then Begin
    if (not empty(EContabilEntrada.Text)) and (AchaPlanoDeConta(450, 30, 'S', EContabilEntrada.Text, '')) Then Begin
      EContabilEntrada.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' | '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
      DM_Tabelas.CDSPagEntrada.Edit;
      DM_Tabelas.CDSPagEntradacontabil.Value :=  DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
      DM_tabelas.CDSPagEntradatipodoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').AsString;
      DM_Tabelas.CDSPagEntrada.Post;
    end
    else
      EContabilEntrada.SetFocus;
  end;
end;

procedure TFrm_Venda.EContabilEntradaEnter(Sender: TObject);
begin
  EContabilEntrada.Text:= DM_Tabelas.CDSPagEntradacontabildescri.AsString;
  atualiza_tela_venda_entrada;

end;

procedure TFrm_Venda.DBEValorExit(Sender: TObject);
var
  varreg, varconta : integer;
  varsoma, resto : double;
begin
  varreg := DM_Tabelas.CDSPagEntrada.RecNo;
  if varreg < XNQtdeEnt.Value Then Begin
    DM_Tabelas.CDSPagEntrada.First;
    varsoma := 0;
    for varconta := 1 to varreg do Begin
      DM_Tabelas.CDSPagEntrada.RecNo := varconta;
      varsoma := varsoma + DM_Tabelas.CDSPagEntradavalor.Value;
    end;
    resto := XNEEntrada.Value - varsoma;
    varsoma := 0;
    for varconta := varreg+1 to DM_Tabelas.CDSPagEntrada.RecordCount do begin
      DM_Tabelas.CDSPagEntrada.RecNo := varconta;
      DM_Tabelas.CDSPagEntrada.Edit;
      DM_Tabelas.CDSPagEntradavalor.Value := truncar(resto/(XNQtdeEnt.Value-varreg),6);
      varsoma:=varsoma+DM_Tabelas.CDSPagEntradavalor.Value;
    end;

    DM_Tabelas.CDSPagEntrada.RecNo := varreg+1;
    DM_Tabelas.CDSPagEntrada.Edit;
    DM_Tabelas.CDSPagEntradavalor.Value := DM_Tabelas.CDSPagEntradavalor.Value+(resto-varsoma);
    DM_Tabelas.CDSPagEntrada.Post;
    DM_Tabelas.CDSPagEntrada.RecNo := varreg;
    if (DM_Tabelas.CDSPagEntrada.RecNo<=XNQtdeEnt.Value) and (DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'N') Then
    Begin
      DM_Tabelas.CDSPagEntrada.Next;
      EContabilEntrada.SetFocus;
    end
    else if (DM_Tabelas.CDSPagEntrada.RecNo<=XNQtdeEnt.Value) and (DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S') Then
    Begin
      DBEBanco.SetFocus;
    end;
  end
  else
  begin
    if (DM_Tabelas.CDSPagEntrada.RecNo<=XNQtdeEnt.Value) and (DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S') Then
    Begin
      DBEBanco.SetFocus;
      exit;
    end;
    XNParcela.SetFocus;
  end;
end;

procedure TFrm_Venda.DBEValorEnter(Sender: TObject);
begin
 // if DM_Tabelas.CDSPagEntrada.RecNo=DM_Tabelas.CDSPagEntrada.RecordCount Then
//    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrm_Venda.DBEDocumentoExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then 
    if not Verif_doc(DBEDocumento.text,True) Then Begin
      DBEDocumento.SetFocus;
      exit;
    end;
end;

procedure TFrm_Venda.XVencimentoEntExit(Sender: TObject);
begin
  if XVencimentoEnt.DateValue<=0 Then
    XVencimentoEnt.SetFocus;
end;

procedure TFrm_Venda.DBCBDocVendaExit(Sender: TObject);
begin
  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.XNEEntradaExit(Sender: TObject);
begin
  XNParcela.Value := DBEVrvenda.Value - XNEEntrada.Value;


  if XNEEntrada.Value = 0 Then
    GBParcelas.SetFocus;
end;

procedure TFrm_Venda.GBParcelasEnter(Sender: TObject);
begin
  if XNParcela.Value>0 Then
    XNQtdeParc.SetFocus
  else
    BtGravar.SetFocus;
end;

procedure TFrm_Venda.CBPeriodoParcExit(Sender: TObject);
var
  varvezes, varfinal, vardias: integer;
  varsoma : double;
  varvenci : Tdatetime;
  vartipo : string;
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then Begin
    if empty(CBPeriodoParc.Text) Then Begin
      CBPeriodoParc.SetFocus;
      Exit;
    end;
    varvenci :=  XVencimentoParc.DateValue;
    DM_Tabelas.CDSPagParcela.Close;
    DM_Tabelas.CDSPagParcela.CreateDataSet;
    varfinal := strtoint(floattostr(XNQtdeParc.Value));
    if StrIsNumero(CBPeriodoParc.Text) Then Begin
      vardias := strtoint(CBPeriodoParc.Text);
      vartipo := 'N';
    end
    else Begin
      if CBPeriodoParc.ItemIndex<0 Then Begin
        CBPeriodoParc.SetFocus;
        Exit;
      end;
      vartipo := 'S';
      case CBPeriodoParc.ItemIndex of
        0 : vardias := 1;
        1 : vardias := 7;
        2 : vardias := 15;
        3 : vardias := 1;
        4 : vardias := 2;
        5 : vardias := 3;
        6 : vardias := 4;
        7 : vardias := 6;
        8 : vardias := 12;
        9 : vardias := 24;
      end;
      if CBPeriodoParc.ItemIndex<=2 Then
        vartipo := 'N';
    end;
    DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger,[]);
    varsoma := 0;
    for varvezes := 1 to varfinal do Begin
      DM_Tabelas.CDSPagParcela.Insert;
      DM_Tabelas.CDSPagParcelaordem.Value := inttostrZero(varvezes,3)+'/'+inttostrZero(varfinal,3);
      DM_Tabelas.CDSPagParcelavalor.Value := truncar(XNParcela.Value/varfinal,6);
      DM_Tabelas.CDSPagParcelacontabil.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger;
      DM_Tabelas.CDSPagParcelatipodoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString;
      if varvezes = 1 Then
         DM_Tabelas.CDSPagParcelavencimento.Value := XVencimentoParc.DateValue
      else Begin
        if vartipo = 'S' then
          DM_Tabelas.CDSPagParcelavencimento.Value := IncMonth(XVencimentoParc.DateValue,(varvezes*vardias)-(vardias))
        else
          DM_Tabelas.CDSPagParcelavencimento.Value := varvenci+vardias;
        varvenci := DM_Tabelas.CDSPagParcelavencimento.Value;
      end;
      DM_Tabelas.CDSPagParcela.Post;
      varsoma := varsoma + DM_Tabelas.CDSPagParcelavalor.Value;
    end;
    DM_Tabelas.CDSPagParcela.First;
    DM_Tabelas.CDSPagParcela.Edit;
    DM_Tabelas.CDSPagParcelavalor.Value := DM_Tabelas.CDSPagParcelavalor.Value+(XNParcela.Value-varsoma);
    DM_Tabelas.CDSPagParcela.Post;
  end;
  if Pag_Venda.PageIndex = 1 Then
    atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.XVencimentoParcExit(Sender: TObject);
var
fina,parcelas:double;
begin
  parcelas:=0;
  fina:=0;
  if XVencimentoParc.DateValue<=0 Then
    XVencimentoParc.SetFocus;
  if (alltrim(DBEPRice.text)<>'0,000') and (alltrim(DBEPRice.text)<>'0,0000') and (alltrim(DBEPRice.text)<>'0') and (alltrim(DBEPRice.text)<>'') then
  begin
    DM_Tabelas.ZQPrice.close;
    DM_Tabelas.ZQPrice.SQL.Clear;
    DM_Tabelas.ZQPrice.SQL.Add('select * from Price where taxa='+quotedstr(StringReplace(alltrim(DBEPRice.text),',','.',[rfReplaceAll]))+' and mes='+quotedstr(XNQtdeParc.Text));
    DM_Tabelas.ZQPrice.open;
    if DM_Tabelas.ZQPrice.RecordCount>0 then
    begin
      fina:=XNParcela.Value;
//      parcelas:=fina*DM_Tabelas.ZQPricecoeficiente_multiplicador.Value;
      if DM_Tabelas.ZQprice.FieldByName('Coeficiente_amortizador').AsFloat>0 then
         XNParcela.Value:=fina*(DM_Tabelas.ZQprice.FieldByName('Coeficiente_amortizador').AsFloat/1000);
      XNParcela.Value:=XNParcela.Value*XNQtdeParc.Value;
      parcelas:=0;
      fina:=0;
    end;
    DM_Tabelas.ZQPrice.close;
    DM_Tabelas.ZQPrice.SQL.Clear;
    DM_Tabelas.ZQPrice.SQL.Add('select * from Price  order by taxa,mes');
    DM_Tabelas.ZQPrice.open;
    DM_Tabelas.ZQPrice.close;
  end;
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then
  begin
    DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat:=XNParcela.Value+XNEEntrada.Value;
  end;
{  DM_Tabelas.ZQPrice.SQL.Clear;
  DM_Tabelas.ZQPrice.SQL.Add('select * from Price ');
  DM_Tabelas.ZQPrice.open;}

end;

procedure TFrm_Venda.DBEDocumentoParcExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then
    if not Verif_doc(DBEDocumentoParc.text,True) Then Begin
      DBEDocumentoParc.SetFocus;
      exit;
    end;
end;

procedure TFrm_Venda.DBGPartelasDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBEOrdemParc.Text = DM_Tabelas.CDSPagParcelaordem.Text Then Begin
    DBGPartelas.Canvas.Brush.Color :=$006CFFFF;
    DBGPartelas.Canvas.Font.Color := $00A80000;
    DBGPartelas.Canvas.Font.Style := [FsBold];
  end;
  DBGPartelas.DefaultDrawDataCell(Rect, DBGPartelas.columns[datacol].field, State);
  if Column.Field = DM_Tabelas.CDSPagParcelaMarcar then
  begin
    DBGPartelas.Canvas.FillRect(Rect);
    if DM_Tabelas.CDSPagParcelaMarcar.Value = 0 then
      ImageList1.Draw(DBGPartelas.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBGPartelas.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
  end;


end;

procedure TFrm_Venda.DBEVrvendaEnter(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsEdit,DsInsert] Then
     DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat:=DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat-DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat;
end;

procedure TFrm_Venda.DBEVrvendaExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat<(DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat-DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat) Then
    Showmessage('O valor de venda é menor que o valor de tabela + Comissão...');
end;

procedure TFrm_Venda.DBCBDocVendaParcExit(Sender: TObject);
begin
  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.EContabilEntradaParcEnter(Sender: TObject);
begin
  EContabilEntradaParc.Text:= DM_Tabelas.CDSPagparcelacontabildescri.AsString;
  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.EContabilEntradaParcExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsEdit,DsInsert] Then Begin
    if (not empty(EContabilEntradaParc.Text)) and (AchaPlanoDeConta(450, 30, 'S', EContabilEntradaParc.Text, '')) Then Begin
      EContabilEntradaParc.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' | '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
      DM_Tabelas.cdSPagParcela.Edit;
      DM_Tabelas.cdSPagParcelacontabil.Value :=  DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
      DM_tabelas.cdSPagParcelatipodoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').AsString;
      DM_Tabelas.cdSPagParcela.Post;
    end
    else
      EContabilEntradaParc.SetFocus;
  end;

end;

procedure TFrm_Venda.DBEOrdemParcChange(Sender: TObject);
begin
  Label44.Caption := copy(DM_Tabelas.CDSPagparcelaordem.AsString,1,3);
end;

procedure TFrm_Venda.DBEValorParcEnter(Sender: TObject);
begin
//  if DM_Tabelas.CDSPagparcela.RecNo=DM_Tabelas.CDSPagparcela.RecordCount Then
//    Perform(WM_NEXTDLGCTL, 0, 0);

end;

procedure TFrm_Venda.DBEValorParcExit(Sender: TObject);
var
  varreg, varconta : integer;
  varsoma, resto : double;
begin
  varreg := DM_Tabelas.cdSPagParcela.RecNo;
  if varreg < XNQtdeParc.Value Then Begin
    DM_Tabelas.cdSPagParcela.First;
    varsoma := 0;
    for varconta := 1 to varreg do Begin
      DM_Tabelas.cdSPagParcela.RecNo := varconta;
      varsoma := varsoma + DM_Tabelas.cdSPagParcelavalor.Value;
    end;
    resto := XNParcela.Value - varsoma;
    varsoma := 0;
    for varconta := varreg+1 to DM_Tabelas.cdSPagParcela.RecordCount do begin
      DM_Tabelas.cdSPagParcela.RecNo := varconta;
      DM_Tabelas.cdSPagParcela.Edit;
      DM_Tabelas.cdSPagParcelavalor.Value := truncar(resto/(XNQtdeParc.Value-varreg),6);
      varsoma:=varsoma+DM_Tabelas.cdSPagParcelavalor.Value;
    end;
    DM_Tabelas.cdSPagParcela.RecNo := varreg+1;
    DM_Tabelas.cdSPagParcela.Edit;
    DM_Tabelas.cdSPagParcelavalor.Value := DM_Tabelas.cdSPagParcelavalor.Value+(resto-varsoma);
    DM_Tabelas.cdSPagParcela.Post;
    DM_Tabelas.cdSPagParcela.RecNo := varreg;
    if (DM_Tabelas.cdSPagParcela.RecNo<XNQtdeParc.Value) and (DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'N') Then Begin
      DM_Tabelas.cdSPagParcela.Next;
      EContabilEntradaParc.SetFocus;
    end;
  end;
end;

procedure TFrm_Venda.DBEEntradaMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DBEEntrada.SelectedField.fieldname = 'Marcar' Then
  Begin
    DM_Tabelas.CDSPagEntrada.Edit;
    if  DM_Tabelas.CDSPagEntradaMarcar.Value = 0 Then
    begin
      DM_Tabelas.CDSPagEntradaMarcar.Value := 1;
    end
    else
    begin
      DM_Tabelas.CDSPagEntradaMarcar.Value := 0;
    end;
    DM_Tabelas.CDSPagEntrada.post;
  end;

  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.DBGPartelasKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if DBGPartelas.SelectedField.fieldname = 'Marcar' Then
  Begin
    DM_Tabelas.CDSPagParcela.Edit;
    if  DM_Tabelas.CDSPagParcelaMarcar.Value = 0 Then
    begin
      DM_Tabelas.CDSPagParcelaMarcar.Value := 1;
    end
    else
    begin
      DM_Tabelas.CDSPagParcelaMarcar.Value := 0;
    end;
    DM_Tabelas.CDSPagParcela.post;
  end;

  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.DBGPartelasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DBGPartelas.SelectedField.fieldname = 'Marcar' Then
  Begin
    DM_Tabelas.CDSPagParcela.Edit;
    if  DM_Tabelas.CDSPagParcelaMarcar.Value = 0 Then
    begin
      DM_Tabelas.CDSPagParcelaMarcar.Value := 1;
    end
    else
    begin
      DM_Tabelas.CDSPagParcelaMarcar.Value := 0;
    end;
    DM_Tabelas.CDSPagParcela.post;
  end;
  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.DBGPartelasEnter(Sender: TObject);
begin
  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.DBEDonoEntradaExit(Sender: TObject);
begin
 if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then
    if (DM_Tabelas.CDSPagEntrada.RecNo<XNQtdeEnt.Value) Then Begin
      DM_Tabelas.CDSPagEntrada.Next;
      EContabilEntrada.SetFocus;
    end
    else
      XNParcela.SetFocus;
end;

procedure TFrm_Venda.DBEDonoParcelaExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then
    if (DM_Tabelas.CDSPagParcela.RecNo<XNQtdeParc.Value) Then Begin
      DM_Tabelas.CDSPagParcela.Next;
      EContabilEntradaParc.SetFocus;
    end;
end;

procedure TFrm_Venda.JvXPBar1Items0Click(Sender: TObject);
var
  varvenda : integer;
begin
  Panel1.Visible:=true;
  Application.ProcessMessages;

  Frm_Venda.FormStyle:=fsNormal;
  Application.ProcessMessages;
//  Timer1.Enabled:=true;
  JvXPBar1.Visible := False;
  Atualizatebelaspagamento;

  varvenda := DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
  DM_TAbelas.ZQVenda.Filter := 'idvenda='+inttostr(varvenda);
  DM_Tabelas.ZQVenda.Filtered := True;
  DM_TAbelas.ZQCompr_Dados.Open;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_TAbelas.ZQCompr_Resp_Dados.Open;
  DM_TAbelas.ZQCompr_Resp_Dados.Refresh;
  DM_TAbelas.ZQCompr_Conjuge.Open;
  DM_TAbelas.ZQCompr_Conjuge.Refresh;
  DM_Tabelas.ZQCompr_Dados.First;
  if DM_Tabelas.ZQLoteamento.Active=false then
     DM_Tabelas.ZQLoteamento.Open;
  DM_Tabelas.ZQLoteamento.Locate('idloteamento',DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger,[]);
  DM_Tabelas.ZQCidade.Open;
  try
    if Frm_Loteamento=nil then
       Frm_Loteamento := TFrm_Loteamento.Create(self);
    Frm_Loteamento.atualiza_lotes;

  finally
    Frm_Loteamento:=nil;
    Frm_Loteamento.Free;
  end;
  DM_TAbelas.CDSAss1.Close;
  DM_TAbelas.CDSAss1.CreateDataSet;
{  DM_TAbelas.CDSAss1.Insert;
  DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamentonomeloteamento.Value;
  DM_Tabelas.CDSAss1cargo.Value :=' Vendedor';
//

  if alltrim(DM_TAbelas.ZQLoteamentonomeloteamento.AsString)<>'ALICIO VILAR PONTES' then
     DM_Tabelas.CDSAss1responsavel.Value := 'PP: '+DM_Tabelas.CDSIncorpnomeparti.Value;
 }


  DM_Tabelas.zqprocuradores.Filtered:=false;
  DM_Tabelas.zqprocuradores.Filter:='Assina=''S''';
  DM_Tabelas.zqprocuradores.filtered:=true;



  if (DM_Tabelas.CDSIncorpassina.Value='S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString<>'S') then
  begin
    DM_TAbelas.CDSAss1.Insert;
    DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
    DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
    if alltrim(DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString)<>'ALICIO VILAR PONTES' then
       DM_Tabelas.CDSAss1responsavel.Value := DM_Tabelas.CDSIncorpnomeparti.Value;
    DM_Tabelas.CDSAss1conjuge.Value := DM_Tabelas.CDSIncorpconjuge.Value;
    DM_Tabelas.CDSAss1idconjuge.Value := DM_Tabelas.CDSIncorpidconjuge.Value;
    DM_Tabelas.CDSAss1assina.Value := DM_Tabelas.CDSIncorpassina.Value;
    DM_TAbelas.CDSAss1.Post;
  end;

  if (DM_Tabelas.CDSIncorpassina.Value='S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString='S') then
  begin
    DM_Tabelas.zqprocuradores.First;
    while not DM_Tabelas.zqprocuradores.Eof do
    begin
      DM_TAbelas.CDSAss1.Insert;
      DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
      DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
      DM_Tabelas.CDSAss1responsavel.Value := DM_Tabelas.ZQProcuradores.FieldByName('nome').AsString;

      DM_Tabelas.CDSAss1conjuge.Value := DM_Tabelas.CDSIncorpconjuge.Value;
      DM_Tabelas.CDSAss1idconjuge.Value := DM_Tabelas.CDSIncorpidconjuge.Value;
      DM_Tabelas.CDSAss1assina.Value := DM_Tabelas.CDSIncorpassina.Value;


      DM_TAbelas.CDSAss1.Post;
      DM_Tabelas.zqprocuradores.Next;
    end;
  end;

  if (DM_Tabelas.CDSIncorpassina.Value<>'S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString='S') then
  begin
    DM_Tabelas.zqprocuradores.First;
    while not DM_Tabelas.zqprocuradores.Eof do
    begin
      DM_TAbelas.CDSAss1.Insert;
      DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
      DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
      DM_Tabelas.CDSAss1responsavel.Value := DM_Tabelas.ZQProcuradores.FieldByName('nome').AsString;
  //    DM_Tabelas.CDSAss1assina.Value := DM_Tabelas.zqprocuradoresassina.Value;
      DM_TAbelas.CDSAss1.Post;
      DM_Tabelas.zqprocuradores.Next;
    end;
  end;


  if (DM_Tabelas.CDSIncorpassina.Value<>'S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString<>'S') then
  begin
    DM_TAbelas.CDSAss1.Insert;
    DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
    DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
    if alltrim(DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString)<>'ALICIO VILAR PONTES' then
       DM_Tabelas.CDSAss1responsavel.Value := DM_Tabelas.CDSIncorpnomeparti.Value;
    DM_TAbelas.CDSAss1.Post;
  end;


  DM_Tabelas.zqprocuradores.Filtered:=false;


//  DM_TAbelas.CDSAss1.Post;
  DM_TAbelas.CDSAss2.Close;
  DM_TAbelas.CDSAss2.CreateDataSet;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha1').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha1').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha2').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha2').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha3').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha3').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha4').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha4').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  DM_TAbelas.CDSAss.Close;
  DM_TAbelas.CDSAss.CreateDataSet;
  while not DM_Tabelas.ZQCompr_Dados.Eof do
  begin
    if DM_Tabelas.ZQCompr_Dados.FieldByName('tipopessoa').AsString<>'I' then
    begin
      DM_TAbelas.CDSAss.Insert;
      DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_Dados.FieldByName('nome_parte').AsString;
      DM_Tabelas.CDSAsscargo.Value :='Comprador(a)';
      DM_TAbelas.CDSAss.Post;
    end;
    while not DM_TAbelas.ZQCompr_Resp_Dados.Eof do begin
      DM_TAbelas.CDSAss.Insert;
      DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_Resp_Dados.FieldByName('nomerepres').AsString;
      DM_Tabelas.CDSAsscargo.Value :='Responsável';
      DM_TAbelas.CDSAss.Post;
      DM_TAbelas.ZQCompr_Resp_Dados.Next;
    end;
    while not DM_TAbelas.ZQCompr_Conjuge.Eof do
    begin
      //17/06/2016
//      if DM_Tabelas.ZQCompr_Dadostipopessoa.Value='F' then
      if (DM_Tabelas.ZQCompr_Dados.FieldByName('tipopessoa').AsString='F') or (DM_Tabelas.ZQCompr_Dados.FieldByName('tipopessoa').AsString='I') then
      begin
        if DM_Tabelas.ZQCompr_conjuge.FieldByName('assina').AsString = 'S' Then
        Begin
          DM_TAbelas.CDSAss.Insert;
          DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_conjuge.FieldByName('nome').AsString;
          DM_Tabelas.CDSAsscargo.Value :='Conjuge';
          DM_TAbelas.CDSAss.Post;
        end;
      end;
      DM_TAbelas.ZQCompr_Conjuge.Next;
    end;
    DM_Tabelas.ZQCompr_Dados.Next;
  end;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_TAbelas.ZQCompr_Resp_Dados.Refresh;
  DM_TAbelas.ZQCompr_Conjuge.Refresh;
  PosicionarDadosParaQuadroResumo;
  if Frm_QuadroResumo=nil then
    Frm_QuadroResumo:=TFrm_QuadroResumo.Create(Application);
  Frm_QuadroResumo.npg.Value:=strtofloat(pgn.Text);
  Frm_QuadroResumo.RLReport1.Previewmodal;
  Frm_QuadroResumo:=nil;
  DM_Tabelas.ZQVenda.Filtered := False;
  Panel1.Visible:=false;
  Application.ProcessMessages;
end;

procedure TFrm_Venda.JvXPBar1Items1Click(Sender: TObject);
begin
  if not Verif_senha('Vendas','Relatórios Administrativo','') Then Exit;
  JvXPBar1.Visible := False;
  Panel1.Visible:=true;
  Application.ProcessMessages;
  if Frm_RelVenda=nil then
     Frm_RelVenda:=TFrm_RelVenda.Create(Application);
  Frm_RelVenda.Top := Frm_Venda.Top+100;
  Frm_RelVenda.Left := Frm_Venda.Left-4+(Frm_Venda.Width-Frm_RelVenda.Width);
  AbrirModal(Self, Frm_RelVenda);
//  Frm_RelVenda.Show;
  Frm_RelVenda:=nil;
//  Atualiza;
  Panel1.Visible:=false;
  Application.ProcessMessages;

end;

procedure TFrm_Venda.JvXPBar1Exit(Sender: TObject);
begin
  JvXPBar1.Visible := False;
end;

procedure TFrm_Venda.BTPesquisarClick(Sender: TObject);
begin
  Frm_AchaVenda.Top := Frm_Venda.Top+100;
  Frm_AchaVenda.Left := Frm_Venda.Left+(Frm_Venda.Width-Frm_AchaVenda.Width)  ;
  AbrirModal(Self, Frm_AchaVenda);
  tprice.Text:=DM_Tabelas.ZQVenda.FieldByName('tabela_Price').Text+'0';  
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.DBEdit11Enter(Sender: TObject);
begin
 if DM_Tabelas.ZQVenda.State in [DsInsert,DsEdit] then
 begin
   if fp.ItemIndex=0 then
      DM_Tabelas.ZQVenda.FieldByName('forma_reajuste').AsString:='PARCELAS FIXAS'
   else if fp.ItemIndex=1 then
   begin
      if simnao('Reajuste pelo IGPM ?','SIM') then
         DM_Tabelas.ZQVenda.FieldByName('forma_reajuste').AsString:='IGPM - FGV (ANUAL)'
      else
         DM_Tabelas.ZQVenda.FieldByName('forma_reajuste').AsString:='IPCA- 15 (ANUAL)';
   end;
 end;
end;

procedure TFrm_Venda.DBEdit11Exit(Sender: TObject);
begin
  if (DM_Tabelas.ZQVenda.State in [DsInsert]) Then Begin
    Pag_Venda.PageIndex := 1;
   // DBEVrvenda.SetFocus;
   XDBNumEdit3.SetFocus;
  end;

end;

procedure TFrm_Venda.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  frm_principal.lblvnd.Caption:='S';
  DM_TAbelas.ZQCompr_Dados.close;
  DM_Tabelas.ZQReBxHi.close;
  DM_Tabelas.ZQincorp_loteame.close;
  DM_Tabelas.ZQLoteamento.close;
  DM_tabelas.ZQNumOrdem.close;
  DM_Tabelas.ZQImovel.close;
  DM_Tabelas.ZQVendedor.close;
  DM_Tabelas.ZQRecebimento.close;
  DM_Tabelas.ZQComprador.close;
  DM_Tabelas.ZQTipoDoc.close;
  DM_tabelas.ZQQuadras.close;
  DM_tabelas.ZQContaBancaria.close;
  DM_tabelas.ZQCheque.close;
  DM_Tabelas.zqprocuradores.close;
  DM_Tabelas.ZQCidade.close;
  DM_TAbelas.ZQmemorial.close;
  ZQRecebimento.close;
  frm_principal.lblvnd.Caption:='N';
  DM_Tabelas.ZQVenda.close;
  Frm_Venda:=nil;
  Action:=CaFree;
end;

procedure TFrm_Venda.Pag_VendaClick(Sender: TObject);
begin
  Panel1.Visible:=true;
  Application.ProcessMessages;

  Atualizatebelaspagamento;

  Panel1.Visible:=false;

  if DM_Tabelas.ZQVenda.Active=false then
     DM_Tabelas.ZQVenda.open;


  if DM_Tabelas.ZQTipoDoc.Active=false then
     DM_Tabelas.ZQTipoDoc.open;

  Application.ProcessMessages;
  if Pag_Venda.PageIndex=2 then
  begin
    XCorrigido.Value:=0;
    iprice.Clear;
    sldpago.Value:=0;
    vr_price.Value:=0;
    Ultimo_Igpm.Value:=0;
    xpcorrecao_igpm.Value:=0;
    xvcorrecao_igpm.Value:=0;
    xdacrescimo.Value:=0;
    xvracrescimo.Value:=0;
    xddesconto.Value:=0;
    xvrdesconto.Value:=0;
    xdliquidado.Value:=0;
    vrfinal.Value:=0;
    tprice.Text:=DM_Tabelas.ZQVenda.FieldByName('tabela_Price').Text+'0';
    ultimo_igpm.SetFocus;
    XCorrigido.Value:=0;
//    xpcorrecao_igpm.SetFocus;

    if simnao('Deseja utilizar a Tabela Price Para Mora/Dia ?','SIM') then
       Lmora.Caption:='S'
    else
       Lmora.Caption:='N';

   ZQInandimplencia.close;
   ZQInandimplencia.sql.Clear;
//   ZQInandimplencia.sql.Add('select sum(saldo) as Parcela_corrigida, tabela_price,origem,documento,saldo,quadralote,dt_vencimento,venda_idvenda,v.mora,v.multa  FROM recebimento as r left join venda as v on r.venda_idvenda=v.idvenda where r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVendaidvenda.Text)+' and r.dt_vencimento<CURDATE() and r.saldo>0 and r.documento like''%P%''');
   ZQInandimplencia.sql.Add('select sum(saldo) as Parcela_corrigida, tabela_price,origem,documento,saldo,quadralote,dt_vencimento,venda_idvenda,v.mora,v.multa  FROM recebimento as r left join venda as v on r.venda_idvenda=v.idvenda where r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and r.dt_vencimento<CURDATE() and r.saldo>0 ');
   ZQInandimplencia.open;
   ZQInandimplencia.First;
   if (ZQInandimplencia.RecordCount>0) and (ZQInandimplencia.FieldByName('Parcela_corrigida').AsFloat>0) then
   begin
     mensagem('Cliente com Parcela(s) Atrasada(s)');
     Showmessage('Aguarde! Vou Atualizar os Saldos Atrazados.');
     ZQInandimplencia.close;
     ZQInandimplencia.sql.Clear;
     if Lmora.Caption='N' then
     begin
       ZQInandimplencia.sql.Add('select TIMESTAMPDIFF(DAY,dt_vencimento,now()) as dias,Round(((((saldo/30)*v.Mora)/100)*(TIMESTAMPDIFF(DAY,dt_vencimento,now()))),2) as mora_dia, Round(Round((saldo*v.Multa),2)/100,2) as mora_contrato,');
       ZQInandimplencia.sql.Add('(saldo+(Round(Round((saldo*v.Multa),2)/100,2))+(Round(((((saldo/30)*v.Mora)/100)*(TIMESTAMPDIFF(DAY,dt_vencimento,now()))),2))) as parcela_corrigida, tabela_price,origem,documento,saldo,quadralote,dt_vencimento,venda_idvenda,v.mora,v.multa  ');
//       ZQInandimplencia.sql.Add(' FROM recebimento as r left join venda as v on r.venda_idvenda=v.idvenda  where r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVendaidvenda.Text)+' and r.dt_vencimento<CURDATE() and r.saldo>0 and r.documento like ''%P%''');
       ZQInandimplencia.sql.Add(' FROM recebimento as r left join venda as v on r.venda_idvenda=v.idvenda  where r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and r.dt_vencimento<CURDATE() and r.saldo>0 ');
       ZQInandimplencia.open;
     end
     else
     begin
       ZQInandimplencia.sql.Add('select TIMESTAMPDIFF(DAY,dt_vencimento,now()) as dias,Round(((((saldo/30)*tabela_price)/100)*(TIMESTAMPDIFF(DAY,dt_vencimento,now()))),2) as mora_dia, Round(Round((saldo*v.Multa),2)/100,2) as mora_contrato,');
       ZQInandimplencia.sql.Add('(saldo+(Round(Round((saldo*v.Multa),2)/100,2))+(Round(((((saldo/30)*tabela_price)/100)*(TIMESTAMPDIFF(DAY,dt_vencimento,now()))),2))) as parcela_corrigida, tabela_price,origem,documento,saldo,quadralote,dt_vencimento,venda_idvenda,v.mora,v.multa  ');
//       ZQInandimplencia.sql.Add(' FROM recebimento as r left join venda as v on r.venda_idvenda=v.idvenda  where r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVendaidvenda.Text)+' and r.dt_vencimento<CURDATE() and r.saldo>0 and r.documento like ''%P%''');
       ZQInandimplencia.sql.Add(' FROM recebimento as r left join venda as v on r.venda_idvenda=v.idvenda  where r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and r.dt_vencimento<CURDATE() and r.saldo>0 ');
       ZQInandimplencia.open;

     end;
     ZQInandimplencia.first;
     while not ZQInandimplencia.Eof do
     begin
       XCorrigido.Value:=XCorrigido.Value+ZQInandimplencia.FieldByName('Parcela_corrigida').AsFloat;
       ZQInandimplencia.Next;
     end;
   //  Pag_Venda.PageIndex:=0;                        aq
   end;
   ZQInandimplencia.close;
  end;
  if Pag_Venda.PageIndex=3 then
  begin
    DM_Tabelas.ZQmemorial.Close;
    DM_Tabelas.ZQmemorial.SQL.Clear;
    DM_Tabelas.ZQmemorial.SQL.Add('Select idmemorial_descritivo,idvenda,descricao from memorial_descritivo');
    DM_Tabelas.ZQmemorial.Open;
  end
  else
  begin
    DM_Tabelas.ZQmemorial.Close;
  end;
  if Pag_Venda.PageIndex=4 then
  begin
    DM_Tabelas.ZQproposta.Open;
    DM_Tabelas.ZQRecebimento_proposta.close;
    DM_Tabelas.ZQRecebimento_proposta.SQL.Clear;
    DM_Tabelas.ZQRecebimento_proposta.SQL.Add('Select * from Recebimento_proposta order by DT_Vencimento');
    DM_Tabelas.ZQRecebimento_proposta.Open;

    DM_Tabelas.ZQComprador_proposta.close;
    DM_Tabelas.ZQComprador_proposta.MasterSource:=DM_Tabelas.DS_proposta;
    DM_Tabelas.ZQComprador_proposta.MasterFields:='idproposta';
    DM_Tabelas.ZQComprador_proposta.LinkedFields:='venda_idvenda';
    DM_Tabelas.ZQComprador_proposta.SQL.clear;
    DM_Tabelas.ZQComprador_proposta.SQL.Add('Select * from comprador_proposta ');
    DM_Tabelas.ZQComprador_proposta.Open;

    DM_Tabelas.ZQVendedor_proposta.close;
    DM_Tabelas.ZQVendedor_proposta.MasterSource:=DM_Tabelas.DS_proposta;
    DM_Tabelas.ZQVendedor_proposta.MasterFields:='idproposta';
    DM_Tabelas.ZQVendedor_proposta.LinkedFields:='venda_idvenda';
    DM_Tabelas.ZQVendedor_proposta.SQL.clear;
    DM_Tabelas.ZQVendedor_proposta.SQL.Add('Select *  from vendedor_proposta ');
    DM_Tabelas.ZQVendedor_proposta.Open;


    DM_Tabelas.ZQCheque_proposta.close;
    DM_Tabelas.ZQCheque_proposta.MasterSource:=DM_Tabelas.DS_ZQRecebimento_proposta;
    DM_Tabelas.ZQCheque_proposta.MasterFields:='idrecebimento_proposta';
    DM_Tabelas.ZQCheque_proposta.LinkedFields:='idvinculo';
    DM_Tabelas.ZQCheque_proposta.SQL.clear;
    DM_Tabelas.ZQCheque_proposta.SQL.Add('select * from cheque_rec_proposta ');
    DM_Tabelas.ZQCheque_proposta.Open;


  end
  else
  begin
    DM_Tabelas.ZQComprador_proposta.close;
    DM_Tabelas.ZQVendedor_proposta.close;
    DM_Tabelas.ZQCheque_proposta.close;
    DM_Tabelas.ZQproposta.close;
    DM_Tabelas.ZQRecebimento_proposta.Close;
    DM_Tabelas.ZQVenda.Filtered := false;
  end;

end;

procedure TFrm_Venda.JvXPBar1Click(Sender: TObject);
begin
  JvXPBar1.Visible := False;
end;

procedure TFrm_Venda.Timer1Timer(Sender: TObject);
begin
  Frm_Venda.FormStyle:=fsStayOnTop;
  Timer1.Enabled:=false;
end;

procedure TFrm_Venda.DBEPRiceEnter(Sender: TObject);
begin
  ZQPrice.Open;
  ZQPrice.First;
  while not ZQPrice.Eof do
  begin
   DBEPRice.Items.Add(transform(ZQprice.FieldByName('Taxa').AsFloat,'##0.0000'));
   ZQPrice.Next;
  end;
  ZQPrice.close; 

  if DM_Tabelas.zqvenda.state in [DsInsert] then
  begin
    DBEPRice.text:=DM_Tabelas.ZQConfiguracoes.FieldByName('price').AsString;
    if not empty(DM_Tabelas.ZQConfiguracoes.FieldByName('price').AsString) then
      DM_Tabelas.ZQVenda.FieldByName('tabela_Price').AsFloat:=strtofloat(DM_Tabelas.ZQConfiguracoes.FieldByName('price').AsString);
  end;
end;

procedure TFrm_Venda.tpriceEnter(Sender: TObject);
begin
  ZQPrice.Open;
  ZQPrice.First;
  tprice.Items.Clear;
  while not ZQPrice.Eof do
  begin
   tprice.Items.Add(transform(ZQprice.FieldByName('Taxa').AsFloat,'##0.0000'));
   ZQPrice.Next;
  end;
  ZQPrice.close;
end;

procedure TFrm_Venda.tpriceExit(Sender: TObject);
var
S:string;
begin
  S:=floattostr(xdsldvicenda.Value-xdsldatrasadas.Value);
  DM_Tabelas.ZQPrice.close;
  DM_Tabelas.ZQPrice.SQL.Clear;
  DM_Tabelas.ZQPrice.SQL.Add('select idPrice,Taxa,Mes,coeficiente_multiplicador,Coeficiente_amortizador from Price where taxa='+quotedstr(StringReplace(alltrim(tprice.text),',','.',[rfReplaceAll]))+' and mes='+quotedstr(S));
  DM_Tabelas.ZQPrice.open;
  if DM_Tabelas.ZQPrice.RecordCount>0 then
  begin
    iprice.Text:=DM_Tabelas.ZQprice.FieldByName('Coeficiente_amortizador').Text;
    vr_price.Value:=ultparc.Value/(DM_Tabelas.ZQprice.FieldByName('Coeficiente_amortizador').AsFloat/1000);
  end;
  DM_Tabelas.ZQPrice.close;
  DM_Tabelas.ZQPrice.SQL.Clear;
  DM_Tabelas.ZQPrice.SQL.Add('select idPrice,Taxa,Mes,coeficiente_multiplicador,Coeficiente_amortizador from Price  order by taxa,mes');
  DM_Tabelas.ZQPrice.open;
  DM_Tabelas.ZQPrice.close;
  xdliquidado.Value:=vr_price.Value+Xcorrigido.Value;
  vrfinal.Value:= DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat+xdliquidado.Value;
end;

procedure TFrm_Venda.xvracrescimoExit(Sender: TObject);
begin
  if xvracrescimo.Value>0 then
  begin
    xdliquidado.Value:=xdliquidado.Value+xvracrescimo.Value+Xcorrigido.Value;
    vrfinal.Value:= sldpago.Value+xdliquidado.Value;
  end;
end;

procedure TFrm_Venda.xvrdescontoExit(Sender: TObject);
begin
  if xvrdesconto.Value>0 then
  begin
    xdliquidado.Value:=(xdliquidado.Value-xvrdesconto.Value)+Xcorrigido.Value;
    vrfinal.Value:= sldpago.Value+xdliquidado.Value;
  end;
  dxButton1.Enabled:=true;
  dxButton2.Enabled:=true;
  dxButton5.Enabled:=true;
end;

procedure TFrm_Venda.xdliquidadoEnter(Sender: TObject);
begin
  vrfinal.Value:= sldpago.Value+xdliquidado.Value;
end;

procedure TFrm_Venda.xdacrescimoExit(Sender: TObject);
begin
  if xdacrescimo.Value>0 then
  begin
    xvracrescimo.Value:=(xdliquidado.Value*xdacrescimo.Value)/100;
  end;

end;

procedure TFrm_Venda.XDBNumEdit3Enter(Sender: TObject);
var
gera:Boolean;
begin
   if DM_Tabelas.zqvenda.state in [DsInsert] then
   begin
       if simnao('Gerar Comissão ?','SIM') then
       begin
         if DM_Tabelas.ZQVenda.FieldByName('Perc_comissao').AsFloat>0 then
         begin
           DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat:=(DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat*DM_Tabelas.ZQVenda.FieldByName('Perc_comissao').AsFloat)/100;
         end;
       end
       else
       begin
         DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat:=0;
       end;
   end;
   if DM_Tabelas.zqvenda.state in [DsEdit] then
   begin
     if DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat>0 then
     begin
       if DM_Tabelas.ZQVenda.FieldByName('Perc_comissao').AsFloat>0 then
       begin
         DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat:=(DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat*DM_Tabelas.ZQVenda.FieldByName('Perc_comissao').AsFloat)/100;
       end;
     end
     else if simnao('Gerar Comissão ?','SIM') then
     begin
         if DM_Tabelas.ZQVenda.FieldByName('Perc_comissao').AsFloat>0 then
         begin
           DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat:=(DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat*DM_Tabelas.ZQVenda.FieldByName('Perc_comissao').AsFloat)/100;
         end;
     end
     else
     begin
         DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat:=0;
     end;
   end;
end;

procedure TFrm_Venda.xddescontoExit(Sender: TObject);
begin
  if xddesconto.Value>0 then
  begin
    xvrdesconto.Value:=(xdliquidado.Value*xddesconto.Value)/100;
  end;

end;

procedure TFrm_Venda.vrfinalEnter(Sender: TObject);
begin
  vrfinal.Value:= sldpago.Value+xdliquidado.Value;
end;

procedure TFrm_Venda.Ultimo_IgpmEnter(Sender: TObject);
var
S:String;
begin
  S:=floattostr(xdsldvicenda.Value-xdsldatrasadas.Value);
  DM_Tabelas.ZQPrice.close;
  DM_Tabelas.ZQPrice.SQL.Clear;
  DM_Tabelas.ZQPrice.SQL.Add('select idPrice,Taxa,Mes,coeficiente_multiplicador,Coeficiente_amortizador from Price where taxa='+quotedstr(StringReplace(alltrim(tprice.text),',','.',[rfReplaceAll]))+' and mes='+quotedstr(S));
  DM_Tabelas.ZQPrice.open;
  if DM_Tabelas.ZQPrice.RecordCount>0 then
  begin
    iprice.Text:=DM_Tabelas.ZQprice.FieldByName('Coeficiente_amortizador').Text;
    vr_price.Value:=ultparc.Value/(DM_Tabelas.ZQprice.FieldByName('Coeficiente_amortizador').AsFloat/1000);
  end;
  DM_Tabelas.ZQPrice.close;
  DM_Tabelas.ZQPrice.SQL.Clear;
  DM_Tabelas.ZQPrice.SQL.Add('select idPrice,Taxa,Mes,coeficiente_multiplicador,Coeficiente_amortizador from Price  order by taxa,mes');
  DM_Tabelas.ZQPrice.open;
  DM_Tabelas.ZQPrice.close;
  xdliquidado.Value:=vr_price.Value+Xcorrigido.Value;
  sldpago.Value:=(XDentrada.Value+xdparcelas.Value)-xdsldremanescente.Value;
  vrfinal.Value:= sldpago.Value+xdliquidado.Value;
end;

procedure TFrm_Venda.Ultimo_IgpmExit(Sender: TObject);
var
vr,vr2:string;
pr:integer;
begin
  vr:=StringReplace(alltrim(transform(ultparc.Value,'0.00')),',','.',[rfReplaceAll]);
  ZQparcelas12.close;
  ZQparcelas12.SQL.Clear;
  ZQparcelas12.SQL.Add('select origem,documento,round(valor,2) as vlr,saldo,substr(ordem,-7,3) as orde  FROM recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and round(r.valor,2)='+quotedstr(alltrim(vr))+' and r.saldo=0 and r.documento like ''%-P-%''');
  ZQparcelas12.open;

  if ZQparcelas12.RecordCount>0 then
  begin
    if ZQparcelas12.recordcount>=12 then
      pr:=ZQparcelas12.recordcount-12
    else
      pr:=ZQparcelas12.recordcount;

    xpcorrecao_igpm.value:=(Ultimo_Igpm.Value/12)*pr;
    ZQparcelas12.close;
  end
  else
  begin
    ZQparcelas12.close;
    ZQparcelas12.SQL.Clear;
    ZQparcelas12.SQL.Add('select origem,documento,round(valor,2) as vlr,saldo,substr(ordem,-7,3) as orde  FROM recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and r.saldo=0 and r.documento like ''%-P-%'' order by vlr');
    ZQparcelas12.open;
    ZQparcelas12.Last;   // pegar o ultimo valor baixado
    if ZQparcelas12.RecordCount>0 then
    begin
      vr2:=StringReplace(alltrim(transform(ZQparcelas12.FieldByName('vlr').AsFloat,'0.00')),',','.',[rfReplaceAll]);
      ZQparcelas12.close;
      ZQparcelas12.SQL.Clear;
      ZQparcelas12.SQL.Add('select origem,documento,round(valor,2) as vlr,saldo,substr(ordem,-7,3) as orde  FROM recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and round(r.valor,2)='+quotedstr(alltrim(vr2))+' and r.saldo=0 and r.documento like ''%-P-%''');
      ZQparcelas12.open;
      if ZQparcelas12.RecordCount>0 then
      begin
        if ZQparcelas12.recordcount>=12 then
          pr:=ZQparcelas12.recordcount-12
        else
          pr:=ZQparcelas12.recordcount;

        xpcorrecao_igpm.value:=(Ultimo_Igpm.Value/12)*pr;
        ZQparcelas12.close;

      end;
    end;
  end;
//  xpcorrecao_igpm.value:=(Ultimo_Igpm.Value/12)*xdparpagas.Value;

  if xpcorrecao_igpm.Value>0 then
  begin
    xvcorrecao_igpm.Value:=(vr_price.Value*xpcorrecao_igpm.Value)/100;
  end;
  xdliquidado.Value:=xdliquidado.Value+xvcorrecao_igpm.Value+Xcorrigido.Value;
  vrfinal.Value:= sldpago.Value+xdliquidado.Value;
end;

procedure TFrm_Venda.dxButton2Click(Sender: TObject);
begin
  if simnao('Deseja Efetuar a Quitação ?','SIM') then
  begin
    if Frmquitacao=nil then
       Frmquitacao:=TFrmquitacao.Create(Self);

    Frmquitacao.Elote.Text:= DM_Tabelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').AsString;
    Frmquitacao.xvequta.Value:=xdliquidado.Value;

    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and quadralote='+quotedstr(Frmquitacao.Elote.text)+' and saldo>0');
    DM_tabelas.ZQRecebimento.open;


    Frmquitacao.show;
  end;

end;

procedure TFrm_Venda.dxButton1Click(Sender: TObject);
var
varvenda:integer;
begin
  try
   varvenda := DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
   DM_Tabelas.ZQVenda.Filtered := false;
   DM_TAbelas.ZQVenda.Filter := 'idvenda='+inttostr(varvenda);
   DM_Tabelas.ZQVenda.Filtered := True;
   ZQparcelas_pagas.close;
   ZQparcelas_pagas.SQL.Clear;
   ZQparcelas_pagas.SQL.Add('Select tipdoc,ordem,documento,contabil,dt_vencimento,valor,quadralote,idrecebimento,saldo,origem,adversa,venda_idvenda, sum(valor) as parcela from recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) ');
   ZQparcelas_pagas.SQL.Add('  and r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and  r.documento like ''%-P-%'' and r.saldo=0 group by r.idrecebimento');
   ZQparcelas_pagas.open;
   DM_Tabelas.CDSPagParcelavalor.DisplayFormat:='###,###,##0.00';
   DM_Tabelas.CDSPagEntradavalor.DisplayFormat:='###,###,##0.00';

   ZQresumo_parcelas.close;
   ZQresumo_parcelas.SQL.Clear;
//   ZQresumo_parcelas.SQL.Add('select origem,documento,round(valor,2) as vr,saldo,count(valor) as ct,(round(valor,2)*count(valor)) as soma FROM recebimento where origem='+quotedstr('V-'+DM_Tabelas.ZQVendaidvenda.Text)+' and venda_idvenda='+quotedstr(DM_Tabelas.ZQVendaidvenda.Text)+' and saldo=0 and documento like ''%-P-%'' group by round(valor,2)');
   ZQresumo_parcelas.SQL.Add('select origem,documento,round(valor,2) as vr,saldo,count(valor) as ct,(round(valor,2)*count(valor)) as soma FROM recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and r.saldo=0 and r.documento like ''%-P-%'' group by round(r.valor,2)');
   ZQresumo_parcelas.open;

   ZQResumo_entrada.close;
   ZQResumo_entrada.SQL.Clear;
//   ZQResumo_entrada.SQL.Add('select origem,documento,round(valor,2) as vr,saldo,count(valor) as ct,(round(valor,2)*count(valor)) as soma FROM recebimento where origem='+quotedstr('V-'+DM_Tabelas.ZQVendaidvenda.Text)+' and venda_idvenda='+quotedstr(DM_Tabelas.ZQVendaidvenda.Text)+' and saldo=0 and documento like ''%-E-%'' group by round(valor,2)');
   ZQResumo_entrada.SQL.Add('select origem,documento,round(valor,2) as vr,saldo,count(valor) as ct,(round(valor,2)*count(valor)) as soma FROM recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and r.saldo=0 and r.documento like ''%-E-%'' group by round(r.valor,2)');
   ZQResumo_entrada.open;

   ZQtot_parcelas.close;
   ZQtot_parcelas.SQL.Clear;
//   ZQtot_parcelas.SQL.Add('select origem,documento,round(valor,2) as vr,saldo,count(valor) as ct,(round(valor,2)*count(valor)) as soma FROM recebimento where origem='+quotedstr('V-'+DM_Tabelas.ZQVendaidvenda.Text)+' and venda_idvenda='+quotedstr(DM_Tabelas.ZQVendaidvenda.Text)+' and saldo=0 group by venda_idvenda');
   ZQtot_parcelas.SQL.Add('select origem,documento,round(valor,2) as vr,saldo,count(valor) as ct,(round(valor,2)*count(valor)) as soma FROM recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text)+' and r.saldo=0 group by r.venda_idvenda');
   ZQtot_parcelas.open;

   DM_Tabelas.ZQEmpresa.open;
   try
     if FrmResumo_quita=nil then
        FrmResumo_quita:=TFrmResumo_quita.Create(nil);
//    FrmResumo_quita.QRResumo_quita.OnPreview:=Frm_principal.Prev;
     FrmResumo_quita.QRResumo_quita.PreviewModal;
   finally
     FreeAndNil(FrmResumo_quita);
   end;
  finally
    DM_Tabelas.CDSPagParcelavalor.DisplayFormat:='###,###,##0.000000';
    DM_Tabelas.CDSPagEntradavalor.DisplayFormat:='###,###,##0.000000';
    DM_Tabelas.ZQEmpresa.close;
    DM_Tabelas.ZQvnd_quadro.close;
    DM_Tabelas.ZQVenda.Filtered := false;
    DM_TAbelas.ZQVenda.locate('idvenda',varvenda,[]);
    DM_Tabelas.ZQVenda.refresh;
//    DM_TAbelas.ZQCompr_Dados.close;
    if DM_TAbelas.ZQCompr_Dados.Active=true then
       DM_TAbelas.ZQCompr_Dados.Refresh;
    DM_TAbelas.ZQCompr_Resp_Dados.close;
    DM_TAbelas.ZQCompr_Conjuge.close;
    DM_Tabelas.ZQCidade.close;
    DM_TAbelas.CDSAss1.Close;
    DM_TAbelas.CDSAss2.Close;
    DM_TAbelas.CDSAss.Close;
    ZQparcelas_pagas.close;
    ZQresumo_parcelas.close;
    ZQResumo_entrada.close;
    ZQtot_parcelas.close;
  end;
end;

procedure TFrm_Venda.DBGrid2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //bloqueia a tecla ctrl+del
  if ((Shift = [ssCtrl]) and (key = vk_delete)) THEN
     Abort;
end;

procedure TFrm_Venda.DBGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //bloqueia a tecla ctrl+del
  if ((Shift = [ssCtrl]) and (key = vk_delete)) THEN
     Abort;
end;

procedure TFrm_Venda.DBGrid2DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var
  Check: Integer;
  R: TRect;
begin
  if Column.FieldName = 'Marcar' then //COLOQUE O NOME DA COLUNA, NO MEU CASO 'MARCAR'
  begin
    DBGrid2.Canvas.FillRect(Rect);
    Check := 0;
    if DM_TAbelas.CDSCompradorTempmarcar.AsBoolean = true then
      Check := DFCS_CHECKED
    else
      Check := 0;
    R:=Rect;
    InflateRect(R,-2,-2); {Diminue o tamanho do CheckBox}
    DrawFrameControl(DBGrid2.Canvas.Handle,R,DFC_BUTTON, DFCS_BUTTONCHECK or Check);
  end;
end;

procedure TFrm_Venda.DBGrid2CellClick(Column: TColumn);
begin
//  if Column.FieldName = 'Marcar' then
//  begin
    DM_TAbelas.CDSCompradorTemp.Edit;
    if DM_TAbelas.CDSCompradorTempmarcar.AsBoolean = true then
      DM_TAbelas.CDSCompradorTempmarcar.AsBoolean := false
    else
      DM_TAbelas.CDSCompradorTempmarcar.AsBoolean := true;
    DM_TAbelas.CDSCompradorTemp.Post
//  end;
end;

procedure TFrm_Venda.DBComboBox1Exit(Sender: TObject);
begin
 if DM_Tabelas.ZQVenda.State in [DsInsert,DsEdit] then
 begin
   fp.SetFocus;
 end;  
end;

procedure TFrm_Venda.DBEPRiceExit(Sender: TObject);
begin
  if DM_Tabelas.zqvenda.state in [DsInsert] then
  begin
     if (alltrim(DBEPRice.text)='0,000') or (alltrim(DBEPRice.text)='0,0000') or (alltrim(DBEPRice.text)='0') or (alltrim(DBEPRice.text)='') then
     begin
       DBEEntrada.Columns[0].Visible:=true;
       DBGPartelas.Columns[0].Visible:=true;
       btndes.Visible:=true;
       Button1.Visible:=true;
     end
     else
     begin
       DBEEntrada.Columns[0].Visible:=false;
       DBGPartelas.Columns[0].Visible:=false;
       btndes.Visible:=false;
       Button1.Visible:=false;
     end;
  end;
end;

procedure TFrm_Venda.DBEEntradaColEnter(Sender: TObject);
begin
  if DBEEntrada.SelectedField = DM_Tabelas.CDSPagEntradamarcar then
    DBEEntrada.Options := DBEEntrada.Options - [dgEditing]
  else
    DBEEntrada.Options := DBEEntrada.Options + [dgEditing];
    
end;

procedure TFrm_Venda.DBGPartelasColEnter(Sender: TObject);
begin
  if DBGPartelas.SelectedField = DM_Tabelas.CDSPagParcelamarcar then
    DBGPartelas.Options := DBGPartelas.Options - [dgEditing]
  else
    DBGPartelas.Options := DBGPartelas.Options + [dgEditing];

end;

procedure TFrm_Venda.btndesClick(Sender: TObject);
begin
  DM_Tabelas.CDSPagEntrada.First;
  while not DM_Tabelas.CDSPagEntrada.Eof do
  begin
    DM_Tabelas.CDSPagEntrada.Edit;
    if  DM_Tabelas.CDSPagEntradamarcar.Value = 0 Then
      DM_Tabelas.CDSPagEntradamarcar.Value := 1
    else
      DM_Tabelas.CDSPagEntradamarcar.Value := 0;
    DM_Tabelas.CDSPagEntrada.post;
    DM_Tabelas.CDSPagEntrada.Next;
  end;
  DM_Tabelas.CDSPagEntrada.First;
end;

procedure TFrm_Venda.Button1Click(Sender: TObject);
begin
  DM_Tabelas.CDSPagParcela.First;
  while not DM_Tabelas.CDSPagParcela.Eof do
  begin
    DM_Tabelas.CDSPagParcela.Edit;
    if  DM_Tabelas.CDSPagParcelamarcar.Value = 0 Then
      DM_Tabelas.CDSPagParcelamarcar.Value := 1
    else
      DM_Tabelas.CDSPagParcelamarcar.Value := 0;
    DM_Tabelas.CDSPagParcela.post;
    DM_Tabelas.CDSPagParcela.Next;
  end;
  DM_Tabelas.CDSPagParcela.First;
end;

procedure TFrm_Venda.dxButton3Click(Sender: TObject);
begin
  if Frm_RelVenda=nil then
     Frm_RelVenda:=TFrm_RelVenda.Create(Self);
  Frm_RelVenda.Equadra.Text:=DM_Tabelas.ZQVenda.FieldByName('quadra').AsString;
  Frm_RelVenda.Elote.Text:=DM_Tabelas.ZQVenda.FieldByName('lote').AsString;

  JvXPBar1.Visible := False;
  Panel1.Visible:=true;
  Application.ProcessMessages;

  Frm_RelVenda.Top := Frm_Venda.Top+100;
  Frm_RelVenda.Left := Frm_Venda.Left-4+(Frm_Venda.Width-Frm_RelVenda.Width);

  AbrirModal(Self, Frm_RelVenda);
  Frm_RelVenda:=nil;
//  Frm_RelVenda.Show;
//  Atualiza;
  Panel1.Visible:=false;
  Application.ProcessMessages;

end;

procedure TFrm_Venda.JvXPBar1Items2Click(Sender: TObject);
begin
  JvXPBar1.Visible := False;
  Panel1.Visible:=true;
  Application.ProcessMessages;
  DM_Tabelas.ZQmemorial.Open;
  try
    if Frm_QuadroMemorial=nil then
      Frm_QuadroMemorial:=TFrm_QuadroMemorial.Create(Application);
    Frm_QuadroMemorial.RLReport1.Preview;
  finally
    DM_Tabelas.ZQmemorial.close;
  end;
  Panel1.Visible:=false;
  Application.ProcessMessages;


end;

procedure TFrm_Venda.JvXPBar1Items3Click(Sender: TObject);
var
  varvenda : integer;
begin
  Panel1.Visible:=true;
  Application.ProcessMessages;

  Frm_Venda.FormStyle:=fsNormal;
  Application.ProcessMessages;
//  Timer1.Enabled:=true;
  JvXPBar1.Visible := False;
  Atualizatebelaspagamento;

  varvenda := DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
  DM_TAbelas.ZQVenda.Filter := 'idvenda='+inttostr(varvenda);
  DM_Tabelas.ZQVenda.Filtered := True;
  DM_TAbelas.ZQCompr_Dados.Open;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_TAbelas.ZQCompr_Resp_Dados.Open;
  DM_TAbelas.ZQCompr_Resp_Dados.Refresh;
  DM_TAbelas.ZQCompr_Conjuge.Open;
  DM_TAbelas.ZQCompr_Conjuge.Refresh;
  DM_Tabelas.ZQCompr_Dados.First;
  if DM_Tabelas.ZQLoteamento.Active=false then
     DM_Tabelas.ZQLoteamento.Open;
  DM_Tabelas.ZQLoteamento.Locate('idloteamento',DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger,[]);
  DM_Tabelas.ZQCidade.Open;
  try
    if Frm_Loteamento=nil then
       Frm_Loteamento := TFrm_Loteamento.Create(self);
    Frm_Loteamento.atualiza_lotes;

  finally
    Frm_Loteamento:=nil;
    Frm_Loteamento.Free;
  end;

  DM_TAbelas.CDSAss1.Close;
  DM_TAbelas.CDSAss1.CreateDataSet;
{  DM_TAbelas.CDSAss1.Insert;
  DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamentonomeloteamento.Value;
  DM_Tabelas.CDSAss1cargo.Value :=' Vendedor';
//

  if alltrim(DM_TAbelas.ZQLoteamentonomeloteamento.AsString)<>'ALICIO VILAR PONTES' then
     DM_Tabelas.CDSAss1responsavel.Value := 'PP: '+DM_Tabelas.CDSIncorpnomeparti.Value;
 }


  DM_Tabelas.zqprocuradores.Filtered:=false;
  DM_Tabelas.zqprocuradores.Filter:='Assina=''S''';
  DM_Tabelas.zqprocuradores.filtered:=true;



  if (DM_Tabelas.CDSIncorpassina.Value='S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString<>'S') then
  begin
    DM_TAbelas.CDSAss1.Insert;
    DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
    DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
    if alltrim(DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString)<>'ALICIO VILAR PONTES' then
       DM_Tabelas.CDSAss1responsavel.Value := DM_Tabelas.CDSIncorpnomeparti.Value;
    DM_Tabelas.CDSAss1conjuge.Value := DM_Tabelas.CDSIncorpconjuge.Value;
    DM_Tabelas.CDSAss1idconjuge.Value := DM_Tabelas.CDSIncorpidconjuge.Value;
    DM_Tabelas.CDSAss1assina.Value := DM_Tabelas.CDSIncorpassina.Value;
    DM_TAbelas.CDSAss1.Post;
  end;

  if (DM_Tabelas.CDSIncorpassina.Value='S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString='S') then
  begin
    DM_Tabelas.zqprocuradores.First;
    while not DM_Tabelas.zqprocuradores.Eof do
    begin
      DM_TAbelas.CDSAss1.Insert;
      DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
      DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
      DM_Tabelas.CDSAss1responsavel.Value := DM_Tabelas.ZQProcuradores.FieldByName('nome').AsString;

      DM_Tabelas.CDSAss1conjuge.Value := DM_Tabelas.CDSIncorpconjuge.Value;
      DM_Tabelas.CDSAss1idconjuge.Value := DM_Tabelas.CDSIncorpidconjuge.Value;
      DM_Tabelas.CDSAss1assina.Value := DM_Tabelas.CDSIncorpassina.Value;


      DM_TAbelas.CDSAss1.Post;
      DM_Tabelas.zqprocuradores.Next;
    end;
  end;

  if (DM_Tabelas.CDSIncorpassina.Value<>'S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString='S') then
  begin
    DM_Tabelas.zqprocuradores.First;
    while not DM_Tabelas.zqprocuradores.Eof do
    begin
      DM_TAbelas.CDSAss1.Insert;
      DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
      DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
      DM_Tabelas.CDSAss1responsavel.Value := DM_Tabelas.ZQProcuradores.FieldByName('nome').AsString;
  //    DM_Tabelas.CDSAss1assina.Value := DM_Tabelas.zqprocuradoresassina.Value;
      DM_TAbelas.CDSAss1.Post;
      DM_Tabelas.zqprocuradores.Next;
    end;
  end;


  if (DM_Tabelas.CDSIncorpassina.Value<>'S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString<>'S') then
  begin
    DM_TAbelas.CDSAss1.Insert;
    DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
    DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
    if alltrim(DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString)<>'ALICIO VILAR PONTES' then
       DM_Tabelas.CDSAss1responsavel.Value := DM_Tabelas.CDSIncorpnomeparti.Value;
    DM_TAbelas.CDSAss1.Post;
  end;


  DM_Tabelas.zqprocuradores.Filtered:=false;


//  DM_TAbelas.CDSAss1.Post;
  DM_TAbelas.CDSAss2.Close;
  DM_TAbelas.CDSAss2.CreateDataSet;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha1').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha1').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha2').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha2').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha3').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha3').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha4').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha4').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  DM_TAbelas.CDSAss.Close;
  DM_TAbelas.CDSAss.CreateDataSet;
  while not DM_Tabelas.ZQCompr_Dados.Eof do
  begin
    if DM_Tabelas.ZQCompr_Dados.FieldByName('tipopessoa').AsString<>'I' then
    begin
      DM_TAbelas.CDSAss.Insert;
      DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_Dados.FieldByName('nome_parte').AsString;
      DM_Tabelas.CDSAsscargo.Value :='Comprador(a)';
      DM_TAbelas.CDSAss.Post;
    end;
    while not DM_TAbelas.ZQCompr_Resp_Dados.Eof do begin
      DM_TAbelas.CDSAss.Insert;
      DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_Resp_Dados.FieldByName('nomerepres').AsString;
      DM_Tabelas.CDSAsscargo.Value :='Responsável';
      DM_TAbelas.CDSAss.Post;
      DM_TAbelas.ZQCompr_Resp_Dados.Next;
    end;
    while not DM_TAbelas.ZQCompr_Conjuge.Eof do
    begin
      //17/06/2016
//      if DM_Tabelas.ZQCompr_Dadostipopessoa.Value='F' then
      if (DM_Tabelas.ZQCompr_Dados.FieldByName('tipopessoa').AsString='F') or (DM_Tabelas.ZQCompr_Dados.FieldByName('tipopessoa').AsString='I') then
      begin
        if DM_Tabelas.ZQCompr_conjuge.FieldByName('assina').AsString = 'S' Then
        Begin
          DM_TAbelas.CDSAss.Insert;
          DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_conjuge.FieldByName('nome').AsString;
          DM_Tabelas.CDSAsscargo.Value :='Conjuge';
          DM_TAbelas.CDSAss.Post;
        end;
      end;
      DM_TAbelas.ZQCompr_Conjuge.Next;
    end;
    DM_Tabelas.ZQCompr_Dados.Next;
  end;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_TAbelas.ZQCompr_Resp_Dados.Refresh;
  DM_TAbelas.ZQCompr_Conjuge.Refresh;
  PosicionarDadosParaQuadroResumo;
  if Frm_QuadroResumo2=nil then
    Frm_QuadroResumo2:=TFrm_QuadroResumo2.Create(Application);
  Frm_QuadroResumo2.npg.Value:=strtofloat(pgn.Text);
  Frm_QuadroResumo2.RLReport1.Preview;
  DM_Tabelas.ZQVenda.Filtered := False;
  Panel1.Visible:=false;
  Application.ProcessMessages;

end;

procedure TFrm_Venda.JvXPBar2Items0Click(Sender: TObject);
var
path:string;
xcodpart:integer;
begin
  Pag_Venda.PageIndex:=4;
  JvXPBar2.Visible:=false;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('delete FROM participante_proposta');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('ALTER TABLE participante_proposta AUTO_INCREMENT = 1');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('insert  into  participante_proposta select * from participante');
  ZQLimpar_participante_proposta.ExecSQL;
  ZQLimpar_participante_proposta.close;


//13/01/2017

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('delete FROM cheque_rec_proposta');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('ALTER TABLE cheque_rec_proposta AUTO_INCREMENT = 1');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('delete FROM comprador_proposta');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('ALTER TABLE comprador_proposta AUTO_INCREMENT = 1');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('delete FROM conjuge_proposta');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('ALTER TABLE conjuge_proposta AUTO_INCREMENT = 1');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('insert  into  conjuge_proposta select * from conjuge');
  ZQLimpar_participante_proposta.ExecSQL;
  ZQLimpar_participante_proposta.close;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('delete FROM recebimento_proposta');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('ALTER TABLE recebimento_proposta AUTO_INCREMENT = 1');
  ZQLimpar_participante_proposta.ExecSQL;


  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('delete FROM vendedor_proposta');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;
  ZQLimpar_participante_proposta.SQL.Add('ALTER TABLE vendedor_proposta AUTO_INCREMENT = 1');
  ZQLimpar_participante_proposta.ExecSQL;

  ZQLimpar_participante_proposta.close;
  ZQLimpar_participante_proposta.SQL.Clear;


//  Pag_Venda.OnClick(Self);
  OpenDialog1.Execute;
  path:= ExtractFilePath(OpenDialog1.FileName); //(ExtractFilePath(Application.ExeName));
  CdsImovel.close;
  CdsImovel.LoadFromFile(path+'imoveis.xml');
  CdsImovel.open;
  ZQImovel.close;
  //ZQimovel.open;


  cds_proposta.close;
  cds_proposta.LoadFromFile(path+'proposta.xml');
  cds_proposta.open;
  Zqproposta.close;
//  ZQProposta.open;

  cds_comprador.close;
  cds_comprador.LoadFromFile(path+'comprador.xml');
  cds_comprador.open;
  ZQcomprador.close;
  //ZQcomprador.open;

  cds_comprador2.IndexName:='';
  cds_comprador2.Close;
  cds_comprador2.CreateDataSet;
  cds_comprador2.Open;
  cds_comprador2.IndexName:='';
  cds_comprador.First;
  while not cds_comprador.Eof do
  begin
    cds_comprador2.Insert;
    cds_comprador2idcomprador.Value:=cds_compradoridcomprador.Value;
    cds_comprador2promissario.Value:=cds_compradorpromissario.Value;
    cds_comprador2percentual.Value:=cds_compradorpercentual.Value;
    cds_comprador2Marcar.Value:=cds_compradorMarcar.Value;
    cds_comprador2venda_idvenda.Value:=cds_compradorvenda_idvenda.Value;
    cds_comprador2idproposta.Value:=cds_compradoridproposta.Value;
    cds_comprador2paticipante_idpaticipante.Value:=cds_compradorpaticipante_idpaticipante.Value;
    cds_comprador2.Post;
    cds_comprador.next;
  end;



  cds_vendedor.close;
  cds_vendedor.LoadFromFile(path+'vendedor.xml');
  cds_vendedor.open;
  ZQvendedor.close;
//  ZQvendedor.open;



  cds_recebimento.close;
  cds_recebimento.LoadFromFile(path+'recebimento.xml');
  cds_recebimento.open;
  ZQRecebimento_proposta.close;
//  ZQRecebimento_proposta.open;


  cds_cheque.close;
  cds_cheque.LoadFromFile(path+'cheque.xml');
  cds_cheque.open;
  ZQcheques.close;
  //ZQcheques.open;

  cds_participante.close;
  cds_participante.LoadFromFile(path+'participante.xml');
  cds_participante.open;

  cds_participante2.MasterSource:=nil;
  cds_participante2.MasterFields:='';
  cds_participante2.IndexName:='';

  cds_participante2.close;
  cds_participante2.CreateDataSet;
  cds_participante2.open;

  cds_participante.First;
  Gauge2.Progress:=0;
  Gauge2.MaxValue:=cds_participante.RecordCount;

  while not cds_participante.Eof do
  begin
    Gauge2.Progress:=cds_participante.RecNo;
    cds_participante2.Insert;
    cds_participante2idpaticipante.Value:=cds_participanteidpaticipante.Value;
//    cds_participante2idnovo.Value:=xcodpart;
    cds_participante2nome_parte.Value:=cds_participantenome_parte.Value;
    cds_participante2doc1.Value:=cds_participantedoc1.Value;
    cds_participante2doc2.Value:=cds_participantedoc2.Value;
    cds_participante2endereco.Value:=cds_participanteendereco.Value;
    cds_participante2bairro.Value:=cds_participantebairro.Value;
    cds_participante2cidade.Value:=cds_participantecidade.Value;
    cds_participante2cep.Value:=cds_participantecep.Value;
    cds_participante2ende_cob.Value:=cds_participanteende_cob.Value;
    cds_participante2bairro_cob.Value:=cds_participantebairro_cob.Value;
    cds_participante2cidade_cob.Value:=cds_participantecidade_cob.Value;
    cds_participante2cep_cob.Value:=cds_participantecep_cob.Value;
    cds_participante2Fone1.Value:=cds_participanteFone1.Value;
    cds_participante2fone2.Value:=cds_participantefone2.Value;
    cds_participante2fone3.Value:=cds_participantefone3.Value;
    cds_participante2tipopessoa.Value:=cds_participantetipopessoa.Value;
    if cds_participanteaniversario.Value>0 then
       cds_participante2aniversario.Value:=cds_participanteaniversario.Value;
    cds_participante2nacionalidade.Value:=cds_participantenacionalidade.Value;
    cds_participante2email.Value:=cds_participanteemail.Value;
    cds_participante2naturalidade.Value:=cds_participantenaturalidade.Value;
    cds_participante2localdetrab.Value:=cds_participantelocaldetrab.Value;
    cds_participante2profissao.Value:=cds_participanteprofissao.Value;
    cds_participante2estadocivil.Value:=cds_participanteestadocivil.Value;
    cds_participante2renda.Value:=cds_participanterenda.Value;
    cds_participante2observacao.Value:=cds_participanteobservacao.Value;
    cds_participante2codrecinc.Value:=cds_participantecodrecinc.Value;
    cds_participante2codrecBx.Value:=cds_participantecodrecBx.Value;
    cds_participante2codpaginc.Value:=cds_participantecodpaginc.Value;
    cds_participante2codpagBx.Value:=cds_participantecodpagBx.Value;
    cds_participante2fone4.Value:=cds_participantefone4.Value;
    cds_participante2complemento.Value:=cds_participantecomplemento.Value;
    cds_participante2complemento_cob.Value:=cds_participantecomplemento_cob.Value;
    if cds_participantecadastrado.Value>0 then
       cds_participante2cadastrado.Value:=cds_participantecadastrado.Value;
    cds_participante2Nome_Firma.Value:=cds_participanteNome_Firma.Value;
    cds_participante2Inativo.Value:=cds_participanteInativo.Value;
    cds_participante2Sexo.Value:=cds_participanteSexo.Value;
    cds_participante2cc.Value:=cds_participantecc.Value;
    cds_participante2cex.Value:=cds_participantecex.Value;
    cds_participante2jb.Value:=cds_participantejb.Value;
    cds_participante2judicial.Value:=cds_participantejudicial.Value;
    cds_participante2xquadra.Value:=cds_participantexquadra.Value;
    cds_participante2xlote.Value:=cds_participantexlote.Value;
    cds_participante2plantao.Value:=cds_participanteplantao.Value;
    cds_participante2.post;

    cds_participante.Next;

  end;



  cds_conjuge.close;
  cds_conjuge.LoadFromFile(path+'conjuge.xml');
  cds_conjuge.open;
  ZQConjuge.close;

  cds_conjuge2.close;
  cds_conjuge2.CreateDataSet;
  cds_conjuge2.open;

  Gauge2.Progress:=0;
  Gauge2.MaxValue:=cds_participante.RecordCount;
  cds_participante.DisableControls;
  cds_participante.First;
  while not cds_participante.Eof do
  begin
    application.ProcessMessages;
    Label85.caption:='Verificando Participantes...';
    Gauge2.Progress:=cds_participante.RecNo;
    DM_Tabelas.ZQparticipante_proposta.close;
    DM_Tabelas.ZQparticipante_proposta.SQL.Clear;
    DM_Tabelas.ZQparticipante_proposta.SQL.Add('Select * from Participante_proposta where nome_parte='+quotedstr(cds_participantenome_parte.Text));
//    DM_Tabelas.ZQparticipante_proposta.SQL.Add('Select * from Participante_proposta where doc1='+quotedstr(cds_participantedoc1.Value)+' and  nome_parte='+quotedstr(cds_participantenome_parte.Text));
    DM_Tabelas.ZQparticipante_proposta.open;
    if DM_Tabelas.ZQparticipante_proposta.RecordCount=0 then
    begin
      DM_Tabelas.ZQparticipante_proposta.Insert;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('nome_parte').AsString:=cds_participantenome_parte.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('doc1').AsString:=cds_participantedoc1.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('doc2').AsString:=cds_participantedoc2.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('endereco').AsString:=cds_participanteendereco.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('bairro').AsString:=cds_participantebairro.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('cidade').AsInteger:=cds_participantecidade.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('cep').AsString:=cds_participantecep.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('ende_cob').AsString:=cds_participanteende_cob.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('bairro_cob').AsString:=cds_participantebairro_cob.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('cidade_cob').AsLargeInt:=cds_participantecidade_cob.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('cep_cob').AsString:=cds_participantecep_cob.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('Fone1').AsString:=cds_participanteFone1.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('fone2').AsString:=cds_participantefone2.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('fone3').AsString:=cds_participantefone3.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('tipopessoa').AsString:=cds_participantetipopessoa.Value;
      if cds_participanteaniversario.Value>0 then
         DM_Tabelas.ZqParticipante_proposta.FieldByName('aniversario').AsDateTime:=cds_participanteaniversario.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('nacionalidade').AsString:=cds_participantenacionalidade.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('email').AsString:=cds_participanteemail.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('naturalidade').AsInteger:=cds_participantenaturalidade.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('localdetrab').AsString:=cds_participantelocaldetrab.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('profissao').AsString:=cds_participanteprofissao.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('ESTADOCIVIL').AsString:=cds_participanteestadocivil.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('renda').AsFloat:=cds_participanterenda.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('observacao').AsString:=cds_participanteobservacao.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('codrecinc').AsInteger:=cds_participantecodrecinc.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('codrecBx').AsInteger:=cds_participantecodrecBx.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('codpaginc').AsInteger:=cds_participantecodpaginc.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('codpagBx').AsInteger:=cds_participantecodpagBx.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('fone4').AsString:=cds_participantefone4.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('complemento').AsString:=cds_participantecomplemento.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('complemento_cob').AsString:=cds_participantecomplemento_cob.Value;
      if cds_participantecadastrado.Value>0 then
         DM_Tabelas.ZqParticipante_proposta.FieldByName('cadastrado').AsDateTime:=cds_participantecadastrado.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('Nome_Firma').AsString:=cds_participanteNome_Firma.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('Inativo').AsString:=cds_participanteInativo.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('sexo').AsString:=cds_participanteSexo.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('cc').AsDateTime:=cds_participantecc.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('cex').AsDateTime:=cds_participantecex.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('jb').AsDateTime:=cds_participantejb.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('judicial').AsString:=cds_participantejudicial.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('xquadra').AsString:=cds_participantexquadra.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('xlote').AsString:=cds_participantexlote.Value;
      DM_Tabelas.ZqParticipante_proposta.FieldByName('plantao').AsString:=cds_participanteplantao.Value;

      DM_Tabelas.ZQparticipante_proposta.post;
      xcodpart:=DM_Tabelas.ZqParticipante_proposta.FieldByName('idpaticipante').AsLargeInt;


      cds_conjuge.Filtered:=false;
      cds_conjuge.Filter:='idparti='+quotedstr(cds_participanteidpaticipante.Text);
      cds_conjuge.Filtered:=true;


      cds_conjuge.First;
      if cds_conjuge.RecordCount>0 then
      begin
        application.ProcessMessages;
        cds_conjuge2.Insert;
        cds_conjuge2nome.Value:=cds_conjugenome.Value;
        cds_conjuge2doc1.Value:=cds_Conjugedoc1.Value;
        cds_conjuge2doc2.Value:=cds_Conjugedoc2.Value;
        if cds_Conjugenasc.Value>0 then
           cds_conjuge2nasc.Value:=cds_Conjugenasc.Value;
        cds_conjuge2natural.Value:=cds_Conjugenatural.Value;
        cds_conjuge2idparti.Value:=xcodpart;//cds_Conjugeidparti.Value;
        cds_conjuge2email.Value:=cds_Conjugeemail.Value;
        cds_conjuge2assina.Value:=cds_Conjugeassina.Value;
        cds_conjuge2cargo.Value:=cds_Conjugecargo.Value;
        if cds_Conjugeadmissa.Value>0 then
           cds_conjuge2admissa.Value:=cds_Conjugeadmissa.Value;
        cds_conjuge2renda.Value:=cds_Conjugerenda.Value;
        cds_conjuge2regime.Value:=cds_Conjugeregime.Value;
        cds_conjuge2comarca.Value:=cds_Conjugecomarca.Value;
        cds_conjuge2cidade_cart.Value:=cds_Conjugecidade_cart.Value;
        cds_conjuge2folha.Value:=cds_Conjugefolha.Value;
        cds_conjuge2livro.Value:=cds_Conjugelivro.Value;
        cds_conjuge2cartorio.Value:=cds_Conjugecartorio.Value;
        if cds_Conjugedata.Value>0 then
           cds_conjuge2data.Value:=cds_Conjugedata.Value;
        cds_conjuge2.post;
      end;
      cds_conjuge.Filtered:=false;
    end;
    cds_participante.Next;
  end;



  Gauge2.Progress:=0;
  Gauge2.MaxValue:=cds_conjuge2.RecordCount;
  cds_conjuge2.DisableControls;
  cds_conjuge2.First;
  if cds_conjuge2.RecordCount>0 then
  begin
    application.ProcessMessages;
    Label85.caption:='Verificando Conjuges...';

    application.ProcessMessages;
    DM_Tabelas.ZQConjuge_proposta.Open;
    while not cds_conjuge2.Eof do
    begin
        Gauge2.Progress:=cds_conjuge2.RecNo;
        application.ProcessMessages;        
        DM_Tabelas.ZQConjuge_proposta.Insert;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('nome').AsString:=cds_conjuge2nome.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('doc1').AsString:=cds_Conjuge2doc1.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('doc2').AsString:=cds_Conjuge2doc2.Value;
        if cds_Conjuge2nasc.Value>0 then
           DM_Tabelas.ZQConjuge_proposta.FieldByName('nasc').AsDateTime:=cds_Conjuge2nasc.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('natural').AsInteger:=cds_Conjuge2natural.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('idparti').AsInteger:=cds_Conjuge2idparti.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('email').AsString:=cds_Conjuge2email.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('assina').AsString:=cds_Conjuge2assina.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('cargo').AsString:=cds_Conjuge2cargo.Value;
        if cds_Conjuge2admissa.Value>0 then
           DM_Tabelas.ZQConjuge_proposta.FieldByName('admissa').AsDateTime:=cds_Conjuge2admissa.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('renda').AsFloat:=cds_Conjuge2renda.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('regime').AsString:=cds_Conjuge2regime.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('comarca').AsFloat:=cds_Conjuge2comarca.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('cidade_cart').AsFloat:=cds_Conjuge2cidade_cart.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('folha').AsString:=cds_Conjuge2folha.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('livro').AsString:=cds_Conjuge2livro.Value;
        DM_Tabelas.ZQConjuge_proposta.FieldByName('cartorio').AsString:=cds_Conjuge2cartorio.Value;
        if cds_Conjuge2data.Value>0 then
           DM_Tabelas.ZQConjuge_proposta.FieldByName('data').AsDateTime:=cds_Conjuge2data.Value;
        DM_Tabelas.ZQConjuge_proposta.post;
        cds_conjuge2.Next;        
    end;
  end;
  Gauge2.Progress:=0;

  cds_participante.EnableControls;
  DM_Tabelas.ZQparticipante_proposta.close;
  DM_Tabelas.ZQparticipante_proposta.SQL.Clear;
  DM_Tabelas.ZQparticipante_proposta.SQL.Add('Select * from Participante_proposta ');
  DM_Tabelas.ZQparticipante_proposta.open;


  application.ProcessMessages;
  Gauge2.Progress:=0;
  Gauge2.MaxValue:=CdsImovel.RecordCount;
  CdsImovel.DisableControls;
  CdsImovel.First;
  if CdsImovel.RecordCount>0 then
  begin
    while not CdsImovel.Eof do
    begin
      application.ProcessMessages;
      Label85.caption:='Verificando Imóveis...                ';
      Gauge2.Progress:=CdsImovel.RecNo;
      DM_Tabelas.ZQimovel.close;
      DM_Tabelas.ZQimovel.SQL.Clear;
      DM_Tabelas.ZQimovel.SQL.Add('Select  loteamento_idloteamento,quadra,lote,lado,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,esquinanome,valorvenal,Obervacaoloteamento,');
      DM_Tabelas.ZQimovel.SQL.Add('        cornomapa,idimovel,disponivel,medidafrente,medidafundo,medidadedireita,esquinalinha,esquinacurva,medidaesquerda,matri,matricula,proposta,valorCusto ');
      DM_Tabelas.ZQimovel.SQL.Add(' from imovel where idimovel='+quotedstr(CdsImovelidimovel.Text));
      DM_Tabelas.ZQimovel.open;
      if DM_Tabelas.ZQimovel.RecordCount>0 then
      begin
        if (DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString='SIM') then
        begin
          DM_Tabelas.ZQimovel.Edit;
          DM_Tabelas.ZQImovel.FieldByName('proposta').AsString:=CdsImovelproposta.value;
          DM_Tabelas.ZQimovel.post;
        end;
      end;
      CdsImovel.Next;
    end
  end;
  CdsImovel.EnableControls;
  DM_Tabelas.ZQimovel.close;
  DM_Tabelas.ZQimovel.SQL.Clear;
  DM_Tabelas.ZQimovel.SQL.Add('Select  loteamento_idloteamento,quadra,lote,lado,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,esquinanome,valorvenal,Obervacaoloteamento,');
  DM_Tabelas.ZQimovel.SQL.Add('        cornomapa,idimovel,disponivel,medidafrente,medidafundo,medidadedireita,esquinalinha,esquinacurva,medidaesquerda,matri,matricula,proposta,valorCusto ');
  DM_Tabelas.ZQimovel.SQL.Add(' from imovel ');
  DM_Tabelas.ZQimovel.open;


  Gauge2.Progress:=0;
  Gauge2.MaxValue:=cds_proposta.RecordCount;
  cds_proposta.DisableControls;
  cds_proposta.First;


  DM_Tabelas.ZQComprador_proposta.close;
  DM_Tabelas.ZQComprador_proposta.MasterSource:=nil;
  DM_Tabelas.ZQComprador_proposta.MasterFields:='';
  DM_Tabelas.ZQComprador_proposta.LinkedFields:='';
  DM_Tabelas.ZQComprador_proposta.open;

  DM_Tabelas.ZQVendedor_proposta.close;
  DM_Tabelas.ZQVendedor_proposta.MasterSource:=nil;
  DM_Tabelas.ZQVendedor_proposta.MasterFields:='';
  DM_Tabelas.ZQVendedor_proposta.LinkedFields:='';
  DM_Tabelas.ZQVendedor_proposta.open;

  DM_Tabelas.ZQCheque_proposta.close;
  DM_Tabelas.ZQCheque_proposta.MasterSource:=nil;
  DM_Tabelas.ZQCheque_proposta.MasterFields:='';
  DM_Tabelas.ZQCheque_proposta.LinkedFields:='';
  DM_Tabelas.ZQCheque_proposta.open;

  cds_participante2.IndexName:='cds_participante2Index1';
  cds_participante2.MasterSource:=ds_cds_comprador2;
  cds_participante2.MasterFields:='paticipante_idpaticipante';
  

  DM_Tabelas.ZQProposta.open;
  while not cds_proposta.Eof do
  begin
    Gauge2.Progress                           :=cds_proposta.RecNo;
    application.ProcessMessages;
    Label85.caption                           :='Verificando Propostas...                  ';
    application.ProcessMessages;
    DM_Tabelas.ZQProposta.Insert;
    DM_Tabelas.ZQProposta.FieldByName('dataproposta').AsDateTime   :=cds_propostadataproposta.Value;
    DM_Tabelas.ZQProposta.FieldByName('imovel').AsLargeInt         :=cds_propostaimovel.Value;
    DM_Tabelas.ZQProposta.FieldByName('valorproposta').AsFloat  :=cds_propostavalorproposta.Value;
    DM_Tabelas.ZQProposta.FieldByName('forma_reajuste').AsString :=cds_propostaforma_reajuste.Value;
    DM_Tabelas.ZQProposta.FieldByName('tabela_Price').AsFloat   :=cds_propostatabela_Price.Value;
    DM_Tabelas.ZQProposta.FieldByName('Escriturado').AsString    :=cds_propostaEscriturado.Value;
    DM_Tabelas.ZQProposta.post;


    cds_comprador2.IndexName:='idpropostaIndex1';
    cds_comprador2.First;
    while not cds_comprador2.Eof do
    begin
      application.ProcessMessages;
      Label85.caption:='Verificando Compradores...                  ';
      application.ProcessMessages;
      if cds_comprador2idproposta.Value=cds_propostaidproposta.Value then
      begin
        DM_Tabelas.ZQComprador_proposta.insert;
        DM_Tabelas.ZQComprador_proposta.FieldByName('venda_idvenda').AsLargeInt:=DM_Tabelas.ZQProposta.FieldByName('idproposta').AsLargeInt;

        DM_Tabelas.ZQparticipante_proposta.close;
        DM_Tabelas.ZQparticipante_proposta.sql.clear;
        DM_Tabelas.ZQparticipante_proposta.sql.add('select * from participante_proposta where nome_parte='+quotedstr(cds_participante2nome_parte.value));
        DM_Tabelas.ZQparticipante_proposta.open;

        DM_Tabelas.ZQComprador_proposta.FieldByName('paticipante_idpaticipante').AsLargeInt:=DM_Tabelas.ZqParticipante_proposta.FieldByName('idpaticipante').AsLargeInt;
        DM_Tabelas.ZQComprador_proposta.FieldByName('promissario').AsString:=cds_comprador2promissario.Value;
        DM_Tabelas.ZQComprador_proposta.FieldByName('percentual').AsFloat:=cds_comprador2percentual.Value;
        DM_Tabelas.ZQComprador_proposta.FieldByName('Marcar').AsInteger:=cds_comprador2Marcar.Value;
//        DM_Tabelas.ZQComprador_propostanome_participante.value:=DM_Tabelas.ZQparticipante_propostanome_parte.Value;
        DM_Tabelas.ZQComprador_proposta.FieldByName('idproposta').AsLargeInt:=DM_Tabelas.ZQProposta.FieldByName('idproposta').AsLargeInt;
        DM_Tabelas.ZQComprador_proposta.post;
      end;
      cds_comprador2.Next;
    end;
    DM_Tabelas.ZQparticipante_proposta.close;
    DM_Tabelas.ZQparticipante_proposta.sql.clear;
    DM_Tabelas.ZQparticipante_proposta.sql.add('select * from participante_proposta ');
    DM_Tabelas.ZQparticipante_proposta.open;

//    cds_comprador2.Filtered:=false;
    cds_recebimento.MasterSource:=ds_cds_proposta;
    cds_recebimento.MasterFields:='idproposta';
    //cds_recebimento.IndexFieldNames:='venda_idvenda';

    DM_Tabelas.ZQRecebimento_proposta.close;
    DM_Tabelas.ZQRecebimento_proposta.SQL.clear;
    //DM_Tabelas.ZQRecebimento_proposta.SQL.Add('Select *  from Recebimento_proposta where venda_idvenda='+quotedstr(cds_propostaidproposta.Text));
    DM_Tabelas.ZQRecebimento_proposta.SQL.Add('Select * from Recebimento_proposta ');
    DM_Tabelas.ZQRecebimento_proposta.Open;
    cds_recebimento.First;
    while not cds_recebimento.Eof do
    begin
      application.ProcessMessages;
      Label85.caption:='Verificando Parcelas...                  ';


      DM_Tabelas.ZQRecebimento_proposta.insert;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('documento').AsString:=cds_Recebimentodocumento.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('cliente').AsInteger:=cds_Recebimentocliente.Value;
      // DM_Tabelas.ZQRecebimento_propostausuario.value:=cds_Recebimentousuario.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('Dt_Entrada').AsDateTime:=cds_RecebimentoDt_Entrada.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('Dt_Vencimento').AsDateTime:=cds_RecebimentoDt_Vencimento.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('Valor').AsFloat:=cds_RecebimentoValor.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('Observ').AsString:=cds_RecebimentoObserv.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('VrDoc').AsFloat:=cds_RecebimentoVrDoc.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('ordem').AsString:=cds_Recebimentoordem.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('TipDoc').AsString:=cds_RecebimentoTipDoc.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('saldo').AsFloat:=cds_Recebimentosaldo.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('marcar').AsString:=cds_Recebimentomarcar.Value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('RefBaixa').AsInteger:=cds_RecebimentoRefBaixa.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('refvinda').AsInteger:=cds_Recebimentorefvinda.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('contabil').AsInteger:=cds_Recebimentocontabil.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('empresa').AsInteger:=cds_Recebimentoempresa.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('custodaparcela').AsFloat:=cds_Recebimentocustodaparcela.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('origem').AsString:=cds_Recebimentoorigem.value;

      DM_Tabelas.ZQparticipante_proposta.close;
      DM_Tabelas.ZQparticipante_proposta.SQL.Clear;
      DM_Tabelas.ZQparticipante_proposta.SQL.Add('select * from Participante_proposta where nome_parte='+quotedstr(cds_recebimentonomeadversa.Value));
      DM_Tabelas.ZQparticipante_proposta.open;


      DM_Tabelas.ZQRecebimento_proposta.FieldByName('nomeadversa').AsString:=DM_Tabelas.ZqParticipante_proposta.FieldByName('nome_parte').AsString;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('adversa').AsInteger:=DM_Tabelas.ZqParticipante_proposta.FieldByName('idpaticipante').AsLargeInt;


      DM_Tabelas.ZQRecebimento_proposta.FieldByName('recpag').AsString:=cds_Recebimentorecpag.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('numordem').AsInteger:=cds_Recebimentonumordem.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('idloteamento').AsInteger:=cds_Recebimentoidloteamento.value;

      DM_Tabelas.ZQRecebimento_proposta.FieldByName('quadralote').AsString:=cds_Recebimentoquadralote.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('numboleto').AsString:=cds_Recebimentonumboleto.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('Substituicao').AsString:=cds_RecebimentoSubstituicao.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('sq').AsLargeInt:=cds_Recebimentosq.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('somar').AsString:=cds_Recebimentosomar.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('Reajustado').AsString:=cds_RecebimentoReajustado.value;
      if cds_RecebimentoData_reajuste.value>0 then
         DM_Tabelas.ZQRecebimento_proposta.FieldByName('Data_reajuste').AsDateTime:=cds_RecebimentoData_reajuste.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('Proximo_Reajuste').AsString:=cds_RecebimentoProximo_Reajuste.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('Parcelas_fixas').AsString:=cds_RecebimentoParcelas_fixas.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('juros').AsFloat:=cds_Recebimentojuros.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('descontos').AsFloat:=cds_Recebimentodescontos.value;
      if cds_RecebimentoData_Quitacao.value>0 then
         DM_Tabelas.ZQRecebimento_proposta.FieldByName('Data_Quitacao').AsDateTime:=cds_RecebimentoData_Quitacao.value;
      DM_Tabelas.ZQRecebimento_proposta.FieldByName('venda_idvenda').AsInteger:=DM_Tabelas.ZQProposta.FieldByName('idproposta').AsLargeInt;
      DM_Tabelas.ZQRecebimento_proposta.post;

      cds_cheque.Filtered:=false;
      cds_cheque.Filter:='idvinculo='+quotedstr(cds_Recebimentoidrecebimento_proposta.Text);
      cds_cheque.Filtered:=true;
      if cds_cheque.RecordCount>0 then
      begin
        cds_cheque.First;
        while not cds_cheque.Eof do
        begin
          application.ProcessMessages;
          Label85.caption:='Verificando Cheques...                  ';
          DM_Tabelas.ZQCheque_proposta.insert;
          DM_Tabelas.ZQCheque_proposta.FieldByName('idvinculo').AsInteger:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('idrecebimento_proposta').AsLargeInt;
          DM_Tabelas.ZQCheque_proposta.FieldByName('idrecebimento').AsInteger:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('idrecebimento_proposta').AsLargeInt;
          DM_Tabelas.ZQCheque_proposta.FieldByName('banco').AsString:=cds_chequebanco.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('agencia').AsString:=cds_chequeagencia.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('dono').AsString:=cds_chequedono.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('CPF_CNPJ').AsString:=cds_chequeCPF_CNPJ.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('valor').AsFloat:=cds_chequevalor.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('emissao').AsDateTime:=cds_chequeemissao.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('deposito').AsDateTime:=cds_chequedeposito.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('alias').AsString:=cds_chequealias.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('prorrogado').AsDateTime:=cds_chequeprorrogado.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('conta').AsString:=cds_chequeconta.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('numero').AsString:=cds_chequenumero.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('Substituicao').AsString:=cds_chequeSubstituicao.Value;
          DM_Tabelas.ZQCheque_proposta.FieldByName('sq').AsLargeInt:=cds_chequesq.Value;

          DM_Tabelas.ZQCheque_proposta.post;
          cds_cheque.Next;
        end;
      end;
      cds_recebimento.Next;
    end;
    cds_vendedor.First;
    while not cds_vendedor.Eof do
    begin
      application.ProcessMessages;
      Label85.caption:='Verificando Vendedores...                  ';
      if cds_vendedoridproposta.Value=cds_propostaidproposta.Value  then
      begin
        DM_Tabelas.ZQVendedor_proposta.insert;
        DM_Tabelas.ZQVendedor_proposta.FieldByName('venda_idvenda').AsLargeInt:=DM_Tabelas.ZQProposta.FieldByName('idproposta').AsLargeInt;
        DM_Tabelas.ZQVendedor_proposta.FieldByName('corretor_idcorretor').AsLargeInt:=cds_vendedorcorretor_idcorretor.Value;
        DM_Tabelas.ZQVendedor_proposta.FieldByName('idproposta').AsLargeInt:=DM_Tabelas.ZQProposta.FieldByName('idproposta').AsLargeInt;
        DM_Tabelas.ZQVendedor_proposta.post;
      end;
      cds_vendedor.Next;
    end;
    cds_proposta.Next;
    cds_recebimento.MasterSource:=nil;
    cds_recebimento.MasterFields:='';
//    cds_recebimento.IndexFieldNames:='';
  end;
  Gauge2.Progress:=0;
  cds_proposta.EnableControls;

  DM_Tabelas.ZQparticipante_proposta.close;
  DM_Tabelas.ZQparticipante_proposta.SQL.Clear;
  DM_Tabelas.ZQparticipante_proposta.SQL.Add('select * from Participante_proposta');
  DM_Tabelas.ZQparticipante_proposta.open;

  DM_Tabelas.ZQComprador_proposta.close;
  DM_Tabelas.ZQComprador_proposta.MasterSource:=DM_Tabelas.DS_proposta;
  DM_Tabelas.ZQComprador_proposta.MasterFields:='idproposta';
  DM_Tabelas.ZQComprador_proposta.LinkedFields:='venda_idvenda';
  DM_Tabelas.ZQComprador_proposta.SQL.clear;
  DM_Tabelas.ZQComprador_proposta.SQL.Add('Select *  from comprador_proposta ');
  DM_Tabelas.ZQComprador_proposta.Open;

  DM_Tabelas.ZQVendedor_proposta.close;
  DM_Tabelas.ZQVendedor_proposta.MasterSource:=DM_Tabelas.DS_proposta;
  DM_Tabelas.ZQVendedor_proposta.MasterFields:='idproposta';
  DM_Tabelas.ZQVendedor_proposta.LinkedFields:='venda_idvenda';
  DM_Tabelas.ZQVendedor_proposta.SQL.clear;
  DM_Tabelas.ZQVendedor_proposta.SQL.Add('Select *  from vendedor_proposta ');
  DM_Tabelas.ZQVendedor_proposta.Open;


  DM_Tabelas.ZQCheque_proposta.close;
  DM_Tabelas.ZQCheque_proposta.SQL.clear;
  DM_Tabelas.ZQCheque_proposta.SQL.Add('select * from cheque_rec_proposta ');
  DM_Tabelas.ZQCheque_proposta.Open;

  DM_Tabelas.ZQCheque_proposta.MasterSource:=DM_Tabelas.DS_ZQRecebimento_proposta;
  DM_Tabelas.ZQCheque_proposta.MasterFields:='idrecebimento_proposta';
  DM_Tabelas.ZQCheque_proposta.LinkedFields:='idvinculo';

  DM_Tabelas.ZQRecebimento_proposta.close;
  DM_Tabelas.ZQRecebimento_proposta.SQL.Clear;
  DM_Tabelas.ZQRecebimento_proposta.SQL.Add('Select * from Recebimento_proposta order by DT_Vencimento');
  DM_Tabelas.ZQRecebimento_proposta.Open;
  Label85.caption:='                                        ';
  Pag_Venda.PageIndex:=4;

  cds_participante2.MasterSource:=nil;
  cds_participante2.MasterFields:='';
  cds_participante2.IndexName:='';

  cds_participante2.close;

  cds_comprador2.Close;

  mensagem('Arquivos Importados!');
  JvXPBar2.Visible:=false;
//  cds_participante2.Filtered:=false;
end;

procedure TFrm_Venda.FormCreate(Sender: TObject);
begin
  EnsureRuntimeFields(Self);
  // para rolar registro no dbgrid com o rolon do mouse
  Application.OnMessage := ApplicationEvents1Message;
end;

procedure TFrm_Venda.JvXPBar2Items1Click(Sender: TObject);
var
caminho,path,comando:string;
begin
  //gerar backup
  JvXPBar2.Visible:=false;
  SaveDialog1.FileName:='plantao.sql';
  SaveDialog1.Execute;
  path:= SaveDialog1.FileName; //(ExtractFilePath(Application.ExeName));
  path:=ExtractFilePath(SaveDialog1.FileName)+'plantao.sql';
  comando:=ExtractFilePath(Application.ExeName);
  comando:=comando+'Exportar.bat '+DM_Tabelas.zconeccao.Database+' '+DM_Tabelas.zconeccao.User+' '+DM_Tabelas.zconeccao.HostName+' '+DM_Tabelas.zconeccao.Password+' '+path;
//  winexec(Pchar(comando),SW_Hide);

  caminho:=ExtractFilePath(Application.ExeName);
                                    //SW_ShowNormal                                                                                                                                                                     //SW_ShowNormal
  if WinExecAndWait32(comando,caminho,SW_Hide) = 0 then
     mensagem('Exportação Foi Realizada!')
  else
     mensagem('Exportação Não Foi Realizada!');
end;

procedure TFrm_Venda.dxButton4Click(Sender: TObject);
var
xidloteamento,varnumordemE,varnumordemp,posi,posi2,xcodpart:integer;
jaE,jaP,xnome:string;
begin

  DM_Tabelas.ZQCheque_proposta.close;
  DM_Tabelas.ZQCheque_proposta.SQL.clear;
  DM_Tabelas.ZQCheque_proposta.SQL.Add('select * from cheque_rec_proposta ');
  DM_Tabelas.ZQCheque_proposta.Open;

  DM_Tabelas.ZQCheque_proposta.MasterSource:=DM_Tabelas.DS_ZQRecebimento_proposta;
  DM_Tabelas.ZQCheque_proposta.MasterFields:='idrecebimento_proposta';
  DM_Tabelas.ZQCheque_proposta.LinkedFields:='idvinculo';

  jaE:='n';
  jaP:='n';
  xnome:='';
  cds_comprador2.IndexName:='';
  cds_comprador2.Close;
  cds_comprador2.CreateDataSet;
  cds_comprador2.Open;

  DM_Tabelas.ZQComprador_proposta.first;
  while not DM_Tabelas.ZQComprador_proposta.Eof do
  begin
    cds_comprador2.Insert;
    cds_comprador2idcomprador.Value:=DM_Tabelas.ZQComprador_proposta.FieldByName('idcomprador').AsLargeInt;
    cds_comprador2promissario.Value:=DM_Tabelas.ZQComprador_proposta.FieldByName('promissario').AsString;
    cds_comprador2percentual.Value:=DM_Tabelas.ZQComprador_proposta.FieldByName('percentual').AsFloat;
    cds_comprador2Marcar.Value:=DM_Tabelas.ZQComprador_proposta.FieldByName('Marcar').AsInteger;
    cds_comprador2venda_idvenda.Value:=DM_Tabelas.ZQComprador_proposta.FieldByName('venda_idvenda').AsLargeInt;
    cds_comprador2idproposta.Value:=DM_Tabelas.ZQComprador_proposta.FieldByName('idproposta').AsLargeInt;
    cds_comprador2paticipante_idpaticipante.Value:=DM_Tabelas.ZQComprador_proposta.FieldByName('paticipante_idpaticipante').AsLargeInt;
    cds_comprador2.Post;
    DM_Tabelas.ZQComprador_proposta.Next;    
  end;
  cds_participante2.MasterSource:=nil;
  cds_participante2.MasterFields:='';
  cds_participante2.IndexName:='';

  cds_participante2.close;
  cds_participante2.CreateDataSet;
  cds_participante2.open;

  cds_comprador2.first;
  while not cds_comprador2.eof do
  begin
    DM_Tabelas.ZQparticipante_proposta.close;
    DM_Tabelas.ZQparticipante_proposta.sql.clear;
    DM_Tabelas.ZQparticipante_proposta.sql.add('select * from participante_proposta where idpaticipante='+quotedstr(cds_comprador2paticipante_idpaticipante.text));
    DM_Tabelas.ZQparticipante_proposta.open;


    cds_participante2.Insert;
    cds_participante2idpaticipante.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('idpaticipante').AsLargeInt;
    cds_participante2nome_parte.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('nome_parte').AsString;
    cds_participante2doc1.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('doc1').AsString;
    cds_participante2doc2.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('doc2').AsString;
    cds_participante2endereco.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('endereco').AsString;
    cds_participante2bairro.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('bairro').AsString;
    cds_participante2cidade.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('cidade').AsInteger;
    cds_participante2cep.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('cep').AsString;
    cds_participante2ende_cob.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('ende_cob').AsString;
    cds_participante2bairro_cob.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('bairro_cob').AsString;
    cds_participante2cidade_cob.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('cidade_cob').AsLargeInt;
    cds_participante2cep_cob.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('cep_cob').AsString;
    cds_participante2Fone1.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('Fone1').AsString;
    cds_participante2fone2.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('fone2').AsString;
    cds_participante2fone3.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('fone3').AsString;
    cds_participante2tipopessoa.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('tipopessoa').AsString;
    if DM_Tabelas.ZqParticipante_proposta.FieldByName('aniversario').AsDateTime>0 then
       cds_participante2aniversario.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('aniversario').AsDateTime;
    cds_participante2nacionalidade.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('nacionalidade').AsString;
    cds_participante2email.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('email').AsString;
    cds_participante2naturalidade.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('naturalidade').AsInteger;
    cds_participante2localdetrab.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('localdetrab').AsString;
    cds_participante2profissao.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('profissao').AsString;
    cds_participante2estadocivil.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('ESTADOCIVIL').AsString;
    cds_participante2renda.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('renda').AsFloat;
    cds_participante2observacao.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('observacao').AsString;
    cds_participante2codrecinc.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('codrecinc').AsInteger;
    cds_participante2codrecBx.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('codrecBx').AsInteger;
    cds_participante2codpaginc.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('codpaginc').AsInteger;
    cds_participante2codpagBx.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('codpagBx').AsInteger;
    cds_participante2fone4.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('fone4').AsString;
    cds_participante2complemento.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('complemento').AsString;
    cds_participante2complemento_cob.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('complemento_cob').AsString;
    if DM_Tabelas.ZqParticipante_proposta.FieldByName('cadastrado').AsDateTime>0 then
       cds_participante2cadastrado.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('cadastrado').AsDateTime;
    cds_participante2Nome_Firma.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('Nome_Firma').AsString;
    cds_participante2Inativo.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('Inativo').AsString;
    cds_participante2Sexo.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('sexo').AsString;
    cds_participante2cc.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('cc').AsDateTime;
    cds_participante2cex.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('cex').AsDateTime;
    cds_participante2jb.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('jb').AsDateTime;
    cds_participante2judicial.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('judicial').AsString;
    cds_participante2xquadra.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('xquadra').AsString;
    cds_participante2xlote.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('xlote').AsString;
    cds_participante2plantao.Value:=DM_Tabelas.ZqParticipante_proposta.FieldByName('plantao').AsString;
    cds_participante2.post;
    cds_comprador2.next;
  end;  

  DM_Tabelas.ZQparticipante_proposta.close;
  DM_Tabelas.ZQparticipante_proposta.sql.clear;
  DM_Tabelas.ZQparticipante_proposta.sql.add('select * from participante_proposta ');
  DM_Tabelas.ZQparticipante_proposta.open;

{  cds_participante2.IndexName:='cds_participante2Index1';
  cds_participante2.MasterSource:=ds_cds_comprador2;
  cds_participante2.MasterFields:='paticipante_idpaticipante';}


  cds_comprador2.IndexName:='idpropostaIndex1';
  cds_comprador2.First;

  cds_conjuge2.close;
  cds_conjuge2.CreateDataSet;
  cds_conjuge2.open;

  Gauge2.Progress:=0;
  Gauge2.MaxValue:=cds_participante2.RecordCount;
  cds_participante2.DisableControls;
  cds_participante2.First;
  while not cds_participante2.Eof do
  begin
    application.ProcessMessages;
    Label85.caption:='Verificando Participantes...';
    application.ProcessMessages;
    Gauge2.Progress:=cds_participante2.RecNo;
    DM_Tabelas.ZQparticipante.close;
    DM_Tabelas.ZQparticipante.SQL.Clear;
//    DM_Tabelas.ZQparticipante.SQL.Add('Select * from Participante where doc1='+quotedstr(cds_participante2doc1.Value)+' and nome_parte='+quotedstr(cds_participante2nome_parte.Value));
    DM_Tabelas.ZQparticipante.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
    DM_Tabelas.ZQparticipante.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
    DM_Tabelas.ZQparticipante.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico ');
    DM_Tabelas.ZQparticipante.SQL.Add('  from Participante where nome_parte='+quotedstr(cds_participante2nome_parte.Value));
    DM_Tabelas.ZQparticipante.open;
    if DM_Tabelas.ZQparticipante.RecordCount=0 then
    begin
      DM_Tabelas.ZQparticipante.Insert;
      DM_Tabelas.ZQParticipante.FieldByName('nome_parte').AsString:=cds_participante2nome_parte.Value;
      DM_Tabelas.ZqParticipante.FieldByName('doc1').AsString:=cds_participante2doc1.Value;
      DM_Tabelas.ZqParticipante.FieldByName('doc2').AsString:=cds_participante2doc2.Value;
      DM_Tabelas.ZqParticipante.FieldByName('endereco').AsString:=cds_participante2endereco.Value;
      DM_Tabelas.ZqParticipante.FieldByName('bairro').AsString:=cds_participante2bairro.Value;
      DM_Tabelas.ZqParticipante.FieldByName('cidade').AsInteger:=cds_participante2cidade.Value;
      DM_Tabelas.ZqParticipante.FieldByName('cep').AsString:=cds_participante2cep.Value;
      DM_Tabelas.ZqParticipante.FieldByName('ende_cob').AsString:=cds_participante2ende_cob.Value;
      DM_Tabelas.ZqParticipante.FieldByName('bairro_cob').AsString:=cds_participante2bairro_cob.Value;
      DM_Tabelas.ZqParticipante.FieldByName('cidade_cob').AsLargeInt:=cds_participante2cidade_cob.Value;
      DM_Tabelas.ZqParticipante.FieldByName('cep_cob').AsString:=cds_participante2cep_cob.Value;
      DM_Tabelas.ZqParticipante.FieldByName('Fone1').AsString:=cds_participante2Fone1.Value;
      DM_Tabelas.ZqParticipante.FieldByName('fone2').AsString:=cds_participante2fone2.Value;
      DM_Tabelas.ZqParticipante.FieldByName('fone3').AsString:=cds_participante2fone3.Value;
      DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString:=cds_participante2tipopessoa.Value;
      if cds_participante2aniversario.Value>0 then
         DM_Tabelas.ZqParticipante.FieldByName('aniversario').AsDateTime:=cds_participante2aniversario.Value;
      DM_Tabelas.ZqParticipante.FieldByName('nacionalidade').AsString:=cds_participante2nacionalidade.Value;
      DM_Tabelas.ZqParticipante.FieldByName('email').AsString:=cds_participante2email.Value;
      DM_Tabelas.ZqParticipante.FieldByName('naturalidade').AsInteger:=cds_participante2naturalidade.Value;
      DM_Tabelas.ZqParticipante.FieldByName('localdetrab').AsString:=cds_participante2localdetrab.Value;
      DM_Tabelas.ZqParticipante.FieldByName('profissao').AsString:=cds_participante2profissao.Value;
      DM_Tabelas.ZqParticipante.FieldByName('ESTADOCIVIL').AsString:=cds_participante2estadocivil.Value;
      DM_Tabelas.ZqParticipante.FieldByName('renda').AsFloat:=cds_participante2renda.Value;
      DM_Tabelas.ZqParticipante.FieldByName('observacao').AsString:=cds_participante2observacao.Value;
      DM_Tabelas.ZqParticipante.FieldByName('codrecinc').AsInteger:=cds_participante2codrecinc.Value;
      DM_Tabelas.ZqParticipante.FieldByName('codrecBx').AsInteger:=cds_participante2codrecBx.Value;
      DM_Tabelas.ZqParticipante.FieldByName('codpaginc').AsInteger:=cds_participante2codpaginc.Value;
      DM_Tabelas.ZqParticipante.FieldByName('codpagBx').AsInteger:=cds_participante2codpagBx.Value;
      DM_Tabelas.ZqParticipante.FieldByName('fone4').AsString:=cds_participante2fone4.Value;
      DM_Tabelas.ZqParticipante.FieldByName('complemento').AsString:=cds_participante2complemento.Value;
      DM_Tabelas.ZqParticipante.FieldByName('complemento_cob').AsString:=cds_participante2complemento_cob.Value;
      if cds_participante2cadastrado.Value>0 then
         DM_Tabelas.ZqParticipante.FieldByName('cadastrado').AsDateTime:=cds_participante2cadastrado.Value;
      DM_Tabelas.ZqParticipante.FieldByName('Nome_Firma').AsString:=cds_participante2Nome_Firma.Value;
      DM_Tabelas.ZqParticipante.FieldByName('Inativo').AsString:=cds_participante2Inativo.Value;
      DM_Tabelas.ZqParticipante.FieldByName('sexo').AsString:=cds_participante2Sexo.Value;
      DM_Tabelas.ZqParticipante.FieldByName('cc').AsDateTime:=cds_participante2cc.Value;
      DM_Tabelas.ZqParticipante.FieldByName('cex').AsDateTime:=cds_participante2cex.Value;
      DM_Tabelas.ZqParticipante.FieldByName('jb').AsDateTime:=cds_participante2jb.Value;
      DM_Tabelas.ZqParticipante.FieldByName('judicial').AsString:=cds_participante2judicial.Value;
      DM_Tabelas.ZqParticipante.FieldByName('xquadra').AsString:=cds_participante2xquadra.Value;
      DM_Tabelas.ZqParticipante.FieldByName('xlote').AsString:=cds_participante2xlote.Value;
      DM_Tabelas.ZqParticipante.FieldByName('plantao').AsString:=cds_participante2plantao.Value;
      DM_Tabelas.ZQparticipante.post;

      xcodpart:=DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt;

      DM_Tabelas.ZQcomprador_proposta.close;
      DM_Tabelas.ZQcomprador_proposta.sql.clear;
      DM_Tabelas.ZQcomprador_proposta.sql.add('select * from comprador_proposta where paticipante_idpaticipante='+quotedstr(cds_participante2idpaticipante.text));
      DM_Tabelas.ZQcomprador_proposta.open;
      DM_Tabelas.ZQcomprador_proposta.First;
      while not DM_Tabelas.ZQcomprador_proposta.Eof do
      begin
        DM_Tabelas.ZQcomprador_proposta.Edit;
        DM_Tabelas.ZQComprador_proposta.FieldByName('paticipante_idpaticipante').AsLargeInt:=DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt;
        DM_Tabelas.ZQcomprador_proposta.Post;
        DM_Tabelas.ZQcomprador_proposta.Next;
      end;

      DM_Tabelas.ZQConjuge_proposta.close;
      DM_Tabelas.ZQConjuge_proposta.SQL.Clear;
      DM_Tabelas.ZQConjuge_proposta.SQL.Add('select * from conjuge_proposta where idparti='+quotedstr(cds_participante2idpaticipante.Text));
      DM_Tabelas.ZQConjuge_proposta.open;


      DM_Tabelas.ZQConjuge_proposta.First;
      if DM_Tabelas.ZQConjuge_proposta.RecordCount>0 then
      begin
        application.ProcessMessages;
        cds_conjuge2.Insert;
        cds_conjuge2nome.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('nome').AsString;
        cds_conjuge2doc1.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('doc1').AsString;
        cds_conjuge2doc2.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('doc2').AsString;
        if DM_Tabelas.ZQConjuge_proposta.FieldByName('nasc').AsDateTime>0 then
           cds_conjuge2nasc.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('nasc').AsDateTime;
        cds_conjuge2natural.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('natural').AsInteger;
        cds_conjuge2idparti.Value:=xcodpart;//cds_Conjugeidparti.Value;
        cds_conjuge2email.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('email').AsString;
        cds_conjuge2assina.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('assina').AsString;
        cds_conjuge2cargo.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('cargo').AsString;
        if DM_Tabelas.ZQConjuge_proposta.FieldByName('admissa').AsDateTime>0 then
           cds_conjuge2admissa.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('admissa').AsDateTime;
        cds_conjuge2renda.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('renda').AsFloat;
        cds_conjuge2regime.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('regime').AsString;
        cds_conjuge2comarca.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('comarca').AsFloat;
        cds_conjuge2cidade_cart.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('cidade_cart').AsFloat;
        cds_conjuge2folha.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('folha').AsString;
        cds_conjuge2livro.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('livro').AsString;
        cds_conjuge2cartorio.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('cartorio').AsString;
        if DM_Tabelas.ZQConjuge_proposta.FieldByName('data').AsDateTime>0 then
           cds_conjuge2data.Value:=DM_Tabelas.ZQConjuge_proposta.FieldByName('data').AsDateTime;
        cds_conjuge2.post;
      end;
    end;
    cds_participante2.Next;
  end;
  Gauge2.Progress:=0;
  Gauge2.MaxValue:=cds_conjuge2.RecordCount;
  cds_conjuge2.DisableControls;
  cds_conjuge2.First;
  if cds_conjuge2.RecordCount>0 then
  begin
    application.ProcessMessages;
    Label85.caption:='Verificando Conjuges...';
    application.ProcessMessages;
    DM_Tabelas.ZQConjuge.Open;
    while not cds_conjuge2.Eof do
    begin
        Gauge2.Progress:=cds_conjuge2.RecNo;
        application.ProcessMessages;
        DM_Tabelas.ZQConjuge.Insert;
        DM_Tabelas.ZQConjuge.FieldByName('nome').AsString:=cds_conjuge2nome.Value;
        DM_Tabelas.ZQConjuge.FieldByName('doc1').AsString:=cds_Conjuge2doc1.Value;
        DM_Tabelas.ZQConjuge.FieldByName('doc2').AsString:=cds_Conjuge2doc2.Value;
        if cds_Conjuge2nasc.Value>0 then
           DM_Tabelas.ZQConjuge.FieldByName('nasc').AsDateTime:=cds_Conjuge2nasc.Value;
        DM_Tabelas.ZQConjuge.FieldByName('natural').AsInteger:=cds_Conjuge2natural.Value;
        DM_Tabelas.ZQConjuge.FieldByName('idparti').AsInteger:=cds_Conjuge2idparti.Value;
        DM_Tabelas.ZQConjuge.FieldByName('email').AsString:=cds_Conjuge2email.Value;
        DM_Tabelas.ZQConjuge.FieldByName('assina').AsString:=cds_Conjuge2assina.Value;
        DM_Tabelas.ZQConjuge.FieldByName('cargo').AsString:=cds_Conjuge2cargo.Value;
        if cds_Conjuge2admissa.Value>0 then
           DM_Tabelas.ZQConjuge.FieldByName('admissa').AsDateTime:=cds_Conjuge2admissa.Value;
        DM_Tabelas.ZQConjuge.FieldByName('renda').AsFloat:=cds_Conjuge2renda.Value;
        DM_Tabelas.ZQConjuge.FieldByName('regime').AsString:=cds_Conjuge2regime.Value;
        DM_Tabelas.ZQConjuge.FieldByName('comarca').AsFloat:=cds_Conjuge2comarca.Value;
        DM_Tabelas.ZQConjuge.FieldByName('cidade_cart').AsFloat:=cds_Conjuge2cidade_cart.Value;
        DM_Tabelas.ZQConjuge.FieldByName('folha').AsString:=cds_Conjuge2folha.Value;
        DM_Tabelas.ZQConjuge.FieldByName('livro').AsString:=cds_Conjuge2livro.Value;
        DM_Tabelas.ZQConjuge.FieldByName('cartorio').AsString:=cds_Conjuge2cartorio.Value;
        if cds_Conjuge2data.Value>0 then
           DM_Tabelas.ZQConjuge.FieldByName('data').AsDateTime:=cds_Conjuge2data.Value;
        DM_Tabelas.ZQConjuge.post;
        cds_conjuge2.Next;
    end;
  end;
  Gauge2.Progress:=0;

  cds_participante2.EnableControls;
  DM_Tabelas.ZQparticipante.close;
  DM_Tabelas.ZQparticipante.SQL.Clear;
  DM_Tabelas.ZQparticipante.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  DM_Tabelas.ZQparticipante.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  DM_Tabelas.ZQparticipante.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico ');
  DM_Tabelas.ZQparticipante.SQL.Add(' from Participante ');
  DM_Tabelas.ZQparticipante.open;

  DM_Tabelas.ZQparticipante_proposta.close;
  DM_Tabelas.ZQparticipante_proposta.sql.clear;
  DM_Tabelas.ZQparticipante_proposta.sql.add('select * from participante_proposta ');
  DM_Tabelas.ZQparticipante_proposta.open;

  DM_Tabelas.ZQConjuge_proposta.close;
  DM_Tabelas.ZQConjuge_proposta.SQL.Clear;
  DM_Tabelas.ZQConjuge_proposta.SQL.Add('select * from conjuge_proposta ');
  DM_Tabelas.ZQConjuge_proposta.open;

  application.ProcessMessages;
  Label85.caption:='Verificando Imóvel...';
  application.ProcessMessages;


  DM_Tabelas.ZQimovel.close;
  DM_Tabelas.ZQimovel.SQL.Clear;
  DM_Tabelas.ZQimovel.SQL.Add('Select  loteamento_idloteamento,quadra,lote,lado,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,esquinanome,valorvenal,Obervacaoloteamento,');
  DM_Tabelas.ZQimovel.SQL.Add('        cornomapa,idimovel,disponivel,medidafrente,medidafundo,medidadedireita,esquinalinha,esquinacurva,medidaesquerda,matri,matricula,proposta,valorCusto ');
  DM_Tabelas.ZQimovel.SQL.Add('  from imovel where idimovel='+quotedstr(DM_Tabelas.ZQProposta.FieldByName('imovel').Text)+' and disponivel<>''SIM''');
  DM_Tabelas.ZQimovel.open;
  if DM_Tabelas.ZQimovel.RecordCount>0 then
  begin
    mensagem('Imóvel Não Disponível !');
    exit;
  end;

  DM_Tabelas.ZQimovel.close;
  DM_Tabelas.ZQimovel.SQL.Clear;
  DM_Tabelas.ZQimovel.SQL.Add('Select  loteamento_idloteamento,quadra,lote,lado,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,esquinanome,valorvenal,Obervacaoloteamento,');
  DM_Tabelas.ZQimovel.SQL.Add('        cornomapa,idimovel,disponivel,medidafrente,medidafundo,medidadedireita,esquinalinha,esquinacurva,medidaesquerda,matri,matricula,proposta,valorCusto ');
  DM_Tabelas.ZQimovel.SQL.Add('  from imovel ');
  DM_Tabelas.ZQimovel.open;

  application.ProcessMessages;
  Label85.caption:='Transferindo Proposta...';
  application.ProcessMessages;

  DM_Tabelas.ZQVenda.Insert;
  DM_Tabelas.ZQVenda.FieldByName('datavenda').AsDateTime:=DM_Tabelas.ZQProposta.FieldByName('dataproposta').AsDateTime;
  DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt:=DM_Tabelas.ZQProposta.FieldByName('imovel').AsLargeInt;
  DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat:=DM_Tabelas.ZQProposta.FieldByName('valorproposta').AsFloat;
  if not empty(DM_Tabelas.ZQProposta.FieldByName('forma_reajuste').AsString) then
      DM_Tabelas.ZQVenda.FieldByName('forma_reajuste').AsString:=DM_Tabelas.ZQProposta.FieldByName('forma_reajuste').AsString
  else
      DM_Tabelas.ZQVenda.FieldByName('forma_reajuste').AsString:='IGPM - FGV (ANUAL)';
  DM_Tabelas.ZQVenda.FieldByName('tabela_Price').AsFloat:=DM_Tabelas.ZQProposta.FieldByName('tabela_Price').AsFloat;
  DM_Tabelas.ZQVenda.FieldByName('Escriturado').AsString:=DM_Tabelas.ZQProposta.FieldByName('Escriturado').AsString;
  DM_Tabelas.ZQVenda.post;

  DM_Tabelas.ZQRecebimento_proposta.First;
  application.ProcessMessages;
  Label85.caption:='Transferindo as Parcelas...';
  application.ProcessMessages;
 // DM_Tabelas.ZQRecebimento_proposta.DisableControls;
  Gauge2.MaxValue:=DM_Tabelas.ZQRecebimento_proposta.RecordCount;
  Gauge2.Progress:=0;
  while not DM_Tabelas.ZQRecebimento_proposta.Eof do
  begin
    Gauge2.Progress:=DM_Tabelas.ZQRecebimento_proposta.RecNo;
    posi:=pos('-',DM_Tabelas.ZQRecebimento_proposta.FieldByName('documento').AsString);
    inc(posi);

    DM_Tabelas.ZQRecebimento.insert;
    DM_Tabelas.ZQRecebimento.FieldByName('documento').AsString:=alltrim(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-'+copy(DM_Tabelas.ZQRecebimento_proposta.FieldByName('documento').AsString,posi,20));
    DM_Tabelas.ZQRecebimento.FieldByName('cliente').AsInteger:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('cliente').AsInteger;
    DM_Tabelas.ZQRecebimento.FieldByName('usuario').AsInteger:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('usuario').AsInteger;
    DM_Tabelas.ZQRecebimento.FieldByName('Dt_Entrada').AsDateTime:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('Dt_Entrada').AsDateTime;
    DM_Tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('Dt_Vencimento').AsDateTime;
    DM_Tabelas.ZQRecebimento.FieldByName('Valor').AsFloat:=abs(DM_Tabelas.ZQRecebimento_proposta.FieldByName('Valor').AsFloat);
    DM_Tabelas.ZQRecebimento.FieldByName('Observ').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('Observ').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('VrDoc').AsFloat:=abs(DM_Tabelas.ZQRecebimento_proposta.FieldByName('VrDoc').AsFloat);


    posi:=Pos('-',DM_Tabelas.ZQRecebimento_proposta.FieldByName('documento').AsString);
    inc(posi);
    if (copy(DM_Tabelas.ZQRecebimento_proposta.FieldByName('documento').AsString,posi,1)='E') and (jaE='n') Then
    begin
      DM_tabelas.ZQNumOrdem.Insert;
      DM_Tabelas.ZQNumOrdem.Post;
      varnumordemE := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').AsLargeInt;
      jaE:='s';
    end;


    if (copy(DM_Tabelas.ZQRecebimento_proposta.FieldByName('documento').AsString,posi,1)='P') and (jaP='n') Then
    begin
      DM_tabelas.ZQNumOrdem.Insert;
      DM_Tabelas.ZQNumOrdem.Post;
      varnumordemp := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').AsLargeInt;
      jaP:='s';
    end;
    posi2:=Pos('-',DM_Tabelas.ZQRecebimento_proposta.FieldByName('documento').AsString);
    inc(posi2);

    posi:=Pos('-',DM_Tabelas.ZQRecebimento_proposta.FieldByName('ordem').AsString);
    inc(posi);

    if (copy(DM_Tabelas.ZQRecebimento_proposta.FieldByName('documento').AsString,posi2,1)='E') then
    begin
      DM_Tabelas.ZQRecebimento.FieldByName('ordem').AsString:=inttostr(varnumordemE)+'-'+alltrim(copy(DM_Tabelas.ZQRecebimento_proposta.FieldByName('ordem').AsString,posi,10));
      DM_Tabelas.ZQRecebimento.FieldByName('numordem').AsInteger:=varnumordemE;
    end;

    if (copy(DM_Tabelas.ZQRecebimento_proposta.FieldByName('documento').AsString,posi2,1)='P') then
    begin
      DM_Tabelas.ZQRecebimento.FieldByName('ordem').AsString:=inttostr(varnumordemP)+'-'+alltrim(copy(DM_Tabelas.ZQRecebimento_proposta.FieldByName('ordem').AsString,posi,10));
      DM_Tabelas.ZQRecebimento.FieldByName('numordem').AsInteger:=varnumordemP;
    end;

    DM_Tabelas.ZQRecebimento.FieldByName('TipDoc').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('TipDoc').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('saldo').AsFloat:=abs(DM_Tabelas.ZQRecebimento_proposta.FieldByName('saldo').AsFloat);
    DM_Tabelas.ZQRecebimento.FieldByName('marcar').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('marcar').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('RefBaixa').AsInteger:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('RefBaixa').AsInteger;
    DM_Tabelas.ZQRecebimento.FieldByName('refvinda').AsInteger:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('refvinda').AsInteger;
    DM_Tabelas.ZQRecebimento.FieldByName('contabil').AsInteger:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('contabil').AsInteger;
    DM_Tabelas.ZQRecebimento.FieldByName('empresa').AsInteger:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('empresa').AsInteger;
    DM_Tabelas.ZQRecebimento.FieldByName('custodaparcela').AsFloat:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('custodaparcela').AsFloat;
    DM_Tabelas.ZQRecebimento.FieldByName('origem').AsString:='V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text;


    DM_Tabelas.ZQparticipante.close;
    DM_Tabelas.ZQparticipante.sql.clear;
    DM_Tabelas.ZQparticipante.sql.add('select * from participante where nome_parte='+quotedstr(DM_Tabelas.ZQRecebimento_proposta.FieldByName('nomeadversa').AsString));
    DM_Tabelas.ZQparticipante.open;

    DM_Tabelas.ZQRecebimento.FieldByName('adversa').AsInteger:=DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt;
    DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').AsString:=DM_Tabelas.ZQParticipante.FieldByName('nome_parte').AsString;

    DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('recpag').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('idloteamento').AsInteger:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('idloteamento').AsInteger;
    DM_Tabelas.ZQRecebimento.FieldByName('quadralote').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('quadralote').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('numboleto').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('numboleto').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('Substituicao').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('Substituicao').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('sq').AsLargeInt:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('sq').AsLargeInt;
    DM_Tabelas.ZQRecebimento.FieldByName('somar').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('somar').AsString;

    DM_Tabelas.ZQRecebimento.FieldByName('Reajustado').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('Reajustado').AsString;
    if DM_Tabelas.ZQRecebimento_proposta.FieldByName('Data_reajuste').AsDateTime>0 then
       DM_Tabelas.ZQRecebimento.FieldByName('Data_reajuste').AsDateTime:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('Data_reajuste').AsDateTime;
    DM_Tabelas.ZQRecebimento.FieldByName('Proximo_Reajuste').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('Proximo_Reajuste').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('Parcelas_fixas').AsString:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('Parcelas_fixas').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('juros').AsFloat:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('juros').AsFloat;
    DM_Tabelas.ZQRecebimento.FieldByName('descontos').AsFloat:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('descontos').AsFloat;
    if DM_Tabelas.ZQRecebimento_proposta.FieldByName('Data_Quitacao').AsDateTime>0 then
       DM_Tabelas.ZQRecebimento.FieldByName('Data_Quitacao').AsDateTime:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('Data_Quitacao').AsDateTime;
    DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').AsInteger:=DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
    DM_Tabelas.ZQRecebimento.post;



{    DM_Tabelas.ZQRecebimento.close;
    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_tabelas.ZQRecebimento.SQL.Add('INSERT INTO  Recebimento (idloteamento,documento,tipdoc,cliente,dt_entrada,dt_vencimento,ordem,saldo,Valor,VrDoc,contabil,origem,adversa,nomeadversa,'+
                                     'recpag,numordem,venda_idvenda,quadralote,Observ,Parcelas_fixas,Proximo_Reajuste) VALUES ('+
                                      quotedstr(inttostr(DM_Tabelas.ZQRecebimento_propostaidloteamento.value))+', REPLACE('+quotedstr(DM_Tabelas.ZQRecebimento_propostadocumento.value)+', SUBSTRING('+quotedstr(DM_Tabelas.ZQRecebimento_propostadocumento.value)+',1,POSITION(''-'' in '+quotedstr(DM_Tabelas.ZQRecebimento_propostadocumento.value)+')-1),'+quotedstr(DM_Tabelas.ZQVendaidvenda.Text)+'),'+quotedstr(DM_Tabelas.ZQRecebimento_propostaTipDoc.value)+','+
                                      quotedstr(inttostr(DM_Tabelas.ZQRecebimento_propostacliente.value))+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.ZQRecebimento_propostaDt_Entrada.value))+','+
                                      quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.ZQRecebimento_propostaDt_Vencimento.value))+','+quotedstr(DM_Tabelas.ZQRecebimento_propostaordem.value)+','+
                                      quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_Tabelas.ZQRecebimento_propostasaldo.value),'.')))+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_Tabelas.ZQRecebimento_propostaValor.value),'.')))+','+
                                      quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_Tabelas.ZQRecebimento_propostaVrDoc.value),'.')))+','+quotedstr(inttostr(DM_Tabelas.ZQRecebimento_propostacontabil.value))+','+
                                      quotedstr('V-'+inttostr(DM_Tabelas.ZQVendaidvenda.Value))+','+quotedstr(inttostr(DM_Tabelas.ZQRecebimento_propostaadversa.value))+','+
                                      quotedstr(DM_Tabelas.ZQRecebimento_propostanomeadversa.value)+',''R'','+quotedstr(inttostr(DM_Tabelas.ZQRecebimento_propostanumordem.value))+','+quotedstr(inttostr(DM_Tabelas.ZQVendaidvenda.Value))+','+
                                      quotedstr(DM_Tabelas.ZQRecebimento_propostaquadralote.value)+','+quotedstr('Empreendimento: '+DM_Tabelas.ZQVendanometoeam.Value+chr(13)+'Quadra: '+DM_Tabelas.ZQVendaquadra.Value+chr(13)+'Lote: '+DM_Tabelas.ZQVendalote.Value)+','+
                                      quotedstr(DM_Tabelas.ZQRecebimento_propostaParcelas_fixas.value)+','+quotedstr(DM_Tabelas.ZQRecebimento_propostaProximo_Reajuste.Value)+')');

    DM_tabelas.ZQRecebimento.ExecSQL;


    DM_Tabelas.ZQRecebimento.close;
    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_tabelas.ZQRecebimento.SQL.Add('select * from recebimento  where documento='+quotedstr(DM_Tabelas.ZQRecebimento_propostadocumento.value));
    DM_tabelas.ZQRecebimento.open;
    DM_tabelas.ZQRecebimento.Last;}
   // if DM_Tabelas.ZQRecebimento_propostaTipDoc.value='CH' then
    begin
      DM_Tabelas.ZQCheque_proposta.first;
      while not DM_Tabelas.ZQCheque_proposta.Eof do
      begin
        DM_Tabelas.ZQCheque.insert;
        DM_Tabelas.ZQCheque.FieldByName('idvinculo').AsInteger:=DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').AsInteger:=DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
        DM_Tabelas.ZQCheque.FieldByName('Banco').AsString:=DM_Tabelas.ZQCheque_proposta.FieldByName('banco').AsString;
        DM_Tabelas.ZQCheque.FieldByName('agencia').AsString:=DM_Tabelas.ZQCheque_proposta.FieldByName('agencia').AsString;
        DM_Tabelas.ZQCheque.FieldByName('Dono').AsString:=DM_Tabelas.ZQCheque_proposta.FieldByName('dono').AsString;
        DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').AsString:=DM_Tabelas.ZQCheque_proposta.FieldByName('CPF_CNPJ').AsString;
        DM_Tabelas.ZQCheque.FieldByName('valor').AsFloat:=DM_Tabelas.ZQCheque_proposta.FieldByName('valor').AsFloat;
        DM_Tabelas.ZQCheque.FieldByName('emissao').AsDateTime:=DM_Tabelas.ZQCheque_proposta.FieldByName('emissao').AsDateTime;
        DM_Tabelas.ZQCheque.FieldByName('deposito').AsDateTime:=DM_Tabelas.ZQCheque_proposta.FieldByName('deposito').AsDateTime;
        DM_Tabelas.ZQCheque.FieldByName('alias').AsString:=DM_Tabelas.ZQCheque_proposta.FieldByName('alias').AsString;
        DM_Tabelas.ZQCheque.FieldByName('prorrogado').AsDateTime:=DM_Tabelas.ZQCheque_proposta.FieldByName('prorrogado').AsDateTime;
        DM_Tabelas.ZQCheque.FieldByName('conta').AsString:=DM_Tabelas.ZQCheque_proposta.FieldByName('conta').AsString;
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').AsInteger:=DM_Tabelas.ZQCheque_proposta.FieldByName('idrecebimento').AsInteger;
        DM_Tabelas.ZQCheque.FieldByName('idpagamento').AsInteger:=DM_Tabelas.ZQCheque_proposta.FieldByName('idpagamento').AsInteger;
        DM_Tabelas.ZQCheque.FieldByName('numero').AsString:=DM_Tabelas.ZQCheque_proposta.FieldByName('numero').AsString;
  //    DM_Tabelas.ZQChequeidvinculo.Value:=DM_Tabelas.ZQCheque_propostaidvinculo.Value;
        DM_Tabelas.ZQCheque.FieldByName('Substituicao').AsString:=DM_Tabelas.ZQCheque_proposta.FieldByName('Substituicao').AsString;
        DM_Tabelas.ZQCheque.FieldByName('sq').AsLargeInt:=DM_Tabelas.ZQCheque_proposta.FieldByName('sq').AsLargeInt;
        DM_Tabelas.ZQCheque.post;
        DM_Tabelas.ZQCheque_proposta.Next;
      end;
    end;

{    DM_Tabelas.ZQRecebimento.close;
    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_tabelas.ZQRecebimento.SQL.Add('update recebimento set documento=REPLACE('+quotedstr(DM_Tabelas.ZQRecebimento_propostadocumento.value)+', SUBSTRING('+quotedstr(DM_Tabelas.ZQRecebimento_propostadocumento.value)+',1,POSITION(''-'' in '+quotedstr(DM_Tabelas.ZQRecebimento_propostadocumento.value)+')-1),'+quotedstr(DM_Tabelas.ZQVendaidvenda.Text)+')  where documento='+quotedstr(DM_Tabelas.ZQRecebimento_propostadocumento.value));
    DM_tabelas.ZQRecebimento.ExecSQL;}



    DM_Tabelas.ZQRecebimento_proposta.Next;
  end;
  jaE:='n';
  jaP:='n';
  //DM_Tabelas.ZQRecebimento_proposta.EnableControls;
  DM_Tabelas.ZQcomprador_proposta.close;
  DM_Tabelas.ZQcomprador_proposta.sql.clear;
  DM_Tabelas.ZQcomprador_proposta.sql.add('select * from comprador_proposta ');
  DM_Tabelas.ZQcomprador_proposta.open;
  DM_Tabelas.ZQcomprador_proposta.First;

  Gauge2.MaxValue:=DM_Tabelas.ZQComprador_proposta.RecordCount;
  Gauge2.Progress:=0;
  DM_Tabelas.ZQComprador_proposta.first;
  application.ProcessMessages;
  Label85.caption:='Verificando Comprador...';
  application.ProcessMessages;

  while not DM_Tabelas.ZQComprador_proposta.Eof do
  begin
    Gauge2.Progress:=DM_Tabelas.ZQComprador_proposta.RecNo;
    application.ProcessMessages;
    DM_Tabelas.ZQComprador.insert;

  {  DM_Tabelas.ZQRecebimento_proposta.Close;
    DM_Tabelas.ZQRecebimento_proposta.SQL.Clear;
    DM_Tabelas.ZQRecebimento_proposta.SQL.Add('SELECT * FROM recebimento_proposta group by venda_idvenda');
    DM_Tabelas.ZQRecebimento_proposta.open;


    DM_Tabelas.ZQparticipante.close;
    DM_Tabelas.ZQparticipante.sql.clear;
    DM_Tabelas.ZQparticipante.sql.add('select * from participante where nome_parte='+quotedstr(DM_Tabelas.ZQRecebimento_propostanomeadversa.value));
    DM_Tabelas.ZQparticipante.open;

    if DM_Tabelas.ZQparticipante.RecordCount>0 then
    begin
      DM_Tabelas.ZQCompradorpaticipante_idpaticipante.Value:=DM_Tabelas.ZQparticipanteidpaticipante.Value;
    end
    else
    begin
      DM_Tabelas.ZQparticipante_proposta.close;
      DM_Tabelas.ZQparticipante_proposta.sql.clear;
      DM_Tabelas.ZQparticipante_proposta.sql.add('select * from participante_proposta where idpaticipante='+quotedstr(DM_Tabelas.ZQComprador_propostapaticipante_idpaticipante.text));
      DM_Tabelas.ZQparticipante_proposta.open;
      xnome:=DM_Tabelas.ZQparticipante_propostanome_parte.value;
      DM_Tabelas.ZQparticipante.close;
      DM_Tabelas.ZQparticipante.sql.clear;
      DM_Tabelas.ZQparticipante.sql.add('select * from participante where nome_parte='+quotedstr(xnome));
      DM_Tabelas.ZQparticipante.open;
      DM_Tabelas.ZQCompradorpaticipante_idpaticipante.Value:=DM_Tabelas.ZQparticipanteidpaticipante.Value;
    end;}


    DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').AsLargeInt:=DM_Tabelas.ZQComprador_proposta.FieldByName('paticipante_idpaticipante').AsLargeInt;
    DM_Tabelas.ZQcomprador.FieldByName('promissario').AsString:=DM_Tabelas.ZQComprador_proposta.FieldByName('promissario').AsString;
    DM_Tabelas.ZQcomprador.FieldByName('percentual').AsFloat:=DM_Tabelas.ZQComprador_proposta.FieldByName('percentual').AsFloat;
    DM_Tabelas.ZQcomprador.FieldByName('Marcar').AsInteger:= DM_Tabelas.ZQComprador_proposta.FieldByName('Marcar').AsInteger;
    DM_Tabelas.ZQcomprador.FieldByName('venda_idvenda').AsLargeInt:=DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
    DM_Tabelas.ZQComprador.post;
    DM_Tabelas.ZQComprador_proposta.Next;
  end;
  xnome:='';
  Gauge2.MaxValue:=DM_Tabelas.ZQVendedor_proposta.RecordCount;
  Gauge2.Progress:=0;
  DM_Tabelas.ZQVendedor_proposta.first;
  application.ProcessMessages;
  Label85.caption:='Verificando Vendedor...';
  application.ProcessMessages;

  while not DM_Tabelas.ZQVendedor_proposta.Eof do
  begin
    Gauge2.Progress:=DM_Tabelas.ZQVendedor_proposta.RecNo;
    application.ProcessMessages;
    DM_Tabelas.ZQVendedor.insert;
    DM_Tabelas.ZQvendedor.FieldByName('venda_idvenda').AsLargeInt:=DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
    DM_Tabelas.ZQvendedor.FieldByName('corretor_idcorretor').AsLargeInt:=DM_Tabelas.ZQVendedor_proposta.FieldByName('corretor_idcorretor').AsLargeInt;
    DM_Tabelas.ZQVendedor.post;
    DM_Tabelas.ZQVendedor_proposta.Next;
  end;


  DM_Tabelas.ZQparticipante_proposta.close;
  DM_Tabelas.ZQparticipante_proposta.sql.clear;
  DM_Tabelas.ZQparticipante_proposta.sql.add('select * from participante_proposta ');
  DM_Tabelas.ZQparticipante_proposta.open;

  DM_Tabelas.ZQparticipante.close;
  DM_Tabelas.ZQparticipante.sql.clear;
  DM_Tabelas.ZQparticipante.sql.add('select * from participante');
  DM_Tabelas.ZQparticipante.open;
  application.ProcessMessages;
  Label85.caption:='Deletando a Proposta Transferida...';
  application.ProcessMessages;

  DM_Tabelas.ZQRecebimento_proposta.Close;
  DM_Tabelas.ZQRecebimento_proposta.SQL.Clear;
  DM_Tabelas.ZQRecebimento_proposta.SQL.Add('Select * from Recebimento_proposta order by DT_Vencimento');
  DM_Tabelas.ZQRecebimento_proposta.open;


  Gauge2.MaxValue:=DM_Tabelas.ZQVendedor_proposta.RecordCount;
  Gauge2.Progress:=0;
  DM_Tabelas.ZQVendedor_proposta.first;
  while not DM_Tabelas.ZQVendedor_proposta.Eof do
  begin
    Gauge2.Progress:=DM_Tabelas.ZQVendedor_proposta.RecNo;
    application.ProcessMessages;
    DM_Tabelas.ZQVendedor_proposta.delete;
  end;

  Gauge2.MaxValue:=DM_Tabelas.ZQVendedor_proposta.RecordCount;
  Gauge2.Progress:=0;
  DM_Tabelas.ZQComprador_proposta.first;
  while not DM_Tabelas.ZQComprador_proposta.Eof do
  begin
    Gauge2.Progress:=DM_Tabelas.ZQComprador_proposta.RecNo;
    application.ProcessMessages;
    DM_Tabelas.ZQComprador_proposta.delete;
  end;

  Gauge2.MaxValue:=DM_Tabelas.ZQRecebimento_proposta.RecordCount;
  Gauge2.Progress:=0;
  DM_Tabelas.ZQRecebimento_proposta.First;

  xidloteamento:=DM_Tabelas.ZQRecebimento_proposta.FieldByName('idloteamento').AsInteger;

  while not DM_Tabelas.ZQRecebimento_proposta.Eof do
  begin
    Gauge2.Progress:=DM_Tabelas.ZQRecebimento_proposta.RecNo;
    application.ProcessMessages;
    DM_Tabelas.ZQCheque_proposta.first;
    while not DM_Tabelas.ZQCheque_proposta.Eof do
    begin
      DM_Tabelas.ZQCheque_proposta.delete;
    end;
    DM_Tabelas.ZQRecebimento_proposta.delete;
  end;

  DM_Tabelas.ZQimovel.close;
  DM_Tabelas.ZQimovel.SQL.Clear;
  DM_Tabelas.ZQimovel.SQL.Add('Select  loteamento_idloteamento,quadra,lote,lado,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,esquinanome,valorvenal,Obervacaoloteamento,');
  DM_Tabelas.ZQimovel.SQL.Add('        cornomapa,idimovel,disponivel,medidafrente,medidafundo,medidadedireita,esquinalinha,esquinacurva,medidaesquerda,matri,matricula,proposta,valorCusto ');
  DM_Tabelas.ZQimovel.SQL.Add('  from imovel where idimovel='+quotedstr(DM_Tabelas.ZQProposta.FieldByName('imovel').Text));
  DM_Tabelas.ZQimovel.open;
  if DM_Tabelas.ZQimovel.RecordCount>0 then
  begin
    if (DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString='SIM') then
     begin
       DM_Tabelas.ZQimovel.Edit;
       DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString:='NºO';
       DM_Tabelas.ZQImovel.FieldByName('proposta').AsString:='NºO';
       DM_Tabelas.ZQimovel.post;
     end;
  end;
  DM_Tabelas.ZQimovel.close;
  DM_Tabelas.ZQimovel.SQL.Clear;
  DM_Tabelas.ZQimovel.SQL.Add('Select  loteamento_idloteamento,quadra,lote,lado,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,esquinanome,valorvenal,Obervacaoloteamento,');
  DM_Tabelas.ZQimovel.SQL.Add('        cornomapa,idimovel,disponivel,medidafrente,medidafundo,medidadedireita,esquinalinha,esquinacurva,medidaesquerda,matri,matricula,proposta,valorCusto ');
  DM_Tabelas.ZQimovel.SQL.Add('  from imovel ');
  DM_Tabelas.ZQimovel.open;


  if DM_Tabelas.ZQproposta.RecordCount>0 then
     DM_Tabelas.ZQProposta.Delete;


  // aqui corrige o incoporador do loteamento para cobrança bancaria
{  ZQRecebimento_temp.close;
  ZQRecebimento_temp.SQL.Clear;
  ZQRecebimento_temp.SQL.Add('update recebimento,incorporador_loteamento set cliente=incorporador_idincorporador where idloteamento='+quotedstr(inttostr(xidloteamento))+' and loteamento_idloteamento='+quotedstr(inttostr(xidloteamento)));
  ZQRecebimento_temp.ExecSQL;
  ZQRecebimento_temp.close;}


  DM_Tabelas.ZQRecebimento.close;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('select * from recebimento');
  DM_tabelas.ZQRecebimento.open;

  Atualiza;

  application.ProcessMessages;
  Label85.caption:='';
  application.ProcessMessages;
  Gauge2.Progress:=0;
  mensagem('Transferência Completa.');
end;

procedure TFrm_Venda.dxButton5Click(Sender: TObject);
begin
  try
  if frmsimulaaditamento=nil  then
     frmsimulaaditamento:=Tfrmsimulaaditamento.create(self);
     AbrirModal(Self, frmsimulaaditamento);
  finally
    FreeAndNil(frmsimulaaditamento);
  end;
end;

procedure TFrm_Venda.Button2Click(Sender: TObject);
begin
  if mdata.Text='  /  /    ' then
  begin
    mensagem('Informe a data de Rescisão!');
    mdata.SetFocus;
    exit;
  end;  
  Gauge1.Progress:=0;
  Gauge1.Visible:=true;
  DM_Tabelas.ZQRecebimento.close;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
//  DM_Tabelas.ZQRecebimento.SQL.Add('select * from recebimento where dt_vencimento >= :dt1 and saldo > 0 and documento like '+quotedstr(DM_Tabelas.ZQVendaidvenda.Text+'-%'));
  DM_Tabelas.ZQRecebimento.SQL.Add('select * from recebimento where saldo > 0 and documento like '+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-%'));
  // marcelo pediu para colocar assim. O historico passado permanece
//  DM_Tabelas.ZQRecebimento.ParamByName('dt1').AsDate:=strtodate(mdata.Text);
  DM_Tabelas.ZQRecebimento.open;
  if DM_Tabelas.ZQRecebimento.recordCount>0 then
  begin
    ZQRescisao.Open;
    ZQRescisao.Insert;
    ZQRescisao.FieldByName('idvenda').AsLargeInt:=DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
    ZQRescisao.FieldByName('datavenda').AsDateTime:=DM_Tabelas.ZQVenda.FieldByName('datavenda').AsDateTime;
    ZQRescisao.FieldByName('datarescisao').AsDateTime:=strtodate(mdata.Text);
    ZQRescisao.FieldByName('imovel').AsLargeInt:=DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt;
    ZQRescisao.FieldByName('valorvenda').AsFloat:=DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat;
    ZQRescisao.FieldByName('idcomprador').AsLargeInt:=DM_Tabelas.CDSCompradorTempcodparticipante.Value;
    ZQRescisao.FieldByName('idparticipante').AsLargeInt:=DM_Tabelas.CDSCompradorTempcodparticipante.Value;;
    ZQRescisao.FieldByName('nome_parte').AsString:=DM_Tabelas.CDSCompradorTempnomeparticipante.Value;
    ZQRescisao.FieldByName('historico').AsString:='Rescisão de Contrato da compra do imóvel por:'+DM_Tabelas.CDSCompradorTempnomeparticipante.Value+', '+DM_Tabelas.CDSCompradorTempprofissao.Value+', '+DM_Tabelas.CDSCompradorTempestadocivil.Value+', portador do CPF/CNPJ nº '+DM_Tabelas.CDSCompradorTempdoc1.Value;
    ZQRescisao.FieldByName('hora_rescisao').AsDateTime:=time;
    ZQRescisao.FieldByName('quadralote').AsString:=DM_Tabelas.ZQRecebimento.FieldByName('quadralote').AsString;
    ZQRescisao.post;
    ZQRescisao.close;
    if DM_Tabelas.ZQRecebimento.RecordCount>0 then
    begin
      Gauge1.MaxValue:=DM_Tabelas.ZQRecebimento.RecordCount;
      Gauge1.Progress:=0;
      Gauge1.Visible:=true;
      while not DM_Tabelas.ZQRecebimento.Eof do
      begin
        Gauge1.Progress:=DM_Tabelas.ZQRecebimento.RecNo;
        DM_Tabelas.ZQReBxHi.SQL.clear;
        DM_Tabelas.ZQReBxHi.SQL.Add('delete from recbxhist where idrecib='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
        DM_Tabelas.ZQReBxHi.ExecSQL;
        DM_Tabelas.ZQReceb_Baixa.SQL.clear;
        DM_Tabelas.ZQReceb_Baixa.SQL.Add('delete from Receb_Baixa where refbaixa='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('RefBaixa').Text));
        DM_Tabelas.ZQReceb_Baixa.ExecSQL;
        ZQAditamento.SQL.clear;
        ZQAditamento.SQL.Add('delete from aditamento where idvenda='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').Text));
        ZQAditamento.ExecSQL;

        ZQQuitacao.SQL.clear;
        ZQQuitacao.SQL.Add('delete from quitacao where idvenda='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').Text));
        ZQQuitacao.ExecSQL;

        ZQcomprador_cessao.SQL.clear;
        ZQcomprador_cessao.SQL.Add('delete from comprador_cessao where idvenda='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').Text));
        ZQcomprador_cessao.ExecSQL;


        ZQrecebimento_historico_excluidos.SQL.clear;
        ZQrecebimento_historico_excluidos.SQL.Add('delete from recebimento_historico_excluidos where idrecebimento='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
        ZQrecebimento_historico_excluidos.ExecSQL;

        ZQrecebimento_historico_cessao.SQL.clear;
        ZQrecebimento_historico_cessao.SQL.Add('delete from recebimento_historico_cessao where idrecebimento='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
        ZQrecebimento_historico_cessao.ExecSQL;


        DM_Tabelas.ZQCheque.Close;
        DM_Tabelas.ZQCheque.SQL.Clear;
        DM_Tabelas.ZQCheque.SQL.Add('delete from cheque_rec where idvinculo='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
        DM_Tabelas.ZQCheque.ExecSQL;


        DM_Tabelas.ZQRecebimento.Next;
      end;
      Gauge1.Progress:=0;
      Gauge1.Visible:=false;



{     ZQParcelas_rescisao.close;
     ZQParcelas_rescisao.SQL.Clear;
     ZQParcelas_rescisao.SQL.Add('INSERT INTO `parcelas_rescisao` (`idrecebimento`,`documento`,`cliente`,`usuario`,`Dt_Entrada`,`Dt_Vencimento`,`Valor`,`Observ`,`VrDoc`,`ordem`,`TipDoc`,`saldo`,`marcar`,');
     ZQParcelas_rescisao.SQL.Add('`RefBaixa`,`refvinda`,`contabil`,`empresa`,`custodaparcela`,`origem`,`adversa`,`recpag`,`numordem`,`idloteamento`,`venda_idvenda`,`quadralote`,`numboleto`,`Substituicao`,`sq`,');
     ZQParcelas_rescisao.SQL.Add('`somar`,`nomeadversa`,`Reajustado`,`Data_reajuste`,`Proximo_Reajuste`,`Parcelas_fixas`,`juros`,`descontos`,`Data_Quitacao`,`observ_estorno`,`tip`,`sld_antes_reajuste`,');
     ZQParcelas_rescisao.SQL.Add('`Percentual_reajuste`,`juridico`,`data_juridico`,`dt_nao_pagou_no_mes`,`descricao_juridico`) VALUES ');
     ZQParcelas_rescisao.SQL.Add(' ('+quotedstr(DM_Tabelas.ZQRecebimentoidrecebimento.Text)+','+quotedstr(DM_Tabelas.ZQRecebimentodocumento.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentocliente.Text)+','+quotedstr(DM_Tabelas.ZQRecebimentousuario.Text)+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.ZQRecebimentoDt_Entrada.value))+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.ZQRecebimentodt_vencimento.Value))+',');
     ZQParcelas_rescisao.SQL.Add(quotedstr(TrocaVirgPPto(DM_Tabelas.ZQRecebimentovalor.Text))+','+quotedstr(DM_Tabelas.ZQRecebimentoobserv.Value)+','+quotedstr(TrocaVirgPPto(DM_Tabelas.ZQRecebimentoVrDoc.text))+','+quotedstr(DM_Tabelas.ZQRecebimentoordem.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentotipdoc.Value)+','+quotedstr(TrocaVirgPPto(DM_Tabelas.ZQRecebimentosaldo.text))+',');
     ZQParcelas_rescisao.SQL.Add(quotedstr(DM_Tabelas.ZQRecebimentomarcar.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentorefbaixa.Text)+','+quotedstr(DM_Tabelas.ZQRecebimentorefvinda.Text)+','+quotedstr(DM_Tabelas.ZQRecebimentocontabil.Text)+','+quotedstr(DM_Tabelas.ZQRecebimentoempresa.Text)+','+quotedstr(TrocaVirgPPto(DM_Tabelas.ZQRecebimentocustodaparcela.text))+',');
     ZQParcelas_rescisao.SQL.Add(quotedstr(DM_Tabelas.ZQRecebimentoorigem.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentoadversa.Text)+','+quotedstr(DM_Tabelas.ZQRecebimentorecpag.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentonumordem.Text)+','+quotedstr(DM_Tabelas.ZQRecebimentoidloteamento.Text)+',');
     ZQParcelas_rescisao.SQL.Add(quotedstr(DM_Tabelas.ZQRecebimentovenda_idvenda.Text)+','+quotedstr(DM_Tabelas.ZQRecebimentoquadralote.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentonumboleto.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentosubstituicao.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentosq.Text)+',');
     ZQParcelas_rescisao.SQL.Add(quotedstr(DM_Tabelas.ZQRecebimentosomar.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentonomeadversa.Value)+','+quotedstr(DM_Tabelas.ZQRecebimentoreajustado.Value)+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.ZQRecebimentodata_reajuste.Value))+','+quotedstr(DM_Tabelas.ZQRecebimentoproximo_reajuste.value)+','+quotedstr(DM_Tabelas.ZQRecebimentoparcelas_fixas.value)+',');
     ZQParcelas_rescisao.SQL.Add(quotedstr(TrocaVirgPPto(DM_Tabelas.ZQRecebimentojuros.text))+','+quotedstr(TrocaVirgPPto(DM_Tabelas.ZQRecebimentodescontos.text))+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.ZQRecebimentoData_Quitacao.Value))+','+quotedstr(DM_Tabelas.ZQRecebimentoobserv_estorno.value)+','+quotedstr(DM_Tabelas.ZQRecebimentotip.value)+','+quotedstr(TrocaVirgPPto(DM_Tabelas.ZQRecebimentosld_antes_reajuste.text))+',');
     ZQParcelas_rescisao.SQL.Add(quotedstr(TrocaVirgPPto(DM_Tabelas.ZQRecebimentopercentual_reajuste.text))+','+quotedstr(DM_Tabelas.ZQRecebimentojuridico.value)+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.ZQRecebimentodata_juridico.Value))+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_Tabelas.ZQRecebimentodt_nao_pagou_no_mes.Value))+','+quotedstr(DM_Tabelas.ZQRecebimentodescricao_juridico.value)+')');
     ZQParcelas_rescisao.SQL.Add(' where dt_vencimento >= :dt1 and saldo > 0 and documento like '+quotedstr(DM_Tabelas.ZQVendaidvenda.Text+'-%'));
     ZQParcelas_rescisao.ParamByName('dt1').AsDate:=strtodate(mdata.Text);
     ZQParcelas_rescisao.ExecSQL;
     ZQParcelas_rescisao.close;}




     ZQParcelas_rescisao.close;
     ZQParcelas_rescisao.SQL.Clear;
     ZQParcelas_rescisao.SQL.Add('INSERT INTO `parcelas_rescisao` (`idrecebimento`,`documento`,`cliente`,`usuario`,`Dt_Entrada`,`Dt_Vencimento`,`Valor`,`Observ`,`VrDoc`,`ordem`,`TipDoc`,`saldo`,`marcar`,');
     ZQParcelas_rescisao.SQL.Add('`RefBaixa`,`refvinda`,`contabil`,`empresa`,`custodaparcela`,`origem`,`adversa`,`recpag`,`numordem`,`idloteamento`,`venda_idvenda`,`quadralote`,`numboleto`,`Substituicao`,`sq`,');
     ZQParcelas_rescisao.SQL.Add('`somar`,`nomeadversa`,`Reajustado`,`Data_reajuste`,`Proximo_Reajuste`,`Parcelas_fixas`,`juros`,`descontos`,`Data_Quitacao`,`observ_estorno`,`tip`,`sld_antes_reajuste`,');
     ZQParcelas_rescisao.SQL.Add('`Percentual_reajuste`,`juridico`,`data_juridico`,`dt_nao_pagou_no_mes`,`descricao_juridico`) ');

     ZQParcelas_rescisao.SQL.Add(' SELECT idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,');
     ZQParcelas_rescisao.SQL.Add('RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,');
     ZQParcelas_rescisao.SQL.Add('somar,nomeadversa,Reajustado,Data_reajuste,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip,sld_antes_reajuste,');
     ZQParcelas_rescisao.SQL.Add('Percentual_reajuste,juridico,data_juridico,dt_nao_pagou_no_mes,descricao_juridico FROM Recebimento ');
     ZQParcelas_rescisao.SQL.Add(' where dt_vencimento <= :dt1 and saldo > 0 and documento like '+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-%'));
     ZQParcelas_rescisao.ParamByName('dt1').AsDate:=strtodate(mdata.Text);
     ZQParcelas_rescisao.ExecSQL;
     ZQParcelas_rescisao.close;



      DM_Tabelas.ZQRecebimento.close;
      DM_Tabelas.ZQRecebimento.SQL.Clear;
//      DM_Tabelas.ZQRecebimento.SQL.Add('delete from recebimento where dt_vencimento >= :dt1 and saldo > 0 and documento like '+quotedstr(DM_Tabelas.ZQVendaidvenda.Text+'-%'));
      DM_Tabelas.ZQRecebimento.SQL.Add('delete from recebimento where saldo > 0 and documento like '+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-%'));
     // DM_Tabelas.ZQRecebimento.ParamByName('dt1').AsDate:=strtodate(mdata.Text);
      DM_Tabelas.ZQRecebimento.ExecSQL;
      DM_Tabelas.ZQRecebimento.SQL.Clear;
      DM_Tabelas.ZQRecebimento.SQL.Add('select * from recebimento ');
      DM_Tabelas.ZQRecebimento.Open;

      DM_Tabelas.ZQReBxHi.SQL.Clear;
      DM_Tabelas.ZQReBxHi.SQL.Add('select * from recbxhist ');
      DM_Tabelas.ZQReBxHi.Open;

      DM_Tabelas.ZQReceb_Baixa.SQL.Clear;
      DM_Tabelas.ZQReceb_Baixa.SQL.Add('select * from Receb_Baixa ');
      DM_Tabelas.ZQReceb_Baixa.Open;



      DM_Tabelas.ZQComprador.close;
      DM_Tabelas.ZQComprador.SQL.Clear;
      DM_Tabelas.ZQComprador.SQL.Add('delete from comprador where  venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text));
      DM_Tabelas.ZQComprador.ExecSQL;
      DM_Tabelas.ZQComprador.SQL.Clear;
      DM_Tabelas.ZQComprador.SQL.Add('select * from comprador ');
      DM_Tabelas.ZQComprador.Open;

      DM_Tabelas.ZQVendedor.close;
      DM_Tabelas.ZQVendedor.SQL.Clear;
      DM_Tabelas.ZQVendedor.SQL.Add('delete from Vendedor where  venda_idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text));
      DM_Tabelas.ZQVendedor.ExecSQL;
      DM_Tabelas.ZQVendedor.SQL.Clear;
      DM_Tabelas.ZQVendedor.SQL.Add('select * from Vendedor ');
      DM_Tabelas.ZQVendedor.Open;

      DM_Tabelas.ZQImovel.Locate('idimovel',DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt,[]);
      DM_Tabelas.ZQImovel.Edit;
      DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString := 'SIM';
      DM_Tabelas.ZQImovel.FieldByName('proposta').AsString := 'NºO';
      DM_Tabelas.ZQImovel.Post;

      if DM_TAbelas.ZQmemorial.active= false then
         DM_TAbelas.ZQmemorial.open;
      if DM_TAbelas.ZQmemorial.recordcount>0 then
         DM_TAbelas.ZQmemorial.delete;
      DM_Tabelas.ZQVenda.Delete;

      DM_Tabelas.ZQCheque.Close;
      DM_Tabelas.ZQCheque.SQL.Clear;
      DM_Tabelas.ZQCheque.SQL.Add('select * from cheque_rec ');
      DM_Tabelas.ZQCheque.Open;
      mensagem('Rescisão Efetuada Com Sucesso!');
    end;
  end
  else
  begin
    mensagem('Rescisão Não Efetuada. Parcelas todas pagas e inadimplentes!');
  end;  
  mdata.Clear;
  gbrescisao.Visible:=false;

end;

procedure TFrm_Venda.ZQInandimplenciaCalcFields(DataSet: TDataSet);
var
zmora:double;
begin
{  zmora:=0;
  ZQInandimplenciaDias.Value:= strtoint(DifDias(ZQInandimplenciaDt_Vencimento.Value, date));
  if Lmora.Caption='N' then
     ZQInandimplenciamora_dia.Value:=ExRound(((((ZQInandimplenciasaldo.Value/30)*ZQInandimplenciaMora.value)/100)*ZQInandimplenciaDias.Value),2)
  else
     ZQInandimplenciamora_dia.Value:=ExRound(((((ZQInandimplenciasaldo.Value/30)*ZQInandimplenciatabela_price.value)/100)*ZQInandimplenciaDias.Value),2);
  ZQInandimplenciaMulta_Contrato.Value:=ExRound(ExRound((ZQInandimplenciasaldo.Value*ZQInandimplenciaMulta.value),2)/100,2);
  ZQInandimplenciaparcela_corrigida.Value:=ZQInandimplenciasaldo.Value+ZQInandimplenciaMulta_Contrato.Value+ZQInandimplenciamora_dia.Value;
 }


//  if Lmora.Caption='N' then
//     ZQInandimplenciamora_dia.Value:=ExRound(((((ZQInandimplenciasaldo.Value/30)*DM_Tabelas.ZQConfiguracoesMora.value)/100)*ZQInandimplenciaDias.Value),2)
//  else
//     ZQInandimplenciamora_dia.Value:=ExRound(((((ZQInandimplenciasaldo.Value/30)*ZQInandimplenciatabela_price.value)/100)*ZQInandimplenciaDias.Value),2);
//  ZQInandimplenciaMulta_Contrato.Value:=ExRound(ExRound((ZQInandimplenciasaldo.Value*DM_Tabelas.ZQConfiguracoesMulta.value),2)/100,2);
//  ZQInandimplenciaparcela_corrigida.Value:=ZQInandimplenciasaldo.Value+ZQInandimplenciaMulta_Contrato.Value+ZQInandimplenciamora_dia.Value;
end;


procedure TFrm_Venda.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQParcelas', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas', 'ZQParcelasvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas', 'ZQParcelasquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas', 'ZQParcelasidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas', 'ZQParcelassaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas', 'ZQParcelasorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas', 'ZQParcelasadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas', 'ZQParcelasvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQParcelatotal', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelatotal', 'ZQParcelatotalvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelatotal', 'ZQParcelatotalquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelatotal', 'ZQParcelatotalidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelatotal', 'ZQParcelatotalsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelatotal', 'ZQParcelatotalorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelatotal', 'ZQParcelatotaladversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelatotal', 'ZQParcelatotalvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelatotal', 'ZQParcelatotalparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQparcelas_pagas', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagassaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasdt_vencimento', 'dt_vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagascontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagascontabildescri', 'contabildescri', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'descricao', True);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagastipdoc', 'tipdoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas_pagas', 'ZQparcelas_pagasordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQAditamento', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQAditamento', 'ZQAditamentocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQQuitacao', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQQuitacao', 'ZQQuitacaocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQcomprador_cessao', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador_cessao', 'ZQcomprador_cessaoautocodigo', 'autocodigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQrecebimento_historico_excluidos', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosmotivo', 'motivo', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossequencial', 'sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQrecebimento_historico_cessao', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoSequencial', 'Sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaonumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaodata_cessao', 'data_cessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQrecebimento_historico_cessao', 'ZQrecebimento_historico_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQparcelas12', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas12', 'ZQparcelas12documento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas12', 'ZQparcelas12saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas12', 'ZQparcelas12orde', 'orde', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas12', 'ZQparcelas12origem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparcelas12', 'ZQparcelas12vlr', 'vlr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQresumo_parcelas', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQresumo_parcelas', 'ZQresumo_parcelasorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQresumo_parcelas', 'ZQresumo_parcelasdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQresumo_parcelas', 'ZQresumo_parcelasvr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQresumo_parcelas', 'ZQresumo_parcelassaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQresumo_parcelas', 'ZQresumo_parcelasct', 'ct', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQresumo_parcelas', 'ZQresumo_parcelassoma', 'soma', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQResumo_entrada', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQResumo_entrada', 'ZQResumo_entradaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQResumo_entrada', 'ZQResumo_entradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQResumo_entrada', 'ZQResumo_entradavr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQResumo_entrada', 'ZQResumo_entradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQResumo_entrada', 'ZQResumo_entradact', 'ct', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQResumo_entrada', 'ZQResumo_entradasoma', 'soma', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQtot_parcelas', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQtot_parcelas', 'ZQtot_parcelasorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQtot_parcelas', 'ZQtot_parcelasdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQtot_parcelas', 'ZQtot_parcelasvr', 'vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQtot_parcelas', 'ZQtot_parcelassaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQtot_parcelas', 'ZQtot_parcelasct', 'ct', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQtot_parcelas', 'ZQtot_parcelassoma', 'soma', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQInandimplencia', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciadt_vencimento', 'dt_vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciaMora_dia', 'Mora_dia', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciamulta_contrato', 'multa_contrato', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciaParcela_corrigida', 'Parcela_corrigida', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciaDias', 'Dias', TIntegerField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciatabela_price', 'tabela_price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciamora', 'mora', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQInandimplencia', 'ZQInandimplenciamulta', 'multa', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQimovel', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovellote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovellado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelmedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelmedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
   RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelmedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
   RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimoveldescricao', 'descricao', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelmedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelcornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimoveldisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelmatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelmatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQimovel', 'ZQimovelproposta', 'proposta', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQProposta', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQProposta', 'ZQPropostaidproposta', 'idproposta', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQProposta', 'ZQPropostadataproposta', 'dataproposta', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQProposta', 'ZQPropostaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQProposta', 'ZQPropostavalorproposta', 'valorproposta', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQProposta', 'ZQPropostaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQProposta', 'ZQPropostatabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQProposta', 'ZQPropostaEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQProposta', 'ZQPropostaPerc_comissao', 'Perc_comissao', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQProposta', 'ZQPropostavlr_comissao', 'vlr_comissao', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQcomprador', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador', 'ZQcompradoridcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador', 'ZQcompradorpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador', 'ZQcompradorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador', 'ZQcompradorpromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador', 'ZQcompradorpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador', 'ZQcompradorMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcomprador', 'ZQcompradoridproposta', 'idproposta', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQRecebimento_proposta', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaidrecebimento_proposta', 'idrecebimento_proposta', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostanumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostanomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostadescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRecebimento_proposta', 'ZQRecebimento_propostaData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQcheques', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesbanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesdono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesdeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesalias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesnumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequesSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQcheques', 'ZQchequessq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQConjuge', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugeidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugenome', 'nome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugedoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugedoc2', 'doc2', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugenasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugenatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugeidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugeemail', 'email', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugeassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugeempresa', 'empresa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugecargo', 'cargo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugeadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugerenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugeregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugecomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugecidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugefolha', 'folha', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugelivro', 'livro', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugecartorio', 'cartorio', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQConjuge', 'ZQConjugedata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQvendedor', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQvendedor', 'ZQvendedoridvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQvendedor', 'ZQvendedorcorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQvendedor', 'ZQvendedorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQvendedor', 'ZQvendedoridproposta', 'idproposta', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQvendedor', 'ZQvendedorPcomissao_corretor', 'Pcomissao_corretor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQparticipante', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantenome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantedoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantedoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantebairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantebairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantefone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantefone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantetipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantenacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantenaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantelocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanterenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantefone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteSexo', 'Sexo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantecex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantejb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantejudicial', 'judicial', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantexquadra', 'xquadra', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipantexlote', 'xlote', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQparticipante', 'ZQparticipanteplantao', 'plantao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQRescisao', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaodatarescisao', 'datarescisao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaohora_rescisao', 'hora_rescisao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaoautocodigo', 'autocodigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQRescisao', 'ZQRescisaoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQPrice', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQPrice', 'ZQPriceidPrice', 'idPrice', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQPrice', 'ZQPriceTaxa', 'Taxa', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQPrice', 'ZQPriceMes', 'Mes', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQPrice', 'ZQPricecoeficiente_multiplicador', 'coeficiente_multiplicador', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQPrice', 'ZQPriceCoeficiente_amortizador', 'Coeficiente_amortizador', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Venda, 'ZQParcelas_rescisao', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoobserv_estorno', 'observ_estorno', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaotip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaosld_antes_reajuste', 'sld_antes_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaoPercentual_reajuste', 'Percentual_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaojuridico', 'juridico', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaodata_juridico', 'data_juridico', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaodt_nao_pagou_no_mes', 'dt_nao_pagou_no_mes', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Venda, 'ZQParcelas_rescisao', 'ZQParcelas_rescisaodescricao_juridico', 'descricao_juridico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
