unit RelVenda;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner, Mask, XDate, ExtCtrls, Grids, DBGrids,
    dxButton, CheckLst, DB, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, DBClient, TFlatRadioButtonUnit, IniFiles, Gauges, dxCore2,system.Threading,
  FnpNumericEdit;

type
  TFrm_RelVenda = class(TForm)
    XBanner14: TXBanner;
    Label19: TLabel;
    DXBImprimir: TdxButton;
    DXBFechar: TdxButton;
    GroupBox1: TGroupBox;
    XBanner1: TXBanner;
    Label1: TLabel;
    Label2: TLabel;
    XDEEntradaInicio: TXDateEdit;
    XDEEntradaFinal: TXDateEdit;
    GBEmpree: TGroupBox;
    XBanner3: TXBanner;
    XBanner4: TXBanner;
    Corretor: TGroupBox;
    XBanner6: TXBanner;
    Ecorretor: TEdit;
    DBGrid2: TDBGrid;
    Comprador: TGroupBox;
    XBanner2: TXBanner;
    DBGrid3: TDBGrid;
    EComprador: TEdit;
    CLBEmpree: TCheckListBox;
    DXBMarcEmpree: TdxButton;
    DS_Venda: TDataSource;
    ZQVenda: TZQuery;

























    DS_Empree: TDataSource;
    ZQEmpree: TZQuery;



















    GroupBox3: TGroupBox;
    XBanner8: TXBanner;
    XBanner9: TXBanner;
    Label3: TLabel;
    Equadra: TEdit;
    DBGrid1: TDBGrid;
    Elote: TEdit;
    DS_QuadraLote: TDataSource;
    CDSQuadraLote: TClientDataSet;
    CDSQuadraLoteQuadra: TWideStringField;
    CDSQuadraLotelote: TWideStringField;
    CDSQuadraLoteidlote: TIntegerField;
    DS_Entrada: TDataSource;
    DS_Parcela: TDataSource;
    ZQEntrada: TZQuery;
    ZQParcela: TZQuery;
    DS_Particip: TDataSource;
    ZQParticip: TZQuery;





































    DS_BxEntrada: TDataSource;
    ZQBxEntrada: TZQuery;
    GroupBox4: TGroupBox;
    XBanner5: TXBanner;
    FRBA: TFlatRadioButton;
    FRBS: TFlatRadioButton;
    CDSQuadraLoteidloteamen: TIntegerField;
    DS_Ven1: TDataSource;
    ZQVen1: TZQuery;

























    DS_vendedor: TDataSource;
    ZQVendedor: TZQuery;

























    DS_BsParcela: TDataSource;
    ZQBxParcela: TZQuery;











    Datacheque_rec_e: TDataSource;
    ZQcheque_rec_e: TZQuery;















    DataZQcheque_rec_p: TDataSource;
    ZQcheque_rec_p: TZQuery;















    Databaixa: TDataSource;
    baixa: TZQuery;

















    Datacvenda: TDataSource;
    CBtotais: TCheckBox;
    ZQrecebparcela: TZQuery;
    ZQrecebentrada: TZQuery;


    DataZQrecebentrada: TDataSource;
    DataZQrecebparcela: TDataSource;


    Gauge1: TGauge;
    Datamontarelatorio: TDataSource;
    ZQReceb_BaixaE: TZQuery;



    DS_Receb_baixaE: TDataSource;
    ZQReceb_BaixaP: TZQuery;
    DS_Receb_baixap: TDataSource;




    ZQcomprador_cessao: TZQuery;
    DataZQcomprador_cessao: TDataSource;










    ZQaditamento: TZQuery;
    DataZQaditamento: TDataSource;










    ZQquitcao: TZQuery;
    Datazqquitacao: TDataSource;














































    CBquitados: TCheckBox;
    Datazqquitados: TDataSource;
    ZQquitados: TZQuery;
    ZQVendaDataSource: TDataSource;
    ZQVendaZQuery: TZQuery;

























    CBContabil: TCheckBox;
    cbcessao: TCheckBox;
    DScessao: TDataSource;
    ZQcessao: TZQuery;




























































































































    DS_ZQRelquitados: TDataSource;
    ZQrelquitados: TZQuery;

























































































































    CBPercent: TCheckBox;
    ZQTemp2: TZQuery;



















    ZQtemp: TZQuery;
    DS_zqtemp2: TDataSource;
    CBanalitico: TCheckBox;
    ZQtemp3: TZQuery;
    ZQTemp4: TZQuery;
    DS_zqtemp4: TDataSource;






















    CBanalitico2: TCheckBox;









































































    CBDiarias: TCheckBox;
    ZQdiaria: TZQuery;
    DS_diaria: TDataSource;



















































































































    montarelatorio: TClientDataSet;
    montarelatorioidvenda: TLargeintField;
    montarelatoriodatavenda: TDateField;
    montarelatorioimovel: TLargeintField;
    montarelatoriovalorvenda: TFloatField;
    montarelatorioidimovel: TLargeintField;
    montarelatorioquadra: TStringField;
    montarelatoriolote: TStringField;
    montarelatoriolado: TStringField;
    montarelatoriomedidafrente: TStringField;
    montarelatoriomedidafundo: TStringField;
    montarelatoriomedidaesquerda: TStringField;
    montarelatoriomedidadedireita: TStringField;
    montarelatorioarea: TFloatField;
    montarelatorioconfrontofrente: TStringField;
    montarelatorioconfrontofundo: TStringField;
    montarelatorioconfrontodireita: TStringField;
    montarelatorioconfrontoesquerda: TStringField;
    montarelatorioesquinanome: TStringField;
    montarelatorioesquinalinha: TStringField;
    montarelatorioesquinacurva: TStringField;
    montarelatoriovalorvenal: TFloatField;
    montarelatorioObervacaoloteamento: TMemoField;
    montarelatoriocornomapa: TStringField;
    montarelatorioidpaticipante: TLargeintField;
    montarelatorionome_parte: TStringField;
    montarelatoriodoc1: TStringField;
    montarelatoriodoc2: TStringField;
    montarelatorioendereco: TStringField;
    montarelatoriobairro: TStringField;
    montarelatoriocidade: TIntegerField;
    montarelatoriocep: TStringField;
    montarelatorioende_cob: TStringField;
    montarelatoriobairro_cob: TStringField;
    montarelatoriocidade_cob: TLargeintField;
    montarelatoriocep_cob: TStringField;
    montarelatorioFone1: TStringField;
    montarelatoriofone2: TStringField;
    montarelatoriofone3: TStringField;
    montarelatoriotipopessoa: TStringField;
    montarelatorioaniversario: TDateField;
    montarelatorionacionalidade: TStringField;
    montarelatorioemail: TStringField;
    montarelatorionaturalidade: TIntegerField;
    montarelatoriolocaldetrab: TStringField;
    montarelatorioprofissao: TStringField;
    montarelatorioestadocivil: TStringField;
    montarelatoriorenda: TFloatField;
    montarelatorioobservacao: TMemoField;
    montarelatoriocodrecinc: TIntegerField;
    montarelatoriocodrecBx: TIntegerField;
    montarelatoriocodpaginc: TIntegerField;
    montarelatoriocodpagBx: TIntegerField;
    montarelatoriofone4: TStringField;
    montarelatoriocomplemento: TStringField;
    montarelatoriocomplemento_cob: TStringField;
    montarelatoriocadastrado: TDateField;
    montarelatorioidcomprador: TLargeintField;
    montarelatoriopaticipante_idpaticipante: TLargeintField;
    montarelatoriovenda_idvenda: TLargeintField;
    montarelatoriopromissario: TStringField;
    montarelatoriodisponivel: TStringField;
    montarelatorioidloteamento: TIntegerField;
    montarelatoriovr_entrada: TFloatField;
    montarelatoriovr_parcela: TFloatField;
    montarelatoriojuros: TFloatField;
    montarelatoriodescontos: TFloatField;
    montarelatoriojuros_p: TFloatField;
    montarelatoriodescontos_p: TFloatField;
    Cvenda: TClientDataSet;
    Cvendaidvenda: TLargeintField;
    Cvendadatavenda: TDateField;
    Cvendaimovel: TLargeintField;
    Cvendavalorvenda: TFloatField;
    Cvendaidimovel: TLargeintField;
    Cvendaquadra: TStringField;
    Cvendalote: TStringField;
    Cvendalado: TStringField;
    Cvendamedidafrente: TStringField;
    Cvendamedidafundo: TStringField;
    Cvendamedidaesquerda: TStringField;
    Cvendamedidadedireita: TStringField;
    Cvendaarea: TFloatField;
    Cvendaconfrontofrente: TStringField;
    Cvendaconfrontofundo: TStringField;
    Cvendaconfrontodireita: TStringField;
    Cvendaconfrontoesquerda: TStringField;
    Cvendaesquinanome: TStringField;
    Cvendaesquinalinha: TStringField;
    Cvendaesquinacurva: TStringField;
    Cvendavalorvenal: TFloatField;
    CvendaObervacaoloteamento: TMemoField;
    Cvendacornomapa: TStringField;
    Cvendaidpaticipante: TLargeintField;
    Cvendanome_parte: TStringField;
    Cvendadoc1: TStringField;
    Cvendadoc2: TStringField;
    Cvendaendereco: TStringField;
    Cvendabairro: TStringField;
    Cvendacidade: TIntegerField;
    Cvendacep: TStringField;
    Cvendaende_cob: TStringField;
    Cvendabairro_cob: TStringField;
    Cvendacidade_cob: TLargeintField;
    Cvendacep_cob: TStringField;
    CvendaFone1: TStringField;
    Cvendafone2: TStringField;
    Cvendafone3: TStringField;
    Cvendatipopessoa: TStringField;
    Cvendaaniversario: TDateField;
    Cvendanacionalidade: TStringField;
    Cvendaemail: TStringField;
    Cvendanaturalidade: TIntegerField;
    Cvendalocaldetrab: TStringField;
    Cvendaprofissao: TStringField;
    Cvendaestadocivil: TStringField;
    Cvendarenda: TFloatField;
    Cvendaobservacao: TMemoField;
    Cvendacodrecinc: TIntegerField;
    CvendacodrecBx: TIntegerField;
    Cvendacodpaginc: TIntegerField;
    CvendacodpagBx: TIntegerField;
    Cvendafone4: TStringField;
    Cvendacomplemento: TStringField;
    Cvendacomplemento_cob: TStringField;
    Cvendacadastrado: TDateField;
    Cvendaidcomprador: TLargeintField;
    Cvendapaticipante_idpaticipante: TLargeintField;
    Cvendavenda_idvenda: TLargeintField;
    Cvendapromissario: TStringField;
    Cvendadisponivel: TStringField;
    Cvendaidloteamento: TIntegerField;
    Cvendavr_entrada: TFloatField;
    Cvendavr_parcela: TFloatField;
    Cvendajuros: TFloatField;
    Cvendadescontos: TFloatField;
    Cvendajuros_p: TFloatField;
    Cvendadescontos_p: TFloatField;



    ZQVenda2: TZQuery;
    DS_Zqvenda2: TDataSource;

























































    DS_balao: TDataSource;
    ZQBalao: TZQuery;















    ZQRescisao: TZQuery;
    DSrescisao: TDataSource;
    CBRescisao: TCheckBox;






























































































































    procedure DXBFecharClick(Sender: TObject);
    procedure EcorretorExit(Sender: TObject);
    procedure ECompradorExit(Sender: TObject);
    procedure DXBImprimirClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DXBMarcEmpreeClick(Sender: TObject);
    procedure XDEEntradaInicioExit(Sender: TObject);
    procedure XDEEntradaFinalExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure GroupBox3Enter(Sender: TObject);
    procedure EloteExit(Sender: TObject);
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid3KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure GroupBox3Exit(Sender: TObject);
    procedure DBGrid2KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ZQReceb_BaixaEAfterPost(DataSet: TDataSet);
    procedure ZQReceb_BaixaPAfterPost(DataSet: TDataSet);
    procedure ZQcomprador_cessaoAfterPost(DataSet: TDataSet);
    procedure ZQVen1AfterPost(DataSet: TDataSet);
    procedure ZQEmpreeAfterPost(DataSet: TDataSet);
    procedure ZQVendaAfterPost(DataSet: TDataSet);
    procedure ZQEntradaAfterPost(DataSet: TDataSet);
    procedure ZQParcelaAfterPost(DataSet: TDataSet);
    procedure ZQParticipAfterPost(DataSet: TDataSet);
    procedure ZQVendedorAfterPost(DataSet: TDataSet);
    procedure ZQcheque_rec_eAfterPost(DataSet: TDataSet);
    procedure ZQBxParcelaAfterPost(DataSet: TDataSet);
    procedure ZQBxEntradaAfterPost(DataSet: TDataSet);
    procedure ZQrecebentradaAfterPost(DataSet: TDataSet);
    procedure ZQrecebparcelaAfterPost(DataSet: TDataSet);
    procedure baixaAfterPost(DataSet: TDataSet);
    procedure ZQcheque_rec_pAfterPost(DataSet: TDataSet);
    procedure ZQaditamentoAfterPost(DataSet: TDataSet);
    procedure ZQquitcaoAfterPost(DataSet: TDataSet);
    procedure CBtotaisClick(Sender: TObject);
    procedure CBContabilClick(Sender: TObject);
    procedure cbcessaoClick(Sender: TObject);
    procedure CBanaliticoClick(Sender: TObject);
    procedure CBDiariasClick(Sender: TObject);
    procedure CBRescisaoClick(Sender: TObject);
    procedure CBquitadosClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelVenda: TFrm_RelVenda;
  varschemata : string;
  
implementation

uses tabelas, funcoes, RelVenda01, RelVenda02, RelVenda02_totais,
  RelReceb03, RelReceb02_cessao, RelReceb02_quitado,
  RelVenda02_totais_analitico, RelReceb02_diaria, RelReceb02_quitado2,
  RelReceb02_rescisao, uRuntimeFields;

{$R *.dfm}

procedure TFrm_RelVenda.DXBFecharClick(Sender: TObject);
begin
  DXBImprimir.Enabled:=true;
  Close;
end;

procedure TFrm_RelVenda.EcorretorExit(Sender: TObject);
begin
  if Not empty(Ecorretor.Text) Then Begin
    if AchaCorretor(200, 10,Ecorretor.Text) Then Begin
      DM_Tabelas.CDSVendedorTemp.Insert;
      DM_Tabelas.CDSVendedorTempcodcorretor.Value := DM_TAbelas.ZQAchaCorretor.FieldByName('idcorretor').AsLargeInt;
      DM_Tabelas.CDSVendedorTemp.Post;
    end;
    Ecorretor.Text := '';
    Ecorretor.SetFocus;
  end
end;

procedure TFrm_RelVenda.ECompradorExit(Sender: TObject);
begin
  if not empty(Ecomprador.Text) Then Begin
    if AchaCLiente(200, 10,Ecomprador.Text) Then Begin
      DM_Tabelas.CDSCompradorTemp.Insert;
      DM_Tabelas.CDSCompradorTempcodparticipante.Value := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
      DM_Tabelas.CDSCompradorTemp.Post;
    end;
    Ecomprador.Text := '';
    Ecomprador.SetFocus;
  end;
end;

procedure TFrm_RelVenda.DXBImprimirClick(Sender: TObject);
var
  ql,VarEmpree, VarImovel, varParte, varvendedor, TempTableName : String;
  tm,posi,varcontador, x, y : integer;
begin
   DXBImprimir.Enabled:=false;


  if CBDiarias.Checked=true then
  begin

     VarParte := '';
    DM_Tabelas.CDSCompradorTemp.First;
    While not DM_Tabelas.CDSCompradorTemp.Eof do begin
      varparte := varparte+quotedstr(DM_Tabelas.CDSCompradorTempcodparticipante.Text)+',';
      DM_Tabelas.CDSCompradorTemp.Next;
    end;
    Varparte := copy(varparte,1,length(varparte)-1);
    varvendedor := '';
    DM_Tabelas.CDSVendedorTemp.First;
    while not DM_Tabelas.CDSVendedorTemp.Eof do begin
      varvendedor := varvendedor+quotedstr(DM_Tabelas.CDSVendedorTempcodcorretor.Text)+',';
      DM_Tabelas.CDSVendedorTemp.Next;
    end;
    varvendedor := copy(varvendedor,1,length(varvendedor)-1);
    VarImovel := '';
    CDSQuadraLote.First;
    while not CDSQuadraLote.Eof do begin
      Varimovel := VarImovel + quotedstr(CDSQuadraLoteidlote.Text)+',';
      CDSQuadraLote.Next;
    end;
    Varimovel := copy(VarImovel,1,length(VarImovel)-1);
    VarEmpree:= '';
    varcontador := 0;
    for x := 0 to CLBEmpree.Count-1 do begin
      if CLBEmpree.Checked[x] Then Begin
        inc(varcontador);
        for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
          if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
            break;
        VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
      end;
    end;
    if varcontador<>CLBEmpree.Count Then
      VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
    else
      VarEmpree := '';


    ZQdiaria.Close;
    ZQdiaria.SQL.clear;
    ZQdiaria.SQL.Add(' SELECT v.valorvenda,v.idvenda,v.datavenda,v.imovel,v.forma_reajuste,v.tabela_Price,v.Escriturado,v.marca,v.vlr_comissao,');
    ZQdiaria.SQL.Add('        i.idimovel,i.loteamento_idloteamento,i.quadra,i.lote,i.lado,i.medidafrente,i.medidafundo,i.medidaesquerda,i.medidadedireita,i.area,i.confrontofrente,i.confrontofundo,');
    ZQdiaria.SQL.Add('        i.confrontodireita,i.confrontoesquerda,i.esquinanome,i.esquinalinha,i.esquinacurva,i.valorvenal,i.Obervacaoloteamento,i.cornomapa,i.disponivel,i.matri,i.matricula,i.proposta,');
    ZQdiaria.SQL.Add('        l.idloteamento,l.cidade_idcidade,l.matriculaloteamento,l.datacadastro,l.datainalguracao,l.dataconclusao,l.logotipo,l.mapa,l.testemunha1,l.testemunha2,');
    ZQdiaria.SQL.Add('        l.testemunha3,l.testemunha4,l.Observincorp,l.bairro,l.pasta_mapa,l.nomeloteamento,l.tipo,l.codcontabancaria,l.apelido,l.digito_dif,');
    ZQdiaria.SQL.Add('        vd.idvendedor,vd.corretor_idcorretor,vd.venda_idvenda,');
    ZQdiaria.SQL.Add('        c.idcorretor,c.nomecorretor,c.doc1,c.creci,c.telefone1,c.telefone2,c.endereco_corr,c.bairro_corr,c.cep_corr,c.Obser_corr,c.codcidade,');
    ZQdiaria.SQL.Add('        cp.idcomprador,cp.paticipante_idpaticipante,cp.venda_idvenda,cp.promissario,cp.percentual,cp.Marcar,');
    ZQdiaria.SQL.Add('        pa.idpaticipante,pa.nome_parte,pa.doc1,pa.doc2,pa.endereco,pa.bairro,pa.cidade,pa.cep,pa.ende_cob,pa.bairro_cob,pa.cidade_cob,pa.cep_cob,pa.Fone1,');
    ZQdiaria.SQL.Add('        pa.fone2,pa.fone3,pa.tipopessoa,pa.aniversario,pa.nacionalidade,pa.email,pa.naturalidade,pa.localdetrab,pa.profissao,pa.estadocivil,pa.renda,');
    ZQdiaria.SQL.Add('        pa.observacao,pa.codrecinc,pa.codrecBx,pa.codpaginc,pa.codpagBx,pa.fone4,pa.complemento,pa.complemento_cob,pa.cadastrado,pa.Nome_Firma,pa.Inativo,');
    ZQdiaria.SQL.Add('        pa.Sexo,pa.cc,pa.cex,pa.jb,pa.judicial,pa.xquadra,pa.xlote,pa.Plantao ');
    ZQdiaria.SQL.Add('  FROM venda as v join imovel as i on v.imovel=i.idimovel ');
    ZQdiaria.SQL.Add(' join loteamento as l on l.idloteamento=i.loteamento_idloteamento join vendedor as vd on vd.venda_idvenda=v.idvenda ');
    ZQdiaria.SQL.Add(' join corretor as c on c.idcorretor=vd.corretor_idcorretor join comprador as cp on cp.venda_idvenda=v.idvenda ');
    ZQdiaria.SQL.Add(' join participante as pa on pa.idpaticipante=cp.paticipante_idpaticipante ');
    ZQdiaria.SQL.Add(' where v.datavenda between :dt1 and :dt2');
    if not empty(VarEmpree) Then
      ZQdiaria.SQL.Add(' and l.idloteamento in ('+VarEmpree+')');
    if not empty(varvendedor) Then
       ZQdiaria.SQL.Add(' and vd.corretor_idcorretor in ('+varvendedor+')');
    ZQdiaria.SQL.Add(' group by v.idvenda order by c.nomecorretor,v.datavenda');
    ZQdiaria.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
    ZQdiaria.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;

    ZQdiaria.Open;
    ZQdiaria.First;
    if ZQdiaria.RecordCount>0 then
    begin
      if FrmRelReceb02_diaria=nil then
         FrmRelReceb02_diaria:=TFrmRelReceb02_diaria.Create(Application);
      FrmRelReceb02_diaria.RLLabel2.Caption:='Período de '+XDEEntradaInicio.Text + ' até '+XDEEntradaFinal.Text;
      FrmRelReceb02_diaria.RLReport1.PreviewModal;
      FrmRelReceb02_diaria:=nil;
    end
    else
    begin
      Showmessage('Não há Venda nesse Período.');
    end;
    DXBImprimir.Enabled:=true;
    ZQdiaria.Close;
    exit;
  end;
  //21/06/2012
  if (cbcessao.Checked=true) and (CBRescisao.Checked=false) then
  begin
    VarEmpree:= '';
    varcontador := 0;
    for x := 0 to CLBEmpree.Count-1 do begin
      if CLBEmpree.Checked[x] Then Begin
        inc(varcontador);
        for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
          if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
            break;
        VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
      end;
    end;
    if varcontador<>CLBEmpree.Count Then
      VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
    else
      VarEmpree := '';
    ZQcessao.Close;
    ZQcessao.SQL.clear;
//    ZQcessao.SQL.Add(' SELECT *,COUNT(*) as total FROM recebimento_historico_cessao as rb join loteamento as lt on rb.idloteamento=lt.idloteamento ');
    ZQcessao.SQL.Add(' SELECT cc.idvenda,cc.datacessao,cc.imovel,cc.valorvenda,cc.idcomprador,cc.idparticipante,cc.nome_parte,cc.historico,cc.hora_cessao,cc.autocodigo,cc.datavenda,');
    ZQcessao.SQL.Add('        cm.idcomprador,cm.paticipante_idpaticipante,cm.venda_idvenda,cm.promissario,cm.percentual,cm.Marcar,');
    ZQcessao.SQL.Add('        v.valorvenda,v.idvenda,v.datavenda,v.imovel,v.forma_reajuste,v.tabela_Price,v.Escriturado,v.marca,');
    ZQcessao.SQL.Add('        rc.idrecebimento,rc.documento,rc.cliente,rc.usuario,rc.Dt_Entrada,rc.Dt_Vencimento,rc.Valor,rc.Observ,rc.VrDoc,rc.ordem,rc.TipDoc,rc.saldo,');
    ZQcessao.SQL.Add('        rc.RefBaixa,rc.refvinda,rc.contabil,rc.empresa,rc.custodaparcela,rc.origem,rc.adversa,rc.recpag,rc.numordem,rc.idloteamento,rc.quadralote,rc.numboleto,');
    ZQcessao.SQL.Add('        rc.Substituicao,rc.sq,rc.somar,rc.nomeadversa,rc.Reajustado,rc.Data_reajuste,rc.Proximo_Reajuste,rc.Parcelas_fixas,rc.juros,rc.descontos,rc.Data_Quitacao,');
    ZQcessao.SQL.Add('        lt.idloteamento,lt.cidade_idcidade,lt.matriculaloteamento,lt.datacadastro,lt.datainalguracao,lt.dataconclusao,lt.logotipo,lt.mapa,lt.testemunha1,lt.testemunha2,');
    ZQcessao.SQL.Add('        lt.testemunha3,lt.testemunha4,lt.Observincorp,lt.bairro,lt.pasta_mapa,lt.nomeloteamento,lt.tipo,lt.codcontabancaria,lt.apelido,lt.digito_dif,');
    ZQcessao.SQL.Add('        pt.idpaticipante,pt.nome_parte,pt.doc1,pt.doc2,pt.endereco,pt.bairro,pt.cidade,pt.cep,pt.ende_cob,pt.bairro_cob,pt.cidade_cob,pt.cep_cob,pt.Fone1,');
    ZQcessao.SQL.Add('        pt.fone2,pt.fone3,pt.tipopessoa,pt.aniversario,pt.nacionalidade,pt.email,pt.naturalidade,pt.localdetrab,pt.profissao,pt.estadocivil,pt.renda,pt.observacao,');
    ZQcessao.SQL.Add('        pt.codrecinc,pt.codrecBx,pt.codpaginc,pt.codpagBx,pt.fone4,pt.complemento,pt.complemento_cob,pt.cadastrado,pt.Nome_Firma,pt.Nome_Firma,');
    ZQcessao.SQL.Add('        rc.marcar,rc.venda_idvenda ');
    ZQcessao.SQL.Add(' FROM comprador_cessao as cc join comprador as cm on cc.idvenda=cm.venda_idvenda join venda as v on v.idvenda=cc.idvenda join recebimento as rc on rc.venda_idvenda=cc.idvenda');
    ZQcessao.SQL.Add(' join loteamento as lt on rc.idloteamento=lt.idloteamento join participante as pt on pt.idpaticipante=cm.paticipante_idpaticipante');
    ZQcessao.SQL.Add(' where cc.datacessao between :dt1 and :dt2 ');

    if not empty(VarEmpree) Then
      ZQcessao.SQL.Add(' and rc.idloteamento in ('+VarEmpree+')');

//    ZQcessao.SQL.Add(' group by rb.idloteamento order by lt.apelido');
    ZQcessao.SQL.Add(' group by rc.quadralote order by lt.apelido');
    ZQcessao.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
    ZQcessao.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;
    ZQcessao.Open;
    ZQcessao.First;
    if ZQcessao.RecordCount>0 then
    begin
      if FrmRelReceb02_cessao=nil then
         FrmRelReceb02_cessao:=TFrmRelReceb02_cessao.Create(Application);
      FrmRelReceb02_cessao.RLLabel2.Caption:='Período de '+XDEEntradaInicio.Text + ' até '+XDEEntradaFinal.Text;
      FrmRelReceb02_cessao.RLReport1.PreviewModal;
      FrmRelReceb02_cessao:=nil;
    end
    else
    begin
      Showmessage('Não há Cessão nesse período.');
    end;
    ZQcessao.Close;
  end

  else if (cbcessao.Checked=false) and (CBRescisao.Checked=true) then
  begin
    VarEmpree:= '';
    varcontador := 0;
    for x := 0 to CLBEmpree.Count-1 do begin
      if CLBEmpree.Checked[x] Then Begin
        inc(varcontador);
        for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
          if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
            break;
        VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
      end;
    end;
    if varcontador<>CLBEmpree.Count Then
      VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
    else
      VarEmpree := '';
    ZQRescisao.Close;
    ZQRescisao.SQL.clear;


    ZQRescisao.SQL.Add(' SELECT rs.idvenda,rs.datavenda,rs.datarescisao,rs.imovel,rs.valorvenda,rs.idcomprador,rs.idparticipante,rs.nome_parte,rs.historico,rs.hora_rescisao,rs.autocodigo,rs.quadralote,');
    ZQRescisao.SQL.Add('        rc.idrecebimento,rc.documento,rc.cliente,rc.usuario,rc.Dt_Entrada,rc.Dt_Vencimento,rc.Valor,rc.Observ,rc.VrDoc,rc.ordem,rc.TipDoc,rc.saldo,rc.marcar,rc.RefBaixa,');
    ZQRescisao.SQL.Add('        rc.refvinda,rc.contabil,rc.empresa,rc.custodaparcela,rc.origem,rc.adversa,rc.recpag,rc.numordem,rc.idloteamento,rc.venda_idvenda,rc.quadralote,rc.numboleto,');
    ZQRescisao.SQL.Add('        rc.Substituicao,rc.sq,rc.somar,rc.nomeadversa,rc.Reajustado,rc.Data_reajuste,rc.Proximo_Reajuste,rc.Parcelas_fixas,rc.juros,rc.descontos,rc.Data_Quitacao,');
    ZQRescisao.SQL.Add('        rc.observ_estorno,rc.tip,rc.sld_antes_reajuste,rc.Percentual_reajuste,rc.juridico,rc.data_juridico,rc.dt_nao_pagou_no_mes,rc.descricao_juridico,');
    ZQRescisao.SQL.Add('        lt.idloteamento,lt.cidade_idcidade,lt.matriculaloteamento,lt.datacadastro,lt.datainalguracao,lt.dataconclusao,lt.logotipo,lt.mapa,lt.testemunha1,lt.testemunha2,');
    ZQRescisao.SQL.Add('        lt.testemunha3,lt.testemunha4,lt.Observincorp,lt.bairro,lt.pasta_mapa,lt.nomeloteamento,lt.tipo,lt.codcontabancaria,lt.apelido,lt.digito_dif,');
    ZQRescisao.SQL.Add('        pt.idpaticipante,pt.nome_parte,pt.doc1,pt.doc2,pt.endereco,pt.bairro,pt.cidade,pt.cep,pt.ende_cob,pt.bairro_cob,pt.cidade_cob,pt.cep_cob,pt.Fone1,pt.fone2,');
    ZQRescisao.SQL.Add('        pt.fone3,pt.tipopessoa,pt.aniversario,pt.nacionalidade,pt.email,pt.naturalidade,pt.localdetrab,pt.profissao,pt.estadocivil,pt.renda,pt.observacao,pt.codrecinc,');
    ZQRescisao.SQL.Add('        pt.codrecBx,pt.codpaginc,pt.codpagBx,pt.fone4,pt.complemento,pt.complemento_cob,pt.cadastrado,pt.Nome_Firma,pt.Inativo,pt.Sexo,pt.cc,pt.cex,pt.jb,pt.judicial,pt.xquadra,pt.xlote,pt.Plantao ');
    ZQRescisao.SQL.Add(' FROM  rescisao as rs  join recebimento as rc on rc.venda_idvenda=rs.idvenda');
    ZQRescisao.SQL.Add(' join loteamento as lt on rc.idloteamento=lt.idloteamento join participante as pt on pt.idpaticipante=rs.idparticipante ');
    ZQRescisao.SQL.Add(' where rs.datarescisao between :dt1 and :dt2 ');

    if not empty(VarEmpree) Then
      ZQRescisao.SQL.Add(' and rc.idloteamento in ('+VarEmpree+')');

//    ZQRescisao.SQL.Add(' group by rb.idloteamento order by lt.apelido');
    ZQRescisao.SQL.Add(' group by rc.quadralote order by lt.apelido');
    ZQRescisao.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
    ZQRescisao.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;

   // showmessage(zqrescisao.sql.Text);

    ZQRescisao.Open;
    ZQRescisao.First;
    if ZQRescisao.RecordCount>0 then
    begin
      if not assigned(FrmRelReceb02_rescisao) then
         FrmRelReceb02_rescisao:=TFrmRelReceb02_rescisao.Create(Application);
      FrmRelReceb02_rescisao.RLLabel2.Caption:='Período de '+XDEEntradaInicio.Text + ' até '+XDEEntradaFinal.Text;
      FrmRelReceb02_rescisao.RLReport1.PreviewModal;
      FrmRelReceb02_rescisao:=nil;
    end
    else
    begin
      Showmessage('Não há Rescisão nesse período.');
    end;
    ZQRescisao.Close;
  end

  else if (cbcessao.Checked=false) and (CBRescisao.Checked=false) then
  begin
     if (CDSQuadraLote.RecordCount=0) and (DM_Tabelas.CDSVendedorTemp.RecordCount=0) then
       Comprador.Enabled:=true;

     if (CDSQuadraLote.RecordCount=0) and (DM_Tabelas.CDSCompradorTemp.RecordCount=0) then
       Corretor.Enabled:=true;


     VarParte := '';
    DM_Tabelas.CDSCompradorTemp.First;
    While not DM_Tabelas.CDSCompradorTemp.Eof do begin
      varparte := varparte+quotedstr(DM_Tabelas.CDSCompradorTempcodparticipante.Text)+',';
      DM_Tabelas.CDSCompradorTemp.Next;
    end;
    Varparte := copy(varparte,1,length(varparte)-1);
    varvendedor := '';
    DM_Tabelas.CDSVendedorTemp.First;
    while not DM_Tabelas.CDSVendedorTemp.Eof do begin
      varvendedor := varvendedor+quotedstr(DM_Tabelas.CDSVendedorTempcodcorretor.Text)+',';
      DM_Tabelas.CDSVendedorTemp.Next;
    end;
    varvendedor := copy(varvendedor,1,length(varvendedor)-1);
    VarImovel := '';
    CDSQuadraLote.First;
    while not CDSQuadraLote.Eof do begin
      Varimovel := VarImovel + quotedstr(CDSQuadraLoteidlote.Text)+',';
      CDSQuadraLote.Next;
    end;
    Varimovel := copy(VarImovel,1,length(VarImovel)-1);
    VarEmpree:= '';
    varcontador := 0;
    for x := 0 to CLBEmpree.Count-1 do begin
      if CLBEmpree.Checked[x] Then Begin
        inc(varcontador);
        for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
          if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
            break;
        VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
      end;
    end;
    if varcontador<>CLBEmpree.Count Then
      VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
    else
      VarEmpree := '';

    ZQVen1.close;
    ZQVen1.SQL.Clear;
    ZQVen1.SQL.Add(' Select v.idvenda,v.datavenda,v.imovel,v.valorvenda,');
    ZQVen1.SQL.Add('        i.idimovel as idimovel,i.loteamento_idloteamento,i.quadra,i.lote,i.lado,i.medidafrente,i.medidafundo,i.medidaesquerda,i.medidadedireita,i.area,i.confrontofrente,i.confrontofundo,i.confrontodireita,');
    ZQVen1.SQL.Add('        i.confrontoesquerda,i.esquinanome,i.esquinalinha,i.esquinacurva,i.valorvenal,i.Obervacaoloteamento,i.cornomapa,i.disponivel ');
    ZQVen1.SQL.Add(' from venda as v join imovel as i on i.idimovel=v.imovel ');
    ZQVen1.SQL.Add(' where datavenda between :dt1 and :dt2');
    
    if not empty(VarImovel) Then
      ZQVen1.SQL.Add(' and idImovel in ('+VarImovel+')')
    else if (not empty(varParte)) or (not empty(varvendedor)) Then
    Begin
      if not empty(varParte) Then
        ZQVen1.SQL.Add(' and exists(select idcomprador,paticipante_idpaticipante,venda_idvenda,promissario,percentual,marcar from comprador  where paticipante_idpaticipante in ('+varParte+') and venda_idvenda=idvenda)');
      if not empty(varvendedor) Then
        ZQVen1.SQL.Add(' and exists(select idvendedor,corretor_idcorretor,venda_idvenda from vendedor  where corretor_idcorretor in ('+varvendedor+') and venda_idvenda=idvenda)');
    end;
    if not empty(VarEmpree) Then
      ZQVen1.SQL.Add(' and loteamento_idloteamento in ('+VarEmpree+')');
    ZQVen1.SQL.Add(' group by loteamento_idloteamento');
    ZQVen1.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
    ZQVen1.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;
    ZQVen1.Open;
    ZQVen1.First;
    VarEmpree := '';
    while not ZQVen1.eof do begin
      VarEmpree := VarEmpree + quotedstr(ZQVen1.FieldByName('loteamento_idloteamento').Text)+',';
      ZQVen1.Next;
    end;
    VarEmpree := Copy(VarEmpree,1,length(varempree)-1);

    ZQEmpree.close;
    ZQEmpree.SQL.Clear;
    ZQEmpree.SQL.Add('select idloteamento,cidade_idcidade,matriculaloteamento,datacadastro,datainalguracao,dataconclusao,logotipo,mapa,testemunha1,testemunha2,');
    ZQEmpree.SQL.Add('       testemunha3,testemunha4,Observincorp,bairro,pasta_mapa,nomeloteamento,tipo,codcontabancaria,apelido,digito_dif ');
    ZQEmpree.SQL.Add(' from loteamento');
    if not empty(VarEmpree) then
      ZQEmpree.SQL.Add('where idloteamento in  ('+VarEmpree+')');
    ZQEmpree.Open;
    ZQquitados.close;


   // 03/01/2012
    if CBContabil.Checked=true then
    begin
      VarEmpree:= '';
      varcontador := 0;
      for x := 0 to CLBEmpree.Count-1 do
      begin
         if CLBEmpree.Checked[x] Then
         Begin
         inc(varcontador);
          for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
              if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
              break;
          VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
         end;
      end;
      if varcontador<>CLBEmpree.Count Then
         VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
      else
         VarEmpree := '';

      VarImovel := '';
      CDSQuadraLote.First;
      while not CDSQuadraLote.Eof do begin
        ql:=CDSQuadraLotequadra.Text+'-'+CDSQuadraLotelote.Text;
        Varimovel := VarImovel + quotedstr(ql)+',';
        CDSQuadraLote.Next;
      end;
      Varimovel := copy(VarImovel,1,length(VarImovel)-1);

      if FrmRelReceb03=nil then
         FrmRelReceb03:=TFrmRelReceb03.Create(Application);
        tm:=length(VarEmpree);
      if tm>3 then
      begin
        FrmRelReceb03.RLDBImage1.Visible:=false;
        FrmRelReceb03.RLDBText1.Visible:=false;
        FrmRelReceb03.RLBand4.Visible:=true;
      end
      else
      begin
        FrmRelReceb03.RLDBImage1.Visible:=true;
        FrmRelReceb03.RLDBText1.Visible:=true;
        FrmRelReceb03.RLBand4.Visible:=false;
      end;
      FrmRelReceb03.ZQParticip.open;
      FrmRelReceb03.ZQcontabil.close;
      FrmRelReceb03.ZQcontabil.SQL.Clear;

      FrmRelReceb03.ZQcontabil.SQL.Add(' SELECT RB.idreceb_baixa,RB.Dt_rec,RB.Vr_rec,RB.TipDoc,RB.Docum,RB.Juros_Vr,RB.Desc_Vr,RB.RefBaixa,RB.obsebx,RB.Codcontabil,RB.dataref,RB.vencimento,RB.Substituicao,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        RB.sq,RB.VR_AB,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        RH.idrecbxhist,RH.refer,RH.idrecib,RH.valor,RH.descon,RH.juros,RH.percent_usado,RH.data,RH.sq,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        RE.idrecebimento,RE.documento,RE.cliente,RE.usuario,RE.Dt_Entrada,RE.Dt_Vencimento,RE.Valor,RE.Observ,RE.VrDoc,RE.ordem,RE.TipDoc,RE.saldo,RE.marcar,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        RE.RefBaixa,RE.refvinda,RE.contabil,RE.empresa,RE.custodaparcela,RE.origem,RE.adversa,RE.recpag,RE.numordem,RE.idloteamento,RE.venda_idvenda,RE.quadralote,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        RE.numboleto,RE.Substituicao,RE.sq,RE.somar,RE.nomeadversa,RE.Reajustado,RE.Data_reajuste,RE.Proximo_Reajuste,RE.Parcelas_fixas,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        v.idvenda,v.datavenda,v.imovel,v.valorvenda,v.forma_reajuste,v.tabela_Price,v.Escriturado,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        c.idcomprador,c.paticipante_idpaticipante,c.venda_idvenda,c.promissario,c.percentual,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        pt.idpaticipante,pt.nome_parte,pt.doc1,pt.doc2,pt.endereco,pt.bairro,pt.cidade,pt.cep,pt.ende_cob,pt.bairro_cob,pt.cidade_cob,pt.cep_cob,pt.Fone1,pt.fone2,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        pt.fone3,pt.tipopessoa,pt.aniversario,pt.nacionalidade,pt.email,pt.naturalidade,pt.localdetrab,pt.profissao,pt.estadocivil,pt.renda,pt.observacao,pt.codrecinc,');
      FrmRelReceb03.ZQcontabil.SQL.Add('        pt.codrecBx,pt.codpaginc,pt.codpagBx,pt.fone4,pt.complemento,pt.complemento_cob,pt.cadastrado,pt.Nome_Firma, ');
      FrmRelReceb03.ZQcontabil.SQL.Add('        cb.n_banco,cb.n_agencia,cb.n_agencia_v,cb.n_conta,cb.n_conta_v,cb.nomebanco ');
      FrmRelReceb03.ZQcontabil.SQL.Add(' FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa ');
      FrmRelReceb03.ZQcontabil.SQL.Add(' join recebimento as RE on RH.idrecib=RE.idrecebimento join venda as v on v.idvenda=re.venda_idvenda ');
      FrmRelReceb03.ZQcontabil.SQL.Add(' join comprador as c on c.venda_idvenda=re.venda_idvenda join participante as pt on pt.idpaticipante=c.paticipante_idpaticipante ');
      FrmRelReceb03.ZQcontabil.SQL.Add(' join imovel as i on i.idimovel=v.imovel ');
      FrmRelReceb03.ZQcontabil.SQL.Add(' join incorporador_loteamento as l on l.loteamento_idloteamento=i.loteamento_idloteamento ');
      FrmRelReceb03.ZQcontabil.SQL.Add(' join conta_bancaria cb on cb.idconta_bancaria=l.codcontabancaria ');
      FrmRelReceb03.ZQcontabil.SQL.Add(' where rb.dataref between :dt1 and :dt2 ');
      if not empty(VarImovel) Then
         FrmRelReceb03.ZQcontabil.SQL.Add(' and re.quadralote in ('+VarImovel+')');
      if (not empty(varParte)) or (not empty(varvendedor)) Then
      Begin
        if not empty(varParte) Then
           FrmRelReceb03.ZQcontabil.SQL.Add(' and exists(select idcomprador,paticipante_idpaticipante,venda_idvenda,promissario,percentual,marcar from comprador  where paticipante_idpaticipante in ('+varParte+') and venda_idvenda=idvenda)');
        if not empty(varvendedor) Then
           FrmRelReceb03.ZQcontabil.SQL.Add(' and exists(select ve.idvendedor,ve.corretor_idcorretor,ve.venda_idvenda from vendedor as ve  where corretor_idcorretor in ('+varvendedor+') and ve.venda_idvenda=re.venda_idvenda)');
      end;
      if not empty(VarEmpree) Then
          FrmRelReceb03.ZQcontabil.SQL.Add(' and re.idloteamento in ('+VarEmpree+')');
      FrmRelReceb03.ZQcontabil.SQL.Add(' group by re.idloteamento, re.quadralote ');//re.nomeadversa tava no lugar do pt.nome_parte
      FrmRelReceb03.ZQcontabil.SQL.Add(' order by re.idloteamento, pt.nome_parte ');
      FrmRelReceb03.ZQcontabil.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
      FrmRelReceb03.ZQcontabil.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;
      FrmRelReceb03.ZQcontabil.Open;
      FrmRelReceb03.ZQcontabil.First;
      FrmRelReceb03.ZQcontabil.recordcount;


     //FrmRelReceb03.ZQparcelas.open;
      FrmRelReceb03.ZQparcelas.close;
      FrmRelReceb03.ZQparcelas.SQL.Clear;
      FrmRelReceb03.ZQparcelas.SQL.Add(' SELECT RB.idreceb_baixa,RB.Dt_rec,RB.Vr_rec,RB.TipDoc,RB.Docum,RB.Juros_Vr,RB.Desc_Vr,RB.RefBaixa,RB.obsebx,RB.Codcontabil,RB.dataref,RB.vencimento,RB.Substituicao,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        RB.sq,RB.VR_AB,RB.Valor_parcela,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        RH.idrecbxhist,RH.refer,RH.idrecib,RH.valor,RH.descon,RH.juros,RH.percent_usado,RH.data,RH.sq,RH.Valor_parcela,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        RE.idrecebimento,RE.documento,RE.cliente,RE.usuario,RE.Dt_Entrada,RE.Dt_Vencimento,RE.Valor,RE.Observ,RE.VrDoc,RE.ordem,RE.TipDoc,RE.saldo,RE.marcar,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        RE.RefBaixa,RE.refvinda,RE.contabil,RE.empresa,RE.custodaparcela,RE.origem,RE.adversa,RE.recpag,RE.numordem,RE.idloteamento,RE.venda_idvenda,RE.quadralote,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        RE.numboleto,RE.Substituicao,RE.sq,RE.somar,RE.nomeadversa,RE.Reajustado,RE.Data_reajuste,RE.Proximo_Reajuste,RE.Parcelas_fixas,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        v.idvenda,v.datavenda,v.imovel,v.valorvenda,v.forma_reajuste,v.tabela_Price,v.Escriturado,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        c.idcomprador,c.paticipante_idpaticipante,c.venda_idvenda,c.promissario,c.percentual,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        pt.idpaticipante,pt.nome_parte,pt.doc1,pt.doc2,pt.endereco,pt.bairro,pt.cidade,pt.cep,pt.ende_cob,pt.bairro_cob,pt.cidade_cob,pt.cep_cob,pt.Fone1,pt.fone2,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        pt.fone3,pt.tipopessoa,pt.aniversario,pt.nacionalidade,pt.email,pt.naturalidade,pt.localdetrab,pt.profissao,pt.estadocivil,pt.renda,pt.observacao,pt.codrecinc,');
      FrmRelReceb03.ZQparcelas.SQL.Add('        pt.codrecBx,pt.codpaginc,pt.codpagBx,pt.fone4,pt.complemento,pt.complemento_cob,pt.cadastrado,pt.Nome_Firma ');
      FrmRelReceb03.ZQparcelas.SQL.Add(' FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa ');
      FrmRelReceb03.ZQparcelas.SQL.Add(' join recebimento as RE on RH.idrecib=RE.idrecebimento join venda as v on v.idvenda=re.venda_idvenda ');
      FrmRelReceb03.ZQparcelas.SQL.Add(' join comprador as c on c.venda_idvenda=re.venda_idvenda join participante as pt on pt.idpaticipante=c.paticipante_idpaticipante ');
      FrmRelReceb03.ZQparcelas.SQL.Add(' where rb.dataref between :dt1 and :dt2 ');
      if not empty(VarImovel) Then
         FrmRelReceb03.ZQparcelas.SQL.Add(' and re.quadralote in ('+VarImovel+')')
      else if (not empty(varParte)) or (not empty(varvendedor)) Then Begin
//      if not empty(varParte) Then
  //     FrmRelReceb03.ZQparcelas.SQL.Add(' and exists(select * from comprador  where paticipante_idpaticipante in ('+varParte+') and venda_idvenda=idvenda)');
      if not empty(varvendedor) Then
         FrmRelReceb03.ZQparcelas.SQL.Add(' and exists(select idvendedor,corretor_idcorretor,venda_idvenda from vendedor  where corretor_idcorretor in ('+varvendedor+') and venda_idvenda=re.venda_idvenda)');
      end
      else
      if not empty(VarEmpree) Then
          FrmRelReceb03.ZQparcelas.SQL.Add(' and re.idloteamento in ('+VarEmpree+')');
      FrmRelReceb03.ZQparcelas.SQL.Add('  group by rb.idreceb_baixa order by re.dt_vencimento,re.idloteamento, re.nomeadversa ');
      FrmRelReceb03.ZQparcelas.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
      FrmRelReceb03.ZQparcelas.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;
      FrmRelReceb03.ZQparcelas.Open;
      FrmRelReceb03.ZQparcelas.First;
      FrmRelReceb03.ZQparcelas.recordcount;

      FrmRelReceb03.RLLabel2.Caption:='Período de '+XDEEntradaInicio.Text + ' até '+XDEEntradaFinal.Text;

      ZQcomprador_cessao.MasterSource:=nil;
      ZQcomprador_cessao.MasterFields:='';
      ZQcomprador_cessao.LinkedFields:='';

      ZQquitcao.MasterSource:=nil;
      ZQquitcao.MasterFields:='';
      ZQquitcao.LinkedFields:='';

      ZQaditamento.MasterSource:=nil;
      ZQaditamento.MasterFields:='';
      ZQaditamento.LinkedFields:='';

      FrmRelReceb03.Cabecalho.close;
      FrmRelReceb03.Cabecalho.CreateDataSet;
      FrmRelReceb03.Cabecalho2.close;
      FrmRelReceb03.Cabecalho2.CreateDataSet;
      ZQaditamento.Close;
      ZQquitcao.close;
      ZQcomprador_cessao.close;
      ZQcomprador_cessao.SQL.Clear;
      ZQcomprador_cessao.SQL.add('Select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora_cessao from comprador_cessao where idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text));
      ZQcomprador_cessao.Open;
      Gauge1.MaxValue:=ZQcomprador_cessao.RecordCount;
      ZQcomprador_cessao.First;
      ZQcomprador_cessao.DisableControls;
      while not ZQcomprador_cessao.eof do
      begin
         application.ProcessMessages;
         gauge1.Progress:=ZQcomprador_cessao.RecNo;
         FrmRelReceb03.Cabecalho2.Insert;
         FrmRelReceb03.Cabecalho2idvenda.Value:=ZQcomprador_cessao.FieldByName('idvenda').AsLargeInt;
         FrmRelReceb03.Cabecalho2datavenda.Value:=ZQcomprador_cessao.FieldByName('datavenda').AsDateTime;
         FrmRelReceb03.Cabecalho2data.Value:=ZQcomprador_cessao.FieldByName('datacessao').AsDateTime;
         FrmRelReceb03.Cabecalho2imovel.Value:=ZQcomprador_cessao.FieldByName('imovel').AsLargeInt;
         FrmRelReceb03.Cabecalho2valorvenda.Value:=ZQcomprador_cessao.FieldByName('valorvenda').AsFloat;
         FrmRelReceb03.Cabecalho2idcomprador.Value:=ZQcomprador_cessao.FieldByName('idcomprador').AsLargeInt;
         FrmRelReceb03.Cabecalho2idparticipante.Value:=ZQcomprador_cessao.FieldByName('idparticipante').AsLargeInt;
         FrmRelReceb03.Cabecalho2nome_parte.Value:=ZQcomprador_cessao.FieldByName('nome_parte').AsString;
         FrmRelReceb03.Cabecalho2historico.Value:=ZQcomprador_cessao.FieldByName('historico').AsString;
         FrmRelReceb03.Cabecalho2hora.Value:=ZQcomprador_cessao.FieldByName('hora_cessao').AsDateTime;
         FrmRelReceb03.Cabecalho2.post;
         ZQcomprador_cessao.Next;
      end;
      ZQcomprador_cessao.First;
      ZQcomprador_cessao.EnableControls;
      Gauge1.Visible:=true;
      Gauge1.Progress:=0;
      application.ProcessMessages;
      ZQaditamento.close;
      ZQaditamento.SQL.Clear;
      ZQaditamento.SQL.Add('select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora from aditamento where idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text));
      ZQaditamento.open;
      Gauge1.MaxValue:=ZQaditamento.RecordCount;

      ZQaditamento.First;
      ZQaditamento.DisableControls;
      while not ZQaditamento.eof do
      begin
        application.ProcessMessages;
        gauge1.Progress:=ZQaditamento.RecNo;
        FrmRelReceb03.Cabecalho2.Insert;
        FrmRelReceb03.Cabecalho2idvenda.Value:=zqaditamento.FieldByName('idvenda').AsLargeInt;
  //     FrmRelReceb03.Cabecalho2data.Value:=ZQaditamentodatacessao.Value;
        FrmRelReceb03.Cabecalho2historico.Value:=zqaditamento.FieldByName('historico').AsString;
//       FrmRelReceb03.Cabecalho2hora.Value:=ZQaditamentohora.Value;
        FrmRelReceb03.Cabecalho2.post;
        ZQaditamento.Next;
      end;
      ZQaditamento.First;
      ZQaditamento.EnableControls;
      ZQquitcao.close;
      ZQquitcao.SQL.Clear;
      ZQquitcao.SQL.Add('select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora from quitacao where idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text));
      ZQquitcao.open;
      ZQquitcao.First;
      Gauge1.Visible:=true;
      Gauge1.Progress:=0;
      application.ProcessMessages;
      Gauge1.MaxValue:=ZQquitcao.RecordCount;
      ZQquitcao.DisableControls;
      while not ZQquitcao.eof do
      begin
          application.ProcessMessages;
          gauge1.Progress:=ZQquitcao.RecNo;
          FrmRelReceb03.Cabecalho2.Insert;
          FrmRelReceb03.Cabecalho2idvenda.Value:=ZQquitcao.FieldByName('idvenda').AsLargeInt;
          FrmRelReceb03.Cabecalho2data.Value:=ZQquitcao.FieldByName('datacessao').AsDateTime;
          FrmRelReceb03.Cabecalho2historico.Value:=ZQquitcao.FieldByName('historico').AsString;
          FrmRelReceb03.Cabecalho2hora.Value:=ZQquitcao.FieldByName('hora').AsDateTime;
          FrmRelReceb03.Cabecalho2.post;
          ZQquitcao.Next;
      end;
      ZQquitcao.First;
      ZQquitcao.EnableControls;

      FrmRelReceb03.Cabecalho2.IndexName:='dataidx';
      FrmRelReceb03.Cabecalho2.First;
      Gauge1.Visible:=true;
      Gauge1.Progress:=0;
      application.ProcessMessages;
      Gauge1.MaxValue:=FrmRelReceb03.Cabecalho2.RecordCount;

      if FrmRelReceb03.Cabecalho2.RecordCount>0 then
      begin
        FrmRelReceb03.Cabecalho2.DisableControls;
        while not FrmRelReceb03.Cabecalho2.eof do
        begin
          application.ProcessMessages;        
          gauge1.Progress:=FrmRelReceb03.Cabecalho2.RecNo;
          FrmRelReceb03.Cabecalho.Insert;
          FrmRelReceb03.Cabecalhoidvenda.Value:=FrmRelReceb03.Cabecalho2idvenda.Value;
          FrmRelReceb03.Cabecalhodatavenda.Value:=FrmRelReceb03.Cabecalho2datavenda.Value;
          FrmRelReceb03.Cabecalhodata.Value:=FrmRelReceb03.Cabecalho2data.Value;
          FrmRelReceb03.Cabecalhoimovel.Value:=FrmRelReceb03.Cabecalho2imovel.Value;
          FrmRelReceb03.Cabecalhovalorvenda.Value:=FrmRelReceb03.Cabecalho2valorvenda.Value;
          FrmRelReceb03.Cabecalhoidcomprador.Value:=FrmRelReceb03.Cabecalho2idcomprador.Value;
          FrmRelReceb03.Cabecalhoidparticipante.Value:=FrmRelReceb03.Cabecalho2idparticipante.Value;
          FrmRelReceb03.Cabecalhonome_parte.Value:=FrmRelReceb03.Cabecalho2nome_parte.Value;
          FrmRelReceb03.Cabecalhohistorico.Value:=FrmRelReceb03.Cabecalho2historico.Value;
          FrmRelReceb03.Cabecalhohora.Value:=FrmRelReceb03.Cabecalho2hora.Value;
          FrmRelReceb03.Cabecalho.post;
          FrmRelReceb03.Cabecalho2.Next;
        end;
        FrmRelReceb03.Cabecalho2.EnableControls;
      end;
      Gauge1.Progress:=0;
      application.ProcessMessages;      
      Gauge1.Visible:=false;
      FrmRelReceb03.ZQEmpree.Open;
      if DM_Tabelas.ZqParticipante.Active=false then
         DM_Tabelas.ZqParticipante.Open;


      FrmRelReceb03.ZQrecebimento_historico_excluidos.Close;
      FrmRelReceb03.ZQrecebimento_historico_excluidos.SQL.Clear;
      FrmRelReceb03.ZQrecebimento_historico_excluidos.SQL.Add('SELECT idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,');
      FrmRelReceb03.ZQrecebimento_historico_excluidos.SQL.Add('       origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar,motivo,Proximo_Reajuste,sequencial ');
      FrmRelReceb03.ZQrecebimento_historico_excluidos.SQL.Add(' FROM recebimento_historico_excluidos where motivo =''Q'' and venda_idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text)+' group by adversa');
      FrmRelReceb03.ZQrecebimento_historico_excluidos.open;
      FrmRelReceb03.RLReport1.PreviewModal;
      FrmRelReceb03:=nil;

  {    FrmRelReceb03.ZQcontabil.close;
      FrmRelReceb03.ZQParticip.close;
      FrmRelReceb03.ZQEmpree.close;
      FrmRelReceb03.cabecalho2.close;
      FrmRelReceb03.cabecalho.close;
      FrmRelReceb03.ZQparcelas.close;
      FrmRelReceb03.ZQrecebimento_historico_excluidos.close;
 }
      ZQcomprador_cessao.close;
      ZQaditamento.Close;
      ZQquitcao.Close;

     { ZQcomprador_cessao.MasterSource:=DS_Venda;
      ZQcomprador_cessao.MasterFields:='idvenda';
      ZQcomprador_cessao.LinkedFields:='idvenda';

      ZQquitcao.MasterSource:=DS_Venda;
      ZQquitcao.MasterFields:='idvenda';
      ZQquitcao.LinkedFields:='idvenda';

      ZQaditamento.MasterSource:=DS_Venda;
      ZQaditamento.MasterFields:='idvenda';
       ZQaditamento.LinkedFields:='idvenda';
      }



    //ate aqui
    end
    else
    begin
      if CBquitados.Checked=false then
      begin
        if Frm_RelVenda02=nil then
           Frm_RelVenda02:=TFrm_RelVenda02.Create(Application);
        Frm_RelVenda02.RLSubDetail3.Visible:=true;
        Frm_RelVenda02.RLSubDetail4.Visible:=true;
//        Frm_RelVenda02.RLSubDetail5.Visible:=true;

        ZQVenda.close;
        ZQVenda.SQL.Clear;
        ZQVenda.SQL.Add(' Select v.idvenda,v.datavenda,v.imovel,v.valorvenda,v.forma_reajuste,v.tabela_price,v.escriturado,v.marca,v.codigo_contrato_ref,v.Multa,v.Mora,v.Perc_comissao,v.vlr_comissao,');
        ZQVenda.SQL.Add('        i.idimovel as idimovel,i.loteamento_idloteamento,i.quadra,i.lote,i.lado,i.medidafrente,i.medidafundo,i.medidaesquerda,i.medidadedireita,i.area,i.confrontofrente,i.confrontofundo,i.confrontodireita,');
        ZQVenda.SQL.Add('        i.confrontoesquerda,i.esquinanome,i.esquinalinha,i.esquinacurva,i.valorvenal,i.Obervacaoloteamento,i.cornomapa,i.disponivel ');
        ZQVenda.SQL.Add(' from venda as v join imovel as i on i.idimovel=v.imovel');
        ZQVenda.SQL.Add(' where datavenda between :dt1 and :dt2');
        if not empty(VarImovel) Then
          ZQVenda.SQL.Add(' and idImovel in ('+VarImovel+')')
        else Begin
          if not empty(varParte) Then
            ZQVenda.SQL.Add(' and exists(select idcomprador,paticipante_idpaticipante,venda_idvenda,promissario,percentual,marcar from comprador  where paticipante_idpaticipante in ('+varParte+') and venda_idvenda=idvenda)');
          if not empty(varvendedor) Then
            ZQVenda.SQL.Add(' and exists(select idvendedor,corretor_idcorretor,venda_idvenda from vendedor  where corretor_idcorretor in ('+varvendedor+') and venda_idvenda=idvenda) order by quadra,lote');
        end;
        ZQVenda.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
        ZQVenda.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;
        ZQVenda.Open;
      end
      else
      begin
        if Frm_RelVenda02=nil then
           Frm_RelVenda02:=TFrm_RelVenda02.Create(Application);

        Frm_RelVenda02.RLSubDetail3.Visible:=false;
        Frm_RelVenda02.RLSubDetail4.Visible:=false;
        Frm_RelVenda02.RLSubDetail5.Visible:=false;


         // coloquei 28/11/2011
        ZQquitados.SQL.Clear;
        ZQquitados.SQL.Add('CREATE OR REPLACE VIEW `'+varschemata+'`.`relquitados` AS ');
        ZQquitados.SQL.Add(' Select v.idvenda,v.datavenda,v.multa as mult,v.mora as mor,v.imovel,v.valorvenda,v.forma_reajuste,v.tabela_Price,v.Escriturado,v.marca,v.codigo_contrato_ref,v.Perc_comissao,v.vlr_comissao,i.idimovel,i.loteamento_idloteamento,');
        ZQquitados.SQL.Add(' i.quadra,i.lote,i.lado,i.medidafrente,i.medidafundo,i.medidaesquerda,i.medidadedireita,i.area,i.confrontofrente,i.confrontofundo,i.confrontodireita,i.confrontoesquerda,i.esquinanome,');
        ZQquitados.SQL.Add(' i.esquinalinha,i.esquinacurva,i.valorvenal,i.Obervacaoloteamento,i.cornomapa,i.disponivel,i.matri,i.matricula,i.proposta,i.valorCusto,r.idrecebimento,r.documento,r.cliente,r.usuario,');
        ZQquitados.SQL.Add(' r.Dt_Entrada,r.Dt_Vencimento,r.Valor,r.Observ,r.VrDoc,r.ordem,r.TipDoc,r.saldo,r.marcar,r.RefBaixa,r.refvinda,r.contabil,r.empresa,r.custodaparcela,r.origem,r.adversa,r.recpag,r.numordem,');
        ZQquitados.SQL.Add(' r.idloteamento,r.venda_idvenda,r.quadralote,r.numboleto,r.Substituicao,r.sq,r.somar,r.nomeadversa,r.Reajustado,r.Data_reajuste,r.Proximo_Reajuste,r.Parcelas_fixas,r.juros,r.descontos,');
        ZQquitados.SQL.Add(' r.Data_Quitacao,r.observ_estorno,r.tip,r.sld_antes_reajuste,r.Percentual_reajuste,r.juridico,r.data_juridico,r.dt_nao_pagou_no_mes,r.descricao_juridico,r.multa,r.mora,');
        ZQquitados.SQL.Add(' c.paticipante_idpaticipante, pt.idpaticipante, pt.nome_parte,sum(r.saldo) as sld ');
        ZQquitados.SQL.Add(' from venda as v join imovel as i on i.idimovel=v.imovel join recebimento as r on r.venda_idvenda=v.idvenda ');

        ZQquitados.SQL.Add(' join comprador as c on c.venda_idvenda=v.idvenda join participante as pt on pt.idpaticipante=c.paticipante_idpaticipante ');

        ZQquitados.SQL.Add(' where v.datavenda between :dt1 and :dt2 ');
        if not empty(VarEmpree) Then
          ZQquitados.SQL.Add(' and r.idloteamento in ('+VarEmpree+')');

        if not empty(VarImovel) Then
          ZQquitados.SQL.Add(' and i.idImovel in ('+VarImovel+')')
        else Begin
          if not empty(varParte) Then
            ZQquitados.SQL.Add(' and exists(select idcomprador,paticipante_idpaticipante,venda_idvenda,promissario,percentual,marcar from comprador  where paticipante_idpaticipante in ('+varParte+') and venda_idvenda=idvenda)');
          if not empty(varvendedor) Then
            ZQquitados.SQL.Add(' and exists(select idvendedor,corretor_idcorretor,venda_idvenda from vendedor  where corretor_idcorretor in ('+varvendedor+') and venda_idvenda=idvenda) order by quadra,lote');
        end;
        ZQquitados.SQL.Add(' group by r.quadralote ');
        ZQquitados.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
        ZQquitados.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;
        ZQquitados.ExecSQL;

        ZQVenda.close;
        ZQVenda.SQL.clear;
        ZQVenda.SQL.Add('Select * from relquitados where sld=0 ');
        ZQVenda.Open;
         ZQVenda.RecordCount;
       // até aqui

        {    ZQVenda.close;
        ZQVenda.SQL.Clear;
        ZQVenda.SQL.Add(' Select * from venda as v join relquitados as q on q.idvenda=v.idvenda join imovel as i on i.idimovel=v.imovel join recebimento as r on r.venda_idvenda=v.idvenda ');
//      ZQVenda.SQL.Add(' where v.datavenda between :dt1 and :dt2 and r.substituicao=''Q''');
        ZQVenda.SQL.Add(' where v.datavenda between :dt1 and :dt2 ');
        if not empty(VarImovel) Then
          ZQVenda.SQL.Add(' and i.idImovel in ('+VarImovel+')')
        else Begin
          if not empty(varParte) Then
            ZQVenda.SQL.Add(' and exists(select * from comprador  where paticipante_idpaticipante in ('+varParte+') and venda_idvenda=idvenda)');
          if not empty(varvendedor) Then
            ZQVenda.SQL.Add(' and exists(select * from vendedor  where corretor_idcorretor in ('+varvendedor+') and venda_idvenda=idvenda) order by quadra,lote');
        end;
        ZQVenda.SQL.Add(' group by r.quadralote ');
        ZQVenda.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
        ZQVenda.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;
        ZQVenda.Open;}

      end;
      
      ZQEntrada.Open;
      ZQParcela.Open;

      ZQBxEntrada.close;
      ZQBxEntrada.SQL.Clear;
      ZQBxEntrada.SQL.Add('SELECT  *,  ');
      ZQBxEntrada.SQL.Add('        sum(valor) as vr, sum(descon) as de, sum(juros) as jr FROM recbxhist  where  exists (SELECT * FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ) group by idrecib');
      ZQBxEntrada.Open;


      ZQBxParcela.close;
      ZQBxParcela.SQL.Clear;
      ZQBxParcela.SQL.Add('SELECT  *,  ');
      ZQBxParcela.SQL.Add('        sum(valor) as vr, sum(descon) as de, sum(juros) as jr FROM recbxhist  ');
      ZQBxParcela.SQL.Add(' where  exists (SELECT * FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ) group by idrecib ');
      ZQBxParcela.Open;

      ZQVendedor.Open;

      ZQParticip.Open;

      ZQcheque_rec_p.open;
      ZQcheque_rec_e.Open;
      baixa.open;



      if CBtotais.Checked=false then
      begin
        ZQcomprador_cessao.MasterSource:=nil;
        ZQcomprador_cessao.MasterFields:='';
        ZQcomprador_cessao.LinkedFields:='';

        ZQquitcao.MasterSource:=nil;
        ZQquitcao.MasterFields:='';
        ZQquitcao.LinkedFields:='';

        ZQaditamento.MasterSource:=nil;
        ZQaditamento.MasterFields:='';
        ZQaditamento.LinkedFields:='';

        Frm_RelVenda02.Cabecalho.close;
        Frm_RelVenda02.Cabecalho.CreateDataSet;
        Frm_RelVenda02.Cabecalho2.close;
        Frm_RelVenda02.Cabecalho2.CreateDataSet;
        ZQaditamento.Close;

        ZQquitcao.close;
        ZQcomprador_cessao.close;
        ZQBalao.close;


        Gauge1.Visible:=true;
        Gauge1.Progress:=0;
        application.ProcessMessages;
        if CBquitados.Checked=false then
        begin

          ZQcomprador_cessao.close;
          ZQcomprador_cessao.SQL.Clear;
          ZQcomprador_cessao.SQL.add('Select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora_cessao from comprador_cessao where idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text));
          ZQcomprador_cessao.Open;


          Gauge1.MaxValue:=ZQcomprador_cessao.RecordCount;
          application.ProcessMessages;


          ZQcomprador_cessao.First;
          ZQcomprador_cessao.DisableControls;
          while not ZQcomprador_cessao.eof do
          begin
            gauge1.Progress:=ZQcomprador_cessao.RecNo;
            application.ProcessMessages;
            Frm_RelVenda02.Cabecalho2.Insert;
            Frm_RelVenda02.Cabecalho2idvenda.Value:=ZQcomprador_cessao.FieldByName('idvenda').AsLargeInt;
            Frm_RelVenda02.Cabecalho2datavenda.Value:=ZQcomprador_cessao.FieldByName('datavenda').AsDateTime;
            Frm_RelVenda02.Cabecalho2data.Value:=ZQcomprador_cessao.FieldByName('datacessao').AsDateTime;
            Frm_RelVenda02.Cabecalho2imovel.Value:=ZQcomprador_cessao.FieldByName('imovel').AsLargeInt;
            Frm_RelVenda02.Cabecalho2valorvenda.Value:=ZQcomprador_cessao.FieldByName('valorvenda').AsFloat;
            Frm_RelVenda02.Cabecalho2idcomprador.Value:=ZQcomprador_cessao.FieldByName('idcomprador').AsLargeInt;
            Frm_RelVenda02.Cabecalho2idparticipante.Value:=ZQcomprador_cessao.FieldByName('idparticipante').AsLargeInt;
            Frm_RelVenda02.Cabecalho2nome_parte.Value:=ZQcomprador_cessao.FieldByName('nome_parte').AsString;
            Frm_RelVenda02.Cabecalho2historico.Value:=ZQcomprador_cessao.FieldByName('historico').AsString;
            Frm_RelVenda02.Cabecalho2hora.Value:=ZQcomprador_cessao.FieldByName('hora_cessao').AsDateTime;
            Frm_RelVenda02.Cabecalho2.post;
            ZQcomprador_cessao.Next;
          end;
          ZQcomprador_cessao.First;
          ZQcomprador_cessao.EnableControls;
          Gauge1.Visible:=true;
          Gauge1.Progress:=0;
          application.ProcessMessages;
          ZQaditamento.close;
          ZQaditamento.SQL.Clear;
          ZQaditamento.SQL.Add('select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora from aditamento where idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text));
          ZQaditamento.open;

          Gauge1.MaxValue:=ZQaditamento.RecordCount;

          ZQaditamento.First;
          ZQaditamento.DisableControls;
          while not ZQaditamento.eof do
          begin
              gauge1.Progress:=ZQaditamento.RecNo;
              application.ProcessMessages;
              Frm_RelVenda02.Cabecalho2.Insert;
              Frm_RelVenda02.Cabecalho2idvenda.Value:=zqaditamento.FieldByName('idvenda').AsLargeInt;
//            Frm_RelVenda02.Cabecalho2data.Value:=ZQaditamentodatacessao.Value;
              Frm_RelVenda02.Cabecalho2historico.Value:=zqaditamento.FieldByName('historico').AsString;
//            Frm_RelVenda02.Cabecalho2hora.Value:=ZQaditamentohora.Value;
              Frm_RelVenda02.Cabecalho2.post;
              ZQaditamento.Next;
          end;
          ZQaditamento.First;
          ZQaditamento.EnableControls;
          ZQquitcao.close;
          ZQquitcao.SQL.Clear;
          ZQquitcao.SQL.Add('select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora from quitacao where idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text));
          ZQquitcao.open;

          ZQquitcao.First;
          Gauge1.Visible:=true;
          Gauge1.Progress:=0;
          Gauge1.MaxValue:=ZQquitcao.RecordCount;
          application.ProcessMessages;
          ZQquitcao.DisableControls;
          while not ZQquitcao.eof do
          begin
            gauge1.Progress:=ZQquitcao.RecNo;
            application.ProcessMessages;
            Frm_RelVenda02.Cabecalho2.Insert;
            Frm_RelVenda02.Cabecalho2idvenda.Value:=ZQquitcao.FieldByName('idvenda').AsLargeInt;
            Frm_RelVenda02.Cabecalho2data.Value:=ZQquitcao.FieldByName('datacessao').AsDateTime;
            Frm_RelVenda02.Cabecalho2historico.Value:=ZQquitcao.FieldByName('historico').AsString;
            Frm_RelVenda02.Cabecalho2hora.Value:=ZQquitcao.FieldByName('hora').AsDateTime;
            Frm_RelVenda02.Cabecalho2.post;
            ZQquitcao.Next;
          end;
          ZQquitcao.First;
          ZQquitcao.EnableControls;
          ZQBalao.close;
          ZQBalao.SQL.Clear;
          ZQBalao.SQL.Add('select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora from balao where idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text));
          ZQBalao.open;

          ZQBalao.First;
          Gauge1.Visible:=true;
          Gauge1.Progress:=0;
          Gauge1.MaxValue:=ZQBalao.RecordCount;
          application.ProcessMessages;
          ZQBalao.DisableControls;
          while not ZQBalao.eof do
          begin
            gauge1.Progress:=ZQBalao.RecNo;
            application.ProcessMessages;
            Frm_RelVenda02.Cabecalho2.Insert;
            Frm_RelVenda02.Cabecalho2idvenda.Value:=ZQBalao.FieldByName('idvenda').AsLargeInt;
  //          Frm_RelVenda02.Cabecalho2data.Value:=ZQBalaodatacessao.Value;
            Frm_RelVenda02.Cabecalho2historico.Value:=ZQBalao.FieldByName('historico').AsString;
    //        Frm_RelVenda02.Cabecalho2hora.Value:=ZQBalaohora.Value;
            Frm_RelVenda02.Cabecalho2.post;
            ZQBalao.Next;
          end;
          ZQBalao.First;
          ZQBalao.EnableControls;
        end
        else
        begin
          ZQquitcao.close;
          ZQquitcao.SQL.Clear;
          ZQquitcao.SQL.Add('select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora from quitacao where idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text));
          ZQquitcao.open;

          ZQquitcao.First;
          Gauge1.Visible:=true;
          Gauge1.Progress:=0;
          Gauge1.MaxValue:=ZQquitcao.RecordCount;
          application.ProcessMessages;
          ZQquitcao.DisableControls;
          while not ZQquitcao.eof do
          begin
            application.ProcessMessages;
            gauge1.Progress:=ZQquitcao.RecNo;
            Frm_RelVenda02.Cabecalho2.Insert;
            Frm_RelVenda02.Cabecalho2idvenda.Value:=ZQquitcao.FieldByName('idvenda').AsLargeInt;
            Frm_RelVenda02.Cabecalho2data.Value:=ZQquitcao.FieldByName('datacessao').AsDateTime;
            Frm_RelVenda02.Cabecalho2historico.Value:=ZQquitcao.FieldByName('historico').AsString;
            Frm_RelVenda02.Cabecalho2hora.Value:=ZQquitcao.FieldByName('hora').AsDateTime;
            Frm_RelVenda02.Cabecalho2.post;
            ZQquitcao.Next;
          end;
          ZQquitcao.First;
          ZQquitcao.EnableControls;

          ZQBalao.close;
          ZQBalao.SQL.Clear;
          ZQBalao.SQL.Add('select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora from balao where idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text));
          ZQBalao.open;

          ZQBalao.First;
          Gauge1.Visible:=true;
          Gauge1.Progress:=0;
          Gauge1.MaxValue:=ZQBalao.RecordCount;
          application.ProcessMessages;
          ZQBalao.DisableControls;
          while not ZQBalao.eof do
          begin
            gauge1.Progress:=ZQBalao.RecNo;
            application.ProcessMessages;
            Frm_RelVenda02.Cabecalho2.Insert;
            Frm_RelVenda02.Cabecalho2idvenda.Value:=ZQBalao.FieldByName('idvenda').AsLargeInt;
  //          Frm_RelVenda02.Cabecalho2data.Value:=ZQBalaodatacessao.Value;
            Frm_RelVenda02.Cabecalho2historico.Value:=ZQBalao.FieldByName('historico').AsString;
    //        Frm_RelVenda02.Cabecalho2hora.Value:=ZQBalaohora.Value;
            Frm_RelVenda02.Cabecalho2.post;
            ZQBalao.Next;
          end;
          ZQBalao.First;
          ZQBalao.EnableControls;

        end;
        Frm_RelVenda02.Cabecalho2.IndexName:='dataidx';
        Frm_RelVenda02.Cabecalho2.First;
        Gauge1.Visible:=true;
        Gauge1.Progress:=0;
        Gauge1.MaxValue:=Frm_RelVenda02.Cabecalho2.RecordCount;
        application.ProcessMessages;

        if Frm_RelVenda02.Cabecalho2.RecordCount>0 then
        begin
          Frm_RelVenda02.Cabecalho2.DisableControls;
          while not Frm_RelVenda02.Cabecalho2.eof do
          begin
            gauge1.Progress:=Frm_RelVenda02.Cabecalho2.RecNo;
            application.ProcessMessages;
            Frm_RelVenda02.Cabecalho.Insert;
            Frm_RelVenda02.Cabecalhoidvenda.Value:=Frm_RelVenda02.Cabecalho2idvenda.Value;
            Frm_RelVenda02.Cabecalhodatavenda.Value:=Frm_RelVenda02.Cabecalho2datavenda.Value;
            Frm_RelVenda02.Cabecalhodata.Value:=Frm_RelVenda02.Cabecalho2data.Value;
            Frm_RelVenda02.Cabecalhoimovel.Value:=Frm_RelVenda02.Cabecalho2imovel.Value;
            Frm_RelVenda02.Cabecalhovalorvenda.Value:=Frm_RelVenda02.Cabecalho2valorvenda.Value;
            Frm_RelVenda02.Cabecalhoidcomprador.Value:=Frm_RelVenda02.Cabecalho2idcomprador.Value;
            Frm_RelVenda02.Cabecalhoidparticipante.Value:=Frm_RelVenda02.Cabecalho2idparticipante.Value;
            Frm_RelVenda02.Cabecalhonome_parte.Value:=Frm_RelVenda02.Cabecalho2nome_parte.Value;
            Frm_RelVenda02.Cabecalhohistorico.Value:=Frm_RelVenda02.Cabecalho2historico.Value;
            Frm_RelVenda02.Cabecalhohora.Value:=Frm_RelVenda02.Cabecalho2hora.Value;
            Frm_RelVenda02.Cabecalho.post;
            Frm_RelVenda02.Cabecalho2.Next;
          end;
          Frm_RelVenda02.Cabecalho2.EnableControls;
        end;
        Gauge1.Progress:=0;
        Gauge1.Visible:=false;
        application.ProcessMessages;
{        Frm_RelVenda02.Cabecalho.MasterSource:=Frm_RelVenda.DS_Venda;
        Frm_RelVenda02.Cabecalho.MasterFields:='IdVenda';
        Frm_RelVenda02.Cabecalho.IndexFieldNames:='idvenda';
        Frm_RelVenda02.Cabecalho.First;
 }


        if FRBS.Checked Then Begin
          Frm_RelVenda02.RLSubDetail3.Visible := False;
          Frm_RelVenda02.RLSubDetail4.Visible := False;
        end
        else Begin
          Frm_RelVenda02.RLSubDetail3.Visible := True;
          Frm_RelVenda02.RLSubDetail4.Visible := True;
        end;
        if CBquitados.Checked=false then
        begin
          Frm_RelVenda02.RLSubDetail3.Visible:=true;
          Frm_RelVenda02.RLSubDetail4.Visible:=true;
//          Frm_RelVenda02.RLSubDetail5.Visible:=true;
        end
        else
        begin
          Frm_RelVenda02.RLSubDetail3.Visible:=false;
          Frm_RelVenda02.RLSubDetail4.Visible:=false;
          Frm_RelVenda02.RLSubDetail5.Visible:=false;
        end;


        if ZQParticip.FieldByName('judicial').AsString='SIM' then
        begin
          Frm_RelVenda02.RLLabel11.Visible:=true;
          Frm_RelVenda02.RLDBText51.Visible:=true;
          Frm_RelVenda02.RLLabel33.Visible:=true;
          Frm_RelVenda02.RLDBText57.Visible:=true;
          Frm_RelVenda02.RLLabel65.Visible:=true;
          Frm_RelVenda02.RLDBText58.Visible:=true;
          Frm_RelVenda02.RLLabel66.Visible:=true;
          Frm_RelVenda02.RLDBText59.Visible:=true;
          Frm_RelVenda02.RLLabel67.Visible:=true;
          Frm_RelVenda02.RLDBText60.Visible:=true;

        end
        else
        begin
          Frm_RelVenda02.RLLabel11.Visible:=false;
          Frm_RelVenda02.RLDBText51.Visible:=false;
          Frm_RelVenda02.RLLabel33.Visible:=false;
          Frm_RelVenda02.RLDBText57.Visible:=false;
          Frm_RelVenda02.RLLabel65.Visible:=false;
          Frm_RelVenda02.RLDBText58.Visible:=false;
          Frm_RelVenda02.RLLabel66.Visible:=false;
          Frm_RelVenda02.RLDBText59.Visible:=false;
          Frm_RelVenda02.RLLabel67.Visible:=false;
          Frm_RelVenda02.RLDBText60.Visible:=false;

        end;
        if not frbs.Checked then
        begin
          if Frm_RelVenda02=nil then
             Frm_RelVenda02:=TFrm_RelVenda02.Create(Application);
          Frm_RelVenda02.RLReport1.PreviewModal;
          Frm_RelVenda02:=nil;
        end
        else
        begin
          if CBquitados.Checked=true then
          begin
            ZQVenda2.close;
            ZQVenda2.SQL.clear;
            ZQVenda2.SQL.Add('Select * from relquitados where sld=0 ');
            ZQVenda2.Open;
            ZQVenda2.RecordCount;
            ZQVenda2.first;

            if FrmRelReceb02_quitado2=nil then
               FrmRelReceb02_quitado2:=TFrmRelReceb02_quitado2.Create(Application);
            FrmRelReceb02_quitado2.cds_recebimento.Close;
            FrmRelReceb02_quitado2.cds_recebimento.CreateDataSet;
            FrmRelReceb02_quitado2.cds_recebimento.Open;
            Gauge1.Visible:=true;
            Gauge1.Progress:=0;
            Gauge1.MaxValue:=ZQVenda2.RecordCount;
            ZQvenda2.DisableControls;
            while not ZQVenda2.Eof do
            begin
              ZQquitcao.close;
              ZQquitcao.SQL.Clear;
              ZQquitcao.SQL.Add('select idvenda,datavenda,datacessao,imovel,valorvenda,idcomprador,idparticipante,nome_parte,historico,hora from quitacao where idvenda='+quotedstr(ZQVenda2.FieldByName('idvenda').Text));
              ZQquitcao.open;
              ZQquitcao.First;

              Gauge1.Progress:=ZQVenda2.RecNo;
              FrmRelReceb02_quitado2.cds_recebimento.Insert;
              FrmRelReceb02_quitado2.cds_recebimentoidrecebimento_proposta.Value:=ZQVenda2.FieldByName('idrecebimento').AsLargeInt;
              FrmRelReceb02_quitado2.cds_recebimentoadversa.Value:=ZQVenda2.FieldByName('idpaticipante').AsLargeInt;
              FrmRelReceb02_quitado2.cds_recebimentoidloteamento.Value:=ZQVenda2.FieldByName('idloteamento').AsInteger;
              FrmRelReceb02_quitado2.cds_recebimentovenda_idvenda.Value:=ZQVenda2.FieldByName('idvenda').AsLargeInt;
              FrmRelReceb02_quitado2.cds_recebimentoData_Quitacao.Value:=ZQquitcao.FieldByName('datacessao').AsDateTime; //ZQVenda2Data_Quitacao.Value;
              FrmRelReceb02_quitado2.cds_recebimentoquadralote.Value:=ZQVenda2.FieldByName('quadralote').AsString;
              FrmRelReceb02_quitado2.cds_recebimentosld.Value:=ZQVenda2.FieldByName('sld').AsFloat;
              FrmRelReceb02_quitado2.cds_recebimentoValor.Value:=ZQVenda2.FieldByName('valorvenda').AsFloat;
              FrmRelReceb02_quitado2.cds_recebimentoDt_Entrada.Value:=ZQVenda2.FieldByName('datavenda').AsDateTime;
              FrmRelReceb02_quitado2.cds_recebimentoimovel.Value:=ZQVenda2.FieldByName('imovel').AsLargeInt;
              FrmRelReceb02_quitado2.cds_recebimentodocumento.Value:=ZQVenda2.FieldByName('documento').AsString;
              FrmRelReceb02_quitado2.cds_recebimentocliente.Value:=ZQVenda2.FieldByName('cliente').AsInteger;
              FrmRelReceb02_quitado2.cds_recebimentoHistorico.Value:=ZQquitcao.FieldByName('historico').AsString;
              FrmRelReceb02_quitado2.cds_recebimento.Post;
              ZQVenda2.next;
            end;
            Gauge1.Progress:=0;
            Gauge1.Visible:=false;
            ZQVenda2.EnableControls;
            FrmRelReceb02_quitado2.RLReport1.PreviewModal;
            FrmRelReceb02_quitado2:=nil;
            ZQVenda2.close;

          end
          else
          begin
            VarImovel := '';
            CDSQuadraLote.First;
            while not CDSQuadraLote.Eof do begin
              ql:=CDSQuadraLotequadra.Text+'-'+CDSQuadraLotelote.Text;
              Varimovel := VarImovel + quotedstr(ql)+',';
              CDSQuadraLote.Next;
            end;
            Varimovel := copy(VarImovel,1,length(VarImovel)-1);

            VarEmpree:= '';
            varcontador := 0;
            for x := 0 to CLBEmpree.Count-1 do begin
                if CLBEmpree.Checked[x] Then Begin
                    inc(varcontador);
                    for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
                        if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
                         break;
                    VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
                end;
            end;
            if varcontador<>CLBEmpree.Count Then
               VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
            else
               VarEmpree := '';
            ZQrelquitados.Close;
            ZQrelquitados.SQL.clear;

            ZQrelquitados.SQL.Add(' SELECT rb.idrecebimento,rb.documento,rb.cliente,rb.usuario,rb.Dt_Entrada,rb.Dt_Vencimento,rb.Valor,rb.Observ,rb.VrDoc,rb.ordem,rb.TipDoc,rb.saldo,rb.marcar,');
            ZQrelquitados.SQL.Add('        rb.RefBaixa,rb.refvinda,rb.contabil,rb.empresa,rb.custodaparcela,rb.origem,rb.adversa,rb.recpag,rb.numordem,rb.idloteamento,rb.venda_idvenda,rb.quadralote, rb.juros,rb.descontos,rb.Data_Quitacao,');
            ZQrelquitados.SQL.Add('        rb.numboleto,rb.Substituicao,rb.sq,rb.somar,rb.nomeadversa,rb.Reajustado,rb.Data_reajuste,rb.Proximo_Reajuste,rb.Parcelas_fixas,');
            ZQrelquitados.SQL.Add('        RH.idrecbxhist,RH.refer,RH.idrecib,RH.valor,RH.descon,RH.juros,RH.percent_usado,RH.data,RH.sq,RH.Valor_parcela,');
            ZQrelquitados.SQL.Add('        vd.valorvenda,vd.idvenda,vd.datavenda,vd.imovel,vd.forma_reajuste,vd.tabela_Price,vd.Escriturado,vd.marca,');
            ZQrelquitados.SQL.Add('        lt.idloteamento,lt.cidade_idcidade,lt.matriculaloteamento,lt.datacadastro,lt.datainalguracao,lt.dataconclusao,lt.logotipo,lt.mapa,lt.testemunha1,lt.testemunha2,');
            ZQrelquitados.SQL.Add('        lt.testemunha3,lt.testemunha4,lt.Observincorp,lt.bairro,lt.pasta_mapa,lt.nomeloteamento,lt.tipo,lt.codcontabancaria,lt.apelido,lt.digito_dif,');
            ZQrelquitados.SQL.Add('        pt.idpaticipante,pt.nome_parte,pt.doc1,pt.doc2,pt.endereco,pt.bairro,pt.cidade,pt.cep,pt.ende_cob,pt.bairro_cob,pt.cidade_cob,pt.cep_cob,pt.Fone1,');
            ZQrelquitados.SQL.Add('        pt.fone2,pt.fone3,pt.tipopessoa,pt.aniversario,pt.nacionalidade,pt.email,pt.naturalidade,pt.localdetrab,pt.profissao,pt.estadocivil,pt.renda,pt.observacao,');
            ZQrelquitados.SQL.Add('        pt.codrecinc,pt.codrecBx,pt.codpaginc,pt.codpagBx,pt.fone4,pt.complemento,pt.complemento_cob,pt.cadastrado,pt.Nome_Firma,pt.Nome_Firma,');
            ZQrelquitados.SQL.Add('        pt.Inativo,pt.sexo,pt.cc,pt.cex,pt.jb,pt.judicial,pt.xquadra,pt.xlote ');
            ZQrelquitados.SQL.Add('  FROM recebimento as rb join recbxhist as rh on rh.idrecib=rb.idrecebimento ');
            ZQrelquitados.SQL.Add('  join venda as vd on rb.venda_idvenda=vd.idvenda join loteamento as lt on rb.idloteamento=lt.idloteamento');
           // ZQrelquitados.SQL.Add('  join  participante as pt on rb.adversa=pt.idpaticipante');

            ZQquitados.SQL.Add(' join comprador as c on c.venda_idvenda=rb.venda_idvenda join participante as pt on pt.idpaticipante=c.paticipante_idpaticipante ');

//           ZQrelquitados.SQL.Add('   where rb.saldo=0 and vd.datavenda between :dt1 and :dt2 ');
            ZQrelquitados.SQL.Add('   where vd.datavenda between :dt1 and :dt2 ');

            if not empty(VarEmpree) Then
              ZQrelquitados.SQL.Add(' and rb.idloteamento in ('+VarEmpree+')');

            if not empty(VarImovel) Then
               ZQrelquitados.SQL.Add(' and rb.quadralote in ('+VarImovel+')');

            ZQrelquitados.SQL.Add(' group by rb.quadralote order by lt.apelido,pt.nome_parte');
            ZQrelquitados.ParamByName('dt1').AsDate:=XDEEntradaInicio.DateValue;
            ZQrelquitados.ParamByName('dt2').AsDate:=XDEEntradaFinal.DateValue;
            ZQrelquitados.Open;
            ZQrelquitados.First;
           // FrmRelReceb02_quitado.RLLabel2.Caption:='Período de '+XDEEntradaInicio.Text + ' até '+XDEEntradaFinal.Text;
            if FrmRelReceb02_quitado=nil then
               FrmRelReceb02_quitado:=TFrmRelReceb02_quitado.Create(Application);

            FrmRelReceb02_quitado.RLReport1.PreviewModal;
            FrmRelReceb02_quitado:=nil;
            ZQrelquitados.close;
          end;
        end;
       ZQcomprador_cessao.close;
       ZQaditamento.Close;
       ZQquitcao.Close;
       ZQBalao.close;
    // Frm_RelVenda02.RLReport1.Print;

       {ZQcomprador_cessao.MasterSource:=DS_Venda;
       ZQcomprador_cessao.MasterFields:='idvenda';
       ZQcomprador_cessao.LinkedFields:='idvenda';
       ZQquitcao.MasterSource:=DS_Venda;
       ZQquitcao.MasterFields:='idvenda';
       ZQquitcao.LinkedFields:='idvenda';

        ZQaditamento.MasterSource:=DS_Venda;
        ZQaditamento.MasterFields:='idvenda';
        ZQaditamento.LinkedFields:='idvenda';}

      end
      else
      begin

        montarelatorio.close;
        montarelatorio.CreateDataSet;
        ZQVenda.First;
        Gauge1.Visible:=true;
        Gauge1.Progress:=0;
        application.ProcessMessages;

         // 27/01/2014
        { Only the analytic report reads the staging table later through
          tempmensal.  The administrative report uses Cvenda directly, so it
          must not create or populate the obsolete TempTable. }
        if CBanalitico.Checked then
        begin
          TempTableName := varschemata+'.temp';
          ZQTemp.close;
          ZQTemp.sql.clear;
          ZQTemp.sql.add('DROP TABLE IF EXISTS '+TempTableName);
          ZQTemp.execsql;
          ZQTemp.close;
          ZQTemp.sql.clear;
          ZQTemp.sql.add('CREATE TABLE '+TempTableName+' ( ');
        end;
        if CBanalitico.Checked then
        begin
          ZQTemp.sql.add(' idvenda int(10) unsigned,venda_idvenda int(10) unsigned,datavenda date,imovel int(10) unsigned,valorvenda double,idimovel int(10) unsigned, idloteamento int(10) unsigned,');
          ZQTemp.sql.add(' quadra varchar(25),lote varchar(4),valorvenal double, paticipante_idpaticipante int(10) unsigned, idpaticipante int(10) unsigned,');
          ZQTemp.sql.add(' nome_parte VARCHAR(140),vr_entrada double,vr_parcela double,juros double,descontos double,juros_p double,');
          ZQTemp.sql.add(' descontos_p double) ENGINE=MEMORY');
          ZQTemp.execsql;
        end;

        ZQVenda.DisableControls;
        while not ZQVenda.Eof do
        begin
          application.ProcessMessages;
          Gauge1.Progress:=ZQVenda.RecNo * 100 div ZQVenda.RecordCount;
          ZQrecebentrada.close;
          ZQrecebentrada.sql.clear;
//          ZQrecebentrada.sql.add('Select cliente,RefBaixa,idloteamento,quadralote,Substituicao,idrecebimento,venda_idvenda,documento,Valor,ordem,Dt_Vencimento,TipDoc, sum(valor) as vr from recebimento where documento like  ''%-E-%'' and venda_idvenda='+quotedstr(ZQVendaidvenda.text)+' group by documento' );
          ZQrecebentrada.sql.add('Select cliente,refbaixa,venda_idvenda,documento,Valor, sum(valor) as vr from recebimento where documento like  ''%-E-%'' and venda_idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').text)+' group by documento' );
          ZQrecebentrada.open;


          ZQReceb_BaixaE.close;
          ZQReceb_BaixaE.sql.clear;
          ZQReceb_BaixaE.sql.add('Select Juros_Vr,Desc_Vr,refbaixa from Receb_Baixa where refbaixa='+quotedstr(ZQrecebentrada.FieldByName('RefBaixa').Text));
          ZQReceb_BaixaE.open;



          ZQrecebparcela.close;
          ZQrecebparcela.sql.clear;
//          ZQrecebparcela.sql.add('Select cliente,RefBaixa,idloteamento,quadralote,Substituicao,idrecebimento,venda_idvenda,documento,Valor,ordem,Dt_Vencimento,TipDoc, sum(valor) as vr from recebimento where documento like  ''%-P-%'' and venda_idvenda='+quotedstr(ZQVendaidvenda.text)+' group by documento');
          ZQrecebparcela.sql.add('Select cliente,refbaixa,venda_idvenda,documento,Valor, sum(valor) as vr from recebimento where documento like  ''%-P-%'' and venda_idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').text)+' group by documento');
          ZQrecebparcela.open;


          ZQReceb_BaixaP.close;
          ZQReceb_BaixaP.sql.clear;
          ZQReceb_BaixaP.sql.add('Select Juros_Vr,Desc_Vr,refbaixa from Receb_Baixa where refbaixa='+quotedstr(ZQrecebparcela.FieldByName('RefBaixa').Text));
          ZQReceb_BaixaP.open;
          // coloquei no dia 21/08/2019
          ZQParticip.Close;
          ZQParticip.SQL.Clear;
          ZQParticip.SQL.Add('Select * from participante join comprador on paticipante_idpaticipante=idpaticipante where venda_idvenda='+quotedstr(ZQVenda.FieldByName('idvenda').Text)+' order by marcar desc');
          ZQParticip.Open;
          //////

          montarelatorio.Insert;
          montarelatorioidvenda.Value:=ZQVenda.FieldByName('idvenda').AsLargeInt;
          montarelatoriovenda_idvenda.Value:=ZQVenda.FieldByName('idvenda').AsLargeInt;
          montarelatoriodatavenda.Value:=ZQVenda.FieldByName('datavenda').AsDateTime;
          montarelatorioimovel.Value:=ZQVenda.FieldByName('imovel').AsLargeInt;
          montarelatoriovalorvenda.Value:=ZQVenda.FieldByName('valorvenda').AsFloat;
          montarelatorioidimovel.Value:=ZQVenda.FieldByName('imovel').AsLargeInt;
          montarelatorioidloteamento.Value:=ZQVenda.FieldByName('loteamento_idloteamento').AsLargeInt;
          montarelatorioquadra.Value:=ZQVenda.FieldByName('quadra').AsString;
          montarelatoriolote.Value:=ZQVenda.FieldByName('lote').AsString;
          montarelatoriovalorvenal.Value:=ZQVenda.FieldByName('valorvenal').AsFloat;
          montarelatoriopaticipante_idpaticipante.Value:=ZQParticip.FieldByName('idpaticipante').AsLargeInt;
          montarelatorioidpaticipante.Value:=ZQParticip.FieldByName('idpaticipante').AsLargeInt;
          montarelatorionome_parte.Value:=ZQParticip.FieldByName('nome_parte').AsString;
//        montarelatoriovr_entrada.Value:=(ZQrecebentradavr.Value+ZQReceb_BaixaEJuros_Vr.Value)-ZQReceb_BaixaEDesc_Vr.Value;
//        montarelatoriovr_parcela.Value:=(ZQrecebparcelavr.Value+ZQReceb_BaixaPJuros_Vr.Value)-ZQReceb_BaixaPDesc_Vr.Value;
          montarelatoriovr_entrada.Value:=ZQrecebentrada.FieldByName('vr').AsFloat;
          montarelatoriovr_parcela.Value:=ZQrecebparcela.FieldByName('vr').AsFloat;

          montarelatoriojuros.Value:=ZQReceb_BaixaE.FieldByName('Juros_Vr').AsFloat;
          montarelatoriodescontos.Value:=ZQReceb_BaixaE.FieldByName('Desc_Vr').AsFloat;
          montarelatoriojuros_p.Value:=ZQReceb_BaixaP.FieldByName('Juros_Vr').AsFloat;
          montarelatoriodescontos_p.Value:=ZQReceb_BaixaP.FieldByName('Desc_Vr').AsFloat;
          montarelatorio.post;
          ZQVenda.Next;
        end;
        ZQVenda.EnableControls;
        montarelatorio.First;
        Gauge1.Progress:=0;
        application.ProcessMessages;        
        Gauge1.Visible:=false;
        ZQReceb_BaixaE.close;
        ZQReceb_BaixaP.close;


        Cvenda.close;
        Cvenda.CreateDataSet;
        montarelatorio.First;
        Gauge1.Visible:=true;
        Gauge1.Progress:=0;
        application.ProcessMessages;
        montarelatorio.DisableControls;
        while not montarelatorio.Eof do
        begin
          Gauge1.Progress:=montarelatorio.RecNo * 100 div montarelatorio.RecordCount;
          application.ProcessMessages;

          if CBanalitico.Checked then
          begin
            ZQTemp2.close;
            ZQTemp2.SQL.Clear;
            ZQTemp2.sql.add('INSERT INTO '+TempTableName+' (idvenda,venda_idvenda,datavenda,imovel,valorvenda,idimovel,idloteamento,quadra,lote,');
            ZQTemp2.sql.add('valorvenal,paticipante_idpaticipante,idpaticipante,nome_parte,vr_entrada,vr_parcela,juros,');
            ZQTemp2.sql.add('descontos,juros_p,descontos_p) VALUES ');
            ZQTemp2.sql.add('('+quotedstr(montarelatorioidvenda.text)+','+quotedstr(montarelatoriovenda_idvenda.text)+','+quotedstr(FormatDateTime('yyyy-mm-dd',montarelatoriodatavenda.Value))+',');
            ZQTemp2.sql.add(quotedstr(montarelatorioimovel.text)+','+quotedstr(StringReplace(floattostr(montarelatoriovalorvenda.Value),',','.',[rfReplaceAll]))+','+quotedstr(montarelatorioidimovel.text)+',');
            ZQTemp2.sql.add(quotedstr(montarelatorioidloteamento.text)+','+quotedstr(montarelatorioquadra.Value)+','+quotedstr(montarelatoriolote.Value)+',');
            ZQTemp2.sql.add(quotedstr(StringReplace(floattostr(montarelatoriovalorvenal.Value),',','.',[rfReplaceAll]))+','+quotedstr(montarelatoriopaticipante_idpaticipante.text)+','+quotedstr(montarelatorioidpaticipante.text)+',');
            ZQTemp2.sql.add(quotedstr(montarelatorionome_parte.Value)+','+quotedstr(StringReplace(floattostr(montarelatoriovr_entrada.Value),',','.',[rfReplaceAll]))+','+quotedstr(StringReplace(floattostr(montarelatoriovr_parcela.Value),',','.',[rfReplaceAll]))+',');
            ZQTemp2.sql.add(quotedstr(StringReplace(floattostr(montarelatoriojuros.Value),',','.',[rfReplaceAll]))+','+quotedstr(StringReplace(floattostr(montarelatoriodescontos.Value),',','.',[rfReplaceAll]))+','+quotedstr(StringReplace(floattostr(montarelatoriojuros_p.Value),',','.',[rfReplaceAll]))+',');
            ZQTemp2.sql.add(quotedstr(StringReplace(floattostr(montarelatoriodescontos_p.Value),',','.',[rfReplaceAll]))+')');
            ZQTemp2.ExecSQL;
          end;



          Cvenda.Insert;
          Cvendaidvenda.Value:=montarelatorioidvenda.Value;
          Cvendavenda_idvenda.Value:=montarelatoriovenda_idvenda.Value;
          Cvendadatavenda.Value:=montarelatoriodatavenda.Value;
          Cvendaimovel.Value:=montarelatorioimovel.Value;
          Cvendavalorvenda.Value:=montarelatoriovalorvenda.Value;
          Cvendaidimovel.Value:=montarelatorioidimovel.Value;
          Cvendaidloteamento.Value:=montarelatorioidloteamento.Value;
          Cvendaquadra.Value:=montarelatorioquadra.Value;
          Cvendalote.Value:=montarelatoriolote.Value;
          Cvendavalorvenal.Value:=montarelatoriovalorvenal.Value;
          Cvendapaticipante_idpaticipante.Value:=montarelatoriopaticipante_idpaticipante.Value;
          Cvendaidpaticipante.Value:=montarelatorioidpaticipante.Value;
          Cvendanome_parte.Value:=montarelatorionome_parte.Value;
          Cvendavr_entrada.Value:=montarelatoriovr_entrada.Value;
          Cvendavr_parcela.Value:=montarelatoriovr_parcela.Value;
          Cvendajuros.Value:=montarelatoriojuros.Value;
          Cvendadescontos.Value:=montarelatoriodescontos.Value;
          Cvendajuros_p.Value:=montarelatoriojuros_p.Value;
          Cvendadescontos_p.Value:=montarelatoriodescontos_p.Value;
          Cvenda.post;
          montarelatorio.Next;
        end;
        montarelatorio.EnableControls;
        //28/01/2014
        if CBanalitico.Checked=true then
        begin
          ZQTemp3.close;
          ZQTemp3.SQL.Clear;
          //    ZQTemp3.SQL.Add('DROP VIEW IF EXISTS `'+varschemata+'`.`relmensal`;');
          ZQTemp3.SQL.Add('CREATE OR REPLACE VIEW `'+varschemata+'`.`tempmensal` AS ');
          ZQTemp3.SQL.Add(' ( select *,month(datavenda) as Mes, case month(datavenda)');
          ZQTemp3.SQL.Add( 'when 1 then ''Janeiro'' when 2 then ''Fevereiro'' when 3 then ''Março''');
          ZQTemp3.SQL.Add(' when 4 then ''Abril'' when 5 then ''Maio'' when 6 then ''Junho''');
          ZQTemp3.SQL.Add(' when 7 then ''Julho'' when 8 then ''Agosto'' when 9 then ''Setembro''');
          ZQTemp3.SQL.Add(' when 10 then ''Outubro'' when 11 then ''Novembro'' when 12 then ''Dezembro''');
          //    ZQTemp3.SQL.Add(' end AS MESDESC, CASE when (re.valor <> rb.vr_rec) then re.valor else rb.Vr_rec end as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr,RE.valor,RH.descon,RH.juros, ');
          ZQTemp3.SQL.Add(' end AS MESDESC,year(datavenda) as Ano');
          ZQTemp3.SQL.Add(' from '+varschemata+'.temp where 1=1');
          if not empty(VarEmpree) then
             ZQTemp3.SQL.Add(' and idloteamento in ('+VarEmpree+')');
          if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then
          begin
              ZQTemp3.SQL.Add(' and datavenda between :dt1 and :dt2');
              ZQTemp3.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
              ZQTemp3.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
          end;

  //        ZQTemp3.SQL.Add(' group by quadra,lote order by ano,mes');
          ZQTemp3.SQL.Add(' order by ano,mes )');
          //    ZQTemp3.Open;
          //  ZQTemp3.First;
          ZQTemp3.ExecSQL;


          ZQTemp4.close;
          ZQTemp4.sql.clear;
          ZQTemp4.sql.Add('select * from tempmensal group by quadra,lote order by ano,mes');
          ZQTemp4.Open;

          ZQEntrada.close;
          ZQParcela.close;
          ZQParcela.MasterSource:=DS_zqtemp4;
          ZQEntrada.MasterSource:=DS_zqtemp4;
          ZQEntrada.Open;
          ZQParcela.Open;


          if Frm_RelVenda02_totais_analitico=nil then
             Frm_RelVenda02_totais_analitico:=TFrm_RelVenda02_totais_analitico.Create(Application);
          Frm_RelVenda02_totais_analitico.inicial.Caption:=XDEEntradaInicio.Text;
          Frm_RelVenda02_totais_analitico.Final.Caption:=XDEEntradaFinal.Text;
          if CBanalitico2.Checked=false then
          begin
            Frm_RelVenda02_totais_analitico.RLSubDetail3.Visible:=false;
            Frm_RelVenda02_totais_analitico.RLSubDetail4.Visible:=false;
          end
          else
          begin
            Frm_RelVenda02_totais_analitico.RLSubDetail3.Visible:=true;
            Frm_RelVenda02_totais_analitico.RLSubDetail4.Visible:=true;
          end;

          Frm_RelVenda02_totais_analitico.RLReport1.PreviewModal;
          Frm_RelVenda02_totais_analitico:=nil;
          ZQEntrada.close;
          ZQParcela.close;
          ZQParcela.MasterSource:=DS_Venda;
          ZQEntrada.MasterSource:=DS_Venda;
          ZQEntrada.Open;
          ZQParcela.Open;
          cvenda.close;
          cvenda.CreateDataSet;

        end;
        Cvenda.First;
        Gauge1.Progress:=0;
        Gauge1.Visible:=false;
        montarelatorio.close;
        application.ProcessMessages;
        if Cvenda.recordcount>0 then
        begin
          if Frm_RelVenda02_totais=nil then
             Frm_RelVenda02_totais:=TFrm_RelVenda02_totais.Create(Application);
          Frm_RelVenda02_totais.inicial.Caption:=XDEEntradaInicio.Text;
          Frm_RelVenda02_totais.Final.Caption:=XDEEntradaFinal.Text;
          Frm_RelVenda02_totais.RLReport1.PreviewModal;
          Frm_RelVenda02_totais:=nil;
          //  Frm_RelVenda02_totais.RLReport1.Print;
        end
        else
        begin
          if CBanalitico.Checked=false then
             showmessage('Nenhuma Venda Nesse Período!');
        end;
      end;
      baixa.close;
      Cvenda.Close;
      ZQrecebentrada.Close;
      ZQrecebparcela.Close;




     CDSQuadraLote.First;
     CDSQuadraLote.DisableControls;
     while not CDSQuadraLote.Eof do
     begin
       CDSQuadraLote.Delete;
     end;
     CDSQuadraLote.EnableControls;
     DM_Tabelas.CDSVendedorTemp.First;
     DM_Tabelas.CDSVendedorTemp.DisableControls;
     while not DM_Tabelas.CDSVendedorTemp.Eof do
     begin
       DM_Tabelas.CDSVendedorTemp.Delete;
     end;
     DM_Tabelas.CDSVendedorTemp.EnableControls;
     DM_Tabelas.CDSCompradorTemp.First;
     DM_Tabelas.CDSCompradorTemp.DisableControls;
     while not DM_Tabelas.CDSCompradorTemp.Eof do
     begin
       DM_Tabelas.CDSCompradorTemp.Delete;
     end;
     DM_Tabelas.CDSCompradorTemp.EnableControls;
     if (CDSQuadraLote.RecordCount=0) and (DM_Tabelas.CDSVendedorTemp.RecordCount=0) then
       Comprador.Enabled:=true;

     if (CDSQuadraLote.RecordCount=0) and (DM_Tabelas.CDSCompradorTemp.RecordCount=0) then
       Corretor.Enabled:=true;

    end;
    DXBImprimir.Enabled:=true;
    GBEmpree.Enabled := True;
  end;
  DXBImprimir.Enabled:=true;
end;

procedure TFrm_RelVenda.FormShow(Sender: TObject);
Var
  VarPath : string;
  ArqIni2 : tIniFile;
  Task : Itask;
begin
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni2 := tIniFile.Create(varpath+'siai.Ini');
  try
    varschemata := ArqIni2.ReadString('BANCO_DE_DADOS','SCHEMATA', varschemata );
  finally
    ArqIni2.Free;
  end;
  DM_Tabelas.CDSVendedorTemp.Close;
  DM_Tabelas.CDSVendedorTemp.CreateDataSet;
  DM_Tabelas.CDSCompradorTemp.Close;
  DM_Tabelas.CDSCompradorTemp.CreateDataSet;
  CLBEmpree.Clear;
  DM_Tabelas.ZQLoteamento.close;
  if DM_Tabelas.ZQLoteamento.active=false then
     DM_Tabelas.ZQLoteamento.open;
  DM_Tabelas.ZQLoteamento.First;
//  Task := TTask.create(procedure
//                       begin
                          while not DM_Tabelas.ZQLoteamento.Eof do begin
                            CLBEmpree.Items.Add(DM_TAbelas.ZQLoteamento.FieldByName('apelido').AsString+' | '+DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Text);
                            CLBEmpree.Checked[CLBEmpree.Count-1] := True;
                            DM_Tabelas.ZQLoteamento.Next;
                          end;
//                       end);
//  Task.Start;
//  XDEEntradaInicio.DateValue := date;

//  XDEEntradaFinal.DateValue := date;
  CDSQuadraLote.Close;
  CDSQuadraLote.CreateDataSet;
  GBEmpree.Enabled := True;
end;

procedure TFrm_RelVenda.DXBMarcEmpreeClick(Sender: TObject);
Var
  x, varcontador : Integer;
begin
  varcontador := CLBEmpree.Count-1;
  if DXBMarcEmpree.Caption = 'Desmarcar' Then Begin
    for x := 0 to varcontador do
      CLBEmpree.Checked[x] := False;
    DXBMarcEmpree.Caption := 'Marcar';
  end
  Else
  Begin
    for x := 0 to varcontador do
      CLBEmpree.Checked[x] := True;
    DXBMarcEmpree.Caption := 'Desmarcar';
  End;
end;

procedure TFrm_RelVenda.XDEEntradaInicioExit(Sender: TObject);
begin
  if  empty(XDEEntradaInicio.Text) Then Begin
    XDEEntradaInicio.SetFocus;
    exit;
  end;
  if XDEEntradaInicio.DateValue>XDEEntradaFinal.DateValue Then
    XDEEntradaFinal.DateValue := XDEEntradaInicio.DateValue;
end;

procedure TFrm_RelVenda.XDEEntradaFinalExit(Sender: TObject);
begin
  if  empty(XDEEntradaFinal.Text) Then Begin
    XDEEntradaFinal.SetFocus;
    exit;
  end;
  if XDEEntradaInicio.DateValue>XDEEntradaFinal.DateValue Then
    XDEEntradaInicio.DateValue:=XDEEntradaFinal.DateValue;

end;

procedure TFrm_RelVenda.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    if not (ActiveControl is TDBGrid) then begin
      Key := #0;
      Perform(WM_NEXTDLGCTL, 0, 0);
    end;
end;

procedure TFrm_RelVenda.GroupBox3Enter(Sender: TObject);
begin
  Equadra.SetFocus;
end;

procedure TFrm_RelVenda.EloteExit(Sender: TObject);
begin
  if not empty(Elote.Text+EQuadra.Text) Then Begin
    ZQVen1.SQL.Clear;
    ZQVen1.SQL.Clear;
    ZQVen1.SQL.Add(' Select v.idvenda,v.datavenda,v.imovel,v.valorvenda,');
    ZQVen1.SQL.Add('        i.idimovel as idimovel,i.loteamento_idloteamento,i.quadra,i.lote,i.lado,i.medidafrente,i.medidafundo,i.medidaesquerda,i.medidadedireita,i.area,i.confrontofrente,i.confrontofundo,i.confrontodireita,');
    ZQVen1.SQL.Add('        i.confrontoesquerda,i.esquinanome,i.esquinalinha,i.esquinacurva,i.valorvenal,i.Obervacaoloteamento,i.cornomapa,i.disponivel ');
    ZQVen1.SQL.Add(' from venda as v join imovel as i on i.idimovel=v.imovel');
    ZQVen1.SQL.Add(' where quadra='+quotedstr(Equadra.Text)+' and lote='+quotedstr(Elote.Text));
    ZQVen1.Open;
    if ZQVen1.RecordCount>0 Then
    Begin
      CDSQuadraLote.Insert;
      CDSQuadraLoteidlote.Value := ZQVen1.FieldByName('idimovel').AsLargeInt;
      CDSQuadraLoteQuadra.Value := Equadra.Text;
      CDSQuadraLotelote.Value := Elote.Text;
      CDSQuadraLote.Post;
      if CBContabil.Checked=false then
      begin
        if ZQVen1.FieldByName('datavenda').AsDateTime<XDEEntradaInicio.DateValue Then
          XDEEntradaInicio.DateValue := ZQVen1.FieldByName('datavenda').AsDateTime;
        if ZQVen1.FieldByName('datavenda').AsDateTime>XDEEntradaFinal.DateValue Then
          XDEEntradaFinal.DateValue := ZQVen1.FieldByName('datavenda').AsDateTime;
      end;
    end
    else
      Showmessage('Esta Quadra e lote não foram vendidos!!!!');
    Equadra.Text := '';
    Elote.text := '';
    Equadra.SetFocus;
  end
  else
    if CDSQuadraLote.RecordCount>0 Then
      DXBImprimir.SetFocus;
end;

procedure TFrm_RelVenda.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSQuadraLote.RecordCount>0) Then Begin
    CDSQuadraLote.Delete;
    if CDSQuadraLote.RecordCount=0 Then Begin
      GBEmpree.Enabled := True;
      Comprador.Enabled := True;
      Corretor.Enabled := True;
    end;
  end;
end;

procedure TFrm_RelVenda.DBGrid3KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (DM_Tabelas.CDSCompradorTemp.RecordCount>0) Then Begin
    DM_Tabelas.CDSCompradorTemp.Delete;
  end;
end;

procedure TFrm_RelVenda.GroupBox3Exit(Sender: TObject);
begin
  if CDSQuadraLote.RecordCount>0 Then Begin
    GBEmpree.Enabled := False;
    Comprador.Enabled := False;
    Corretor.Enabled := False;
  end
  else Begin
    GBEmpree.Enabled := True;
    Comprador.Enabled := True;
    Corretor.Enabled := True;
  end;
end;

procedure TFrm_RelVenda.DBGrid2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (DM_Tabelas.CDSVendedorTemp.RecordCount>0) Then Begin
    DM_Tabelas.CDSVendedorTemp.Delete;
  end;
end;

procedure TFrm_RelVenda.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQdiaria.Close;
  ZQcheque_rec_p.close;
  ZQcheque_rec_e.close;
  ZQReceb_BaixaE.close;
  ZQReceb_BaixaP.close;
  ZQcomprador_cessao.close;
  ZQVen1.close;
  ZQEmpree.close;
  ZQVenda.close;
  ZQEntrada.close;
  ZQParcela.close;
  ZQParticip.close;
  ZQVendedor.close;
  ZQcheque_rec_e.close;
  ZQBxParcela.close;
  ZQBxEntrada.close;
  ZQrecebentrada.close;
  ZQrecebparcela.close;
  baixa.close;
  ZQcheque_rec_p.close;
  ZQaditamento.close;
  ZQquitcao.close;
  ZQBalao.close;
  ZQquitados.close;
  Frm_RelVenda:=nil;
  Action:=CaFree;
end;

procedure TFrm_RelVenda.ZQReceb_BaixaEAfterPost(DataSet: TDataSet);
begin
//  ZQReceb_BaixaE.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQReceb_BaixaPAfterPost(DataSet: TDataSet);
begin
//  ZQReceb_BaixaP.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQcomprador_cessaoAfterPost(DataSet: TDataSet);
begin
  ZQcomprador_cessao.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQVen1AfterPost(DataSet: TDataSet);
begin
  ZQVen1.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQEmpreeAfterPost(DataSet: TDataSet);
begin
  ZQEmpree.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQVendaAfterPost(DataSet: TDataSet);
begin
  ZQVenda.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQEntradaAfterPost(DataSet: TDataSet);
begin
 // ZQEntrada.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQParcelaAfterPost(DataSet: TDataSet);
begin
 // ZQParcela.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQParticipAfterPost(DataSet: TDataSet);
begin
  //ZQParticip.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQVendedorAfterPost(DataSet: TDataSet);
begin
  ZQVendedor.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQcheque_rec_eAfterPost(DataSet: TDataSet);
begin
  ZQcheque_rec_e.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQBxParcelaAfterPost(DataSet: TDataSet);
begin
  ZQBxParcela.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQBxEntradaAfterPost(DataSet: TDataSet);
begin
   ZQBxEntrada.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQrecebentradaAfterPost(DataSet: TDataSet);
begin
  ZQrecebentrada.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQrecebparcelaAfterPost(DataSet: TDataSet);
begin
  ZQrecebparcela.ApplyUpdates;
end;

procedure TFrm_RelVenda.baixaAfterPost(DataSet: TDataSet);
begin
  baixa.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQcheque_rec_pAfterPost(DataSet: TDataSet);
begin
  ZQcheque_rec_p.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQaditamentoAfterPost(DataSet: TDataSet);
begin
  ZQaditamento.ApplyUpdates;
end;

procedure TFrm_RelVenda.ZQquitcaoAfterPost(DataSet: TDataSet);
begin
  ZQquitcao.ApplyUpdates;
end;

procedure TFrm_RelVenda.CBtotaisClick(Sender: TObject);
begin
  if (CBquitados.Checked=true) or (CBDiarias.Checked=true) or (CBtotais.Checked=true)  or (CBContabil.Checked=true) or (cbcessao.Checked=true) or (cbrescisao.Checked=true) then
  begin
    XDEEntradaInicio.DateText:=datetostr(date);
    XDEEntradaFinal.DateText:=datetostr(date);
    GroupBox1.Visible:=true;
    FRBS.Visible:=false;
    if CBtotais.Checked=true then
    begin
       CBanalitico.Visible:=true;
    end
    else
    begin
       CBanalitico.Checked:=false;
       CBanalitico.Visible:=false;
       CBanalitico2.Checked:=false;
       CBanalitico2.Visible:=false;

    end;
  end
  else
  begin
    CBanalitico.Checked:=false;
    CBanalitico.Visible:=false;
    CBanalitico2.Checked:=false;
    CBanalitico2.Visible:=false;
    GroupBox1.Visible:=false;
    XDEEntradaInicio.DateText:='01/01/1900';
    XDEEntradaFinal.DateText:='31/12/2100';
    FRBS.Visible:=true;
  end;
end;

procedure TFrm_RelVenda.CBContabilClick(Sender: TObject);
begin
  if  (CBquitados.Checked=true) or (CBDiarias.Checked=true) or (CBtotais.Checked=true)  or (CBContabil.Checked=true) or (cbcessao.Checked=true) or (cbrescisao.Checked=true) then
  begin
    XDEEntradaInicio.DateText:=datetostr(date);
    XDEEntradaFinal.DateText:=datetostr(date);
    CBquitados.Checked:=false;
    CBquitados.Enabled:=false;
    CBtotais.Checked:=false;
    CBtotais.Enabled:=false;
 //   equadra.Enabled:=false;
 //   elote.Enabled:=false;
//    ecomprador.Enabled:=false;
    ecorretor.Enabled:=false;
    GroupBox1.Caption:='Data da Baixa';
    GroupBox1.Visible:=true;
    FRBS.Visible:=false;
//    GroupBox3.Enabled:=false;
//    Comprador.Enabled:=false;
    Corretor.Enabled:=false;
{    if CDSQuadraLote.RecordCount>0 then
    begin
      CDSQuadraLote.First;
      while not CDSQuadraLote.Eof do
      begin
        CDSQuadraLote.Delete;
      end;
    end;}
    if DM_Tabelas.CDSCompradorTemp.RecordCount>0 then
    begin
      DM_Tabelas.CDSCompradorTemp.First;
      DM_Tabelas.CDSCompradorTemp.DisableControls;
      while not DM_Tabelas.CDSCompradorTemp.Eof do
      begin
        DM_Tabelas.CDSCompradorTemp.Delete;
      end;
      DM_Tabelas.CDSCompradorTemp.EnableControls;
    end;

    if DM_Tabelas.CDSVendedorTemp.RecordCount>0 then
    begin
      DM_Tabelas.CDSVendedorTemp.First;
      DM_Tabelas.CDSVendedorTemp.DisableControls;
      while not DM_Tabelas.CDSVendedorTemp.Eof do
      begin
        DM_Tabelas.CDSVendedorTemp.Delete;
      end;
      DM_Tabelas.CDSVendedorTemp.EnableControls;
    end;
    XDEEntradaInicio.SetFocus;
  end
  else
  begin
    GroupBox1.Caption:='Entrada      ';
    GroupBox1.Visible:=false;
    XDEEntradaInicio.DateText:='01/01/1900';
    XDEEntradaFinal.DateText:='31/12/2100';
    FRBS.Visible:=true;
    CBquitados.Enabled:=true;
    CBquitados.Checked:=false;
    CBtotais.Enabled:=true;
    CBtotais.Checked:=false;
    equadra.Enabled:=true;
    elote.Enabled:=true;
    Comprador.Enabled:=true;
    Corretor.Enabled:=true;
    ecomprador.Enabled:=true;
    ecorretor.Enabled:=true;
    GroupBox3.Enabled:=true;
    DXBImprimir.SetFocus;
  end;

end;

procedure TFrm_RelVenda.cbcessaoClick(Sender: TObject);
begin
  if  (CBquitados.Checked=true) or (CBDiarias.Checked=true) or (CBtotais.Checked=true)  or (CBContabil.Checked=true) or (cbcessao.Checked=true) or (CBRescisao.Checked=true) then
  begin
    XDEEntradaInicio.DateText:=datetostr(date);
    XDEEntradaFinal.DateText:=datetostr(date);
    GroupBox1.Visible:=true;
    FRBS.Visible:=false;
  end
  else
  begin
    GroupBox1.Visible:=false;
    XDEEntradaInicio.DateText:='01/01/1900';
    XDEEntradaFinal.DateText:='31/12/2100';
    FRBS.Visible:=true;
  end;
end;

procedure TFrm_RelVenda.CBanaliticoClick(Sender: TObject);
begin
  if cbanalitico.Checked=true then
  begin
    CBanalitico2.Visible:=true;
  end
  else
  begin
    CBanalitico2.Checked:=false;
    CBanalitico2.Visible:=false;
  end;
end;

procedure TFrm_RelVenda.CBDiariasClick(Sender: TObject);
begin
  if  (CBquitados.Checked=true) or (CBDiarias.Checked=true) or (CBtotais.Checked=true)  or (CBContabil.Checked=true) or (cbcessao.Checked=true) or (cbrescisao.Checked=true) then
  begin
    XDEEntradaInicio.DateText:=datetostr(date);
    XDEEntradaFinal.DateText:=datetostr(date);
    GroupBox1.Visible:=true;
    GroupBox3.Enabled:=false;
    Comprador.Enabled:=false;
    FRBS.Visible:=false;
  end
  else
  begin
    GroupBox1.Visible:=false;
    GroupBox3.Enabled:=true;
    Comprador.Enabled:=true;
    XDEEntradaInicio.DateText:='01/01/1900';
    XDEEntradaFinal.DateText:='31/12/2100';
    FRBS.Visible:=true;
  end;

end;

procedure TFrm_RelVenda.CBquitadosClick(Sender: TObject);
begin
  if  (CBquitados.Checked=true) or (CBDiarias.Checked=true) or (CBtotais.Checked=true)  or (CBContabil.Checked=true) or (cbcessao.Checked=true) or (CBRescisao.Checked=true) then
  begin
    XDEEntradaInicio.DateText:=datetostr(date);
    XDEEntradaFinal.DateText:=datetostr(date);
    GroupBox1.Visible:=true;
    FRBS.Visible:=false;
  end
  else
  begin
    GroupBox1.Visible:=false;
    XDEEntradaInicio.DateText:='01/01/1900';
    XDEEntradaFinal.DateText:='31/12/2100';
    FRBS.Visible:=true;
  end;

  if (CBquitados.Checked=true) then
      FRBS.Visible:=true;

end;

procedure TFrm_RelVenda.CBRescisaoClick(Sender: TObject);
begin
  if  (CBquitados.Checked=true) or (CBDiarias.Checked=true) or (CBtotais.Checked=true)  or (CBContabil.Checked=true) or (cbcessao.Checked=true) or (CBRescisao.Checked=true) then
  begin
    XDEEntradaInicio.DateText:=datetostr(date);
    XDEEntradaFinal.DateText:=datetostr(date);
    GroupBox1.Visible:=true;
    FRBS.Visible:=false;
  end
  else
  begin
    GroupBox1.Visible:=false;
    XDEEntradaInicio.DateText:='01/01/1900';
    XDEEntradaFinal.DateText:='31/12/2100';
    FRBS.Visible:=true;
  end;

end;


procedure TFrm_RelVenda.AfterConstruction;
begin
  inherited AfterConstruction;
  { ZQVen1 is used by the quadra/lote filter.  Its design-time SQL used to
    request idimovel directly from venda, although that field belongs to
    imovel.  Prepare a valid metadata query before runtime fields are built. }
  ZQVen1.Close;
  ZQVen1.SQL.Text :=
    'select v.idvenda,v.datavenda,v.imovel,v.valorvenda,' +
    'i.idimovel as idimovel,i.loteamento_idloteamento,i.quadra,i.lote,' +
    'i.lado,i.medidafrente,i.medidafundo,i.medidaesquerda,i.medidadedireita,' +
    'i.area,i.confrontofrente,i.confrontofundo,i.confrontodireita,' +
    'i.confrontoesquerda,i.esquinanome,i.esquinalinha,i.esquinacurva,' +
    'i.valorvenal,i.Obervacaoloteamento,i.cornomapa,i.disponivel ' +
    'from venda as v join imovel as i on i.idimovel=v.imovel';
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQVenda', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendalote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendalado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendamedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendamedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendamedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendamedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendavalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendacornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendadisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendatabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendaEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendamarca', 'marca', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda', 'ZQVendacodigo_contrato_ref', 'codigo_contrato_ref', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQEmpree', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreeidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreecidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreematriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreedatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreedatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreedataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreelogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreemapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreetestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreetestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreetestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreetestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreeObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreebairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreepasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreenomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreetipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreeapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEmpree', 'ZQEmpreeQuadra', 'Quadra', TWideStringField, fkLookup, 30, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQImovel', 'loteamento_idloteamento', 'quadra', True);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQEntrada', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQEntrada', 'ZQEntradasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQParcela', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParcela', 'ZQParcelasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQParticip', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticiptipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticiplocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticiprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticippaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticippromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipsexo', 'sexo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipcex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipjb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipjudicial', 'judicial', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticippercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipxquadra', 'xquadra', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQParticip', 'ZQParticipxlote', 'xlote', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQBxEntrada', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradavr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradajr', 'jr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradade', 'de', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxEntrada', 'ZQBxEntradadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQVen1', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1quadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1lote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1lado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQVendedor', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedoridcorretor', 'idcorretor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedornomecorretor', 'nomecorretor', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedordoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedorcreci', 'creci', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedortelefone1', 'telefone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedortelefone2', 'telefone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedorendereco_corr', 'endereco_corr', TWideStringField, fkData, 70, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedorbairro_corr', 'bairro_corr', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedorcep_corr', 'cep_corr', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedorObser_corr', 'Obser_corr', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedorcodcidade', 'codcidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedoridvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedorcorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendedor', 'ZQVendedorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQBxParcela', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParcelavr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParcelajr', 'jr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParcelade', 'de', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParcelaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParcelarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParcelaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParcelavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParceladescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParcelajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParcelapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBxParcela', 'ZQBxParceladata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQcheque_rec_e', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_eidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_ebanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_eagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_edono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_eCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_evalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_eemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_edeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_ealias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_eprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_econta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_eidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_eidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_enumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_e', 'ZQcheque_rec_eidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQcheque_rec_p', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pbanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pdono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pdeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_palias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pnumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcheque_rec_p', 'ZQcheque_rec_pidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'baixa', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixavr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'baixa', 'baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQrecebparcela', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebparcela', 'ZQrecebparcelanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebparcela', 'ZQrecebparcelavr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebparcela', 'ZQrecebparcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebparcela', 'ZQrecebparceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebparcela', 'ZQrecebparcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebparcela', 'ZQrecebparcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebparcela', 'ZQrecebparcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQrecebentrada', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebentrada', 'ZQrecebentradanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebentrada', 'ZQrecebentradavr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebentrada', 'ZQrecebentradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebentrada', 'ZQrecebentradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebentrada', 'ZQrecebentradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebentrada', 'ZQrecebentradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrecebentrada', 'ZQrecebentradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQReceb_BaixaE', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQReceb_BaixaE', 'ZQReceb_BaixaEJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQReceb_BaixaE', 'ZQReceb_BaixaEDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQReceb_BaixaE', 'ZQReceb_BaixaERefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQReceb_BaixaP', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQReceb_BaixaP', 'ZQReceb_BaixaPJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQReceb_BaixaP', 'ZQReceb_BaixaPDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQReceb_BaixaP', 'ZQReceb_BaixaPRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQcomprador_cessao', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcomprador_cessao', 'ZQcomprador_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQaditamento', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQaditamento', 'ZQaditamentohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQquitcao', True);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitcao', 'ZQquitcaohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQquitados', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosdatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadostabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosmarcar', 'marcar', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadoslote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadoslado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosmedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosmedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosmedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosmedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadoscornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosdisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosmatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosmatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosmarcar_1', 'marcar_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadossomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosdescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosobserv_estorno', 'observ_estorno', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadostip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadossld', 'sld', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadospaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQquitados', 'ZQquitadosnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQVendaZQuery', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaLargeintField', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaDateField', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaLargeintField2', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaFloatField', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaLargeintField3', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaLargeintField4', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField2', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField3', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField4', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField5', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField6', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField7', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaFloatField2', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField8', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField9', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField10', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField11', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField12', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField13', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField14', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaFloatField3', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaMemoField', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField15', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVendaZQuery', 'ZQVendaStringField16', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQcessao', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoautocodigo', 'autocodigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoidcomprador_1', 'idcomprador_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaopaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaovenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaopromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoidvenda_1', 'idvenda_1', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodatavenda_1', 'datavenda_1', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoimovel_1', 'imovel_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaovalorvenda_1', 'valorvenda_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaotabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaonumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoidloteamento_1', 'idloteamento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaomatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaologotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaomapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaotestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaotestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaotestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaotestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaobairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaopasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaonomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaotipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaonome_parte_1', 'nome_parte_1', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaodoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaobairro_1', 'bairro_1', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaobairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaofone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaofone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaotipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaonacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaonaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaolocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaorenda', 'renda', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaofone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaocadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaoNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaomarcar_1', 'marcar_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQcessao', 'ZQcessaovenda_idvenda_1', 'venda_idvenda_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQrelquitados', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadossomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosvalor_1', 'valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosjuros_1', 'juros_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadospercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadossq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadostabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosmatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoslogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosmapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadostestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadostestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadostestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadostestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosbairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadospasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosnomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadostipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosnome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosbairro_1', 'bairro_1', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadostipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoslocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadossexo', 'sexo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadoscex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosjb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosjudicial', 'judicial', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosxquadra', 'xquadra', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosxlote', 'xlote', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosdataquitacao', 'dataquitacao', TDateField, fkLookup, 0, 0, False, '', '', '', '', 0, 'venda_idvenda', 'ZQquitcao', 'idvenda', 'datacessao', True);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosHistorico', 'Historico', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'venda_idvenda', 'ZQquitcao', 'idvenda', 'historico', True);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQrelquitados', 'ZQrelquitadosidloteamento_1', 'idloteamento_1', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQTemp4', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4venda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4idloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4quadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4lote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4paticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4idpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4nome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4vr_entrada', 'vr_entrada', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4vr_parcela', 'vr_parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4juros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4descontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4juros_p', 'juros_p', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4descontos_p', 'descontos_p', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4Mes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4MESDESC', 'MESDESC', TWideStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQTemp4', 'ZQTemp4Ano', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQdiaria', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariatabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariamarca', 'marca', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarialoteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarialote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarialado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariamedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariamedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariamedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariamedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariavalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariadisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariamatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariamatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaproposta', 'proposta', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariamatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariadatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariadatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariadataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarialogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariamapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariatestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariatestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariatestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariatestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariabairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariapasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarianomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariatipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariadigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaidvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariavenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaidcorretor', 'idcorretor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarianomecorretor', 'nomecorretor', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariadoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacreci', 'creci', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariatelefone1', 'telefone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariatelefone2', 'telefone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaendereco_corr', 'endereco_corr', TWideStringField, fkData, 70, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariabairro_corr', 'bairro_corr', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacep_corr', 'cep_corr', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaObser_corr', 'Obser_corr', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacodcidade', 'codcidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariapaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariavenda_idvenda_1', 'venda_idvenda_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariapromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariapercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarianome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariadoc1_1', 'doc1_1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariadoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariabairro_1', 'bairro_1', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariabairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariafone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariafone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariatipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarianacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarianaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiarialocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariarenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariafone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaSexo', 'Sexo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariacex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariajb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariajudicial', 'judicial', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaxquadra', 'xquadra', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaxlote', 'xlote', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariaPlantao', 'Plantao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQdiaria', 'ZQdiariavlr_comissao', 'vlr_comissao', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQVenda2', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2idvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2forma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2tabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2Escriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2marca', 'marca', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2idimovel', 'idimovel', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2quadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2lote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2lado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2medidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2medidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2medidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2medidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2confrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2confrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2confrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2confrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2esquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2esquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2esquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2Obervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2cornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2disponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2matri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2matricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2proposta', 'proposta', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2valorCusto', 'valorCusto', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2documento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2Observ', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2ordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2TipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2marcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2quadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2nomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2Data_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2sld', 'sld', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2nome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQVenda2', 'ZQVenda2idpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQBalao', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQBalao', 'ZQBalaohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_RelVenda, 'ZQRescisao', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodatarescisao', 'datarescisao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaohora_rescisao', 'hora_rescisao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoautocodigo', 'autocodigo', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoquadralote_1', 'quadralote_1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoobserv_estorno', 'observ_estorno', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaotip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaosld_antes_reajuste', 'sld_antes_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoPercentual_reajuste', 'Percentual_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaojuridico', 'juridico', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodata_juridico', 'data_juridico', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodt_nao_pagou_no_mes', 'dt_nao_pagou_no_mes', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodescricao_juridico', 'descricao_juridico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoidloteamento_1', 'idloteamento_1', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaomatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaologotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaomapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaotestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaotestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaotestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaotestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaobairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaopasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaonomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaotipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaonome_parte_1', 'nome_parte_1', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaodoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaobairro_1', 'bairro_1', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaobairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaofone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaofone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaotipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaonacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaonaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaolocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaorenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaofone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoSexo', 'Sexo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaocex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaojb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaojudicial', 'judicial', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoxquadra', 'xquadra', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoxlote', 'xlote', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda, 'ZQRescisao', 'ZQRescisaoPlantao', 'Plantao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
