unit Relparticipante_empre_completa;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport, Data.DB,
  ZAbstractRODataset, ZAbstractDataset, ZDataset;

type
  TFrmRelparticipante_empre_completa = class(TForm)
    RLExpressionParser1: TRLExpressionParser;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLPDFFilter1: TRLPDFFilter;
    ZQAux: TZQuery;
    Ds_Aux: TDataSource;
    RLReport1: TRLReport;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLLabel3: TRLLabel;
    RLDBText1: TRLDBText;
    RLBand1: TRLBand;
    RLSystemInfo1: TRLSystemInfo;
    RLBand5: TRLBand;
    RLSystemInfo2: TRLSystemInfo;
    RLBand6: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel4: TRLLabel;
    RLLabel5: TRLLabel;
    RLSystemInfo4: TRLSystemInfo;
    RLLabel8: TRLLabel;
    RLSystemInfo3: TRLSystemInfo;
    RLSystemInfo5: TRLSystemInfo;
    RLLabel6: TRLLabel;
    RLSystemInfo6: TRLSystemInfo;
    RLLabel7: TRLLabel;
    RLDraw1: TRLDraw;
    RLDraw4: TRLDraw;
    RLDraw5: TRLDraw;
    RLBand3: TRLBand;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLLabel2: TRLLabel;
    RLLabel1: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLDBText7: TRLDBText;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLDBText8: TRLDBText;
    RLLabel15: TRLLabel;
    RLDBText9: TRLDBText;
    RLDBText10: TRLDBText;
    RLLabel16: TRLLabel;
    RLDBText11: TRLDBText;
    RLLabel17: TRLLabel;
    RLDBText12: TRLDBText;
    RLLabel18: TRLLabel;
    RLDBText13: TRLDBText;
    RLLabel19: TRLLabel;
    RLDBText14: TRLDBText;
    RLLabel20: TRLLabel;
    RLDBText15: TRLDBText;
    RLLabel21: TRLLabel;
    RLDBText16: TRLDBText;
    RLLabel22: TRLLabel;
    RLDBText17: TRLDBText;
    RLLabel23: TRLLabel;
    RLDBText18: TRLDBText;
    RLLabel24: TRLLabel;
    RLDBText19: TRLDBText;
    RLLabel25: TRLLabel;
    RLDBText20: TRLDBText;
    RLDraw2: TRLDraw;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1_offAfterPrint(Sender: TObject);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelparticipante_empre_completa: TFrmRelparticipante_empre_completa;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, Participante, principal, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelparticipante_empre_completa.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand3.Color  := clWhite;
  if ZQAux.recno mod 2 = 0 then
    RLBand3.Color  := $00EAEAEA;
  if Trim(ZQAux.FieldByName('nomecidade').AsString) <> '' then
    RLDBText9.DataField := 'nomecidade'
  else
    RLDBText9.DataField := 'nomecidadenatural';
  if Trim(ZQAux.FieldByName('nomeestado').AsString) <> '' then
    RLDBText10.DataField := 'nomeestado'
  else
    RLDBText10.DataField := 'nomeestadonatural';
end;

procedure TFrmRelparticipante_empre_completa.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Close;
  DM_Tabelas.ZQCidade.Open;
  ZQAux.close;
  ZQAux.SQL.clear;
  ZQAux.SQL.Add('select i.idimovel, i.loteamento_idloteamento,v.idvenda,c.paticipante_idpaticipante,c.venda_idvenda,p.idpaticipante,p.nome_parte,p.fone1,p.fone2,p.fone3,');
  ZQAux.SQL.Add(' p.endereco,p.bairro,p.cidade,p.naturalidade,p.cep,p.complemento,p.doc1,p.doc2,p.inativo,p.aniversario,p.estadocivil,p.profissao,p.observacao,p.email,');
  ZQAux.SQL.Add(' l.idloteamento, l.apelido ');
  ZQAux.SQL.Add('from imovel as i  join venda as v on v.imovel=i.idimovel join comprador as c on c.venda_idvenda=v.idvenda join participante as p on p.idpaticipante=c.paticipante_idpaticipante ');
  ZQAux.SQL.Add('join loteamento as l on l.idloteamento=i.loteamento_idloteamento');
//  ZQAux.SQL.Add('group by i.loteamento_idloteamento, v.idvenda, p.idpaticipante order by l.apelido,p.nome_parte ');
  ZQAux.SQL.Add('group by i.loteamento_idloteamento, p.idpaticipante order by l.apelido,p.nome_parte');
  ZQAux.open;
end;

procedure TFrmRelparticipante_empre_completa.RLReport1_offAfterPrint(Sender: TObject);
begin
  ZQAux.close;
end;

procedure TFrmRelparticipante_empre_completa.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel7.Caption:='Usuario:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelparticipante_empre_completa.AfterConstruction;
begin
  inherited AfterConstruction;
  RLReport1.BeforePrint := RLReport1BeforePrint;
  EnsureRuntimeFields(Self);
end;

initialization
  if Classes.GetClass('TRLReport') = nil then
    Classes.RegisterClass(TRLReport);
  RegisterRuntimeDataSet(TFrmRelparticipante_empre_completa, 'ZQAux', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxfone1', 'fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxnomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxnomeestado', 'nomeestado', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxnomecidadenatural', 'nomecidadenatural', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'naturalidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxnomeestadonatural', 'nomeestadonatural', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'naturalidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxinativo', 'inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre_completa, 'ZQAux', 'ZQAuxemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
