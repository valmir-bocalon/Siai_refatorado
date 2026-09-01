unit Relparticipante_empre;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport, Data.DB,
  ZAbstractRODataset, ZAbstractDataset, ZDataset;

type
  TFrmRelparticipante_empre = class(TForm)
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
    RLBand3: TRLBand;
    RLBand4: TRLBand;
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
    RLLabel2: TRLLabel;
    RLLabel1: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDraw1: TRLDraw;
    RLDraw2: TRLDraw;
    RLDraw3: TRLDraw;
    RLDraw4: TRLDraw;
    RLDraw5: TRLDraw;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1_offAfterPrint(Sender: TObject);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand4BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelparticipante_empre: TFrmRelparticipante_empre;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, Participante, principal, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelparticipante_empre.RLBand4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand4.Color  := clWhite;
  if ZQAux.recno mod 2 = 0 then
    RLBand4.Color  := $00EAEAEA;
end;

procedure TFrmRelparticipante_empre.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  ZQAux.close;
  ZQAux.SQL.clear;
  ZQAux.SQL.Add('select i.idimovel, i.loteamento_idloteamento,v.idvenda,c.paticipante_idpaticipante,c.venda_idvenda,p.idpaticipante,p.nome_parte,p.fone1,p.fone2,p.fone3, l.idloteamento, l.apelido ');
  ZQAux.SQL.Add('from imovel as i  join venda as v on v.imovel=i.idimovel join comprador as c on c.venda_idvenda=v.idvenda join participante as p on p.idpaticipante=c.paticipante_idpaticipante ');
  ZQAux.SQL.Add('join loteamento as l on l.idloteamento=i.loteamento_idloteamento ');
//  ZQAux.SQL.Add('group by i.loteamento_idloteamento, v.idvenda, p.idpaticipante order by l.apelido,p.nome_parte ');
  ZQAux.SQL.Add('group by i.loteamento_idloteamento, p.idpaticipante order by l.apelido,p.nome_parte');
  ZQAux.open;
end;

procedure TFrmRelparticipante_empre.RLReport1_offAfterPrint(Sender: TObject);
begin
  ZQAux.close;
end;

procedure TFrmRelparticipante_empre.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel7.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelparticipante_empre.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmRelparticipante_empre, 'ZQAux', False);
  RegisterRuntimeField(TFrmRelparticipante_empre, 'ZQAux', 'ZQAuxidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre, 'ZQAux', 'ZQAuxnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre, 'ZQAux', 'ZQAuxfone1', 'fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre, 'ZQAux', 'ZQAuxfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre, 'ZQAux', 'ZQAuxfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre, 'ZQAux', 'ZQAuxidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre, 'ZQAux', 'ZQAuxapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelparticipante_empre, 'ZQAux', 'ZQAuxloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);

end.
