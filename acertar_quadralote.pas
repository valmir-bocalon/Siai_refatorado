unit acertar_quadralote;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Grids,
  DBGrids, StdCtrls, ComCtrls;

type
  TFracerto = class(TForm)
    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    Button1: TButton;
    Button2: TButton;
    PB: TProgressBar;
    DataZQreceb_baixa: TDataSource;
    ZQRecebimento: TZQuery;
    DS_Recebimento: TDataSource;



















    ZQreceb_baixa: TZQuery;















    Datarecbxhist: TDataSource;
    ZQrecbxhist: TZQuery;










    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Fracerto: TFracerto;

implementation

{$R *.dfm}
uses tabelas,Imoveis, RelREMESSA, uRuntimeFields;

procedure TFracerto.Button1Click(Sender: TObject);
var
pt,od:string;
cc,tr:Integer;
pr,dv:double;
begin
  tr:=0;
  dv:=0;
  pr:=0;
  pt:='';
  ZQrecebimento.First;
  tr:=ZQRecebimento.FieldByName('RefBaixa').Value;
  od:=ZQRecebimento.FieldByName('ordem').Value;
  pt:=ZQRecebimento.FieldByName('quadralote').Value;
  cc:=1;
  pb.max:=ZQrecebimento.recordcount;
  pb.position:=0;
  while not ZQrecebimento.Eof do
  begin
    pb.position:=ZQrecebimento.recno;
{    if DM_Tabelas.ZQrecebimentoquadralote.Value='' then
    begin
      tr:=pos('-',DM_Tabelas.ZQRecebimentoorigem.value);
      inc(tr);
      od:= trim(copy(DM_Tabelas.ZQRecebimentoorigem.value,tr,5));
      DM_Tabelas.ZQvenda.Filtered:=false;
      DM_Tabelas.ZQvenda.Filter:='idvenda='+quotedstr(od);
      DM_Tabelas.ZQvenda.Filtered:=true;
      if DM_Tabelas.ZQvenda.RecordCount>0 then
      begin
        DM_Tabelas.ZQRecebimento.Edit;
        DM_Tabelas.ZQRecebimentoquadralote.Value:=DM_Tabelas.ZQVendaquadra.Value+'-'+DM_Tabelas.ZQVendalote.Value;
        DM_Tabelas.ZQRecebimentovenda_idvenda.Value:=strtoint(od);
        DM_Tabelas.ZQRecebimento.post;
      end;
    end;}
    if (pt=ZQRecebimento.FieldByName('quadralote').Value) and (od=ZQRecebimento.FieldByName('ordem').Value) and (tr=ZQRecebimento.FieldByName('RefBaixa').Value) and (cc=1) then
    begin
      ZQrecebimento.Next;
      inc(cc);
    end;
    if (pt=ZQRecebimento.FieldByName('quadralote').Value) and (od=ZQRecebimento.FieldByName('ordem').Value) and (tr=ZQRecebimento.FieldByName('RefBaixa').Value) and (cc>1) then
    begin
      ZQreceb_baixa.Filtered:=false;
      ZQreceb_baixa.Filter:='idreceb_baixa='+quotedstr(ZQRecebimento.FieldByName('idreceb_baixa').Text);
      ZQreceb_baixa.Filtered:=true;
      if ZQreceb_baixa.RecordCount>0 then
         ZQreceb_baixa.Delete;

      ZQrecbxhist.Filtered:=false;
      ZQrecbxhist.Filter:='idrecbxhist='+quotedstr(ZQRecebimento.FieldByName('idrecbxhist').Text);
      ZQrecbxhist.Filtered:=true;
      if ZQrecbxhist.RecordCount>0 then
         ZQrecbxhist.Delete;

      ZQrecebimento.Next;
    end
    else
    begin
      tr:=ZQRecebimento.FieldByName('RefBaixa').Value;
      od:=ZQRecebimento.FieldByName('ordem').Value;
      pt:=ZQRecebimento.FieldByName('quadralote').Value;
      cc:=1;
    end;  
  end;
  close;
end;

procedure TFracerto.Button2Click(Sender: TObject);
begin
  close;
end;

procedure TFracerto.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 ZQrecebimento.close;
 ZQreceb_baixa.Close;
 ZQrecbxhist.close;
// DM_Tabelas.ZQvenda.close;

end;

procedure TFracerto.FormShow(Sender: TObject);
begin
{ DM_Tabelas.ZQrecebimento.close;
 DM_Tabelas.zqrecebimento.sql.clear;
 DM_Tabelas.zqrecebimento.sql.add('select * from recebimento order by quadralote');}
 zqrecebimento.open;
 ZQreceb_baixa.open;
 ZQrecbxhist.open;
// DM_Tabelas.ZQvenda.open;

end;


procedure TFracerto.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFracerto, 'ZQRecebimento', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentorefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQRecebimento', 'ZQRecebimentoidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFracerto, 'ZQreceb_baixa', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQreceb_baixa', 'ZQreceb_baixaVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFracerto, 'ZQrecbxhist', False);
  RegisterRuntimeField(TFracerto, 'ZQrecbxhist', 'ZQrecbxhistidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQrecbxhist', 'ZQrecbxhistrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQrecbxhist', 'ZQrecbxhistidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQrecbxhist', 'ZQrecbxhistvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQrecbxhist', 'ZQrecbxhistdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQrecbxhist', 'ZQrecbxhistjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQrecbxhist', 'ZQrecbxhistpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQrecbxhist', 'ZQrecbxhistdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFracerto, 'ZQrecbxhist', 'ZQrecbxhistsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
