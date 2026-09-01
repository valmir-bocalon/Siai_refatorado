unit UFrmDash;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, VclTee.TeeGDIPlus, Data.DB,
  VCLTee.TeEngine, VCLTee.TeeTools, VCLTee.TeePageNumTool, VCLTee.Series,
  Vcl.Samples.Gauges, Vcl.ExtCtrls, VCLTee.TeeProcs, VCLTee.Chart,
  VCLTee.DBChart, dxCore2, dxButton, Vcl.StdCtrls, Vcl.Mask, XDate, XBanner,
  Datasnap.DBClient, ZAbstractRODataset, ZAbstractDataset, ZDataset, ACBrBase,
  ACBrEnterTab,jpeg;

type
  TFrmDash = class(TForm)
    CDSResultg: TClientDataSet;
    CDSResultgdata: TDateField;
    CDSResultgareceber: TFloatField;
    CDSResultgapagar: TFloatField;
    CDSResultgrecebido: TFloatField;
    CDSResultgpago: TFloatField;
    DS_result: TDataSource;
    XBanner9: TXBanner;
    Label28: TLabel;
    dtinicio: TXDateEdit;
    Label31: TLabel;
    dtfinal: TXDateEdit;
    Label29: TLabel;
    EParticipante: TEdit;
    Ecodcli: TEdit;
    dxButton4: TdxButton;
    XDVar: TXDateEdit;
    DBChart1: TDBChart;
    Gauge1: TGauge;
    dxButton5: TdxButton;
    ChartTool1: TPageNumTool;
    ZQCalculoEst: TZQuery;
    DS_CalculoEst: TDataSource;
    CDSResultA: TClientDataSet;
    CDSResultAData: TDateField;
    CDSResultAordem: TIntegerField;
    CDSResultAareceber: TFloatField;
    CDSResultArecebido: TFloatField;
    CDSResultAapagar: TFloatField;
    CDSResultApago: TFloatField;
    CDSResultgordem: TIntegerField;
    CDSResultV: TClientDataSet;
    Ds_resultV: TDataSource;
    CDSResultVdata: TDateField;
    CDSResultVordem: TIntegerField;
    CDSResultVrecebido: TFloatField;
    CDSResultVareceber: TFloatField;
    CDSResultgmes: TStringField;
    CDSResultVmes: TStringField;
    CDSResult: TClientDataSet;
    DS_CDSResult: TDataSource;
    DBChart3: TDBChart;
    Gauge3: TGauge;
    dxButton2: TdxButton;
    PageNumTool2: TPageNumTool;
    dxButton3: TdxButton;
    DS_ResultP: TDataSource;
    CDSResultP: TClientDataSet;
    ACBrEnterTab1: TACBrEnterTab;
    CDSResultdata: TDateField;
    CDSResultordem: TIntegerField;
    CDSResultTotalLotes: TIntegerField;
    CDSResulttotal_disponivel: TIntegerField;
    CDSResultTotal_nao_disponivel: TIntegerField;
    CDSResultTexto: TStringField;
    Series3: TBarSeries;
    Series5: TBarSeries;
    Series2: TBarSeries;
    CDSResultPdata: TDateField;
    CDSResultPordem: TIntegerField;
    CDSResultPTotalParcela: TIntegerField;
    CDSResultPaberto: TIntegerField;
    CDSResultPatrasado: TIntegerField;
    CDSResultPbaixado: TIntegerField;
    CDSResultPtexto: TStringField;
    Series6: TBarSeries;
    Series7: TBarSeries;
    CDSResultPpercaberto: TStringField;
    CDSResultPpercatrasado: TStringField;
    CDSResultPpercbaixado: TStringField;
    CDSResultpdisponivel: TStringField;
    CDSResultpnaodisponivel: TStringField;
    Image1: TImage;
    dxButton7: TdxButton;
    DBChart2: TDBChart;
    Gauge2: TGauge;
    dxButton1: TdxButton;
    BarSeries4: TBarSeries;
    BarSeries5: TBarSeries;
    BarSeries6: TBarSeries;
    PageNumTool1: TPageNumTool;
    CDSResultT: TClientDataSet;
    DS_ResultT: TDataSource;
    CDSResultTdata: TDateField;
    CDSResultTordem: TIntegerField;
    CDSResultTTotalParcela: TFloatField;
    CDSResultTaberto: TFloatField;
    CDSResultTatrasado: TFloatField;
    CDSResultTbaixado: TFloatField;
    CDSResultTtexto: TStringField;
    CDSResultTpercaberto: TStringField;
    CDSResultTpercatrasado: TStringField;
    CDSResultTpercbaixado: TStringField;
    BarSeries2: TBarSeries;
    Series1: TBarSeries;
    Series4: TBarSeries;
    procedure dxButton4Click(Sender: TObject);
    procedure EParticipanteExit(Sender: TObject);
    procedure dxButton5Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure dxButton1Click(Sender: TObject);
    procedure dxButton3Click(Sender: TObject);
    procedure dxButton7Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmDash: TFrmDash;

implementation

uses Tabelas, Principal,Funcoes , uRuntimeFields;

{$R *.dfm}

procedure TFrmDash.dxButton1Click(Sender: TObject);
begin
  DBChart3.Print;
end;

procedure TFrmDash.dxButton3Click(Sender: TObject);
begin
  close;
end;

procedure TFrmDash.dxButton4Click(Sender: TObject);
var
paberto,patrasado,pbaixado,pdisponivel,pnao:double;
dt:tdatetime;
ordem,passou:integer;
begin
   pdisponivel:=0;
   pnao:=0;
   ordem:=0;
   paberto:=0;
   patrasado:=0;
   pbaixado:=0;
   CDSResult.Close;
   CDSResult.CreateDataSet;
   CDSResult.Open;

   CDSResultP.Close;
   CDSResultP.CreateDataSet;
   CDSResultP.Open;

   CDSResultT.Close;
   CDSResultT.CreateDataSet;
   CDSResultT.Open;


   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('DROP TEMPORARY TABLE IF EXISTS lotes');
   ZQCalculoEst.ExecSQL;

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('CREATE TEMPORARY TABLE IF NOT EXISTS Lotes AS (SELECT ');
   ZQCalculoEst.SQL.Add(' i.idimovel,i.loteamento_idloteamento,l.idloteamento,l.apelido, ');
   ZQCalculoEst.SQL.Add('       count(i.loteamento_idloteamento) as totallotes, ');
   ZQCalculoEst.SQL.Add('       sum(i.disponivel=''SIM'') as totdisponivel,sum(i.disponivel<>''SIM'') as totNaodisponivel ');
   ZQCalculoEst.SQL.Add(' FROM imovel i inner join loteamento l on l.idloteamento=i.loteamento_idloteamento where 1=1 ');
   if Ecodcli.Text<>Emptystr then
      ZQCalculoEst.SQL.Add(' and i.loteamento_idloteamento in ('+Ecodcli.Text+')');
   ZQCalculoEst.SQL.Add(' group by i.loteamento_idloteamento)');
 //  showmessage(ZQCalculoEst.SQL.Text);
   ZQCalculoEst.ExecSQL;

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('select * from lotes');
   ZQCalculoEst.open;


  if (ZQCalculoEst.FieldByName('totdisponivel').AsInteger<>0) then
      pdisponivel:= (ZQCalculoEst.FieldByName('totdisponivel').AsInteger*100)/ZQCalculoEst.FieldByName('totallotes').AsInteger
   else
     pdisponivel:=0;

   if (ZQCalculoEst.FieldByName('totnaodisponivel').AsInteger<>0) then
      pnao:= (ZQCalculoEst.FieldByName('totnaodisponivel').AsInteger*100)/ZQCalculoEst.FieldByName('totallotes').AsInteger
   else
     pnao:=0;

   CDSResult.Insert;
   CDSResultordem.Value:=1;
   CDSResultTotalLotes.Value          := ZQCalculoEst.FieldByName('totallotes').AsInteger;
   CDSResulttotal_disponivel.Value    := ZQCalculoEst.FieldByName('totdisponivel').AsInteger;
   CDSResultTotal_nao_disponivel.Value:= ZQCalculoEst.FieldByName('totnaodisponivel').AsInteger;
   CDSResultPdisponivel.Value         := FormatFloat('##0.00',pdisponivel)+' %';
   CDSResultPnaodisponivel.Value      := FormatFloat('##0.00',pnao)+' %';
   CDSResultTexto.Value               := 'Totais';
   CDSResult.Post;


   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('DROP TEMPORARY TABLE IF EXISTS tot_parcelas');
   ZQCalculoEst.ExecSQL;


   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('CREATE TEMPORARY TABLE IF NOT EXISTS tot_parcelas AS (SELECT ');
   ZQCalculoEst.SQL.Add(' l.idloteamento,l.apelido,count(r.idloteamento) as totalParcelas, sum(r.saldo>0) as abertos, ');
   ZQCalculoEst.SQL.Add(' sum(r.saldo>0 and dt_vencimento<now()) as atrasados, sum(r.saldo=0) as baixados, ');
   ZQCalculoEst.SQL.Add(' concat(FORMAT(((sum(r.saldo>0)*100)/count(r.idloteamento)),2,''de_DE''),'' %'') as percaberto, ');
   ZQCalculoEst.SQL.Add(' concat(FORMAT(((sum(r.saldo>0 and dt_vencimento<now())*100)/count(r.idloteamento)),2,''de_DE''),'' %'') as percatrasado, ');
   ZQCalculoEst.SQL.Add(' concat(FORMAT(((sum(r.saldo=0)*100)/count(r.idloteamento)),2,''de_DE''),'' %'') as percbaixa ');
   ZQCalculoEst.SQL.Add('FROM recebimento r inner join loteamento l on l.idloteamento=r.idloteamento ');
   ZQCalculoEst.SQL.Add('  where dt_vencimento between :dt1 and :dt2');
   if Ecodcli.Text<>Emptystr then
      ZQCalculoEst.SQL.Add(' and r.idloteamento in ('+Ecodcli.Text+')');
   ZQCalculoEst.SQL.Add(' group by r.idloteamento order by l.apelido)');
   ZQCalculoEst.ParamByName('dt1').AsDate :=dtinicio.DateValue;
   ZQCalculoEst.ParamByName('dt2').AsDate :=dtfinal.DateValue;

 //  showmessage(ZQCalculoEst.SQL.Text);
   ZQCalculoEst.ExecSQL;

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('select * from tot_parcelas');
   ZQCalculoEst.open;
   ZQCalculoEst.First;

   CDSResultP.Insert;
   CDSResultPordem.Value:=1;
   CDSResultPTotalParcela.Value   := ZQCalculoEst.FieldByName('totalParcelas').AsInteger;
   CDSResultPaberto.Value         := ZQCalculoEst.FieldByName('abertos').AsInteger;
   CDSResultPatrasado.Value       := ZQCalculoEst.FieldByName('atrasados').AsInteger;
   CDSResultPbaixado.Value        := ZQCalculoEst.FieldByName('baixados').AsInteger;
   CDSResultPpercaberto.Value     := ZQCalculoEst.FieldByName('percaberto').AsString;
   CDSResultPpercatrasado.Value   := ZQCalculoEst.FieldByName('percatrasado').AsString;
   CDSResultPpercbaixado.Value    := ZQCalculoEst.FieldByName('percbaixa').AsString;
   CDSResultPTexto.Value          := 'Totais';
   CDSResultP.Post;


   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('DROP TEMPORARY TABLE IF EXISTS total_valor');
   ZQCalculoEst.ExecSQL;

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('CREATE TEMPORARY TABLE IF NOT EXISTS total_valor AS (SELECT ');
   ZQCalculoEst.SQL.Add('       l.idloteamento,l.apelido,round(sum(r.valor),2) as totalParcelas, ');
   ZQCalculoEst.SQL.Add('       round(sum(CASE WHEN (r.saldo > 0) THEN r.valor else 0 end),2) as aberto, ');
   ZQCalculoEst.SQL.Add('       round(sum(CASE WHEN (r.saldo > 0) and (r.dt_vencimento<now()) THEN r.valor else 0 end),2) as atrasado, ');
   ZQCalculoEst.SQL.Add('       round(sum(CASE WHEN (r.saldo = 0) THEN r.valor else 0 end),2) as pago ');
   ZQCalculoEst.SQL.Add(' FROM recebimento r inner join loteamento l on l.idloteamento=r.idloteamento ');
   ZQCalculoEst.SQL.Add('  where dt_vencimento between :dt1 and :dt2');
   if Ecodcli.Text<>Emptystr then
      ZQCalculoEst.SQL.Add(' and r.idloteamento in ('+Ecodcli.Text+')');
   ZQCalculoEst.SQL.Add(' group by r.idloteamento order by l.apelido)');
   ZQCalculoEst.ParamByName('dt1').AsDate :=dtinicio.DateValue;
   ZQCalculoEst.ParamByName('dt2').AsDate :=dtfinal.DateValue;

 //  showmessage(ZQCalculoEst.SQL.Text);
   ZQCalculoEst.ExecSQL;

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('select * from total_valor');
   ZQCalculoEst.open;
   ZQCalculoESt.First;


   if (ZQCalculoEst.FieldByName('aberto').AsFloat<>0) then
      paberto:= (ZQCalculoEst.FieldByName('aberto').AsFloat*100)/ZQCalculoEst.FieldByName('totalParcelas').AsFloat
   else
     paberto:=0;

   if (ZQCalculoEst.FieldByName('atrasado').AsFloat<>0) then
      patrasado:= (ZQCalculoEst.FieldByName('atrasado').AsFloat*100)/ZQCalculoEst.FieldByName('totalParcelas').AsFloat
   else
     patrasado:=0;

   if (ZQCalculoEst.FieldByName('pago').AsFloat<>0) then
      pbaixado:= (ZQCalculoEst.FieldByName('pago').AsFloat*100)/ZQCalculoEst.FieldByName('totalParcelas').AsFloat
   else
     pbaixado:=0;

   CDSResultT.Insert;
   CDSResultTordem.Value:=1;
   CDSResultTTotalParcela.Value   := ZQCalculoEst.FieldByName('totalParcelas').AsFloat;
   CDSResultTaberto.Value         := ZQCalculoEst.FieldByName('aberto').AsFloat;
   CDSResultTatrasado.Value       := ZQCalculoEst.FieldByName('atrasado').AsFloat;
   CDSResultTbaixado.Value        := ZQCalculoEst.FieldByName('pago').AsFloat;
   CDSResultTpercaberto.Value     := formatfloat('##0.00',paberto)+' %';
   CDSResultTpercatrasado.Value   := formatfloat('##0.00',patrasado)+' %';
   CDSResultTpercbaixado.Value    := formatfloat('##0.00',pbaixado)+' %';
   CDSResultTTexto.Value          := 'Totais';
   CDSResultT.Post;


end;
procedure TFrmDash.dxButton5Click(Sender: TObject);
begin
  DBChart1.Print;
end;

procedure TFrmDash.dxButton7Click(Sender: TObject);
begin
  Image1.Picture.Assign(CapturaTela);
  Image1.Picture.SaveToFile(ExtractFilePath(Application.ExeName)+'Dashboard.jpg');
  showmessage('O printe foi gravado na pasta do sistema com nome "Dashboard.jpg"');
end;

procedure TFrmDash.EParticipanteExit(Sender: TObject);
begin
  if not empty(EParticipante.Text) Then Begin
    if Acha_loteamento(EParticipante.Text) Then Begin
      EParticipante.Text := DM_tabelas.ZQAchaLotea.FieldByName('apelido').AsString;
      Ecodcli.Text := DM_tabelas.ZQAchaLotea.FieldByName('idloteamento').Text;
    end
    else
      Ecodcli.Text := '';
  end
  else
    Ecodcli.Text := '';
end;



procedure TFrmDash.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  CDSResult.Close;
  CDSResultP.Close;
  CDSResultV.Close;
  //CDSResultE.Close;  CDSResultS.Close;

end;


procedure TFrmDash.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
