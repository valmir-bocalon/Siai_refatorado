unit UFrmDash;

interface

uses ButtonDxArround, Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Samples.Gauges,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Mask, XDate, XBanner,
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
    dxButton4: TdxButtonArround;
    XDVar: TXDateEdit;
    PnlGrafico1: TPanel;
    Gauge1: TGauge;
    dxButton5: TdxButtonArround;
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
    PnlGrafico3: TPanel;
    Gauge3: TGauge;
    dxButton2: TdxButtonArround;
    dxButton3: TdxButtonArround;
    DS_ResultP: TDataSource;
    CDSResultP: TClientDataSet;
    ACBrEnterTab1: TACBrEnterTab;
    CDSResultdata: TDateField;
    CDSResultordem: TIntegerField;
    CDSResultTotalLotes: TIntegerField;
    CDSResulttotal_disponivel: TIntegerField;
    CDSResultTotal_nao_disponivel: TIntegerField;
    CDSResultTexto: TStringField;
    CDSResultPdata: TDateField;
    CDSResultPordem: TIntegerField;
    CDSResultPTotalParcela: TIntegerField;
    CDSResultPaberto: TIntegerField;
    CDSResultPatrasado: TIntegerField;
    CDSResultPbaixado: TIntegerField;
    CDSResultPtexto: TStringField;
    CDSResultPpercaberto: TStringField;
    CDSResultPpercatrasado: TStringField;
    CDSResultPpercbaixado: TStringField;
    CDSResultpdisponivel: TStringField;
    CDSResultpnaodisponivel: TStringField;
    Image1: TImage;
    dxButton7: TdxButtonArround;
    PnlGrafico2: TPanel;
    Gauge2: TGauge;
    dxButton1: TdxButtonArround;
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
    procedure dxButton4Click(Sender: TObject);
    procedure EParticipanteExit(Sender: TObject);
    procedure dxButton5Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure dxButton1Click(Sender: TObject);
    procedure dxButton3Click(Sender: TObject);
    procedure dxButton7Click(Sender: TObject);
  private
    { Private declarations }

    FPaintBoxGrafico1: TPaintBox;
    FPaintBoxGrafico2: TPaintBox;
    FPaintBoxGrafico3: TPaintBox;
    procedure CriarGraficosRuntime;
    procedure PintarGrafico1(Sender: TObject);
    procedure PintarGrafico2(Sender: TObject);
    procedure PintarGrafico3(Sender: TObject);
    procedure InvalidarGraficosRuntime;

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmDash: TFrmDash;

implementation

uses Tabelas, Principal,Funcoes , uRuntimeFields, ChartGenerator;

{$R *.dfm}

procedure TFrmDash.CriarGraficosRuntime;
  procedure CriarPaintBox(Container: TPanel; var PaintBox: TPaintBox;
    PaintHandler: TNotifyEvent);
  begin
    if PaintBox <> nil then
      Exit;
    if Container = nil then
      Exit;
    PaintBox := TPaintBox.Create(Container);
    PaintBox.Parent := Container;
    PaintBox.Align := alClient;
    PaintBox.OnPaint := PaintHandler;
  end;
begin
  CriarPaintBox(PnlGrafico1, FPaintBoxGrafico1, PintarGrafico1);
  CriarPaintBox(PnlGrafico2, FPaintBoxGrafico2, PintarGrafico2);
  CriarPaintBox(PnlGrafico3, FPaintBoxGrafico3, PintarGrafico3);
end;

procedure TFrmDash.PintarGrafico1(Sender: TObject);
var
  Series: TChartSeriesArray;
  Options: TChartRenderOptions;
begin
  if FPaintBoxGrafico1 = nil then
    Exit;
  LoadChartSeriesFromDataSet(CDSResult,
    ['Texto', 'pdisponivel', 'pnaodisponivel'],
    ['Total', 'Disponivel', 'Nao Disponivel'],
    ['TotalLotes', 'total_disponivel', 'Total_nao_disponivel'],
    [clBlue, clLime, clRed], Series);
  Series[0].BarPenWidth := 2;
  Series[0].MarkStyle := cmsValue;
  Series[1].BarPenWidth := 2;
  Series[1].BarStyle := cbsCylinder;
  Series[1].MarkStyle := cmsLabelValue;
  Series[2].BarPenStyle := psDash;
  Series[2].BarStyle := cbsCylinder;
  Series[2].MarkStyle := cmsLabelValue;

  Options := DefaultChartRenderOptions(clWhite);
  Options.UseBackgroundGradient := True;
  Options.BackgroundStartColor := TColor(RGB(228, 239, 255));
  Options.BackgroundEndColor := clWhite;
  Options.BorderColor := clBlue;
  Options.BorderWidth := 2;
  GenerateStyledMultiSeriesChart(FPaintBoxGrafico1.Canvas, Series, ctBar,
    FPaintBoxGrafico1.ClientWidth, FPaintBoxGrafico1.ClientHeight,
    'Lotes', Options, True, True);
end;

procedure TFrmDash.PintarGrafico2(Sender: TObject);
var
  Series: TChartSeriesArray;
  Options: TChartRenderOptions;
begin
  if FPaintBoxGrafico2 = nil then
    Exit;
  LoadChartSeriesFromDataSet(CDSResultT,
    ['texto', 'percaberto', 'percatrasado', 'percbaixado'],
    ['Total', 'Abertos', 'Atrasados', 'Baixados'],
    ['TotalParcela', 'aberto', 'atrasado', 'baixado'],
    [clBlue, clLime, clRed, clOlive], Series);
  Series[0].MarkStyle := cmsValue;
  Series[1].BarPenWidth := 2;
  Series[1].MarkStyle := cmsLabelValue;
  Series[2].MarkStyle := cmsLabelValue;
  Series[3].MarkStyle := cmsLabelValue;

  Options := DefaultChartRenderOptions(clWhite);
  Options.UseBackgroundGradient := True;
  Options.BackgroundStartColor := TColor(RGB(228, 239, 255));
  Options.BackgroundEndColor := clWhite;
  Options.BorderColor := clBlue;
  Options.BorderWidth := 2;
  GenerateStyledMultiSeriesChart(FPaintBoxGrafico2.Canvas, Series, ctBar,
    FPaintBoxGrafico2.ClientWidth, FPaintBoxGrafico2.ClientHeight,
    'Total de Parcelas Valores', Options, True, True);
end;

procedure TFrmDash.PintarGrafico3(Sender: TObject);
var
  Series: TChartSeriesArray;
  Options: TChartRenderOptions;
begin
  if FPaintBoxGrafico3 = nil then
    Exit;
  LoadChartSeriesFromDataSet(CDSResultP,
    ['texto', 'percaberto', 'percatrasado', 'percbaixado'],
    ['Total', 'Abertos', 'Atrasados', 'Baixados'],
    ['TotalParcela', 'aberto', 'atrasado', 'baixado'],
    [clBlue, clLime, clRed, clOlive], Series);
  Series[0].MarkStyle := cmsValue;
  Series[1].BarPenWidth := 2;
  Series[1].MarkStyle := cmsLabelValue;
  Series[2].MarkStyle := cmsLabelValue;
  Series[3].MarkStyle := cmsLabelValue;

  Options := DefaultChartRenderOptions(clWhite);
  Options.UseBackgroundGradient := True;
  Options.BackgroundStartColor := TColor(RGB(228, 239, 255));
  Options.BackgroundEndColor := clWhite;
  Options.BorderColor := clBlue;
  Options.BorderWidth := 2;
  GenerateStyledMultiSeriesChart(FPaintBoxGrafico3.Canvas, Series, ctBar,
    FPaintBoxGrafico3.ClientWidth, FPaintBoxGrafico3.ClientHeight,
    'Total de Parcelas', Options, True, True);
end;

procedure TFrmDash.InvalidarGraficosRuntime;
begin
  if FPaintBoxGrafico1 <> nil then
    FPaintBoxGrafico1.Invalidate;
  if FPaintBoxGrafico2 <> nil then
    FPaintBoxGrafico2.Invalidate;
  if FPaintBoxGrafico3 <> nil then
    FPaintBoxGrafico3.Invalidate;
end;

procedure TFrmDash.dxButton1Click(Sender: TObject);
begin
  if FPaintBoxGrafico3 <> nil then
    PrintPaintBoxChart(FPaintBoxGrafico3, 'Total de Parcelas', True);
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


   InvalidarGraficosRuntime;
end;
procedure TFrmDash.dxButton5Click(Sender: TObject);
begin
  if FPaintBoxGrafico1 <> nil then
    PrintPaintBoxChart(FPaintBoxGrafico1, 'Lotes', True);
end;

procedure TFrmDash.dxButton7Click(Sender: TObject);
var
  Captura: TBitmap;
begin
  Captura := CapturaTela;
  try
    Image1.Picture.Assign(Captura);
    Image1.Picture.SaveToFile(ExtractFilePath(Application.ExeName)+'Dashboard.jpg');
  finally
    Captura.Free;
  end;
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
  CriarGraficosRuntime;
end;

end.
