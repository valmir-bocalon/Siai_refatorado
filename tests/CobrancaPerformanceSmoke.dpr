program CobrancaPerformanceSmoke;
{$APPTYPE CONSOLE}
uses System.SysUtils, System.Classes, Data.DB, Vcl.Forms,
  ZConnection, ZDataset, uSiaiPerformance, RLReport, uSiaiReportPerformance,
  RelCobranca2 in 'RelCobranca2.pas',
  Cad_Recebimento in 'Cad_Recebimento.pas', tabelas in 'tabelas.pas';
type
  TLookupProbe = class(TZQuery)
  public
    Calls: Integer;
    function Lookup(const KeyFields: string; const KeyValues: Variant;
      const ResultFields: string): Variant; override;
  end;
  TProbe = class(TFrmCad_Recebimento)
    procedure AfterConstruction; override;
  end;
  TReportProbe = class(TFrmRelCobranca2)
    procedure AfterConstruction; override;
  end;
  TDMProbe = class(TDM_Tabelas)
    procedure AfterConstruction; override;
  end;
  TCounter = class
    Sales, Summaries, RenderedRows: Integer;
    Calculations: Integer;
    RenderedTotal: Double;
    ReportBeforeCount, ReportAfterCount: Integer;
    CalcOriginal: TDataSetNotifyEvent;
    procedure CalcFields(DataSet: TDataSet);
    procedure DetailBefore(Sender: TObject; var PrintIt: Boolean);
    procedure SaleOpen(DataSet: TDataSet);
    procedure SummaryOpen(DataSet: TDataSet);
    procedure ReportBefore(Sender: TObject; var PrintIt: Boolean);
    procedure ReportAfter(Sender: TObject);
  end;
function TLookupProbe.Lookup(const KeyFields: string; const KeyValues: Variant;
  const ResultFields: string): Variant;
begin
  Inc(Calls);
  Result := inherited Lookup(KeyFields, KeyValues, ResultFields);
end;
procedure TProbe.AfterConstruction;
begin end;
procedure TReportProbe.AfterConstruction;
begin end;
procedure TDMProbe.AfterConstruction;
begin end;
procedure TCounter.SaleOpen(DataSet: TDataSet);
begin Inc(Sales) end;
procedure TCounter.SummaryOpen(DataSet: TDataSet);
begin Inc(Summaries) end;
procedure TCounter.ReportBefore(Sender: TObject; var PrintIt: Boolean);
begin Inc(ReportBeforeCount); PrintIt := False end;
procedure TCounter.ReportAfter(Sender: TObject);
begin Inc(ReportAfterCount) end;
procedure TCounter.CalcFields(DataSet: TDataSet);
begin
  Inc(Calculations);
  if Assigned(CalcOriginal) then
    CalcOriginal(DataSet);
end;
procedure TCounter.DetailBefore(Sender: TObject; var PrintIt: Boolean);
var ExpectedBank: string;
begin
  Inc(RenderedRows);
  RenderedTotal := RenderedTotal + FrmCad_Recebimento.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat;
  if FrmCad_Recebimento.ZQRecebimento.FieldByName('idloteamento').AsInteger=1 then
    ExpectedBank := '001' else ExpectedBank := '237';
  if FrmCad_Recebimento.ZQRecebimento_bancario.FieldByName('n_banco').AsString <> ExpectedBank then
    raise Exception.Create('Wrong bank while generating report detail');
end;
procedure Check(Value: Boolean; const S: string);
begin if not Value then raise Exception.Create(S) end;
procedure Near(Value, Expected: Double; const S: string);
begin Check(Abs(Value - Expected) < 0.001, S + ': ' + FloatToStr(Value)) end;
procedure Fields(Q: TZQuery; WithLookups: Boolean);
var I: Integer; Field: TField;
begin
  Q.FieldDefs.Update;
  for I := 0 to Q.FieldDefs.Count - 1 do Q.FieldDefs[I].CreateField(Q);
  for I := 0 to 3 do
  begin
    Field := TFloatField.Create(Q);
    case I of
      0: Field.FieldName := 'Dias';
      1: Field.FieldName := 'Multa_Contrato';
      2: Field.FieldName := 'mora_dia';
      3: Field.FieldName := 'parcela_corrigida';
    end;
    Field.FieldKind := fkCalculated; Field.DataSet := Q;
  end;
  if WithLookups then
    for I := 0 to 2 do
    begin
      Field := TStringField.Create(Q);
      case I of
        0: Field.FieldName := 'nome_loteamento';
        1: Field.FieldName := 'adversanome';
        2: Field.FieldName := 'CPF';
      end;
      Field.Size := 80; Field.FieldKind := fkCalculated; Field.DataSet := Q;
    end;
end;
var C: TZConnection; F: TProbe; D: TDMProbe; Counter: TCounter;
  I, Queries, Summaries, ParticipantCalls, DevelopmentCalls: Integer;
  Report: TRLReport;
  GridReport: TReportProbe;
  BankMaster: TDataSource;
  PrintIt, MissingSaleRejected: Boolean;
  MixedNormal, MixedPrice, RowNormal, RowPrice: Double;
begin
  try
    Application.Initialize;
    C := TZConnection.Create(nil);
    Counter := TCounter.Create;
    D := nil; F := nil;
    try
      C.Protocol := 'mysql-5'; C.LibraryLocation := 'C:\Windows\SysWOW64\libmysql.dll';
      C.HostName := '127.0.0.1'; C.Port := 33557; C.User := 'root';
      C.Database := 'siai_metadata_test'; C.Connect;
      C.ExecuteDirect('CREATE TABLE venda (idvenda BIGINT, multa DOUBLE, mora DOUBLE, datavenda DATE, imovel INT, valorvenda DOUBLE, forma_reajuste INT, tabela_Price DOUBLE, Escriturado CHAR(1), marca INT, codigo_contrato_ref INT, Perc_comissao DOUBLE, vlr_comissao DOUBLE)');
      C.ExecuteDirect('INSERT INTO venda (idvenda,multa,mora,tabela_price) VALUES (100,10,2,3),(200,5,4,6)');
      C.ExecuteDirect('CREATE TABLE recebimento (idrecebimento BIGINT, venda_idvenda BIGINT, multa DOUBLE, mora DOUBLE, idloteamento INT, adversa INT, quadralote VARCHAR(20), saldo DOUBLE, Dt_Vencimento DATE, tabela_price DOUBLE, nomeadversa VARCHAR(80))');
      C.ExecuteDirect('INSERT INTO recebimento VALUES (1,100,10,2,1,1,''A'',100,CURRENT_DATE-INTERVAL 15 DAY,3,''Client''),(2,100,10,2,1,1,''A'',200,CURRENT_DATE-INTERVAL 15 DAY,3,''Client''),(3,200,5,4,1,2,''B'',300,CURRENT_DATE-INTERVAL 15 DAY,6,''Other'')');
      D := TDMProbe.CreateNew(nil); DM_Tabelas := D;
      D.ZqParticipante := TLookupProbe.Create(D); D.ZqParticipante.Connection := C;
      D.ZqParticipante.SQL.Text := 'SELECT 1 AS idpaticipante, ''Client'' AS nome_parte, ''123'' AS doc1 UNION ALL SELECT 2,''Other'',''456'''; D.ZqParticipante.Open;
      D.ZQLoteamento := TLookupProbe.Create(D); D.ZQLoteamento.Connection := C;
      D.ZQLoteamento.SQL.Text := 'SELECT 1 AS idloteamento, ''Lote'' AS apelido'; D.ZQLoteamento.Open;
      F := TProbe.Create(nil); FrmCad_Recebimento := F;
      F.OnShow := nil; F.OnClose := nil;
      BankMaster := F.ZQRecebimento_bancario.MasterSource;
      F.ZQRecebimento_bancario.MasterSource := nil;
      F.ZQRecebimento.Connection := C;
      F.ZQRecebimento2.Connection := C;
      F.ZQRecebimento3.Connection := C;
      F.ZQRecebimento4.Connection := C;
      F.ZQVenda_cobr.Connection := C;
      F.ZQVenda_cobr.BeforeOpen := Counter.SaleOpen;
      F.ZQRecebimento2.BeforeOpen := Counter.SummaryOpen;
      Check(not Assigned(F.ZQRecebimento.BeforeOpen) and
        not Assigned(F.ZQRecebimento.BeforeRefresh) and
        not Assigned(F.ZQRecebimento.AfterPost) and
        not Assigned(F.ZQRecebimento.AfterDelete), 'Existing lifecycle handler would be overwritten');
      F.ZQRecebimento.BeforeOpen := F.CobrancaAntesReabrir;
      F.ZQRecebimento.BeforeRefresh := F.CobrancaAntesReabrir;
      F.ZQRecebimento.AfterPost := F.CobrancaAntesReabrir;
      F.ZQRecebimento.AfterDelete := F.CobrancaAntesReabrir;
      F.ZQRecebimento.SQL.Text := 'SELECT * FROM recebimento WHERE adversa=1 ORDER BY idrecebimento';
      Fields(F.ZQRecebimento, True);
      Counter.CalcOriginal := F.ZQRecebimento.OnCalcFields;
      F.ZQRecebimento.OnCalcFields := Counter.CalcFields;
      F.ZQRecebimento4.SQL.Text := 'SELECT * FROM recebimento WHERE 1=0';
      Fields(F.ZQRecebimento4, False);
      F.Lmora.Caption := 'N'; F.Label52.Caption := '';
      F.ZQRecebimento.Open;
      Check(F.ZQRecebimento.AutoCalcFields, 'Calculated fields are disabled');
      F.ZQRecebimento.First;
      Counter.Calculations := 0;
      F.DBcobrancaEnter(nil);
      Check(Counter.Calculations <= F.ZQRecebimento.RecordCount + 2,
        'A title was calculated more than once during the complete recalculation');
      Near(F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat,111,
        'Single-pass recalculation changed first title');
      Check(F.ZQRecebimento4.AutoCalcFields,
        'Totalizer did not restore calculated fields');
      F.ZQRecebimento4.First;
      Near(F.ZQRecebimento4.FieldByName('parcela_corrigida').AsFloat,111,
        'Totalizer calculation changed the first title');
      Near(F.XCorrigido.Value,333,'Totalizer direct calculation');
      Writeln('PASS: complete recalculation evaluates each title once');
      Queries := Counter.Sales;
      for I := 1 to 20 do begin F.ZQRecebimento.Last; F.ZQRecebimento.First end;
      Check(Counter.Sales = Queries, 'Repeated navigation reloaded the same sale');
      Near(F.ZQRecebimento.FieldByName('Multa_Contrato').AsFloat,10,'Fine');
      Near(F.ZQRecebimento.FieldByName('mora_dia').AsFloat,1,'Interest');
      Near(F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat,111,'Corrected');
      Writeln('PASS: repeated navigation adds zero sale queries; fine/interest/corrected values preserved');
      D.ZqParticipante.Last;
      ParticipantCalls := TLookupProbe(D.ZqParticipante).Calls;
      DevelopmentCalls := TLookupProbe(D.ZQLoteamento).Calls;
      F.ZQRecebimento.Last; F.ZQRecebimento.First;
      ParticipantCalls := TLookupProbe(D.ZqParticipante).Calls-ParticipantCalls;
      DevelopmentCalls := TLookupProbe(D.ZQLoteamento).Calls-DevelopmentCalls;
      Check((ParticipantCalls>0) and (ParticipantCalls=DevelopmentCalls),
        'Participant lookup was not consolidated into one access');
      Check(D.ZqParticipante.FieldByName('idpaticipante').AsInteger=2,'Lookup moved global participant');
      Check(F.ZQRecebimento.FieldByName('adversanome').AsString='Client','Combined name lookup');
      Check(F.ZQRecebimento.FieldByName('CPF').AsString='123','Combined document lookup');
      Writeln('PASS: one participant lookup supplies name and document without moving global cursor');
      F.DBcobrancaEnter(nil);
      Near(F.XCorrigido.Value,333,'Totals');
      F.RGSelecao.OnClick := nil; F.RGSelecao.ItemIndex := 0;
      F.cbmontante.Checked := False; F.fpg.Text := '000';
      F.fpgExit(nil);
      Check(F.CDnegocio.Active, 'Negotiation was not created');
      Queries := Counter.Sales; Summaries := Counter.Summaries;
      for I := 1 to 10 do F.DBcobrancaEnter(nil);
      Check(Counter.Sales=Queries,'Focus reloaded sale');
      Check(Counter.Summaries=Summaries,'Focus rebuilt totals');
      Check(F.CDnegocio.Active, 'Focus erased negotiation');
      Near(F.CDnegociovr_parcela.AsFloat,333,'Preserved negotiation');
      Check(D.ZqParticipante.Active and D.ZQLoteamento.Active, 'Global lookup was closed');
      Writeln('PASS: ten focus returns add zero sale/summary queries, globals remain active');
      F.Lmora.Caption := 'S'; F.DBcobrancaEnter(nil);
      Check(Counter.Summaries=Summaries+1,'Price mode did not invalidate totals');
      Near(F.XCorrigido.Value,334.5,'Price totals');
      Writeln('PASS: Price mode recalculates with preserved formula');
      F.Lmora.Caption := 'N';
      C.ExecuteDirect('UPDATE venda SET mora=4 WHERE idvenda=100');
      F.ZQRecebimento.Refresh;
      F.DBcobrancaEnter(nil);
      Near(F.XCorrigido.Value,336,'Refreshed rate');
      Writeln('PASS: refresh reloads changed sale rates');
      F.fpg.Text := '000'; F.fpgExit(nil);
      Summaries := Counter.Summaries;
      F.ZQRecebimento.Last; F.DBcobrancaEnter(nil);
      Check(Counter.Summaries=Summaries,'Same participant navigation rebuilt totals');
      Near(F.XCorrigido.Value,336,'Totals after same participant navigation');
      Check(F.CDnegocio.Active,'Navigation erased negotiation');
      Near(F.CDnegociovr_parcela.AsFloat,336,'Negotiation after navigation');
      F.ZQRecebimento.Close;
      F.ZQRecebimento.SQL.Text := 'SELECT * FROM recebimento WHERE adversa=2 ORDER BY idrecebimento';
      F.ZQRecebimento.Open; F.DBcobrancaEnter(nil);
      Near(F.XCorrigido.Value,321,'Other customer');
      Check(F.ZQVenda_cobr.FieldByName('idvenda').AsInteger=200,'Local sale identity changed');
      Writeln('PASS: changed receipt and customer reload correct totals and sale');
      FlushPerformanceLog;
      Report := TRLReport.Create(nil);
      try
        Report.BeforePrint := Counter.ReportBefore;
        Report.AfterPrint := Counter.ReportAfter;
        MeasureReport(Report, 'Teste dos eventos do relatorio');
        MeasureReport(Report, 'Teste dos eventos do relatorio');
        PrintIt := True;
        Report.BeforePrint(Report, PrintIt);
        Check(not PrintIt, 'Report cancellation changed');
        Report.AfterPrint(Report);
        Check((Counter.ReportBeforeCount = 1) and (Counter.ReportAfterCount = 1),
          'Report events were replaced or duplicated');
        Writeln('PASS: report timing preserves callbacks and cancellation without duplicate hooks');
      finally Report.Free end;
      { Financial consistency gate: one participant, two contracts and two
        developments. Expected amounts are independent of the UI totals. }
      C.ExecuteDirect('INSERT INTO recebimento VALUES (4,100,10,4,1,3,''C'',100,CURRENT_DATE-INTERVAL 15 DAY,3,''Mixed''),(5,200,5,4,2,3,''D'',300,CURRENT_DATE-INTERVAL 15 DAY,6,''Mixed'')');
      F.ZQRecebimento.Close;
      F.ZQRecebimento.SQL.Text := 'SELECT * FROM recebimento WHERE adversa=3 ORDER BY idrecebimento';
      F.Lmora.Caption := 'N'; F.Label52.Caption := '';
      F.ZQRecebimento.Open; F.DBcobrancaEnter(nil);
      MixedNormal := F.XCorrigido.Value;
      Near(F.nparcelas.Value, 2, 'Mixed contract count');
      Near(F.vencidos.Value, 400, 'Mixed principal');
      F.ZQRecebimento.First;
      Near(F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat,112,'Contract 100 normal');
      RowNormal := F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat;
      F.ZQRecebimento.Next;
      Near(F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat,321,'Contract 200 normal');
      RowNormal := RowNormal + F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat;
      Writeln('CHECK mixed normal: grid=',RowNormal:0:2,' summary=',MixedNormal:0:2,' expected=433.00');
      F.ZQRecebimento.First;
      F.Lmora.Caption := 'S'; F.DBcobrancaEnter(nil);
      MixedPrice := F.XCorrigido.Value;
      F.ZQRecebimento.First;
      Near(F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat,111.5,'Contract 100 Price');
      RowPrice := F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat;
      F.ZQRecebimento.Next;
      Near(F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat,324,'Contract 200 Price');
      RowPrice := RowPrice + F.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat;
      Writeln('CHECK mixed Price: grid=',RowPrice:0:2,' summary=',MixedPrice:0:2,' expected=435.50');
      Near(MixedNormal,433,'Mixed contract summary normal');
      Near(MixedPrice,435.5,'Mixed contract summary Price');
      Writeln('PASS: mixed contracts match independent expected totals in normal and Price modes');
      { Preserve the actual DFM master/detail relation using fictitious banks. }
      C.ExecuteDirect('CREATE TABLE bancos_teste (idloteamento INT, n_banco VARCHAR(10))');
      C.ExecuteDirect('INSERT INTO bancos_teste VALUES (1,''001''),(2,''237'')');
      Check(BankMaster = F.DS_Recebimento, 'Unexpected bank master source in DFM');
      F.ZQRecebimento_bancario.Connection := C;
      F.ZQRecebimento_bancario.SQL.Text := 'SELECT *, ''1234'' AS n_agencia, ''0'' AS n_agencia_v, ''5678'' AS n_conta, ''9'' AS n_conta_v FROM bancos_teste ORDER BY idloteamento';
      F.ZQRecebimento_bancario.MasterSource := BankMaster;
      F.ZQRecebimento_bancario.Open;
      F.ZQRecebimento.First;
      Check(F.ZQRecebimento_bancario.FieldByName('n_banco').AsString='001','Bank for first development');
      F.ZQRecebimento.Next;
      Check(F.ZQRecebimento_bancario.FieldByName('n_banco').AsString='237','Bank for second development');
      F.ZQRecebimento.First;
      Summaries := Counter.Summaries;
      F.FinalizarPesquisaCobranca;
      Check(Counter.Summaries=Summaries+1,'Search did not build totals exactly once');
      Near(F.XCorrigido.Value,435.5,'Finalized search total');
      Check(F.ZQRecebimento.FieldByName('idrecebimento').AsInteger=4,'Search changed current receipt');
      Check(F.ZQVenda_cobr.FieldByName('idvenda').AsInteger=100,'Search changed negotiation sale');
      Check(F.ZQRecebimento_bancario.FieldByName('n_banco').AsString='001','Search changed current bank');
      F.DBcobrancaEnter(nil);
      Check(Counter.Summaries=Summaries+1,'Focus repeated finalized search');
      F.fpg.Text := '000'; F.fpgExit(nil);
      Check(F.CDnegocio.Active,'Mixed negotiation was not created');
      F.ZQRecebimento.Next;
      F.DBcobrancaEnter(nil);
      Check(Counter.Summaries=Summaries+1,'Moving to another contract rebuilt unchanged participant totals');
      Check(F.ZQRecebimento_bancario.FieldByName('n_banco').AsString='237','Bank navigation after search');
      Check(F.ZQVenda_cobr.FieldByName('idvenda').AsInteger=200,'Navigation did not change active sale');
      Near(F.XCorrigido.Value,435.5,'Navigation between contracts changed totals');
      Check(F.CDnegocio.Active,'Contract navigation erased negotiation');
      Near(F.CDnegociovr_parcela.AsFloat,435.5,'Contract navigation changed negotiation');
      Writeln('PASS: search finalization runs once and preserves real DFM bank relation and current sale');
      F.ZQRecebimento.First;
      GridReport := TReportProbe.Create(nil);
      try
        GridReport.OnClose := nil;
        GridReport.RLBand1.BeforePrint := nil;
        GridReport.RLReport1.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText2.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText2.DataField:='nome_loteamento';
  GridReport.RLDBText12.DataSource:=F.DS_ZQRecebimento_bancario;
  GridReport.RLDBText12.DataField:='n_banco';
  GridReport.RLDBText4.DataSource:=F.DS_ZQRecebimento_bancario;
  GridReport.RLDBText4.DataField:='n_agencia';
  GridReport.RLDBText6.DataSource:=F.DS_ZQRecebimento_bancario;
  GridReport.RLDBText6.DataField:='n_agencia_v';
  GridReport.RLDBText9.DataSource:=F.DS_ZQRecebimento_bancario;
  GridReport.RLDBText9.DataField:='n_conta';
  GridReport.RLDBText11.DataSource:=F.DS_ZQRecebimento_bancario;
  GridReport.RLDBText11.DataField:='n_conta_v';
  GridReport.RLDBText18.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText18.DataField:='adversanome';
  GridReport.RLDBText1.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText1.DataField:='CPF';
  GridReport.RLDBText20.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText20.DataField:='quadralote';
  GridReport.RLDBText8.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText8.DataField:='Dt_Vencimento';
  GridReport.RLDBText10.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText10.DataField:='saldo';
  GridReport.RLDBText5.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText5.DataField:='Dias';
  GridReport.RLDBText16.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText16.DataField:='Multa_Contrato';
  GridReport.RLDBText7.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText7.DataField:='mora_dia';
  GridReport.RLDBText3.DataSource:=F.DS_Recebimento;
  GridReport.RLDBText3.DataField:='parcela_corrigida';
  GridReport.RLDBResult1.DataSource:=F.DS_Recebimento;
  GridReport.RLDBResult1.DataField:='parcela_corrigida';
        for I := 0 to GridReport.ComponentCount-1 do
          if (GridReport.Components[I] is TRLBand) and
             (TRLBand(GridReport.Components[I]).BandType=btDetail) then
            TRLBand(GridReport.Components[I]).BeforePrint := Counter.DetailBefore;
        GridReport.RLReport1.ShowProgress := False;
        Check(GridReport.RLReport1.Prepare, 'Report generation failed');
        Check(Counter.RenderedRows=2,'Report did not generate both titles');
        Near(Counter.RenderedTotal,435.5,'Generated report total');
        GridReport.RLReport1.SaveToFile(ExtractFilePath(ParamStr(0))+'cobranca-mixed-test.rpf');
        Writeln('PASS: actual report DFM generates both titles with correct total and bank per development');
      finally GridReport.Free end;
      F.Label52.Caption := '2';
      F.ZQRecebimento.Close;
      F.ZQRecebimento.SQL.Text := 'SELECT * FROM recebimento WHERE adversa=3 AND idloteamento=2';
      F.ZQRecebimento.Open; F.FinalizarPesquisaCobranca;
      Near(F.XCorrigido.Value,324,'Development filter total');
      Near(F.nparcelas.Value,1,'Development filter count');
      Writeln('PASS: development filter selects only its own contract and rates');
      C.ExecuteDirect('INSERT INTO recebimento VALUES (6,999,0,0,2,3,''E'',50,CURRENT_DATE-INTERVAL 15 DAY,0,''Mixed'')');
      F.ZQRecebimento.Close;
      F.ZQRecebimento.SQL.Text := 'SELECT * FROM recebimento WHERE idrecebimento=5';
      F.ZQRecebimento.Open;
      MissingSaleRejected := False;
      try F.FinalizarPesquisaCobranca;
      except on E: EDatabaseError do
        MissingSaleRejected := Pos('Venda do titulo nao encontrada',E.Message)>0;
      end;
      Check(MissingSaleRejected,'Missing sale silently produced a financial total');
      Check(not F.ZQRecebimento4.ControlsDisabled,'Failure left controls disabled');
      Near(F.XCorrigido.Value,0,'Failure left partial totals');
      Check(not F.CDnegocio.Active,'Failure retained a negotiation');
      C.ExecuteDirect('DELETE FROM recebimento WHERE idrecebimento=6');
      F.FinalizarPesquisaCobranca;
      Near(F.XCorrigido.Value,324,'Recovery after missing sale');
      Writeln('PASS: missing sale rejects calculation, restores controls, and allows retry');
      Summaries := Counter.Summaries;
      F.ZQRecebimento.Edit;
      F.ZQRecebimento.FieldByName('saldo').AsFloat := 400;
      F.ZQRecebimento.Post;
      F.DBcobrancaEnter(nil);
      Check(Counter.Summaries=Summaries+1,'Post reused stale totals');
      Near(F.XCorrigido.Value,432,'Total after balance edit');
      Summaries := Counter.Summaries;
      F.ZQRecebimento.Delete;
      F.DBcobrancaEnter(nil);
      Check(Counter.Summaries=Summaries+1,'Delete reused stale totals');
      Near(F.XCorrigido.Value,0,'Deleted last title left totals');
      Writeln('PASS: Post and Delete invalidate totals even when participant is unchanged');
      F.Label52.Caption := '';
      F.ZQRecebimento.Close;
      F.ZQRecebimento.SQL.Text := 'SELECT * FROM recebimento WHERE 1=0';
      F.ZQRecebimento.Open; F.FinalizarPesquisaCobranca;
      Near(F.nparcelas.Value,0,'Empty search count');
      Near(F.XCorrigido.Value,0,'Empty search total');
      Check(not F.CDnegocio.Active,'Empty search retained negotiation');
      Writeln('PASS: empty search clears totals and negotiation');
      F.ZQRecebimento.Close;
      F.ZQRecebimento.SQL.Text := 'SELECT * FROM recebimento WHERE adversa=1 ORDER BY idrecebimento';
      F.ZQRecebimento.Open;
      D.ZqParticipante.Close;
      D.ZqParticipante.SQL.Text := 'SELECT 1 AS idpaticipante, ''Edited'' AS nome_parte, ''789'' AS doc1';
      D.ZqParticipante.Open;
      F.ZQRecebimento.Last; F.ZQRecebimento.First;
      Check(F.ZQRecebimento.FieldByName('adversanome').AsString='Edited','Name stayed cached');
      Check(F.ZQRecebimento.FieldByName('CPF').AsString='789','Document stayed cached');
      D.ZqParticipante.Close;
      D.ZqParticipante.SQL.Text := 'SELECT 1 AS idpaticipante, NULL AS nome_parte, NULL AS doc1';
      D.ZqParticipante.Open;
      F.ZQRecebimento.Last; F.ZQRecebimento.First;
      Check(F.ZQRecebimento.FieldByName('adversanome').AsString='Client','Null name did not fall back to receipt');
      Check(F.ZQRecebimento.FieldByName('CPF').AsString='','Null document leaked prior value');
      D.ZqParticipante.Close;
      D.ZqParticipante.SQL.Text := 'SELECT 2 AS idpaticipante, ''Other'' AS nome_parte, ''456'' AS doc1';
      D.ZqParticipante.Open;
      F.ZQRecebimento.Last; F.ZQRecebimento.First;
      Check(F.ZQRecebimento.FieldByName('adversanome').AsString='Client','Missing participant lost fallback');
      Check(F.ZQRecebimento.FieldByName('CPF').AsString='','Missing participant inherited another document');
      D.ZqParticipante.Close;
      F.ZQRecebimento.Last; F.ZQRecebimento.First;
      Check(F.ZQRecebimento.FieldByName('adversanome').AsString='Client','Closed source lost fallback');
      Check(F.ZQRecebimento.FieldByName('CPF').AsString='','Closed source leaked document');
      Writeln('PASS: edited, null, missing and closed participant sources preserve lookup semantics');
      Writeln('TOTAL PASS: 16');
    finally
      F.Free; FrmCad_Recebimento := nil;
      D.Free; DM_Tabelas := nil;
      Counter.Free; C.Free;
    end;
  except on E: Exception do begin Writeln('FAIL: ',E.ClassName,': ',E.Message); ExitCode:=1 end end;
end.
