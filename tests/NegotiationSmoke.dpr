program NegotiationSmoke;
{$APPTYPE CONSOLE}
uses System.SysUtils, System.Classes, Data.DB, Vcl.Forms, Vcl.ExtCtrls,
  Vcl.DBGrids, Vcl.Controls, Vcl.Dialogs, ZConnection,
  Cad_Recebimento in 'Cad_Recebimento.pas',
  PesqCobranca in 'PesqCobranca.pas';
type
  TProbe = class(TFrmCad_Recebimento)
  public
    procedure AfterConstruction; override;
  end;
  TDialogCloser = class
    Count: Integer;
    procedure Tick(Sender: TObject);
  end;
procedure TProbe.AfterConstruction;
begin
  { Load the real DFM but skip production queries and FormShow. }
end;
procedure TDialogCloser.Tick(Sender: TObject);
begin
  if (Screen.ActiveForm <> nil) and (Screen.ActiveForm.ClassName = 'TMessageForm') then
  begin
    Inc(Count);
    Screen.ActiveForm.ModalResult := mrOk;
  end;
end;
procedure Check(Value: Boolean; const S: string);
begin
  if not Value then raise Exception.Create(S);
end;
var
  F: TProbe;
  C: TZConnection;
  Closer: TDialogCloser;
  Timer: TTimer;
  Passed: Integer;

procedure Pass(const S: string);
begin Inc(Passed); Writeln('PASS: ', S) end;

procedure PrepareRows(const DateSQL: string = 'CURRENT_DATE');
begin
  F.ZQRecebimento.Close;
  F.ZQRecebimento.SQL.Text := 'SELECT id AS idrecebimento, 0 AS Dias, ' + DateSQL +
    ' AS Dt_Vencimento, amount AS saldo, 0 AS Multa_Contrato, 0 AS mora_dia, '+
    'amount AS parcela_corrigida FROM sample WHERE id<=2 ORDER BY id';
  F.ZQRecebimento.Open;
  F.DBcobranca.SelectedRows.Clear;
  F.fpg.Text := '000';
end;

procedure ExpectBlocked(const Name: string);
var Count: Integer;
begin
  Count := Closer.Count;
  F.fpgExit(nil);
  Check(Closer.Count = Count + 1, 'Missing selection/data warning: ' + Name);
  Check(not F.CDnegocio.Active, 'Invalid negotiation remains visible: ' + Name);
  Check(F.XN.Value = 0, 'Stale negotiation total: ' + Name);
  Pass(Name);
end;

procedure TestMasks;
var P: TFrmPesqCobranca; I: Integer;
begin
  P := TFrmPesqCobranca.Create(nil);
  try
    P.OnShow := nil;
    P.DS_TempCliReceb.DataSet := nil;
    for I := P.ZQTempCliReceb.Fields.Count - 1 downto 0 do
      if not SameText(P.ZQTempCliReceb.Fields[I].FieldName, 'Valor') and
         not SameText(P.ZQTempCliReceb.Fields[I].FieldName, 'saldo') then
        P.ZQTempCliReceb.Fields[I].Free;
    P.ZQTempCliReceb.Connection := C;
    P.ZQTempCliReceb.SQL.Text := 'SELECT 1320e0 AS Valor, 45.5e0 AS saldo';
    P.ZQTempCliReceb.Open;
    Check(P.ZQTempCliReceb.FieldByName('Valor').DisplayText = '1.320,00', 'Value mask');
    Check(P.ZQTempCliReceb.FieldByName('saldo').DisplayText = '45,50', 'Balance mask');
    P.ZQTempCliReceb.Close;
    P.ZQTempCliReceb.Open;
    Check(P.ZQTempCliReceb.FieldByName('Valor').DisplayText = '1.320,00', 'Reopen mask');
    Pass('search currency masks and reopen');
  finally P.Free end;
end;

begin
  try
    Application.Initialize;
    UseLatestCommonDialogs := False;
    FormatSettings.DecimalSeparator := ',';
    FormatSettings.ThousandSeparator := '.';
    C := TZConnection.Create(nil);
    Closer := TDialogCloser.Create;
    Timer := TTimer.Create(nil);
    try
      C.Protocol := 'mysql-5';
      C.LibraryLocation := 'C:\Windows\SysWOW64\libmysql.dll';
      C.HostName := '127.0.0.1'; C.Port := 33557;
      C.User := 'root'; C.Database := 'siai_metadata_test'; C.Connect;
      Timer.Interval := 20; Timer.OnTimer := Closer.Tick; Timer.Enabled := True;
      F := TProbe.Create(nil);
      FrmCad_Recebimento := F;
      try
        F.OnShow := nil; F.OnClose := nil;
        F.RGSelecao.OnClick := nil; F.RGSelecao.ItemIndex := 0;
        F.ZQRecebimento.OnCalcFields := nil;
        F.ZQRecebimento_bancario.MasterSource := nil;
        F.ZQRecebimento.Connection := C;
        F.DBcobranca.OnEnter := nil;
        F.DBcobranca.Options := F.DBcobranca.Options + [dgMultiSelect];
        F.Xpagto.DateValue := Date;
        F.cbmontante.Checked := True;
        PrepareRows;
        ExpectBlocked('active row without explicit selection is rejected');
        F.ZQRecebimento.Close;
        ExpectBlocked('closed source is rejected');
        PrepareRows;
        F.DBcobranca.SelectedRows.CurrentRowSelected := True;
        F.fpgExit(nil);
        Check(F.CDnegocio.RecordCount = 1, 'Single selection count');
        Check(F.CDnegociovencimento.AsDateTime = Date, 'Single selection date');
        Check(Abs(F.CDnegociovr_parcela.AsFloat - 12.34) < 0.001, 'Single selection amount');
        Check(not F.CDnegocio.ControlsDisabled, 'Grid controls disabled');
        Pass('one selected installment generates correct date and amount');
        PrepareRows;
        F.DBcobranca.SelectedRows.CurrentRowSelected := True;
        F.ZQRecebimento.Next;
        F.DBcobranca.SelectedRows.CurrentRowSelected := True;
        F.fpgExit(nil);
        Check(F.CDnegocio.RecordCount = 1, 'Same-day consolidation count');
        Check(Abs(F.CDnegociovr_parcela.AsFloat - 69.12) < 0.001, 'Multiple selection total');
        Check(F.ZQRecebimento.RecordCount = 2, 'Source rows lost');
        Pass('two selected installments preserve their total');
        PrepareRows('NULL');
        F.DBcobranca.SelectedRows.CurrentRowSelected := True;
        ExpectBlocked('missing due date is rejected');
        F.cbmontante.Checked := False;
        F.XCorrigido.Value := 300;
        F.fpg.Text := '000'; F.fpgExit(nil);
        Check(F.CDnegocio.RecordCount = 1, 'Unselected-mode count');
        Check(Abs(F.CDnegociovr_parcela.AsFloat - 300) < 0.001, 'Unselected-mode amount');
        Pass('negotiation without per-installment selection still works');
      finally F.Free; FrmCad_Recebimento := nil end;
      TestMasks;
      Writeln('TOTAL PASS: ', Passed);
    finally Timer.Free; Closer.Free; C.Free end;
  except
    on E: Exception do begin Writeln('FAIL: ', E.ClassName, ': ', E.Message); ExitCode := 1 end;
  end;
end.
