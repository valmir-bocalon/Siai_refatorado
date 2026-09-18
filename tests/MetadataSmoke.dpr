program MetadataSmoke;

{$APPTYPE CONSOLE}

uses
  System.SysUtils, System.Classes, Data.DB, ZConnection, ZDataset,
  uSiaiMetadata in 'uSiaiMetadata.pas',
  uRuntimeFields in 'uRuntimeFields.pas';

type
  { A fixture with the startup owner's class name; no real SIAI module is loaded. }
  TDM_Tabelas = class(TComponent);
  TLegacyMetadataFixture = class(TComponent);
  TDerivedMetadataQuery = class(TZQuery);
  TEvents = class
    Opens: Integer;
    procedure BeforeOpen(DataSet: TDataSet);
    procedure Calculate(DataSet: TDataSet);
  end;

var
  Passed: Integer;
  Connection: TZConnection;
  InitialCalculated, InitialNative, NavigatedCalculated: string;

procedure Check(Value: Boolean; const MessageText: string);
begin
  if not Value then raise Exception.Create(MessageText);
end;

procedure Pass(const Name: string);
begin
  Inc(Passed);
  Writeln('PASS: ', Name);
end;

procedure TEvents.BeforeOpen(DataSet: TDataSet);
begin
  Inc(Opens);
end;

procedure TEvents.Calculate(DataSet: TDataSet);
begin
  DataSet.FieldByName('DoubleId').AsInteger := DataSet.FieldByName('id').AsInteger * 2;
end;

procedure SQLCase(const SQL: string; Accepted: Boolean);
var
  ProbeSQL: string;
begin
  Check(TryBuildMetadataSQL(SQL, ProbeSQL) = Accepted, 'SQL eligibility: ' + SQL);
  if Accepted then
    Check(Copy(ProbeSQL, Length(ProbeSQL) - 6, 7) = 'LIMIT 0', 'Missing LIMIT 0')
  else Check(ProbeSQL = '', 'Rejected SQL returned a probe');
  Pass('SQL eligibility: ' + SQL);
end;

procedure CompareDefinitions(Left, Right: TFieldDefs);
var
  I: Integer;
begin
  Check(Left.Count = Right.Count, 'Field count differs');
  for I := 0 to Left.Count - 1 do
  begin
    Check(Left[I].Name = Right[I].Name, 'Name differs');
    Check(Left[I].DataType = Right[I].DataType, 'Type differs: ' + Left[I].Name);
    Check(Left[I].Size = Right[I].Size, 'Size differs: ' + Left[I].Name);
    Check(Left[I].Precision = Right[I].Precision, 'Precision differs: ' + Left[I].Name);
    Check(Left[I].Required = Right[I].Required, 'Required differs: ' + Left[I].Name);
    Check(Left[I].Attributes = Right[I].Attributes, 'Attributes differ: ' + Left[I].Name);
  end;
end;

procedure QueryCase(const SQL: string; ExpectedRows: Integer; WithParam: Boolean = False);
var
  Original, Baseline, Probe: TZQuery;
  Events: TEvents;
  SavedSQL, MetadataSQL, SavedValue: string;
  SavedDefinitions: TFieldDefs;
  Rows, I: Integer;
begin
  Original := TZQuery.Create(nil);
  Baseline := TZQuery.Create(nil);
  Probe := TZQuery.Create(nil);
  Events := TEvents.Create;
  SavedDefinitions := TFieldDefs.Create(Baseline);
  try
    Writeln('CHECK: ', SQL);
    Original.Connection := Connection;
    Original.SQL.Text := SQL;
    Original.BeforeOpen := Events.BeforeOpen;
    if WithParam then Original.ParamByName('minimum').AsInteger := 1;
    Baseline.Connection := Connection;
    Baseline.SQL.Assign(Original.SQL);
    Baseline.Params.Assign(Original.Params);
    Writeln('CHECK: baseline metadata');
    Baseline.FieldDefs.Update;
    SavedSQL := Original.SQL.Text;
    if WithParam then SavedValue := Original.ParamByName('minimum').AsString;
    Original.Prepare;
    Writeln('CHECK: optimized metadata');
    Check(TryUpdateMetadataOnly(Original), 'Metadata optimization rejected test');
    Check(Original.Prepared, 'Original prepared state changed');
    Check(not Original.Active, 'Metadata opened original query');
    Check(Events.Opens = 0, 'Metadata invoked original BeforeOpen');
    Check(Original.SQL.Text = SavedSQL, 'Original SQL changed');
    if WithParam then
      Check(Original.ParamByName('minimum').AsString = SavedValue, 'Parameter changed');
    CompareDefinitions(Baseline.FieldDefs, Original.FieldDefs);
    SavedDefinitions.Assign(Original.FieldDefs);
    Writeln('CHECK: empty probe');
    Check(TryBuildMetadataSQL(SQL, MetadataSQL), 'Could not build probe');
    Probe.Connection := Connection;
    Probe.SQL.Text := MetadataSQL;
    Probe.Params.Assign(Original.Params);
    Probe.Open;
    Check(Probe.IsEmpty, 'Metadata SQL returned rows');
    Probe.Close;
    { Bind persistent fields as the runtime field setup does, then actually open. }
    for I := 0 to Original.FieldDefs.Count - 1 do
      Original.FieldDefs[I].CreateField(Original);
    Writeln('CHECK: original open');
    Original.Open;
    CompareDefinitions(SavedDefinitions, Original.FieldDefs);
    Rows := 0;
    while not Original.Eof do
    begin
      Inc(Rows);
      Original.Next;
    end;
    Check(Rows = ExpectedRows, Format('Original lost rows: %d expected %d', [Rows, ExpectedRows]));
    Check(Events.Opens = 1, 'Original BeforeOpen not preserved');
    Check(not TryUpdateMetadataOnly(Original), 'Active query should not be touched');
    Original.Close;
    Original.Open;
    Check(Original.IsEmpty = (ExpectedRows = 0), 'Reopen lost rows');
    Check(Original.SQL.Text = SavedSQL, 'Reopen changed SQL');
    Pass('metadata equality + original rows + persistent fields + reopen: ' + SQL);
  finally
    Original.Free;
    Baseline.Free;
    Probe.Free;
    Events.Free;
    SavedDefinitions.Free;
  end;
end;

procedure TestFallback;
var
  Q: TZQuery;
begin
  Q := TZQuery.Create(nil);
  try
    Q.Connection := Connection;
    Q.SQL.Text := 'SELECT * FROM sample LIMIT 1';
    Check(not TryUpdateMetadataOnly(Q), 'Existing limit must use old path');
    Q.FieldDefs.Update;
    Q.Open;
    Check(not Q.IsEmpty, 'Fallback lost records');
    Q.Close;
    Q.SQL.Text := 'SELECT * FROM table_that_does_not_exist';
    Check(not TryUpdateMetadataOnly(Q), 'Probe failure must return fallback');
    Check(not Q.Active, 'Failure opened original query');
    Check(Pos('LIMIT', Q.SQL.Text) = 0, 'Failure changed original SQL');
    Pass('fallback preserves original query, including probe failure');
  finally
    Q.Free;
  end;
end;

procedure TestQuerySubclass;
var
  Q: TDerivedMetadataQuery;
begin
  Q := TDerivedMetadataQuery.Create(nil);
  try
    Q.Connection := Connection;
    Q.SQL.Text := 'SELECT id,label FROM sample ORDER BY id';
    Check(TryUpdateMetadataOnly(Q), 'Metadata rejected a TZQuery subclass');
    Check(Q.FieldDefs.Count = 2, 'Subclass metadata field count differs');
    Check(Q.FieldDefs[0].Name = 'id', 'Subclass metadata first field differs');
    Q.Open;
    Check(Q.RecordCount = 3, 'Subclass metadata lost records on open');
    Pass('metadata supports TZQuery subclasses without changing rows');
  finally
    Q.Free;
  end;
end;

procedure TestReadOnlyQuery;
var
  Q: TZReadOnlyQuery;
begin
  Q := TZReadOnlyQuery.Create(nil);
  try
    Q.Connection := Connection;
    Q.SQL.Text := 'SELECT id,label FROM sample ORDER BY id';
    Check(TryUpdateMetadataOnly(Q), 'Metadata rejected TZReadOnlyQuery');
    Check(Q.FieldDefs.Count = 2, 'Read-only metadata field count differs');
    Check(Q.FieldDefs[0].Name = 'id', 'Read-only metadata first field differs');
    Q.Open;
    Check(Q.RecordCount = 3, 'Read-only metadata lost records on open');
    Pass('metadata supports TZReadOnlyQuery without changing rows');
  finally
    Q.Free;
  end;
end;

procedure TestRuntimeFields(Optimized: Boolean);
var
  Owner: TComponent;
  Lookup, Q: TZQuery;
  Events: TEvents;
  SQL, PreviousLabel: string;
  OriginalField: TField;

  procedure RegisterField(const DataSetName, Name: string; FieldClass: TFieldClass;
    Kind: TFieldKind; const LookupName: string = ''; Native: Boolean = False);
  begin
    RegisterRuntimeField(Owner.ClassType, DataSetName, '', Name, FieldClass, Kind,
      80, 0, False, Name, '', '', '', 20, 'id', LookupName, 'id', 'label',
      LookupName <> '', Native);
  end;

begin
  if Optimized then Owner := TDM_Tabelas.Create(nil)
  else Owner := TLegacyMetadataFixture.Create(nil);
  Events := TEvents.Create;
  try
    Lookup := TZQuery.Create(Owner);
    Lookup.Name := 'LookupQuery';
    Lookup.Connection := Connection;
    Lookup.SQL.Text := 'SELECT id,label FROM sample';
    RegisterRuntimeDataSet(Owner.ClassType, Lookup.Name, True);
    RegisterField(Lookup.Name, 'id', TIntegerField, fkData);
    RegisterField(Lookup.Name, 'label', TStringField, fkData);
    Q := TZQuery.Create(Owner);
    Q.Name := 'MainQuery';
    Q.Connection := Connection;
    Q.SQL.Text := 'SELECT id,label,amount FROM sample ORDER BY id';
    SQL := Q.SQL.Text;
    Q.BeforeOpen := Events.BeforeOpen;
    Q.OnCalcFields := Events.Calculate;
    RegisterRuntimeDataSet(Owner.ClassType, Q.Name, True);
    { A deliberately different requested numeric class must be corrected by metadata. }
    RegisterField(Q.Name, 'id', TFloatField, fkData);
    RegisterField(Q.Name, 'label', TStringField, fkData);
    RegisterField(Q.Name, 'amount', TFloatField, fkData);
    RegisterField(Q.Name, 'DoubleId', TIntegerField, fkCalculated);
    RegisterField(Q.Name, 'CalculatedLabel', TStringField, fkCalculated, Lookup.Name);
    RegisterField(Q.Name, 'NativeLabel', TStringField, fkLookup, Lookup.Name, True);
    EnsureRuntimeFields(Owner);
    Check(Q.Active and Lookup.Active, 'OpenAfterSetup was lost');
    Check(Q.SQL.Text = SQL, 'Runtime setup changed business SQL');
    Check(Q.FieldByName('id') is TIntegerField, 'Metadata no longer selects field class');
    Check(Q.RecordCount = 3, 'Runtime setup lost records');
    if not Optimized then
    begin
      InitialCalculated := Q.FieldByName('CalculatedLabel').AsString;
      InitialNative := Q.FieldByName('NativeLabel').AsString;
    end
    else
    begin
      Check(Q.FieldByName('CalculatedLabel').AsString = InitialCalculated,
        'Initial calculated lookup differs from legacy setup');
      Check(Q.FieldByName('NativeLabel').AsString = InitialNative,
        'Initial native lookup differs from legacy setup');
    end;
    Q.Next;
    Q.First;
    if not Optimized then NavigatedCalculated := Q.FieldByName('CalculatedLabel').AsString
    else Check(Q.FieldByName('CalculatedLabel').AsString = NavigatedCalculated,
      'Calculated lookup differs from legacy after navigation');
    Check(Q.FieldByName('DoubleId').AsInteger = 2, 'Previous OnCalcFields changed');
    Check(Q.FieldByName('NativeLabel').AsString = Q.FieldByName('label').AsString,
      'Native lookup changed');
    OriginalField := Q.FieldByName('id');
    EnsureRuntimeFields(Owner);
    Check(Q.FieldByName('id') = OriginalField, 'Repeated setup replaced a global field');
    Check(Events.Opens = 1, 'Open event changed');
    Q.Close;
    Q.Open;
    Check(Q.RecordCount = 3, 'Runtime reopen lost records');
    PreviousLabel := Q.FieldByName('label').AsString;
    Q.Edit;
    Q.FieldByName('label').AsString := 'edited in isolated test';
    Q.Post;
    Q.Close;
    Q.Open;
    Check(Q.FieldByName('label').AsString = 'edited in isolated test', 'Edit/post failed');
    Q.Edit;
    Q.FieldByName('label').AsString := PreviousLabel;
    Q.Post;
    Pass('runtime field classes, calculated/native lookup, auto-open, reuse, reopen and edit/post; optimized=' + BoolToStr(Optimized, True));
  finally
    Owner.Free;
    Events.Free;
  end;
end;

begin
  try
    SQLCase('SELECT * FROM sample', True);
    SQLCase('select a.id,b.id from sample a join sample b on a.id=b.id order by a.id;', True);
    SQLCase('select ''LIMIT;'' as `limit`, ''a''''b'' as text from sample', True);
    SQLCase('select * from sample where id=:minimum', True);
    SQLCase('select (select max(id) from sample) as maximum', True);
    SQLCase('', False);
    SQLCase('UPDATE sample SET id=0', False);
    SQLCase('select * from sample LIMIT 2', False);
    SQLCase('select * from sample UNION select * from sample', False);
    SQLCase('select * from sample FOR UPDATE', False);
    SQLCase('select * from sample LOCK IN SHARE MODE', False);
    SQLCase('select * INTO OUTFILE ''x'' from sample', False);
    SQLCase('select SQL_CALC_FOUND_ROWS * from sample', False);
    SQLCase('select * from sample; select 1', False);
    SQLCase('select * from sample -- comment', False);
    SQLCase('select * from sample # comment', False);
    SQLCase('select /*!50000 SQL_CALC_FOUND_ROWS */ * from sample', False);
    SQLCase('select @value from sample', False);
    SQLCase('select ''unclosed from sample', False);
    SQLCase('select (id from sample', False);
    SQLCase('select ''a\b'' from sample', False);
    Connection := TZConnection.Create(nil);
    try
      { Dedicated disposable server only. Never read the application's INI. }
      Connection.Protocol := 'mysql-5';
      Connection.LibraryLocation := 'C:\Windows\SysWOW64\libmysql.dll';
      Connection.HostName := '127.0.0.1';
      Connection.Port := 33557;
      Connection.User := 'root';
      Connection.Database := 'siai_metadata_test';
      Writeln('CHECK: connect isolated MySQL');
      Connection.Connect;
      QueryCase('SELECT * FROM sample ORDER BY id', 3);
      QueryCase('SELECT a.*,b.* FROM sample a JOIN sample b ON a.id=b.id', 3);
      QueryCase('SELECT category, SUM(amount) AS total, COUNT(*) AS quantity FROM sample GROUP BY category ORDER BY category', 2);
      QueryCase('SELECT * FROM sample WHERE id>:minimum', 2, True);
      QueryCase('SELECT * FROM sample WHERE id<0', 0);
      QueryCase('SELECT a.id,b.label FROM sample a LEFT JOIN sample b ON b.id=a.id+1', 3);
    TestFallback;
    TestQuerySubclass;
    TestReadOnlyQuery;
    TestRuntimeFields(False);
      TestRuntimeFields(True);
    finally
      Connection.Free;
    end;
    Writeln('TOTAL PASS: ', Passed);
  except
    on E: Exception do
    begin
      Writeln('FAIL: ', E.ClassName, ': ', E.Message);
      ExitCode := 1;
    end;
  end;
end.
