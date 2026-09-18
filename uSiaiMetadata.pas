unit uSiaiMetadata;

interface

uses Data.DB;

function TryBuildMetadataSQL(const ASQL: string; out AMetadataSQL: string): Boolean;
function TryUpdateMetadataOnly(ADataSet: TDataSet): Boolean;

implementation

uses System.SysUtils, System.Classes, ZDataset, ZAbstractRODataset,
  uSiaiPerformance;

type
  { A API publica de TZReadOnlyQuery expoe as propriedades necessarias em
    published, mas TZAbstractRODataset as declara como protected. Esta classe
    de acesso permite tratar TZQuery e TZReadOnlyQuery pelo mesmo contrato. }
  TZAbstractRODatasetAccess = class(TZAbstractRODataset);

function TryBuildMetadataSQL(const ASQL: string; out AMetadataSQL: string): Boolean;
var
  S, Token: string;
  I, Start, Depth: Integer;
  Quote: Char;
  Closed, FirstToken: Boolean;
begin
  Result := False;
  AMetadataSQL := '';
  S := Trim(ASQL);
  if S = '' then Exit;
  if S[Length(S)] = ';' then S := TrimRight(Copy(S, 1, Length(S) - 1));
  I := 1;
  Depth := 0;
  FirstToken := True;
  while I <= Length(S) do
  begin
    if CharInSet(S[I], [#9, #10, #13, ' ']) then
    begin
      Inc(I);
      Continue;
    end;
    if CharInSet(S[I], ['A'..'Z', 'a'..'z', '_']) then
    begin
      Start := I;
      repeat Inc(I) until (I > Length(S)) or
        not CharInSet(S[I], ['A'..'Z', 'a'..'z', '_', '0'..'9', '$']);
      Token := UpperCase(Copy(S, Start, I - Start));
      if FirstToken then
      begin
        if Token <> 'SELECT' then Exit;
        FirstToken := False;
      end
      else if (Token = 'LIMIT') or (Token = 'UNION') or
        (Token = 'INTO') or (Token = 'PROCEDURE') or (Token = 'FOR') or
        (Token = 'LOCK') or (Token = 'SQL_CALC_FOUND_ROWS') then Exit;
      Continue;
    end;
    if FirstToken then Exit;
    if CharInSet(S[I], ['''', '"', '`']) then
    begin
      Quote := S[I];
      Closed := False;
      Inc(I);
      while I <= Length(S) do
      begin
        { Backslash interpretation depends on the server SQL mode. }
        if S[I] = '\' then Exit;
        if S[I] = Quote then
        begin
          Inc(I);
          if (I <= Length(S)) and (S[I] = Quote) then Inc(I)
          else
          begin
            Closed := True;
            Break;
          end;
        end
        else Inc(I);
      end;
      if not Closed then Exit;
      Continue;
    end;
    { Decline comments (including executable MySQL comments), variables,
      multiple statements and locking clauses. The caller keeps the old path. }
    if CharInSet(S[I], [';', '#', '@', '\']) then Exit;
    if (I < Length(S)) and
      (((S[I] = '-') and (S[I + 1] = '-')) or
       ((S[I] = '/') and (S[I + 1] = '*'))) then Exit;
    if S[I] = '(' then Inc(Depth);
    if S[I] = ')' then
    begin
      Dec(Depth);
      if Depth < 0 then Exit;
    end;
    Inc(I);
  end;
  if FirstToken or (Depth <> 0) then Exit;
  AMetadataSQL := S + sLineBreak + 'LIMIT 0';
  Result := True;
end;

function TryUpdateMetadataOnly(ADataSet: TDataSet): Boolean;
var
  Source: TZAbstractRODatasetAccess;
  Probe: TZQuery;
  SQL: string;
  LStarted: UInt64;
begin
  Result := False;
  { TZQuery e TZReadOnlyQuery compartilham TZAbstractRODataset. Outros
    datasets continuam fora desta otimizacao. }
  if (ADataSet = nil) or not (ADataSet is TZAbstractRODataset) or
     ADataSet.Active then Exit;
  Source := TZAbstractRODatasetAccess(ADataSet);
  if (Source.Connection = nil) or
    (Pos('mysql', LowerCase(Source.Connection.Protocol)) <> 1) then Exit;
  if (Source.DataSource <> nil) or
    ((Source.MasterSource <> nil) and (Source.Params.Count > 0)) then Exit;
  if not TryBuildMetadataSQL(Source.SQL.Text, SQL) then Exit;
  LStarted := PerformanceStart;
  Probe := TZQuery.Create(nil);
  try
    try
      Probe.Connection := Source.Connection;
      Probe.ReadOnly := Source.ReadOnly;
      Probe.Options := Source.Options;
      Probe.Properties.Assign(Source.Properties);
      Probe.ParamCheck := Source.ParamCheck;
      Probe.ParamChar := Source.ParamChar;
      Probe.SQL.Text := SQL;
      Probe.Params.Assign(Source.Params);
      Probe.FieldDefs.Update;
      if Probe.FieldDefs.Count = 0 then Exit;
      Source.FieldDefs.Assign(Probe.FieldDefs);
      Result := True;
      PerformanceElapsed(Source.Name + ': metadados sem registros', LStarted);
    except
      { Preserve the previous discovery path if the server rejects the probe.
        Do not log SQL, parameter values or connection information. }
      PerformanceElapsed(Source.Name + ': metadados pelo caminho original', LStarted);
    end;
  finally
    Probe.Free;
  end;
end;

end.
