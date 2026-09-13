unit uRecebimentoNomes;

interface

uses
  DB;

procedure PrepararConsultaRecebimentoComNomes(DataSet: TDataSet);

implementation

uses
  SysUtils,
  ZDataset;

procedure PrepararConsultaRecebimentoComNomes(DataSet: TDataSet);
var
  LQuery: TZQuery;
  LSQL: string;
  LSQLLower: string;
  LExtras: string;
  LPrefix: string;
  LFromPos: Integer;
  LParams: TParams;
begin
  if not (DataSet is TZQuery) then
    Exit;

  LQuery := TZQuery(DataSet);
  LSQL := LQuery.SQL.Text;
  LSQLLower := LowerCase(LSQL);
  LFromPos := Pos(' from recebimento', LSQLLower);
  if LFromPos = 0 then
    Exit;

  if Pos(' from recebimento re', LSQLLower) > 0 then
    LPrefix := 're'
  else
    LPrefix := 'Recebimento';

  LExtras := '';
  if Pos('adversanome', LSQLLower) = 0 then
    LExtras := LExtras +
      ', COALESCE((SELECT NULLIF(P.nome_parte, '''') FROM participante P ' +
      'WHERE P.idpaticipante = ' + LPrefix +
      '.adversa), ' + LPrefix + '.nomeadversa, '''') AS adversanome';
  if Pos('nome_loteamento', LSQLLower) = 0 then
    LExtras := LExtras +
      ', COALESCE((SELECT L.apelido FROM loteamento L WHERE L.idloteamento = ' +
      LPrefix + '.idloteamento), '''') AS nome_loteamento';
  if Pos('nomeempreend', LSQLLower) = 0 then
    LExtras := LExtras +
      ', COALESCE((SELECT L.apelido FROM loteamento L WHERE L.idloteamento = ' +
      LPrefix + '.idloteamento), '''') AS nomeempreend';
  if LExtras = '' then
    Exit;

  LParams := TParams.Create(nil);
  try
    LParams.Assign(LQuery.Params);
    LQuery.SQL.Text := Copy(LSQL, 1, LFromPos - 1) + LExtras +
      Copy(LSQL, LFromPos, MaxInt);
    LQuery.Params.Assign(LParams);
  finally
    LParams.Free;
  end;
end;

end.
