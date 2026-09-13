unit CessaoRecebimentos;

interface

uses
  SysUtils, Classes, DB, ZConnection, ZDataset;

type
  { id_antigo identifica o titulo enviado ao banco; id_novo aponta sempre
    para o recebimento atual, inclusive depois de outras cessoes. }
  TConsultaCessao = class
  private
    FConsulta: TZQuery;
    FTemMapa: Boolean;
  public
    constructor Create(AConexao: TZConnection);
    destructor Destroy; override;
    function Resolver(AId: Int64): Int64;
    function TemCessao(AIdAtual: Int64): Boolean;
  end;

procedure PrepararCessaoRecebimentos(AConexao: TZConnection);
procedure RenumerarRecebimentosCessao(AConexao: TZConnection;
  AIdVenda: Int64; ADataCessao: TDateTime);
procedure ExcluirRemessasDaParcela(AConexao: TZConnection; AIdAtual: Int64);

{ Retorna False para evento repetido ou parcela quitada. Quando retorna True,
  deixa uma transacao aberta: o chamador deve confirmar a baixa e a ocorrencia
  juntas, ou executar Rollback. Nunca altera numboleto. }
function IniciarRetornoCessao(AConexao: TZConnection;
  AIdOriginal, AConta: Int64; const ANossoNumero, AOcorrencia, AMotivo: string;
  ADataCredito, ADataOcorrencia: TDateTime; AValor, AJuros, ADesconto: Currency;
  out AIdAtual: Int64): Boolean;

implementation

function NovaConsulta(AConexao: TZConnection): TZQuery;
begin
  Result := TZQuery.Create(nil);
  Result.Connection := AConexao;
end;

function ExisteTabela(Q: TZQuery; const Nome: string): Boolean;
begin
  Q.Close;
  Q.SQL.Text := 'select TABLE_NAME from information_schema.TABLES ' +
    'where TABLE_SCHEMA=DATABASE() and TABLE_NAME=:nome';
  Q.ParamByName('nome').AsString := Nome;
  Q.Open;
  Result := not Q.IsEmpty;
  Q.Close;
end;

procedure ExigirInnoDB(Q: TZQuery; const Tabelas: array of string);
var
  I: Integer;
begin
  for I := Low(Tabelas) to High(Tabelas) do
  begin
    Q.Close;
    Q.SQL.Text := 'select ENGINE from information_schema.TABLES ' +
      'where TABLE_SCHEMA=DATABASE() and TABLE_NAME=:nome';
    Q.ParamByName('nome').AsString := Tabelas[I];
    Q.Open;
    if Q.IsEmpty or not SameText(Q.FieldByName('ENGINE').AsString, 'InnoDB') then
      raise Exception.Create('A operacao exige a tabela ' + Tabelas[I] +
        ' em InnoDB para permitir rollback. Nenhum ID foi trocado.');
  end;
  Q.Close;
end;

procedure CriarEstrutura(Q: TZQuery);
begin
  Q.Close;
  Q.SQL.Text := 'create table if not exists recebimento_id_cessao (' +
    'id_antigo int unsigned not null, id_novo int unsigned not null, ' +
    'idvenda int unsigned not null, data_cessao date not null, ' +
    'criado_em datetime not null, primary key (id_antigo), ' +
    'key idx_cessao_atual (id_novo)) ENGINE=InnoDB';
  Q.ExecSQL;
  Q.SQL.Text := 'create table if not exists recebimento_retorno_cessao (' +
    'evento char(40) not null, id_original int unsigned not null, ' +
    'id_atual int unsigned not null, criado_em datetime not null, ' +
    'primary key (evento), key idx_retorno_atual (id_atual)) ENGINE=InnoDB';
  Q.ExecSQL;
end;

procedure PrepararCessaoRecebimentos(AConexao: TZConnection);
var
  Q: TZQuery;
begin
  if AConexao.InTransaction then
    raise Exception.Create('Conclua a transacao em andamento antes da cessao.');
  Q := NovaConsulta(AConexao);
  try
    ExigirInnoDB(Q, ['recebimento', 'recebimento_historico_excluidos',
      'recebimento_historico_cessao', 'cheque_rec', 'recbxhist']);
    { INSERT/DELETE nao pode disparar efeitos externos ou cascatas. }
    Q.SQL.Text := 'select TRIGGER_NAME from information_schema.TRIGGERS ' +
      'where EVENT_OBJECT_SCHEMA=DATABASE() and EVENT_OBJECT_TABLE=''recebimento''';
    Q.Open;
    if not Q.IsEmpty then
      raise Exception.Create('A tabela recebimento possui triggers. Revise-os antes de trocar os IDs.');
    Q.Close;
    Q.SQL.Text := 'select CONSTRAINT_NAME from information_schema.KEY_COLUMN_USAGE ' +
      'where REFERENCED_TABLE_SCHEMA=DATABASE() and REFERENCED_TABLE_NAME=''recebimento''';
    Q.Open;
    if not Q.IsEmpty then
      raise Exception.Create('A tabela recebimento possui chaves estrangeiras. Revise os vinculos antes da cessao.');
    Q.Close;
    { DDL somente antes da transacao, pois MySQL executa commit implicito. }
    CriarEstrutura(Q);
    ExigirInnoDB(Q, ['recebimento_id_cessao', 'recebimento_retorno_cessao']);
  finally
    Q.Free;
  end;
end;

constructor TConsultaCessao.Create(AConexao: TZConnection);
begin
  inherited Create;
  FConsulta := NovaConsulta(AConexao);
  FTemMapa := ExisteTabela(FConsulta, 'recebimento_id_cessao');
end;

destructor TConsultaCessao.Destroy;
begin
  FConsulta.Free;
  inherited Destroy;
end;

function TConsultaCessao.Resolver(AId: Int64): Int64;
begin
  Result := AId;
  if (AId <= 0) or not FTemMapa then
    Exit;
  FConsulta.Close;
  FConsulta.SQL.Text := 'select id_novo from recebimento_id_cessao where id_antigo=:id';
  FConsulta.ParamByName('id').AsLargeInt := AId;
  FConsulta.Open;
  if not FConsulta.IsEmpty then
    Result := FConsulta.FieldByName('id_novo').AsLargeInt;
  FConsulta.Close;
end;

function TConsultaCessao.TemCessao(AIdAtual: Int64): Boolean;
begin
  Result := False;
  if not FTemMapa then
    Exit;
  FConsulta.Close;
  FConsulta.SQL.Text := 'select id_antigo from recebimento_id_cessao where id_novo=:id limit 1';
  FConsulta.ParamByName('id').AsLargeInt := AIdAtual;
  FConsulta.Open;
  Result := not FConsulta.IsEmpty;
  FConsulta.Close;
end;

procedure AtualizarVinculo(Q: TZQuery; const Tabela, Campo: string;
  AIdAntigo, AIdNovo: Int64);
begin
  Q.Close;
  Q.SQL.Text := 'update ' + Tabela + ' set ' + Campo + '=:novo where ' + Campo + '=:antigo';
  Q.ParamByName('novo').AsLargeInt := AIdNovo;
  Q.ParamByName('antigo').AsLargeInt := AIdAntigo;
  Q.ExecSQL;
end;

procedure RenumerarRecebimentosCessao(AConexao: TZConnection;
  AIdVenda: Int64; ADataCessao: TDateTime);
var
  Q, Parcelas: TZQuery;
  Colunas, Nome, Comparacao: string;
  IdAntigo, IdNovo: Int64;
  TemAutoIncremento: Boolean;
begin
  if not AConexao.InTransaction then
    raise Exception.Create('A troca de IDs deve executar dentro de uma transacao.');
  Q := NovaConsulta(AConexao);
  Parcelas := NovaConsulta(AConexao);
  try
    { Copia as colunas reais, preservando NULL, decimais e campos que nao
      estejam declarados nos datasets do formulario. }
    Q.SQL.Text := 'show columns from recebimento';
    Q.Open;
    Colunas := '';
    Comparacao := '';
    TemAutoIncremento := False;
    while not Q.Eof do
    begin
      Nome := Q.FieldByName('Field').AsString;
      if SameText(Nome, 'idrecebimento') then
        TemAutoIncremento := Pos('auto_increment', LowerCase(Q.FieldByName('Extra').AsString)) > 0
      else
      begin
        Nome := '`' + StringReplace(Nome, '`', '``', [rfReplaceAll]) + '`';
        if Comparacao <> '' then
          Comparacao := Comparacao + ' and ';
        Comparacao := Comparacao + '(binary n.' + Nome + ' <=> binary a.' + Nome + ')';
        if Pos('generated', LowerCase(Q.FieldByName('Extra').AsString)) = 0 then
        begin
          if Colunas <> '' then
            Colunas := Colunas + ',';
          Colunas := Colunas + Nome;
        end;
      end;
      Q.Next;
    end;
    Q.Close;
    if not TemAutoIncremento or (Colunas = '') then
      raise Exception.Create('idrecebimento deve ser AUTO_INCREMENT.');

    Parcelas.SQL.Text := 'select idrecebimento from recebimento ' +
      'where saldo>0 and documento like :venda order by idrecebimento for update';
    Parcelas.ParamByName('venda').AsString := IntToStr(AIdVenda) + '-%';
    Parcelas.Open;
    while not Parcelas.Eof do
    begin
      IdAntigo := Parcelas.FieldByName('idrecebimento').AsLargeInt;
      Q.SQL.Text := 'insert into recebimento (' + Colunas + ') select ' + Colunas +
        ' from recebimento where idrecebimento=:id';
      Q.ParamByName('id').AsLargeInt := IdAntigo;
      Q.ExecSQL;
      if Q.RowsAffected <> 1 then
        raise Exception.Create('Nao foi possivel copiar a parcela ' + IntToStr(IdAntigo));
      Q.SQL.Text := 'select LAST_INSERT_ID() as novo';
      Q.Open;
      IdNovo := Q.FieldByName('novo').AsLargeInt;
      Q.Close;
      Q.SQL.Text := 'select n.idrecebimento from recebimento n, recebimento a ' +
        'where n.idrecebimento=:novo and a.idrecebimento=:antigo and ' + Comparacao;
      Q.ParamByName('novo').AsLargeInt := IdNovo;
      Q.ParamByName('antigo').AsLargeInt := IdAntigo;
      Q.Open;
      if Q.IsEmpty or (IdNovo = IdAntigo) then
        raise Exception.Create('A copia da parcela alterou campos alem do ID. Operacao cancelada.');
      Q.Close;

      AtualizarVinculo(Q, 'recebimento_historico_excluidos', 'idrecebimento', IdAntigo, IdNovo);
      AtualizarVinculo(Q, 'cheque_rec', 'idrecebimento', IdAntigo, IdNovo);
      AtualizarVinculo(Q, 'recbxhist', 'idrecib', IdAntigo, IdNovo);
      { Atualiza tambem os ancestrais; nunca reutiliza um ID de outra cessao. }
      AtualizarVinculo(Q, 'recebimento_id_cessao', 'id_novo', IdAntigo, IdNovo);
      AtualizarVinculo(Q, 'recebimento_retorno_cessao', 'id_atual', IdAntigo, IdNovo);
      Q.SQL.Text := 'insert into recebimento_id_cessao ' +
        '(id_antigo,id_novo,idvenda,data_cessao,criado_em) values (:antigo,:novo,:venda,:data,NOW())';
      Q.ParamByName('antigo').AsLargeInt := IdAntigo;
      Q.ParamByName('novo').AsLargeInt := IdNovo;
      Q.ParamByName('venda').AsLargeInt := AIdVenda;
      Q.ParamByName('data').AsDate := ADataCessao;
      Q.ExecSQL;
      Q.SQL.Text := 'delete from recebimento where idrecebimento=:id';
      Q.ParamByName('id').AsLargeInt := IdAntigo;
      Q.ExecSQL;
      if Q.RowsAffected <> 1 then
        raise Exception.Create('Nao foi possivel substituir a parcela ' + IntToStr(IdAntigo));
      Parcelas.Next;
    end;
  finally
    Parcelas.Free;
    Q.Free;
  end;
end;

procedure ExcluirRemessasDaParcela(AConexao: TZConnection; AIdAtual: Int64);
var
  Q: TZQuery;
begin
  Q := NovaConsulta(AConexao);
  try
    Q.SQL.Text := 'delete rr from remessa_receb rr left join recebimento_id_cessao c ' +
      'on c.id_antigo=rr.idrec where rr.idrec=:id or c.id_novo=:atual';
    Q.ParamByName('id').AsLargeInt := AIdAtual;
    Q.ParamByName('atual').AsLargeInt := AIdAtual;
    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

function IniciarRetornoCessao(AConexao: TZConnection;
  AIdOriginal, AConta: Int64; const ANossoNumero, AOcorrencia, AMotivo: string;
  ADataCredito, ADataOcorrencia: TDateTime; AValor, AJuros, ADesconto: Currency;
  out AIdAtual: Int64): Boolean;
var
  Q: TZQuery;
  Consulta: TConsultaCessao;
  Chave: string;
  Formato: TFormatSettings;
begin
  Result := False;
  if AConexao.InTransaction then
    raise Exception.Create('Conclua a transacao em andamento antes da baixa.');
  Q := NovaConsulta(AConexao);
  Consulta := TConsultaCessao.Create(AConexao);
  try
    ExigirInnoDB(Q, ['recebimento', 'recebimento_id_cessao',
      'recebimento_retorno_cessao', 'recbxhist', 'receb_baixa',
      'caixa', 'contabxrec', 'remessa_receb']);
    AConexao.StartTransaction;
    try
      AIdAtual := Consulta.Resolver(AIdOriginal);
      Q.SQL.Text := 'select saldo from recebimento where idrecebimento=:id for update';
      Q.ParamByName('id').AsLargeInt := AIdAtual;
      Q.Open;
      if Q.IsEmpty then
        raise Exception.Create('Recebimento atual nao encontrado: ' + IntToStr(AIdAtual));
      if Q.FieldByName('saldo').AsFloat <= 0 then
      begin
        Q.Close;
        AConexao.Rollback;
        Exit;
      end;
      Q.Close;
      Formato := TFormatSettings.Create;
      Formato.DecimalSeparator := '.';
      { Independe do nome do arquivo e sobrevive a novas cessoes/limpeza de
        remessas. A conta e o ID original distinguem os titulos bancarios. }
      Chave := IntToStr(AConta) + '|' + IntToStr(AIdOriginal) + '|' +
        Trim(ANossoNumero) + '|' + Trim(AOcorrencia) + '|' + Trim(AMotivo) + '|' +
        FormatDateTime('yyyymmdd', ADataCredito) + '|' +
        FormatDateTime('yyyymmdd', ADataOcorrencia) + '|' +
        FormatFloat('0.00', AValor, Formato) + '|' +
        FormatFloat('0.00', AJuros, Formato) + '|' +
        FormatFloat('0.00', ADesconto, Formato);
      Q.SQL.Text := 'insert ignore into recebimento_retorno_cessao ' +
        '(evento,id_original,id_atual,criado_em) values (SHA1(:evento),:original,:atual,NOW())';
      Q.ParamByName('evento').AsString := Chave;
      Q.ParamByName('original').AsLargeInt := AIdOriginal;
      Q.ParamByName('atual').AsLargeInt := AIdAtual;
      Q.ExecSQL;
      if Q.RowsAffected = 0 then
      begin
        AConexao.Rollback;
        Exit;
      end;
      Result := True;
    except
      AConexao.Rollback;
      raise;
    end;
  finally
    Consulta.Free;
    Q.Free;
  end;
end;

end.
