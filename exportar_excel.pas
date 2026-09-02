unit exportar_excel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ComObj, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  ZConnection, StdCtrls,IniFiles, ComCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    zconeccao: TZConnection;
    ZqParticipante: TZQuery;
















































    DS_Participante: TDataSource;
    ZQCidade: TZQuery;






    DS_Cidade: TDataSource;
    ProgressBar1: TProgressBar;
    procedure Button1Click(Sender: TObject);
    procedure zconeccaoBeforeConnect(Sender: TObject);

  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
uses
  uRuntimeFields;

{$R *.dfm}

function inttostrZero(numero: Double; tamanho : integer) : string;
Var
  Varvezes, Xtama : integer;
  Varzeros : String;
Begin
  varzeros := '';
  Xtama := length(floattostr(numero));
  for varvezes := Xtama+1 to tamanho do
    Varzeros := Varzeros + '0';
  result := Varzeros+floattostr(numero);
End;

function HexToInt(Hex: string): integer;
begin
  Result := StrToInt('$' + Hex);
end;

function decodifica(texto : string) : String;
Var
  VarTamanho,Varconta : integer;
  montalinha1, montalinha2, montalinha3, varletra : string;
  vardivisor : integer;
begin
  vartamanho := length(texto);
  if (vartamanho>0) and (vartamanho mod 3 = 0) then Begin
    vardivisor := strtoint(floattostr(length(texto)/3));
    montalinha1 := copy(texto,1,vardivisor);
    montalinha2 := copy(texto,(vardivisor*1)+1,vardivisor);
    montalinha3 := copy(texto,(vardivisor*2)+1,vardivisor);
    Try
    montalinha1 := inttostrZero(HexToInt(montalinha1),vardivisor);
    except
     exit;
    end;
    Try
      montalinha2 := inttostrZero(HexToInt(montalinha2),vardivisor);
    except
     exit;
    end;
    Try
    montalinha3 := inttostrZero(HexToInt(montalinha3),vardivisor);
    except
     exit;
    end;
    varletra := '';
    for Varconta := 2 to length(montalinha1) do
      varletra := varletra + chr(strtoint(copy(montalinha1,varconta,1)+copy(montalinha2,varconta,1)+copy(montalinha3,varconta,1)));
    Result := varletra;
  end;
end;


procedure TForm1.Button1Click(Sender: TObject);
Var
linha, coluna: integer;
planilha: variant;
valorCampo: string;
begin
   planilha:= CreateOleObject('Excel.Application');
   planilha.Workbooks.add(1);
   planilha.Cells.Select;
   planilha.Selection.NumberFormat := '@';
   planilha.caption:= 'Exportação de dados para o excel';
   planilha.visible:= true;
   ZQCidade.open;
   ZqParticipante.open;
   ZqParticipante.First;
   ProgressBar1.max:=ZqParticipante.RecordCount;
   ProgressBar1.position:=0;
   for linha:= 0 to ZqParticipante.RecordCount-1 do
   begin
     ProgressBar1.position:=ZqParticipante.RecNo;
     for coluna:= 1 to ZqParticipante.FieldCount do
      begin
         valorCampo:= ZqParticipante.Fields[coluna-1].AsString;
         planilha.cells[linha+2,coluna]:= valorCampo;
      end;
     ZqParticipante.Next;
   end;
   for coluna:=1 to ZqParticipante.FieldCount do
   begin
      valorCampo:= ZqParticipante.Fields[coluna-1].DisplayLabel;
      planilha.cells[1,coluna]:= valorCampo;
   end;
   ProgressBar1.position:=0;   
   planilha.columns.AutoFit;
   ZQCidade.close;
   ZqParticipante.close;
   close;

end;


procedure TForm1.zconeccaoBeforeConnect(Sender: TObject);
Var
  ArqIni : tIniFile;
  varhostname, varsenha, varschemata, VarUsuarioSQL, VarPedeSenha, VarDestino,
  VarNomeArquivo, VarPorta, VarPath, VarNomePasta : string;

begin
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  Try
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );
  if varhostname='' then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','HOSTNAME','localhost');
    varhostname := 'localhost';
  end;
  Finally
  end;
  Try
  varschemata := ArqIni.ReadString('BANCO_DE_DADOS','SCHEMATA', varschemata );
  if varschemata='' then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','SCHEMATA','siai');
    varschemata := 'siai';
  end;
  Finally
  end;
  Try
  VarUsuarioSQL := ArqIni.ReadString('BANCO_DE_DADOS','USUARIO', VarUsuarioSQL );
  if VarUsuarioSQL='' then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','USUARIO','root');
    VarUsuarioSQL := 'root';
  end;
  Finally
  end;
  Try
  VarPorta := ArqIni.ReadString('BANCO_DE_DADOS','PORTA', VarPorta );
  if VarPorta='' then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','PORTA','3306');
    VarPorta := '3306';
  end;
  Finally
  end;
  Try
  varsenha := ArqIni.ReadString('BANCO_DE_DADOS','SENHASQL', varsenha );
  if varsenha='' then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','SENHASQL','000001C70385');
    varsenha := '000001C70385';
  end;
  Finally
  end;
  Try
    VarNomePasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
    if VarNomePasta='' then Begin
      ArqIni.WriteString('BANCO_DE_DADOS','pasta_servidor','master');
      VarNomePasta := 'master';
    end;
  Finally
  end;
  Try
  VarDestino := ArqIni.ReadString('BACKUP','DESTINHO', VarDestino );
  if VarDestino='' then Begin
    ArqIni.WriteString('BACKUP','DESTINHO',VarPath);
    VarDestino := VarPath;
  end;
  Finally
  end;
  Try
  VarPedeSenha := ArqIni.ReadString('BACKUP','PEDESENHA', VarPedeSenha );
  if VarPedeSenha='' then Begin
    ArqIni.WriteString('BACKUP','PEDESENHA','NºO');
    VarPedeSenha := 'NºO';
  end;
  Finally
  end;
  Try
    VarNomeArquivo := ArqIni.ReadString('BACKUP','NOME_ARQUIVO', VarNomeArquivo );
    if VarNomeArquivo='' then Begin
      ArqIni.WriteString('BACKUP','NOME_ARQUIVO','LOTES');
      VarNomeArquivo := 'LOTES';
    end;
  Finally
  end;
  zconeccao.HostName := varhostname;
  zconeccao.Password := decodifica(varsenha);
  zconeccao.Database := varschemata;
  zconeccao.User     := VarUsuarioSQL;
  zconeccao.Port     := strtoint(VarPorta);
  ArqIni.Free;
end;


procedure TForm1.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TForm1, 'ZqParticipante', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantenome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantedoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantedoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantebairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '99\.999\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantebairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '99\.999\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantefone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantefone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantetipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantenacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantenomecidade', 'nomecidade', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantenomeestado', 'nomeestado', TStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantenomecidadecob', 'nomecidadecob', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cob', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantenomeestadocob', 'nomeestadocob', TStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cob', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantenaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantelocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantenomecidadenatural', 'nomecidadenatural', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'naturalidade', 'ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantenomeestadonatural', 'nomeestadonatural', TStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'naturalidade', 'ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteESTADOCIVIL', 'ESTADOCIVIL', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanterenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantefone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '\(00\)0000\-0000;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteNome_Firma', 'Nome_Firma', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipanteInativo', 'Inativo', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantesexo', 'sexo', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantecex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantejb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantejudicial', 'judicial', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantexquadra', 'xquadra', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZqParticipante', 'ZqParticipantexlote', 'xlote', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'ZQCidade', False);
  RegisterRuntimeField(TForm1, 'ZQCidade', 'ZQCidadeidcidade', 'idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQCidade', 'ZQCidadenomecid', 'nomecid', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQCidade', 'ZQCidadeddd', 'ddd', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQCidade', 'ZQCidadeemancipacao', 'emancipacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQCidade', 'ZQCidadeestado', 'estado', TStringField, fkData, 2, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQCidade', 'ZQCidadecepgeral', 'cepgeral', TStringField, fkData, 10, 0, False, '', '', '', '99\.999\-999;1;_', 0, '', '', '', '', False);

end.
