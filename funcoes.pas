unit Funcoes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DBCtrls, shellapi,math,jpeg,
  ZAbstractRODataset, ZAbstractDataset,db,DBClient,
  ZDataset, ComObj, ActiveX;

function SomenteNumeros(const S: string): Boolean;
// quebrar texto em tamanho
function QuebrarTexto(Texto: String; TamLinha, MaxLinhas: Integer): TArray<string>;

function RemoveLetras(Const Texto:String):String;
function GetMotherboardSerialNumber: string;
procedure CloneDataSet(Source: TDataSet; Dest: TClientDataSet);
 // paginação dbgrid
function paginaDbgrid(const ZQuery: TZQuery; pagesize, pagenumber : Integer; Tabela,filtro,ordem: String): TZQuery;
function RemoveNumeros(Const Texto:String):String;
function WinExecAndWait32(FileName: String; WorkDir: String; Visibility: integer): integer;
function CalendarioLunar(day, month, year : Integer):string;
function EstacaoDoAno(Data: TDate): String;

function RemoveAcentos(Str:String): String;
{Remove caracteres acentuados de uma string}

function espacos(varx : integer) : string;
Function transforma(varval : Double; vartama, vardec : integer) : string;
Function Acha_Contabancaria(pl, pc : integer; texto : string) : Boolean;
function masccontabil(varmascara: string) : string; // Coloca a mascara contabil na variavel
Function SemMascara(documento : string) : string;
Function Acha_loteamento(varloteamento: string): boolean;
Function Acha_Cidade(varcidade: string): boolean;
Function Verif_senha(Vargrupo,vardescricao,varhistorico: string): boolean;
Function Verif_doc(Dados: string; vermens: boolean ): boolean;  { Verifica se o CNPJ ou CPF é válido }
Function Empty( Dados: string ): boolean;  // verifica se uma variavel string está vazia
function AllTrim( Dados: string ): string; //  Retira os espacos em branco da direita
Function SysComputerName: string; // Retorna o nome do computador
Function atualizar_estrutra : boolean;
Function cripto_data(Ddata : Tdatetime) : string;
function descripto_data(Sdata : string) : TdateTime;
function decodifica(texto : string) : String;
function StrIsInteger(const S: string): boolean;
function inttostrZero(numero: Double; tamanho : Integer) : string;
function HexToInt(Hex: string): integer;
Function CSenha : Boolean;
function diasemana(dia : Tdatetime) :string; //  dia da semana por extenso
function GetFileDate(TheFileName: String): TdateTime;  {Retorna a data e a hora de um arquivo}
Function AchaCLiente(PL, PC: integer; Cli: String) : boolean;
Function AchaCorretor(PL, PC: integer; Corr :String) : boolean;
Function truncar(num : Double; dec : integer) : Double;
function dataporextenso(datadig : TdateTime) : string; // retorna data por extenso
function BuscaTroca(Text,Busca,Troca : string) : string;
Function AchaPlanoDeConta(PL, PC: integer; varclas, pln, varCD :String) : boolean;
Function tiraacento(texto: string) : string;
function StrIsnumero(const S: string): boolean;
function Transform( Conteudo: Extended; const Mascara: string ): string; // Alinha algarismos a direita
function Space(N: integer): string;
Function Calcjuros(dtI,dtF :Tdate; varvalor, varjur :double) : Double;
function tiramascara(varx : string) : string;
//Function extenso(valor : double) : string;
{Retorna o Valor em Extenso}
Function Extenso(Valor : Extended): String;
function numeroporextenso(numero : Integer) : string;
function ExRound(Value: Extended; Decimals: Integer): Extended;
{ Reproduz um string varias vezes }
function Repl(C: string; Tamanho: integer): string;
{ Coloca zeros a esquerda }
function StrZero(N: integer; Tamanho: integer): string;
{ Retorna o Dia de uma data }
function Dia( Data: TDateTime ): string;
{ Retorna o Mes de uma data }
function Mes( Data: TDateTime ): string;
{ Retorna o Mes/Ano de uma data }
function MesAno( Data: TDateTime ): string;
{ Retorna o Ano de uma data }
function Ano( Data: TDateTime ): string;
{ Retorna o ultimo dia do mes }
function UltimoDiaDoMes( MesAno: string ): string;
{ Janela para mensagens }
procedure Mensagem( Texto: string );
procedure AbrirModal(Origem, Formulario: TCustomForm);
{mes por extenso}
function MesExtenso( Mes:Word ) : string;
Function SIMNAO(texto, opcao:string) : Boolean; //
function pergunta(texto,resp : string) :string; // pergunta string, nome do aquirvo por exemplo!!
function CHARREM(Texto     : string): string;
// Remove caracteres de uma string deixando apenas numeros
Function RemoveChar(Const Texto:String):String;
//Removendo espaços/caracteres de uma string
function TrimChar(texto: string; delchar: char): string;
// formatar linha de boleto
function Formatar(Texto : string; TamanhoDesejado : integer; AcrescentarADireita : boolean = true; CaracterAcrescentar : char = ' ') : string;
// executa programa externo
procedure ExecutePrograma(Nome, Parametros: String);
procedure Delay(dwMilliseconds: Longint);

function DifDias(DataVenc:TDateTime; DataAtual:TDateTime): String;
// trocar virgula por ponto
function TrocaVirgPPto(Valor: string): String;
function Calc_Dias(DataVencimento:TDateTime; DataAtual:TDateTime): String;
function CapturaTela: TBitmap;

type
  TFrm_Funcoes = class(TForm)
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Funcoes: TFrm_Funcoes;

implementation

uses principal, tabelas, DigSenha, NovaSenha, AchaCidade, Estrutura,
  ContraSenha, AchaParticipante, Achaloteamento, AchaCorretor,
  AchaPlanoDeContas, Acha_Contabancaria, FINANmsg, PerguntaSIMNAO,
  Pergunta, AchaIgpm, Participante, uRuntimeFields;

{$R *.dfm}


procedure AbrirModal(Origem, Formulario: TCustomForm);
begin
  if Formulario = nil then
    Exit;

  if (Origem <> nil) and (Origem <> Formulario) then
  begin
    Formulario.PopupMode := pmExplicit;
    Formulario.PopupParent := Origem;
  end;

  Formulario.ShowModal;
end;
function SomenteNumeros(const S: string): Boolean;
var
  I: Integer;
begin
  Result := Trim(S) <> '';

  for I := 1 to Length(Trim(S)) do
  begin
    if not (Trim(S)[I] in ['0'..'9']) then
    begin
      Result := False;
      Exit;
    end;
  end;
end;
// quebrar texto em tamanho
function QuebrarTexto(Texto: String; TamLinha, MaxLinhas: Integer): TArray<string>;
var
  Palavras: TArray<string>;
  Linha, Palavra: string;
  i, LinhaAtual: Integer;
begin
  SetLength(Result, 0);

  Texto := Trim(Texto);
  if Texto = '' then Exit;

  Palavras := Texto.Split([' ']);

  Linha := '';
  LinhaAtual := 0;

  for i := 0 to High(Palavras) do
  begin
    Palavra := Palavras[i];

    if Length(Linha + ' ' + Palavra) <= TamLinha then
    begin
      if Linha = '' then
        Linha := Palavra
      else
        Linha := Linha + ' ' + Palavra;
    end
    else
    begin
      SetLength(Result, Length(Result) + 1);
      Result[High(Result)] := UpperCase(Linha);

      Inc(LinhaAtual);
      if (MaxLinhas > 0) and (LinhaAtual >= MaxLinhas) then Exit;

      Linha := Palavra;
    end;
  end;

  if Linha <> '' then
  begin
    SetLength(Result, Length(Result) + 1);
    Result[High(Result)] := UpperCase(Linha);
  end;
end;

//-------SOMENTE NUMEROS-----------------
function RemoveLetras(Const Texto:String):String;
//
// Remove caracteres de uma string deixando apenas numeros
//
var
I: integer;
S: string;
begin
  S := '';
  for I := 1 To Length(Texto) Do
  begin
  if (Texto[I] in ['0'..'9']) then
  begin
  S := S + Copy(Texto, I, 1);
  end;
  end;
  result := S;
end;
function GetMotherboardSerialNumber: string;
var
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
begin
  Result := '';
  try
    // Inicializa a COM library
    CoInitialize(nil);
    try
      // Cria o objeto WMI
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      // Conecta ao serviço WMI
      FWMIService := FSWbemLocator.ConnectServer('.', 'root\CIMV2', '', '');
      // Executa a consulta WMI
      FWbemObjectSet := FWMIService.ExecQuery('SELECT SerialNumber FROM Win32_BaseBoard', 'WQL', 0);
      // Enumera os resultados
      oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
      if oEnum.Next(1, FWbemObject, iValue) = 0 then
      begin
        // Obtém o valor do SerialNumber
        Result := FWbemObject.SerialNumber;
      end;
    finally
      // Libera a COM library
      CoUninitialize;
    end;
  except
    on E: Exception do
      Writeln(E.Classname, ': ', E.Message);
  end;
end;

procedure CloneDataSet(Source: TDataSet; Dest: TClientDataSet);
var
  I: Integer;
begin
  // Define a estrutura do ClientDataSet igual à do dataset de origem
  Dest.FieldDefs.Clear;
  Dest.FieldDefs.Assign(Source.FieldDefs);
  Dest.CreateDataSet;

  // Copia os dados do dataset de origem para o ClientDataSet
  Source.DisableControls;
  try
    Dest.DisableControls;
    try
      Source.First;
      while not Source.Eof do
      begin
        Dest.Append;
        for I := 0 to Source.FieldCount - 1 do
          Dest.Fields[I].Value := Source.Fields[I].Value;
        Dest.Post;
        Source.Next;
      end;
    finally
      Dest.EnableControls;
    end;
  finally
    Source.EnableControls;
  end;
end;

// paginação de dbgrid
function paginaDbgrid(const ZQuery: TZQuery; pagesize, pagenumber : Integer; Tabela,filtro,ordem: String): TZQuery;
var
  SQL: string;
begin
  // Verificar se PageSize e PageNumber são válidos
  if (PageSize <= 0) or (PageNumber <= 0) then
    raise Exception.Create('PageSize e PageNumber devem ser maiores que zero.');
  if (filtro<>emptystr) and (ordem<>emptystr) then
      SQL := Format('SELECT * FROM '+tabela+' where '+filtro+' order by '+ordem+' LIMIT %d OFFSET %d', [PageSize, (PageNumber - 1) * PageSize])
  else if (filtro<>emptystr) and (ordem=emptystr) then
      SQL := Format('SELECT * FROM '+tabela+' where '+filtro+'  LIMIT %d OFFSET %d', [PageSize, (PageNumber - 1) * PageSize])
  else if (filtro=emptystr) and (ordem<>emptystr) then
      SQL := Format('SELECT * FROM '+tabela+' order by '+ordem+' LIMIT %d OFFSET %d', [PageSize, (PageNumber - 1) * PageSize])
  else
     SQL := Format('SELECT * FROM '+tabela+' LIMIT %d OFFSET %d', [PageSize, (PageNumber - 1) * PageSize]);
  ZQuery.Close;
  ZQuery.SQL.Clear;
  ZQuery.SQL.text:=sql;
  ZQuery.open;
end;

function RemoveNumeros(Const Texto:String):String;
//
// Remove numeros de uma string deixando apenas letras
//
var
I: integer;
S: string;
begin
  S := '';
  for I := 1 To Length(Texto) Do
  begin
    if (Texto[I] in ['A'..'Z'])or (Texto[I] =' ') or (Texto[I] ='.') then
    begin
      S := S + Copy(Texto, I, 1);
    end;
  end;
  result := S;
end;
function WinExecAndWait32(FileName: String; WorkDir: String; Visibility: integer): integer;
var
   zAppName: array[0..512] of char;
   zCurDir: array[0..255] of char;
   StartupInfo: TStartupInfo;
   ProcessInfo: TProcessInformation;
begin
  StrPCopy(zAppName,FileName);
  StrPCopy(zCurDir,WorkDir);
  FillChar(StartupInfo,Sizeof(StartupInfo),#0);
  StartupInfo.cb:=Sizeof(StartupInfo);
  StartupInfo.dwFlags:=STARTF_USESHOWWINDOW;
  StartupInfo.wShowWindow:=Visibility;

  if not CreateProcess(nil,zAppName,nil,nil,False,CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS,nil,zCurDir,StartupInfo,ProcessInfo) then
	 Result:=-1
  else
  begin
	 WaitforSingleObject(ProcessInfo.hProcess,INFINITE);
	 GetExitCodeProcess(ProcessInfo.hProcess,DWORD(Result));
  end;

end;



function CalendarioLunar(day, month, year : Integer):string;
 const
   ages: array[0..18] of Integer = (18, 0, 11, 22, 3, 14, 25, 6, 17, 28, 9, 20, 1, 12, 23, 4, 15, 26, 7);
   offsets: array[0..11] of Integer = (-1, 1, 0, 1, 2, 3, 4, 5, 7, 7, 9, 9);
   description: array[0..7] of string= ('Lua Nova (totalmente escura)',
                   'Para Quarto Crescente (aumentando para cheia)',
                   'Quarto Crescente (aumentando para cheia)',
                   'Lua Crescente (aumentando para cheia)',
                   'Cheia (claridade total)',
                   'Para Quarto Minguante (decrescente da cheia)',
                   'Quarto Minguante (diminuição do total)',
                   'Para Lua Nova ');
   months: array [1..12] of string = ('Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Otu', 'Nov', 'Dez');

 var
   days_into_phase,
   light, index: integer;

   status, date : string;

 begin
   if day = 31 then
     day := 1;

   days_into_phase  := ((ages[(year+1) mod 19]) + ( (day + offsets[month-1])));

   if not year < 1900 then
     days_into_phase  := days_into_phase mod 30;

   index := Trunc((days_into_phase + 2) * 16/59.0);

   if index > 7 then
         index := 7;
   status := description[index];

   Light := Trunc(2 * days_into_phase * 100/29);

   if light > 100 then
     light := abs(light-200);

   date := IntToStr(day)+ months[month]+IntToStr(year);

//   Result := 'Fase da lua em '+date+ ' é '+status+ ', claridade = '+IntToStr(light)+'%';
   Result := 'Fase da lua é: '+status+ ', claridade = '+IntToStr(light)+'%';
 end;


function EstacaoDoAno(Data: TDate): String;
 var Mes: String;
 begin
    Mes := FormatDateTime('mmdd', data);
    if (Mes >= '0320') and (Mes < '0621') then
       Result := 'Outono'
    else if (Mes >= '0621') and (Mes < '0922') then
       Result := 'Inverno'
    else if (Mes >= '0922') and (Mes < '1221') then
       Result := 'Primavera'
    else
       Result := 'Verão';
 end;

function RemoveAcentos(Str:String): String;
{Remove caracteres acentuados de uma string}
Const ComAcento = 'àâêôûãõáéíóúçüÀÂÊÔÛÃÕÁÉÍÓÚÇÜºª';
      SemAcento = 'aaeouaoaeioucuAAEOUAOAEIOUCU..';
Var
x : Integer;
Begin
For x := 1 to Length(Str) do
    Begin
    if Pos(Str[x],ComAcento)<>0 Then
       begin
       Str[x] := SemAcento[Pos(Str[x],ComAcento)];
       end;
    end;
Result := Str;
end;


function espacos(varx : integer) : string;
var
  vary : string;
  varz : integer;
Begin
  vary := '';
  for varz :=1 to varx do
    vary := vary + ' ';
  result := vary;
End;

Function transforma(varval : Double; vartama, vardec : integer) : string;
Var
  varstrV, varstrD : string;
  varx, vary: integer;
Begin
  vary := pos(',',floattostr(varval));
  varstrD := '';
  if vary=0 Then
    varstrV := formatfloat('###,###,###,###,###',varval)
  else
  Begin
    varstrV := formatfloat('###,###,###,###,###',strtofloat(copy(floattostr(varval),1,vary-1)));
    varstrD := copy(floattostr(varval),vary+1,length(floattostr(varval)));
  end;
  if length(varstrV)=0 Then
    varstrV := '0';
  if length(varstrD)>vardec Then
    varstrD := copy(varstrD,1,vardec)
  else
    for varx := 1 to vardec-length(varstrD) do
      varstrD := varstrD+'0';
  if vardec>0 Then
    varstrV := varstrV+','+varstrD;
  if length(varstrV)<vartama Then
    for varx := 1 to (vartama-length(varstrV))*2 do
    VarstrV := ' '+varstrV;
  result := varstrV;
End;

Function Acha_Contabancaria(pl, pc : integer; texto : string) : Boolean;
Var
  Varx : integer;
  varT : string;
Begin
  for varx:=1 to length(texto) do
    if copy(texto,varx,1)='-' Then
      break;
  varT := copy(texto,1,varx-1);
  DM_Tabelas.ZQAchaContaBanc.SQL.Clear;
  DM_Tabelas.ZQAchaContaBanc.SQL.Add('Select  idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,nometitular,aberta,senha_interna,senha_conta,ativa,');
  DM_Tabelas.ZQAchaContaBanc.SQL.Add('        jurosemboleto,taxadiaria,valoroupercent,carteira,boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,Apelido,cod_transmissao,');
  DM_Tabelas.ZQAchaContaBanc.SQL.Add('        conta_bancaria.complemento,moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote,');
  DM_Tabelas.ZQAchaContaBanc.SQL.Add('        idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  DM_Tabelas.ZQAchaContaBanc.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento_cob,cadastrado,Nome_Firma,');
  DM_Tabelas.ZQAchaContaBanc.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao ');
  DM_Tabelas.ZQAchaContaBanc.SQL.Add('  from conta_bancaria join participante on idparticipante=idpaticipante');
  DM_Tabelas.ZQAchaContaBanc.SQL.Add(' where nomebanco like '+quotedstr('%'+texto+'%'));
  DM_Tabelas.ZQAchaContaBanc.SQL.Add(' or nometitular like '+quotedstr('%'+texto+'%'));
  DM_Tabelas.ZQAchaContaBanc.SQL.Add(' or n_conta like '+quotedstr(texto+'%'));
  DM_Tabelas.ZQAchaContaBanc.SQL.Add(' or n_conta like '+quotedstr(varT+'%'));
  DM_Tabelas.ZQAchaContaBanc.Open;
  if DM_Tabelas.ZQAchaContaBanc.RecordCount=0 Then
    Result := False
  else if DM_Tabelas.ZQAchaContaBanc.RecordCount=1 Then
    Result := True
  else Begin
    if Frm_Acha_Contabancaria = nil then
      Frm_Acha_Contabancaria := TFrm_Acha_Contabancaria.Create(Application);
    Frm_Acha_Contabancaria.ShowModal;
    if Frm_Acha_Contabancaria.Label2.Caption = 'T' then
      Result := True
    else
      Result := False;
  end;
end;

function masccontabil(varmascara: string) : string;
Begin
  masccontabil := varmascara;
  if Length(varmascara) = 3 then
    masccontabil := copy(varmascara,1,1)+'.'+copy(varmascara,2,2);
  if Length(varmascara) = 6 then
    masccontabil := copy(varmascara,1,1)+'.'+copy(varmascara,2,2)+'.'+copy(varmascara,4,3);
  if Length(varmascara) = 10 then
    masccontabil := copy(varmascara,1,1)+'.'+copy(varmascara,2,2)+'.'+copy(varmascara,4,3)+'.'+copy(varmascara,7,4);
  if Length(varmascara) = 15 then
    masccontabil := copy(varmascara,1,1)+'.'+copy(varmascara,2,2)+'.'+copy(varmascara,4,3)+'.'+copy(varmascara,7,4)+'.'+copy(varmascara,11,5);
End;

Function SemMascara(documento : string) : string;
Var
  Varcont : integer;
  VarDoc : string;
Begin
  VarDoc := '';
  for Varcont := 1 to length(documento) do
    if pos(copy(documento,Varcont,1),'1234567890')>0 Then
      VarDoc := VarDoc + copy(documento,Varcont,1);
  Result := VarDoc;
end;

Function Acha_loteamento(varloteamento: string): boolean;
Begin
  DM_tabelas.ZQAchaLotea.SQL.Clear;
  DM_Tabelas.ZQAchaLotea.SQL.Add('select idloteamento,cidade_idcidade,matriculaloteamento,datacadastro,datainalguracao,dataconclusao,logotipo,mapa,pasta_mapa,testemunha1,testemunha2,testemunha3,');
  DM_Tabelas.ZQAchaLotea.SQL.Add('       testemunha4,Observincorp,bairro,tipo,nomeloteamento,codcontabancaria,apelido,Perc_comissao ');
  DM_Tabelas.ZQAchaLotea.SQL.Add(' from loteamento');
  DM_Tabelas.ZQAchaLotea.SQL.Add(' where nomeloteamento like '+quotedstr('%'+varloteamento+'%'));
  DM_Tabelas.ZQAchaLotea.SQL.Add(' or apelido like '+quotedstr('%'+varloteamento+'%')+' order by nomeloteamento');
  DM_Tabelas.ZQAchaLotea.Open;
  if DM_Tabelas.ZQAchaLotea.RecordCount=1 Then
    Result := True
  else  if DM_Tabelas.ZQAchaLotea.RecordCount=0 Then
    result := False
  else Begin
    Frm_Achaloteamento.ELoteamento.Text := varloteamento;
    Frm_Achaloteamento.ShowModal;
    if (Frm_Achaloteamento.Label1.Caption = 'F') or (DM_tabelas.ZQAchaLotea.RecordCount=0)  Then
       Result := false
    else
       Result := True;
  end;
End;

Function Acha_Cidade(varcidade: string): boolean;
Begin
  DM_Tabelas.ZQAchaCidade.SQL.Clear;
  DM_Tabelas.ZQAchaCidade.SQL.Add('select idcidade,nomecid,ddd,emancipacao,estado,cepgeral from cidade');
  DM_tabelas.ZQAchaCidade.SQL.Add(' Where nomecid like '+quotedstr(varcidade+'%')+' order by nomecid');
  DM_tabelas.ZQAchaCidade.Open;
  if DM_tabelas.ZQAchaCidade.RecordCount=1 Then
    Result := True
  else Begin
    if Frm_AchaCidade = nil then
      Frm_AchaCidade := TFrm_AchaCidade.Create(Application);
    Frm_AchaCidade.ECidade.Text := varcidade;
    Frm_AchaCidade.ShowModal;
    varcidade:=DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
    Frm_AchaCidade.ECidade.Text := varcidade;

    DM_Tabelas.ZQAchaCidade.SQL.Clear;
    DM_Tabelas.ZQAchaCidade.SQL.Add('select idcidade,nomecid,ddd,emancipacao,estado,cepgeral from cidade');
    DM_tabelas.ZQAchaCidade.SQL.Add(' Where nomecid like '+quotedstr(varcidade+'%')+' order by nomecid');
    DM_tabelas.ZQAchaCidade.Open;

    if (Frm_AchaCidade.Label2.Caption = 'F') or (DM_tabelas.ZQAchaCidade.RecordCount=0)  Then
      Result := false
    else
        Result := True;
  end;
end;

Function Verif_senha(Vargrupo,vardescricao,varhistorico: string): boolean;
Begin
  if not DM_Tabelas.ZQLoc_Senha.Locate('grupo,descricao',VarArrayOf([vargrupo,vardescricao]),[]) Then Begin
    DM_Tabelas.ZQLoc_Senha.Insert;
    DM_Tabelas.ZQLoc_Senha.FieldByName('grupo').AsString := vargrupo;
    DM_Tabelas.ZQLoc_Senha.FieldByName('descricao').AsString := vardescricao;
    DM_Tabelas.ZQLoc_Senha.Post;
    Frm_NovaSenha.Label3.Caption := 'Grupo....: '+vargrupo;
    Frm_NovaSenha.Label4.Caption := 'Descricao: '+vardescricao;
    Frm_NovaSenha.ShowModal;
  end;
  result := True;
  if DM_Tabelas.ZQLoc_Senha.FieldByName('pedesenha').AsString = 'S' Then Begin
    Frm_DigSenha.Label4.Caption := 'Grupo: '+Vargrupo+' - Descricao: '+vardescricao;
    Frm_DigSenha.ShowModal;
    if Frm_DigSenha.Label6.Caption = 'F' Then
      Result := False;
    DM_Tabelas.ZQLogin.Insert;
    DM_Tabelas.ZQLogin.FieldByName('local_senha').AsLargeInt := DM_Tabelas.ZQLoc_Senha.FieldByName('idlocais_com_senha').AsLargeInt;
    DM_Tabelas.ZQLogin.FieldByName('data_2').AsDateTime := date;
    DM_Tabelas.ZQLogin.FieldByName('hora').AsDateTime := Time;
    DM_Tabelas.ZQLogin.FieldByName('maquina').AsString :=  SysComputerName();
    DM_Tabelas.ZQLogin.FieldByName('usuario').AsLargeInt := DM_tabelas.ZQUsuTemp.FieldByName('idusuario').AsLargeInt;
    DM_Tabelas.ZQLogin.FieldByName('passou').AsString := Frm_DigSenha.Label6.Caption;
    DM_Tabelas.ZQLogin.FieldByName('detalhe').AsString := varhistorico;
    DM_Tabelas.ZQLogin.Post;
    exit;
  end;
End;

Function Verif_doc(Dados: string; vermens: boolean ): boolean;  { Verifica se o CNPJ ou CPF é válido }
var
  Check, varsemmasc : string;
  Resto, Digito, Digito1, Digito2, varx, varcont : integer;
  Multiplicador1, Multiplicador2, d1, d4 : Integer;
Begin
  varsemmasc := SemMascara(Dados);
  Result := True;
  Resto := 0;
  Digito := 0;
  Digito1 := 0;
  Digito2 := 0;
  if empty(varsemmasc) Then
    exit
  else if length(varsemmasc)= 11 Then Begin
    d1 := 0; d4 := 0; varx := 1;
    varsemmasc := '';
    for varcont := 1 to Length( Dados )-2 do begin
      if Pos( Copy( dados, varcont, 1 ), '/-._ ' ) = 0 then begin
        d1 := d1 + ( 11 - varx ) * StrToInt( Copy( dados, varcont, 1 ) );
        d4 := d4 + ( 12 - varx ) * StrToInt( Copy( dados, varcont, 1 ) );
        varx := varx+1;
        varsemmasc := varsemmasc + Copy(dados,varcont,1);
      end;
    end;
    if empty(varsemmasc) Then Begin
      Result := True;
      exit;
    end;
    resto := (d1 mod 11);
    if resto < 2 then begin
       digito1 := 0;
    end
    else
    begin
      digito1 := 11 - resto;
    end;
    d4 := d4 + 2 * digito1;
    resto := (d4 mod 11);
    if resto < 2 then begin
      digito2 := 0;
    end
    else
    begin
      digito2 := 11 - resto;
    end;
    Check := IntToStr(Digito1) + IntToStr(Digito2);
    if Check <> copy(dados,succ(length(dados)-2),2) then begin
      if vermens Then
        Showmessage('C.P.F. invalido...   Tente outra vez!!');
      Result := False;
    end;
  end
  else if length(varsemmasc)= 14 Then Begin
    for varx := 1 to 12 do begin
      if varx < 5 then Multiplicador1 := 6
      else Multiplicador1 := 14;
      if varx < 6 then Multiplicador2 := 7
      else Multiplicador2 := 15;
      Digito1 := Digito1 + StrToInt(Copy( varsemmasc, varx, 1 ) ) * ( Multiplicador1 - varx );
      Digito2 := Digito2 + StrToInt(Copy( varsemmasc, varx, 1 ) ) * ( Multiplicador2 - varx );
    end;
    Resto := Digito1 mod 11;
    if Resto < 2 then Digito := 0
    else Digito := 11 - Resto;
    Digito2 := Digito2 + ( 2 * Digito );
    Resto := Digito2 mod 11;
    Digito := ( Digito * 10 );
    if Resto >= 2 then
      Digito := Digito + ( 11 - Resto );
    if Digito <> StrToInt( Copy( varsemmasc, 13, 2 ) ) then Begin
      Result := False;
      if vermens Then
        mensagem('C.N.P.J. invalido...   Tente outra vez!!');
    end;
  end
  else Begin
    if vermens Then
      mensagem('Documento invalido...');
    Result := False;
  end;
end;

function Empty( Dados: string ): boolean;
begin
  if Length( Trim( Dados ) ) = 0 then
    Empty := True
  else
    Empty := False;
end;

function AllTrim( Dados: string ): string;
var
  Contar: integer;
begin
  Dados := Trim( Dados );
  for Contar := 1 to Length( Dados ) do
  begin
    if Copy( Dados, Contar, 1 ) <> ' ' then
      Break;
    Dados := Copy( Dados, Contar + 1, Length( Dados ) - 1 );
    Application.ProcessMessages;
  end;
  AllTrim := Dados;
end;

function SysComputerName: string;
var
  I: DWord;
begin
  I := MAX_COMPUTERNAME_LENGTH + 1;
  SetLength(Result, I);
  Windows.GetComputerName(PChar(Result), I);
  Result := string(PChar(Result));
end;

Function atualizar_estrutra : boolean;
Begin
  senhaestrut := false;
  FrmEstrutura.ShowModal;
End;


Function cripto_data(Ddata : Tdatetime) : string;
var
  varx : double;
Begin
  varx := strtofloat(copy(datetostr(Ddata),1,2)+copy(datetostr(Ddata),4,2)+copy(datetostr(Ddata),9,2));
  Result := floattostr(varx*1506);
End;

function descripto_data(Sdata : string) : TdateTime;
Var
  Vx : String;
Begin
  if StrIsInteger(Sdata) Then Begin
    Vx := floattostr(strtofloat(Sdata)/1506);
    if length(vx)=5 Then
      vx := '0'+vx;
    if (pos(',',vx)>0) or (length(vx)<>6) Then
      Result := strtodate('01/01/0001')
    else
      Result := strtodate(copy(vx,1,2)+'/'+copy(vx,3,2)+'/20'+copy(vx,5,2));
  end
  else
    Result := strtodate('01/01/0001');
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

function StrIsInteger(const S: string): boolean;
 var
 tamanho, x : integer;
 l, p : string;
begin
  p := s;                            // Esta função verifica de uma
  tamanho := length(p);
  Result := True;                    // string tem apenas caracteres
  for x := 1 to tamanho do           // numericos.
  begin
    l := copy(p,x,1);
    if  pos(l,'0123456789')= 0 then begin
      Result := FALSE;
      exit;
    end;
  end;
end;

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

Function CSenha : Boolean;
Var
  VarCSenha, VarSSenha: String;
  VarFinal, Varmes : Integer;
  VarProxima : TDateTime;
begin
  DM_Tabelas.ZQEmpresa.Open;
  DM_Tabelas.ZQEmpresa.Edit;
  if empty(DM_Tabelas.ZQEmpresa.FieldByName('limite').AsString) Then Begin
    DM_Tabelas.ZQEmpresa.FieldByName('dia').AsDateTime := date;
    DM_Tabelas.ZQEmpresa.FieldByName('limite').AsString :=  cripto_data(date);
  end;
  {if date<DM_Tabelas.ZQEmpresadia.Value Then
  Begin
    Showmessage('Data do computador está incorreta, acerte a data antes de usar o sistema...');
    Result := False;
    exit;
  end;}
  DM_Tabelas.ZQEmpresa.FieldByName('dia').AsDateTime := date;
  DM_Tabelas.ZQEmpresa.Post;
//  mensagem( datetostr(descripto_data(DM_Tabelas.ZQEmpresalimite.Value)));
  if descripto_data(DM_Tabelas.ZQEmpresa.FieldByName('limite').AsString)<=date+5 Then
  Begin
    FrmContraSenha.ShowModal;
    VarCSenha := FrmContraSenha.XEContraSenha.Text;
    VarSSenha := FrmContraSenha.LMasterSenha.Caption;
    VarFinal := 1;
    VarProxima := date;
    if pos(copy(Trim(VarSSenha),1,1),'123456789')>0 Then
      VarFinal := strtoint(copy(Trim(VarSSenha),1,1));
    if (copy(VarCSenha,2,5)= copy(inttoHex(strtoint(VarSSenha)*VarFinal*20050315,6),2,5)) and (pos(copy(VarCSenha,1,1),'DQVX0123456789')>0) Then Begin
      if copy(VarCSenha,1,1)='D' Then
        VarProxima := date+10
      else if copy(VarCSenha,1,1)='Q' Then
        VarProxima := date+15
      else if copy(VarCSenha,1,1)='V' Then
        VarProxima := date+20
      else
      Begin
        If copy(VarCSenha,1,1)='X' Then
          VarMes := 1000
        else
          Varmes := strtoint(copy(VarCSenha,1,1));
        VarProxima := IncMonth(date,Varmes);
      end;
      DM_Tabelas.ZQEmpresa.Edit;
      DM_Tabelas.ZQEmpresa.FieldByName('limite').AsString := cripto_data(VarProxima);
      DM_Tabelas.ZQEmpresa.Post;
      Result := True;
    End
    else
    Begin
      Showmessage('Contra-Senha incorreta...');
      if descripto_data(DM_Tabelas.ZQEmpresa.FieldByName('limite').AsString)<=date Then
        Result := False;
    End;
  end
  else
  begin
    Result := True;
  end;
End;

function diasemana(dia : Tdatetime) :string;
var
  diasemana:integer;
  strdiasemana:string;
begin
  diasemana:=dayofweek(dia);
  case diasemana of
  1:strdiasemana:='Domingo';
  2:strdiasemana:='Segunda';
  3:strdiasemana:='Terca';
  4:strdiasemana:='Quarta';
  5:strdiasemana:='Quinta';
  6:strdiasemana:='Sexta';
  7:strdiasemana:='Sabado';
  end;
  Result := strdiasemana;
end;

function GetFileDate(TheFileName: string): TdateTime;
var
  DataArq: TDate;
begin
  DataArq:= FileDateToDateTime(FileAge(TheFileName));
  result:=(dataarq);
end;

Function AchaCLiente(PL, PC: integer; Cli: String) : boolean;
Begin
  result := True;
  DM_Tabelas.ZQAchaParticip.close;
  DM_Tabelas.ZQAchaParticip.SQL.Clear;
  DM_Tabelas.ZQAchaParticip.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  DM_Tabelas.ZQAchaParticip.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  DM_Tabelas.ZQAchaParticip.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico');
  DM_Tabelas.ZQAchaParticip.SQL.Add(' from participante');
  DM_Tabelas.ZQAchaParticip.Open;
  if DM_Tabelas.ZQAchaParticip.Locate('idpaticipante',quotedstr(cli),[]) Then
  Begin
    exit;
  end
  else
  begin
    if Frm_AchaParticipante = nil then
      Frm_AchaParticipante := TFrm_AchaParticipante.Create(Application);
    Frm_AchaParticipante.Top := PL;
    Frm_AchaParticipante.Left := PC;
    DM_Tabelas.ZQAchaParticip.close;
    DM_Tabelas.ZQAchaParticip.SQL.Clear;
    DM_Tabelas.ZQAchaParticip.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
    DM_Tabelas.ZQAchaParticip.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
    DM_Tabelas.ZQAchaParticip.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico');
    DM_Tabelas.ZQAchaParticip.SQL.Add(' from participante where nome_parte like '+quotedstr('%'+cli+'%'));
    DM_Tabelas.ZQAchaParticip.SQL.Add(' or doc1 like '+quotedstr('%'+cli+'%')+' order by nome_parte');
    DM_Tabelas.ZQAchaParticip.Open;
    Frm_AchaParticipante.Eparticipante.Text := cli;
    IF DM_Tabelas.ZQAchaParticip.RecordCount>1 Then
    begin
      Frm_AchaParticipante.FormStyle := fsStayOnTop;
      Frm_AchaParticipante.BringToFront;
      Frm_AchaParticipante.ShowModal;
    end;
    IF DM_Tabelas.ZQAchaParticip.RecordCount=0 Then
    begin
      mensagem('Participante não encontrado.');
      exit;
    end;
    if (Frm_AchaParticipante.Label1.Caption = 'F') or (DM_tabelas.ZQAchaParticip.RecordCount=0)  Then
        Result := false
    else
        Result := True;
  end;
End;

Function AchaCorretor(PL, PC: integer; Corr :String) : boolean;
Begin
  result := True;
  if Frm_AchaCorretor = nil then
    Frm_AchaCorretor := TFrm_AchaCorretor.Create(Application);
  Frm_AchaCorretor.Top := PL;
  Frm_AchaCorretor.Left := PC;
  DM_Tabelas.ZQAchacorretor.SQL.Clear;
  DM_Tabelas.ZQAchacorretor.SQL.Add('select idcorretor,nomecorretor,doc1,creci,telefone1,telefone2,endereco_corr,bairro_corr,cep_corr,Obser_corr,codcidade,PComissao_corretor ');
  DM_Tabelas.ZQAchacorretor.SQL.Add('  from corretor');
  DM_Tabelas.ZQAchacorretor.SQL.Add(' where nomecorretor like '+quotedstr('%'+corr+'%')+' or idcorretor='+quotedstr(corr));
  DM_Tabelas.ZQAchacorretor.SQL.Add(' order by nomecorretor');
  DM_Tabelas.ZQAchacorretor.Open;
  Frm_AchaCorretor.Ecorretor.Text := corr;
  IF DM_Tabelas.ZQAchacorretor.RecordCount>1 Then
     Frm_AchaCorretor.ShowModal;
  if (Frm_AchaCorretor.Label1.Caption = 'F') or (DM_tabelas.ZQAchacorretor.RecordCount=0)  Then
      Result := false
  else
      Result := True;
End;

Function truncar(num : Double; dec : integer) : Double;
Var
  VArXX : integer;
  VarNum : string;
  VarRet : Double;
Begin
  VarNum := Floattostr(num);
  For VarXX := 1 to length(VarNum) do Begin
    If copy(VarNum,VarXX,1)=',' Then Begin
      VarRet := strtofloat(copy(VarNum,1,VarXX+dec));
      Result := Varret ;
      exit;
    End;
  end;
  Result := num;
End;

function dataporextenso(datadig : TdateTime) : string;
Var
  Vardata : string;
  varmes : integer;
Begin
  Vardata := copy(datetostr(datadig),1,2)+' de ';
  Varmes := strtoint(copy(datetostr(datadig),4,2));
  case VarMes of
    1: Vardata := Vardata+'Janeiro';
    2: Vardata := Vardata+'Fevereiro';
    3: Vardata := Vardata+'Marco';
    4: Vardata := Vardata+'Abril';
    5: Vardata := Vardata+'Maio';
    6: Vardata := Vardata+'Junho';
    7: Vardata := Vardata+'Julho';
    8: Vardata := Vardata+'Agosto';
    9: Vardata := Vardata+'Setembro';
    10: Vardata := Vardata+'Outubro';
    11: Vardata := Vardata+'Novembro';
    12: Vardata := Vardata+'Dezembro';
  end;
  Vardata:=Vardata+' de '+copy(datetostr(datadig),7,4);
  result := Vardata;
End;

function BuscaTroca(Text,Busca,Troca : string) : string;
{ Substitui um caractere dentro da string}
var n : integer;
begin
for n := 1 to length(Text) do
  begin
  if Copy(Text,n,1) = Busca then
  begin
  Delete(Text,n,1);
  Insert(Troca,Text,n);
  end;
  end;
Result := Text;
end;

Function AchaPlanoDeConta(PL, PC: integer; varclas, pln, varCD :String) : boolean;
Var
  Varxx : integer;
Begin
  for varxx:=1 to length(pln) do
    if copy(pln,varxx,1) = ' ' Then
      Break;
  if (not empty(pln)) and (StrIsInteger(copy(pln,1,varxx-1))) Then Begin
    DM_Tabelas.ZQAchaPlanoDeContas.SQL.Clear;
    DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add('select * from plano_contas');
    DM_Tabelas.ZQAchaPlanoDeContas.Open;
    Varxx := StrtoInt(copy(pln,1,varxx-1));
    if DM_Tabelas.ZQAchaPlanoDeContas.Locate('codigo',Varxx,[]) Then Begin
      Result := True;
      exit;
    end;
  end;
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Clear;
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add('select codigo,doccomum,mascara,descricao,classificacao,valor,cod_reduzido,cred_debi,usoativo,CDPlano,DespFixVar ');
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' from plano_contas');
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' where (codigo like '+quotedstr(pln+'%'));
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' or descricao like '+quotedstr('%'+pln+'%'));
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' or mascara like '+quotedstr(pln+'%')+')');
  if not empty(varCD) then
    DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' and cred_debi ='+quotedstr(varcd));
  if not empty(varclas) Then
    DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' and classificacao = '+quotedstr(varclas));
  DM_Tabelas.ZQAchaPlanoDeContas.Open;
  if DM_Tabelas.ZQAchaPlanoDeContas.RecordCount=1 then
    Result := True
  else if DM_Tabelas.ZQAchaPlanoDeContas.RecordCount=0 then
    Result := False
  else Begin
    if Frm_AchaPlanoDeContas = nil then
      Frm_AchaPlanoDeContas := TFrm_AchaPlanoDeContas.Create(Application);
    Frm_AchaPlanoDeContas.LCD.Caption := varcd;
    Frm_AchaPlanoDeContas.Label3.Caption := varclas;
    Frm_AchaPlanoDeContas.Eplano.Text:=pln;
    Frm_AchaPlanoDeContas.Top := pl;
    Frm_AchaPlanoDeContas.Left := pc;
    Frm_AchaPlanoDeContas.ShowModal;
    if Frm_AchaPlanoDeContas.Label1.Caption = 'F' Then
      Result := False
    else
      Result := True;
  end;
end;

Function tiraacento(texto: string) : string;
Var
  Varx,varachado : integer;
const
  comacento = 'áàâãäéèêëíìîïóòôõöúùûüýÿçÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÝÇ&';
  semacento = 'aaaaaeeeeiiiiooooouuuuyycAAAAAEEEEIIIIOOOOOUUUUYCE';
Begin
  For varx := 1 to length(texto) do Begin
    varachado:=pos(copy(texto,varx,1),comacento);
    if varachado>0 Then
      texto := BuscaTroca(Texto,copy(comacento,varachado,1),copy(semacento,varachado,1));
  end;
  result := texto;
End;

function StrIsNumero(const S: string): boolean;
 var
 tamanho, x : integer;
 l, p : string;
begin
  p := s;                            // Esta função verifica de uma
  tamanho := length(p);
  Result := True;                    // string tem apenas caracteres
  for x := 1 to tamanho do           // numericos.
  begin
    l := copy(p,x,1);
    if  pos(l,'0123456789.')= 0 then begin
      Result := FALSE;
      exit;
    end;
  end;
end;

function Transform( Conteudo: Extended; const Mascara: string ): string;
var
  TamMascara: integer;
  Brancos: string;
  Dados: string;
begin
  TamMascara := Length( Mascara );
  Dados := FormatFloat( Mascara, Conteudo );
  if TamMascara > Length( Dados ) then
  begin
    Brancos := Space( TamMascara - Length( Dados ) );
    Dados := Brancos + Dados;
  end;
  Transform := Dados;
end;

function Space(N: integer): string;
var
  I: integer;
  Dados: string;
begin
  Dados := '';
  for I := 1 to N do
  begin
    Dados := Dados + ' ';
    Application.ProcessMessages;
  end;
  Space := Dados;
end;

Function Calcjuros(dtI,dtF: Tdate; varvalor, varjur: double) : Double;
var
  varva : Double;
Begin
  varva := varvalor;
  if (dti+Dm_tabelas.ZQConfiguracoes.FieldByName('diasdeatrazo').AsInteger<dtf) and (varva>0) and (dtI>0)  Then Begin
    while dtI<dtF do Begin
      if IncMonth(dtI,1)<= dtF Then
        varva := varva + (varva*varjur/100)
      else
        varva := varva + (varva*(varjur/30*(dtF-dtI))/100);
      dtI := IncMonth(dtI,1);
    end;
  end;
  result := varva;
End;

function tiramascara(varx : string) : string;
Var
  varpedaco : string;
  vary : integer;
Begin
  varpedaco := '';
  for vary := 1 to length(varx) do
    if pos(copy(varx,vary,1),'1234567890')>0 Then
      varpedaco := varpedaco +copy(varx,vary,1);
  result := varpedaco;
end;

{Function extenso(valor : double) : string;
var
  varx : integer;
  varstr : string;
  valpartes : array[1..5] of string;
  valpartesext : array[1..5] of string;
const
  Unidades: array[1..19] of string = ('Um', 'Dois', 'Tres', 'Quatro', 'Cinco', 'Seis', 'Sete', 'Oito', 'Nove', 'Dez', 'Onze', 'Doze', 'Treze', 'Quatorze', 'Quinze', 'Dezesseis', 'Dezessete', 'Dezoito', 'Dezenove');
  Dezenas: array[1..9] of string = ('Dez', 'Vinte', 'Trinta', 'Quarenta', 'Cinquenta', 'Sessenta', 'Setenta', 'Oitenta', 'Noventa');
  Centenas: array[1..9] of string = ('Cento', 'Duzentos', 'Trezentos', 'Quatrocentos', 'Quinhentos', 'Seiscentos', 'Setecentos', 'Oitocentos', 'Novecentos');
Begin
  varstr := formatfloat('000000000000.00',valor);
  valpartesext[1]:='';
  valpartesext[2]:='';
  valpartesext[3]:='';
  valpartesext[4]:='';
  valpartesext[5]:='';
  valpartes[1] := copy(varstr,1,3);
  valpartes[2] := copy(varstr,4,3);
  valpartes[3] := copy(varstr,7,3);
  valpartes[4] := copy(varstr,10,3);
  valpartes[5] := '0'+copy(varstr,14,2);
  for varx := 1 to 5 do begin
    if valpartes[varx]='000' then valpartesext[varx] := ''
    else if valpartes[varx]='100' then valpartesext[varx] := 'Cem'
    else Begin
      if (strtoint(copy(valpartes[varx],2,2))<20) and (strtoint(copy(valpartes[varx],2,2))>0) Then valpartesext[varx] := valpartesext[varx] + Unidades[strtoint(copy(valpartes[varx],2,2))]
      else Begin
        if strtoint(copy(valpartes[varx],3,1))>0 Then valpartesext[varx] := valpartesext[varx] + Unidades[strtoint(copy(valpartes[varx],3,1))];
        if strtoint(copy(valpartes[varx],2,1))>0 Then Begin
          if empty(valpartesext[varx]) then
            valpartesext[varx] := valpartesext[varx] + Dezenas[strtoint(copy(valpartes[varx],2,1))]
          else
            valpartesext[varx] := Dezenas[strtoint(copy(valpartes[varx],2,1))] + ' e ' + valpartesext[varx];
        end;
      end;
      if strtoint(copy(valpartes[varx],1,1))>0 Then Begin
        if strtoint(copy(valpartes[varx],1,1))>0 Then Begin
          if empty(valpartesext[varx]) then
            valpartesext[varx] := valpartesext[varx] + Centenas[strtoint(copy(valpartes[varx],1,1))]
          else
            valpartesext[varx] := Centenas[strtoint(copy(valpartes[varx],1,1))] + ' e ' + valpartesext[varx];
        end;
      end;
    end;
    if strtoint(valpartes[varx])=1 Then
    case varx of
      1 : valpartesext[varx] :=  valpartesext[varx] +' Bilhão';
      2 : valpartesext[varx] :=  valpartesext[varx] +' Milhão';
      3 : valpartesext[varx] :=  valpartesext[varx] +' Mil';
      5 : valpartesext[varx] :=  valpartesext[varx] +' Centavo';
    end;
    if strtoint(valpartes[varx])>1 Then
    case varx of
      1 : valpartesext[varx] :=  valpartesext[varx] +' Bilhoes';
      2 : valpartesext[varx] :=  valpartesext[varx] +' Milhoes';
      3 : valpartesext[varx] :=  valpartesext[varx] +' Mil';
      5 : valpartesext[varx] :=  valpartesext[varx] +' Centavos';
    end;
  end;
  if strtoint(copy(varstr,1,12))>=2 Then
    valpartesext[4] := valpartesext[4] + ' Reais'
  else  if strtoint(copy(varstr,1,12))=1 Then
    valpartesext[4] := valpartesext[4] + ' Real';
  if (not empty(valpartesext[4])) and (not empty(valpartesext[5])) Then valpartesext[4]:=valpartesext[4]+' e ';
  if (not empty(valpartesext[3])) and (not empty(valpartesext[4])) Then valpartesext[3]:=valpartesext[3]+' e ';
  if (not empty(valpartesext[2])) and (not empty(valpartesext[3])) Then valpartesext[2]:=valpartesext[2]+' e ';
  if (not empty(valpartesext[1])) and (not empty(valpartesext[2])) Then valpartesext[1]:=valpartesext[1]+' e ';
  Result := valpartesext[1]+valpartesext[2]+valpartesext[3]+valpartesext[4]+valpartesext[5];
end;}



Function Extenso(Valor : Extended): String;
var
Centavos, Centena, Milhar, Milhao, Bilhao, Texto : string;
const
Unidades: array [1..9] of string = ('um', 'dois', 'três', 'quatro',
'cinco',
'seis', 'sete', 'oito',
'nove');
Dez : array [1..9] of string = ('onze', 'doze', 'treze',
'quatorze', 'quinze',
'dezesseis', 'dezessete',
'dezoito', 'dezenove');
Dezenas: array [1..9] of string = ('dez', 'vinte', 'trinta',
'quarenta', 'cinquenta',
'sessenta', 'setenta',
'oitenta', 'noventa');
Centenas: array [1..9] of string = ('cento', 'duzentos',
'trezentos', 'quatrocentos',
'quinhentos', 'seiscentos',
'setecentos',
'oitocentos', 'novecentos');

Function ifs( Expressao: Boolean; CasoVerdadeiro, CasoFalso: String): String;
begin
if Expressao then Result := CasoVerdadeiro else Result :=
CasoFalso;
end;

Function MiniExtenso( Valor: ShortString ): string;
var
Unidade,
Dezena,
Centena: String;
begin
if (Valor[2] = '1') and (Valor[3] <> '0') then
begin
Unidade := Dez[StrToInt(Valor[3])];
Dezena := '';
end
else
begin
if Valor[2] <> '0' then
Dezena := Dezenas[StrToInt(Valor[2])];
if Valor[3] <> '0' then
unidade := Unidades[StrToInt(Valor[3])];
end;
if (Valor[1] = '1') and (Unidade = '') and (Dezena = '') then
Centena := 'cem'
else
if Valor[1] <> '0' then
Centena := Centenas[StrToInt(Valor[1])]
else
Centena := '';
Result := Centena +
ifs( (Centena <> '') and ((Dezena <> '') or (Unidade <> '')),
' e ', '') + Dezena + ifs( (Dezena <> '') and (Unidade <> ''), ' e ', '')+
Unidade;
end;
begin
if Valor = 0 then
begin
Result := '';
Exit;
end;
Texto := FormatFloat( '000000000000.00', Valor );
Centavos := MiniExtenso( '0' + Copy( Texto, 14, 2 ) );
Centena := MiniExtenso( Copy( Texto, 10, 3 ) );
Milhar := MiniExtenso( Copy( Texto, 7, 3 ) );
if Milhar <> '' then
Milhar := Milhar + ' mil';
Milhao := MiniExtenso( Copy( Texto, 4, 3 ) );
if Milhao <> '' then
Milhao := Milhao + ifs( Copy( Texto, 4, 3 ) = '001', ' milhao', ' milhoes');
Bilhao := MiniExtenso( Copy( Texto, 1, 3 ) );
if Bilhao <> '' then
Bilhao := Bilhao + ifs( Copy( Texto, 1, 3 ) = '001', ' bilhao', ' bilhoes');
if (Bilhao <> '') and (Milhao + Milhar + Centena = '') then
Result := Bilhao + ' de Reais'
else
if (Milhao <> '') and (Milhar + Centena = '') then
Result := Milhao + ' de Reais'
else
Result := Bilhao +
ifs( (Bilhao <> '') and (Milhao + Milhar + Centena <> ''),
ifs((Pos(' e ', Bilhao) > 0) or (Pos( ' e ', Milhao + Milhar +
Centena ) > 0 ), ', ', ' e '), '')+Milhao +
ifs( (Milhao <> '') and (Milhar + Centena <> ''),
ifs((Pos(' e ', Milhao) > 0) or
(Pos( ' e ', Milhar + Centena ) > 0 ), ', ', ' e '), '') + Milhar +
ifs( (Milhar <> '') and (Centena <> ''),
ifs(Pos( ' e ', Centena ) > 0, ', ', ' e '), '')+
Centena + ifs( Int(Valor) = 1, ' Real', ' Reais' );
if Centavos <> '' then
Result := Result + ' e ' + Centavos +
ifs( Copy( Texto, 14, 2 )= '01', ' Centavo', ' Centavos' );
end;



function numeroporextenso(numero : Integer) : string;
Var
  varnum : integer;
  varext : String;
Begin
  Varnum := numero;
  case Varnum of
    1: varext := '01 (Uma)';
    2: varext := '02 (Duas)';
    3: varext := '03 (Três)';
    4: varext := '04 (Quatro)';
    5: varext := '05 (Cinco)';
    6: varext := '06 (Seis)';
    7: varext := '07 (Sete)';
    8: varext := '08 (Oito)';
    9: varext := '09 (Nove)';
    10: varext := '10 (Dez)';
    11: varext := '11 (Onze)';
    12: varext := '12 (Doze)';
  end;
  result := varext;
End;


function ExRound(Value: Extended; Decimals: Integer): Extended;
var
  Factor, Fraction: Extended;
begin
  Factor := IntPower(10, Decimals);
  Value := StrToFloat(FloatToStr(Value * Factor));
  Result := Int(Value);
  Fraction := Frac(Value);
  if Fraction >= 0.5 then
    Result := Result + 1
  else if Fraction <= -0.5 then
    Result := Result - 1;
  Result := Result / Factor;
end;

function Repl(C: string; Tamanho: integer): string;
var
  Conteudo: string;
  Contar: integer;
begin
  Conteudo := '';
  for Contar := 1 to Tamanho do
  begin
    Conteudo := Conteudo + C;
    Application.ProcessMessages;
  end;
  Repl := Conteudo;
end;

function StrZero(N: integer; Tamanho: integer): string;
var
  Conteudo: string;
  Diferenca: Integer;
begin
  Conteudo := IntToStr( N );
  Diferenca := Tamanho - Length( Conteudo );
  if Diferenca > 0 then
    Conteudo := Repl( '0', Diferenca ) + Conteudo;
  StrZero := Conteudo;
end;



function Dia( Data: TDateTime ): string;
var
  sAno, sMes, sDia: Word;
begin
  DecodeDate( Data, sAno, sMes, sDia );
  Dia := StrZero( sDia, 2 );
end;

function Mes( Data: TDateTime ): string;
var
  sAno, sMes, sDia: Word;
begin
  DecodeDate( Data, sAno, sMes, sDia );
  Mes := StrZero( sMes, 2 );
end;

function MesAno( Data: TDateTime ): string;
var
  sAno, sMes, sDia: Word;
begin
  DecodeDate( Data, sAno, sMes, sDia );
  MesAno := StrZero( sMes, 2 ) + '/' +
    Copy( StrZero( sAno, 4 ), 1, 4 );
end;

function Ano( Data: TDateTime ): string;
var
  sAno, sMes, sDia: Word;
begin
  DecodeDate( Data, sAno, sMes, sDia );
 // Ano := Copy( StrZero( sAno, 4 ), 3, 4 );
  Ano := IntToStr(SAno);
end;

function UltimoDiaDoMes( MesAno: string ): string;
var
  sMes: string;
  sAno: string;
begin
  sMes := Copy(mesano, 1, 2 );
  if pos('/',mesano)>0 then
     sAno:=ano(strtodate('01/'+sMes+'/'+Copy(mesano, 4, 4 )))
  else
     sAno:=ano(strtodate('01/'+sMes+'/'+Copy(mesano, 3, 4 )));
  if Pos( sMes, '01 03 05 07 08 10 12' ) > 0 then
    UltimoDiaDoMes := '31'
  else
    if sMes <> '02' then
      UltimoDiaDoMes := '30'
    else
      if  (StrToInt( sAno ) mod 4)  = 0 then
        UltimoDiaDoMes := '29'
      else
        UltimoDiaDoMes := '28';
end;

procedure Mensagem( Texto: string );
var
  Formulario: TFormMensagem;
  procedure AjustarTamanho;
  const
    LarguraTextoPadrao = 446;
    AlturaTextoPadrao = 132;
    LarguraFormularioPadrao = 550;
    AlturaFormularioPadrao = 211;
    MargemHorizontal = 104;
    MargemVertical = 79;
    MargemTela = 40;
  var
    Linhas: TStringList;
    Retangulo: TRect;
    LarguraMaxima, AlturaMaxima, LarguraTexto, AlturaTexto,
    LarguraFormulario, AlturaFormulario, MaiorLinha, LarguraLinha,
    I: Integer;
  begin
    Formulario.Frase.Caption := Texto;
    Formulario.Frase.AutoSize := False;
    Formulario.Frase.WordWrap := True;
    Formulario.Frase.Canvas.Font.Assign(Formulario.Frase.Font);

    LarguraMaxima := Screen.WorkAreaWidth - MargemTela;
    if LarguraMaxima < LarguraFormularioPadrao then
      LarguraMaxima := LarguraFormularioPadrao;
    AlturaMaxima := Screen.WorkAreaHeight - MargemTela;
    if AlturaMaxima < AlturaFormularioPadrao then
      AlturaMaxima := AlturaFormularioPadrao;

    MaiorLinha := 0;
    Linhas := TStringList.Create;
    try
      Linhas.Text := Texto;
      for I := 0 to Linhas.Count - 1 do
      begin
        LarguraLinha := Formulario.Frase.Canvas.TextWidth(Linhas[I]);
        if LarguraLinha > MaiorLinha then
          MaiorLinha := LarguraLinha;
      end;
    finally
      Linhas.Free;
    end;

    LarguraTexto := LarguraTextoPadrao;
    if MaiorLinha + 24 > LarguraTexto then
      LarguraTexto := MaiorLinha + 24;
    if LarguraTexto > LarguraMaxima - MargemHorizontal then
      LarguraTexto := LarguraMaxima - MargemHorizontal;
    if LarguraTexto < LarguraTextoPadrao then
      LarguraTexto := LarguraTextoPadrao;

    repeat
      Retangulo := Rect(0, 0, LarguraTexto, 0);
      DrawText(Formulario.Frase.Canvas.Handle, PChar(Texto), Length(Texto),
        Retangulo, DT_CALCRECT or DT_WORDBREAK or DT_NOPREFIX);
      AlturaTexto := Retangulo.Bottom - Retangulo.Top;
      if AlturaTexto < AlturaTextoPadrao then
        AlturaTexto := AlturaTextoPadrao;
      if (AlturaTexto + MargemVertical <= AlturaMaxima) or
         (Formulario.Frase.Font.Size <= 10) then
        Break;
      Formulario.Frase.Font.Size := Formulario.Frase.Font.Size - 1;
    until False;

    LarguraFormulario := LarguraTexto + MargemHorizontal;
    if LarguraFormulario < LarguraFormularioPadrao then
      LarguraFormulario := LarguraFormularioPadrao;
    if LarguraFormulario > LarguraMaxima then
      LarguraFormulario := LarguraMaxima;

    AlturaFormulario := AlturaTexto + MargemVertical;
    if AlturaFormulario < AlturaFormularioPadrao then
      AlturaFormulario := AlturaFormularioPadrao;
    if AlturaFormulario > AlturaMaxima then
      AlturaFormulario := AlturaMaxima;

    Formulario.Width := LarguraFormulario;
    Formulario.Height := AlturaFormulario;
    Formulario.Frase.Width := Formulario.ClientWidth - Formulario.Frase.Left - 31;
    Formulario.Frase.Height := Formulario.ClientHeight -
      Formulario.Frase.Top - Formulario.BtnOk.Height - 21;
    if Formulario.Frase.Height < AlturaTextoPadrao then
      Formulario.Frase.Height := AlturaTextoPadrao;
    Formulario.BtnOk.Left := (Formulario.ClientWidth - Formulario.BtnOk.Width) div 2;
    Formulario.BtnOk.Top := Formulario.ClientHeight -
      Formulario.BtnOk.Height - 10;
  end;
begin
  Formulario := TFormMensagem.Create(Application);
  FormMensagem := Formulario;
  try
    AjustarTamanho;
    MessageBeep(0);
    Formulario.ShowModal;
  finally
    if FormMensagem = Formulario then
      FormMensagem := nil;
    Formulario.Free;
  end;
end;

function MesExtenso( Mes:Word ) : string; const meses : array[0..11] of PChar = ('Janeiro', 'Fevereiro', 'Marco', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro','Outubro', 'Novembro', 'Dezembro');
begin
  result := meses[mes-1];
End;

Function SIMNAO(texto, opcao:string) : Boolean; //
Begin
  if FrmPerguntaSIMNAO = nil then
    FrmPerguntaSIMNAO := TFrmPerguntaSIMNAO.Create(Application);
  FrmperguntaSIMNAO.Label1.Caption := Texto;
  frmperguntaSIMNAO.Label2.Caption := opcao;
  FrmperguntaSIMNAO.ShowModal;
  result := False;
  if  FrmperguntaSIMNAO.Label2.Caption = 'SIM' Then
    result := True;
End;

function pergunta(texto,resp : string) :string;
Begin
  if FrmPergunta = nil then
    FrmPergunta := TFrmPergunta.Create(Application);
  FrmPergunta.label5.caption := texto;
  Frmpergunta.XEdit1.Text := resp;
  Frmpergunta.ShowModal;
  result := Frmpergunta.XEdit1.Text ;
End;


function CHARREM(Texto : string) : string;
var
    Pos_Ponto1, Pos_Ponto2, Pos_Ponto3, Pos_Ponto4: Byte;
    Pos_Ponto5:Byte;
    String1, String2, String3, String4, String5, String6 : String;
    val1,val2,val3,val4,val5: String;
    resultado:string;
begin
 Pos_Ponto1 := Pos('.',Alltrim(Texto));
 String1    := Copy(Alltrim(Texto),Pos_Ponto1 +1 ,(length(Alltrim(Texto))-(Pos_Ponto1-1) ));
 Pos_Ponto2 := Pos('.',String1);

 {fase1:}
 {Do Texto: xxx.yyy.zzz captura -> xxx }
 val1 := Copy(alltrim(Texto), 0, (Pos_Ponto1 - 1));

 {fase2:}
 {Do Texto: xxx.yyy.zzz captura -> yyy }
 val2 := Copy(String1,0,(Pos_Ponto2 - 1 ));

 {fase3:}
 {Do Texto: xxx.yyy.zzz captura -> zzz }
 String4    := Copy(String1,Pos_Ponto2 +1 ,(length(String1)-(Pos_Ponto2-1) ));
 if not Empty(String4) then
 begin
    Pos_Ponto3 := Pos('.',String4);
    val3 := Copy(String4,0,(Pos_Ponto3 - 1 ));
    if Pos_Ponto3<>0 then
    begin
      Resultado:=Val1+val2+val3;
    end
    else
    begin
     Resultado:=Val1+val2+String4;
    end;
 end
 else
 begin
   Resultado:=Val1+val2;
 end;

{fase4:}
 String5    := Copy(String4,Pos_Ponto3 +1 ,(length(String4)-(Pos_Ponto3-1) ));
 if not empty(String5) then
 begin
   Pos_Ponto4 := Pos('.',String5);
   val4 := Copy(String5,0,(Pos_Ponto4 - 1 ));
   if Pos_Ponto4<>0 then
   begin
     Resultado:=Val1+val2+val3+val4;
   end
   else
   begin
     Resultado:=Val1+val2+val3+String5;
   end;
 end
 else
 begin
   Resultado:=Val1+val2+val3;
 end;

{fase5:}
 String6    := Copy(String5,Pos_Ponto4 +1 ,(length(String5)-(Pos_Ponto4-1) ));
 if not empty(String6) then
 begin
   Pos_Ponto5 := Pos('.',String6);
   val5 := Copy(String6,0,(Pos_Ponto5 - 1 ));
   if Pos_Ponto5<>0 then
   begin
     Resultado:=Val1+val2+val3+val4+val5;
   end
   else
   begin
     Resultado:=Val1+val2+val3+val4+String6;
   end;
 end
 else
 begin
   Resultado:=Val1+val2+val3+val4;
 end;

 Pos_Ponto1 := Pos('-',Alltrim(Resultado));
 String1    := Copy(Alltrim(Resultado),Pos_Ponto1 +1 ,(length(Alltrim(Resultado))-(Pos_Ponto1-1) ));
 Pos_Ponto2 := Pos('-',String1);

 {fase6:}
 {Do Texto: xxx.yyy.zzz captura -> xxx }
 val1 := Copy(alltrim(Resultado), 0, (Pos_Ponto1 - 1));

 {fase7:}
 {Do Texto: xxx.yyy.zzz captura -> yyy }
 val2 := Copy(String1,0,(Pos_Ponto2 - 1 ));

 {fase8:}
 {Do Texto: xxx.yyy.zzz captura -> zzz }
 String4    := Copy(String1,Pos_Ponto2 +1 ,(length(String1)-(Pos_Ponto2-1) ));
 if not Empty(String4) then
 begin
    Pos_Ponto3 := Pos('-',String4);
    val3 := Copy(String4,0,(Pos_Ponto3 - 1 ));
    if Pos_Ponto3<>0 then
    begin
      Resultado:=Val1+val2+val3;
    end
    else
    begin
     Resultado:=Val1+val2+String4;
    end;
 end
 else
 begin
   Resultado:=Val1+val2;
 end;

{fase9:}
 String5    := Copy(String4,Pos_Ponto3 +1 ,(length(String4)-(Pos_Ponto3-1) ));
 if not empty(String5) then
 begin
   Pos_Ponto4 := Pos('-',String5);
   val4 := Copy(String5,0,(Pos_Ponto4 - 1 ));
   if Pos_Ponto4<>0 then
   begin
     Resultado:=Val1+val2+val3+val4;
   end
   else
   begin
     Resultado:=Val1+val2+val3+String5;
   end;
 end
 else
 begin
   Resultado:=Val1+val2+val3;
 end;

{fase10:}
 String6    := Copy(String5,Pos_Ponto4 +1 ,(length(String5)-(Pos_Ponto4-1) ));
 if not empty(String6) then
 begin
   Pos_Ponto5 := Pos('-',String6);
   val5 := Copy(String6,0,(Pos_Ponto5 - 1 ));
   if Pos_Ponto5<>0 then
   begin
     Resultado:=Val1+val2+val3+val4+val5;
   end
   else
   begin
     Resultado:=Val1+val2+val3+val4+String6;
   end;
 end
 else
 begin
   Resultado:=Val1+val2+val3+val4;
 end;

 Pos_Ponto1 := Pos(',',Alltrim(Resultado));
 String1    := Copy(Alltrim(Resultado),Pos_Ponto1 +1 ,(length(Alltrim(Resultado))-(Pos_Ponto1-1) ));
 Pos_Ponto2 := Pos(',',String1);

 {fase11:}
 {Do Texto: xxx.yyy.zzz captura -> xxx }
 val1 := Copy(alltrim(Resultado), 0, (Pos_Ponto1 - 1));

 {fase12:}
 {Do Texto: xxx.yyy.zzz captura -> yyy }
 val2 := Copy(String1,0,(Pos_Ponto2 - 1 ));

 {fase13:}
 {Do Texto: xxx.yyy.zzz captura -> zzz }
 String4    := Copy(String1,Pos_Ponto2 +1 ,(length(String1)-(Pos_Ponto2-1) ));
 if not Empty(String4) then
 begin
    Pos_Ponto3 := Pos(',',String4);
    val3 := Copy(String4,0,(Pos_Ponto3 - 1 ));
    if Pos_Ponto3<>0 then
    begin
      Resultado:=Val1+val2+val3;
    end
    else
    begin
     Resultado:=Val1+val2+String4;
    end;
 end
 else
 begin
   Resultado:=Val1+val2;
 end;

{fase14:}
 String5    := Copy(String4,Pos_Ponto3 +1 ,(length(String4)-(Pos_Ponto3-1) ));
 if not empty(String5) then
 begin
   Pos_Ponto4 := Pos(',',String5);
   val4 := Copy(String5,0,(Pos_Ponto4 - 1 ));
   if Pos_Ponto4<>0 then
   begin
     Resultado:=Val1+val2+val3+val4;
   end
   else
   begin
     Resultado:=Val1+val2+val3+String5;
   end;
 end
 else
 begin
   Resultado:=Val1+val2+val3;
 end;

{fase15:}
 String6    := Copy(String5,Pos_Ponto4 +1 ,(length(String5)-(Pos_Ponto4-1) ));
 if not empty(String6) then
 begin
   Pos_Ponto5 := Pos(',',String6);
   val5 := Copy(String6,0,(Pos_Ponto5 - 1 ));
   if Pos_Ponto5<>0 then
   begin
     Resultado:=Val1+val2+val3+val4+val5;
   end
   else
   begin
     Resultado:=Val1+val2+val3+val4+String6;
   end;
 end
 else
 begin
   Resultado:=Val1+val2+val3+val4;
 end;

 Pos_Ponto1 := Pos('/',Alltrim(Resultado));
 String1    := Copy(Alltrim(Resultado),Pos_Ponto1 +1 ,(length(Alltrim(Resultado))-(Pos_Ponto1-1) ));
 Pos_Ponto2 := Pos('/',String1);

 {fase16:}
 {Do Texto: xxx.yyy.zzz captura -> xxx }
 val1 := Copy(alltrim(Resultado), 0, (Pos_Ponto1 - 1));

 {fase17:}
 {Do Texto: xxx.yyy.zzz captura -> yyy }
 val2 := Copy(String1,0,(Pos_Ponto2 - 1 ));

 {fase18:}
 {Do Texto: xxx.yyy.zzz captura -> zzz }
 String4    := Copy(String1,Pos_Ponto2 +1 ,(length(String1)-(Pos_Ponto2-1) ));
 if not Empty(String4) then
 begin
    Pos_Ponto3 := Pos('/',String4);
    val3 := Copy(String4,0,(Pos_Ponto3 - 1 ));
    if Pos_Ponto3<>0 then
    begin
      Resultado:=Val1+val2+val3;
    end
    else
    begin
     Resultado:=Val1+val2+String4;
    end;
 end
 else
 begin
   Resultado:=Val1+val2;
 end;

{fase19:}
 String5    := Copy(String4,Pos_Ponto3 +1 ,(length(String4)-(Pos_Ponto3-1) ));
 if not empty(String5) then
 begin
   Pos_Ponto4 := Pos('/',String5);
   val4 := Copy(String5,0,(Pos_Ponto4 - 1 ));
   if Pos_Ponto4<>0 then
   begin
     Resultado:=Val1+val2+val3+val4;
   end
   else
   begin
     Resultado:=Val1+val2+val3+String5;
   end;
 end
 else
 begin
   Resultado:=Val1+val2+val3;
 end;

{fase20:}
 String6    := Copy(String5,Pos_Ponto4 +1 ,(length(String5)-(Pos_Ponto4-1) ));
 if not empty(String6) then
 begin
   Pos_Ponto5 := Pos('/',String6);
   val5 := Copy(String6,0,(Pos_Ponto5 - 1 ));
   if Pos_Ponto5<>0 then
   begin
     Resultado:=Val1+val2+val3+val4+val5;
   end
   else
   begin
     Resultado:=Val1+val2+val3+val4+String6;
   end;
 end
 else
 begin
   Resultado:=Val1+val2+val3+val4;
 end;
 Result := Resultado;
end;


Function RemoveChar(Const Texto:String):String;
//
// Remove caracteres de uma string deixando apenas numeros
//
var
I: integer;
S: string;
begin
  S := '';
  for I := 1 To Length(Texto) Do
  begin
    if (Texto[I] in ['0'..'9']) then
    begin
      S := S + Copy(Texto, I, 1);
    end;
  end;
  result := S;
end;

//Removendo espaços/caracteres de uma string
function TrimChar(texto: string; delchar: char): string;
var
S: string;
begin
S := texto;
while Pos(delchar,S) > 0 do
Delete(S,Pos(delchar,S),1);
Result := S;
end;
{Remover todas as letras 's':
ShowMessage(TrimChar(Edit1.Text,'s'));}


function Formatar(Texto : string; TamanhoDesejado : integer; AcrescentarADireita : boolean = true; CaracterAcrescentar : char = ' ') : string;
{
   OBJETIVO: Eliminar caracteres inválidos e acrescentar caracteres à esquerda ou à direita do texto original para que a string resultante fique com o tamanho desejado

   Texto : Texto original
   TamanhoDesejado: Tamanho que a string resultante deverá ter
   AcrescentarADireita: Indica se o carácter será acrescentado à direita ou à esquerda
      TRUE - Se o tamanho do texto for MENOR que o desejado, acrescentar carácter à direita
             Se o tamanho do texto for MAIOR que o desejado, eliminar últimos caracteres do texto
      FALSE - Se o tamanho do texto for MENOR que o desejado, acrescentar carácter à esquerda
             Se o tamanho do texto for MAIOR que o desejado, eliminar primeiros caracteres do texto
   CaracterAcrescentar: Carácter que deverá ser acrescentado
}
var
   QuantidadeAcrescentar,
   TamanhoTexto,
   PosicaoInicial,
   i : integer;

begin
   case CaracterAcrescentar of
      '0'..'9','a'..'z','A'..'Z' : ;{Não faz nada}
      else
         CaracterAcrescentar := ' ';
   end;

   Texto := Trim(AnsiUpperCase(Texto));
   TamanhoTexto := Length(Texto);
{$WARNINGS OFF}
   for i := 1 to (TamanhoTexto) do
   begin
      if Pos(Texto[i],' 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ`~''"!@#$%^&*()_-+=|/\{}[]:;,.<>') = 0 then
      begin
         case Texto[i] of
            'Á','À','Â','Ä','Ã' : Texto[i] := 'A';
            'É','È','Ê','Ë' : Texto[i] := 'E';
            'Í','Ì','Î','Ï' : Texto[i] := 'I';
            'Ó','Ò','Ô','Ö','Õ' : Texto[i] := 'O';
            'Ú','Ù','Û','Ü' : Texto[i] := 'U';
            'Ç' : Texto[i] := 'C';
            'Ñ' : Texto[i] := 'N';
            else Texto[i] := ' ';
         end;
      end;
   end;
   QuantidadeAcrescentar := TamanhoDesejado - TamanhoTexto;
   if QuantidadeAcrescentar < 0 then
      QuantidadeAcrescentar := 0;
   if CaracterAcrescentar = '' then
      CaracterAcrescentar := ' ';
   if TamanhoTexto >= TamanhoDesejado then
      PosicaoInicial := TamanhoTexto - TamanhoDesejado + 1
   else
      PosicaoInicial := 1;

   if AcrescentarADireita then
      Texto := Copy(Texto,1,TamanhoDesejado) + StringOfChar(CaracterAcrescentar,QuantidadeAcrescentar)
   else
      Texto := StringOfChar(CaracterAcrescentar,QuantidadeAcrescentar) + Copy(Texto,PosicaoInicial,TamanhoDesejado);

   Result := AnsiUpperCase(Texto);
end;


function DifDias(DataVenc:TDateTime; DataAtual:TDateTime): String;
{Retorna a diferenca de dias entre duas datas}
Var Data: TDateTime;
    dia, mes, ano: Word;
begin
  if DataAtual < DataVenc then
  begin
    Result := 'A data atual não pode ser menor que a data inicial';
  end
  else
  begin
   Data := DataAtual - DataVenc;
   DecodeDate( Data, ano, mes, dia);
//   Result := FloatToStr(Data)+' Dias';
   Result := FloatToStr(Data);
  end;
end;

procedure ExecutePrograma(Nome, Parametros: String);
Var
Comando: Array[0..1024] of Char;
Parms: Array[0..1024] of Char;
begin
StrPCopy (Comando, Nome);
StrPCopy (Parms, Parametros);
ShellExecute (0, Nil, Comando, Parms, Nil, SW_NORMAL);
end;

procedure Delay(dwMilliseconds: Longint);
var  iStart, iStop: DWORD;
begin
  iStart := GetTickCount;
  repeat
    iStop := GetTickCount;
    Application.ProcessMessages;
    Sleep(1);
  until
    (iStop - iStart) >= dwMilliseconds;
  end;


function TrocaVirgPPto(Valor: string): String;
   var i:integer;
begin
    if Valor <>'' then begin
        for i := 0 to Length(Valor) do begin
            if Valor[i]=',' then Valor[i]:='.';

        end;
     end;
     Result := valor;
end;


function Modulo11(Valor: String; Base: Integer = 9; Resto : boolean = false) : string;
var
   Soma : integer;
   Contador, Peso, Digito : integer;
begin
   Soma := 0;
   Peso := 2;
   for Contador := Length(Valor) downto 1 do
   begin
      Soma := Soma + (StrToInt(Valor[Contador]) * Peso);
      if Peso < Base then
         Peso := Peso + 1
      else
         Peso := 2;
   end;

   if Resto then
      Result := IntToStr(Soma mod 11)
   else
   begin
      Digito := 11 - (Soma mod 11);
      if (Digito > 9) then
         Digito := 0;
      Result := IntToStr(Digito);
   end
end;



function Calc_Dias(DataVencimento:TDateTime; DataAtual:TDateTime): String;
Var Data: TDateTime; dia, mes, ano: Word;
idade:currency;
begin
  Data := DataAtual - DataVencimento;
  DecodeDate(Data, Ano, Mes, dia);
  idade := Int(Data);
  Result := floattostr(idade);
end;

function CapturaTela: TBitmap;
var
  dc:hdc;
  cv:TCanvas;
  oBmp: TBitmap;
begin
//  oBmp := TBitmap.Create;
 //  Result := TJPEGImage.Create;
  result := TBitmap.Create;
  result.Width := Screen.Width;
  result.Height := Screen.Height;
  dc := GetDc(0);
  cv := TCanvas.Create;
  cv.Handle := DC;
  result.Canvas.CopyRect(Rect(
    0, 0, Screen.Width, Screen.Height),
    cv, Rect(0,0,Screen.Width, Screen.Height));
  cv.Free;
  ReleaseDC(0, DC);
end;


procedure TFrm_Funcoes.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;
end.
