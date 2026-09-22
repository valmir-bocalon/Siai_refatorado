unit CobrancaBancaria;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, StdCtrls, Grids, DBGrids, Mask,
  DBCtrls, DB, Gauges, XDBEdit, DBClient, ImgList, IniFiles,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, RLSaveDialog, RLFilters,
  RLPDFFilter, RLBoleto, ExtCtrls, System.ImageList, ACBrBoleto,ACBrBoletoConversao,
  ACBrBase, ACBrBoletoFCFortesFr;

type
  TFrmCobrancaBancaria = class(TForm)
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    ImageList1: TImageList;
    DBEEmpre: TDBEdit;
    DBEIdRec: TDBEdit;
    Datacarne: TDataSource;
    zqcarne: TZQuery;


    SaveDialog1: TSaveDialog;
    OpenDialog1: TOpenDialog;
    RLPDFFilter1: TRLPDFFilter;
    RLBRemessa1: TRLBRemessa;
    RLBTitulo1: TRLBTitulo;
    CDParticipante: TClientDataSet;
    CDParticipanteidpaticipante: TLargeintField;
    CDParticipantenome_parte: TStringField;
    CDParticipanteemail: TStringField;
    CDParticipanteconcluido: TBooleanField;
    CDParticipanteEmpreendimento: TStringField;
    CDParticipanteinativo: TStringField;
    CDParticipantenascimento: TDateField;
    CDParticipanteProfissao: TStringField;
    DS_participante: TDataSource;
    ZQAniversariante: TZQuery;



































    DS_ZQAniversariante: TDataSource;
    CDSEmpreendimento: TClientDataSet;
    CDSEmpreendimentoidloteamento: TIntegerField;
    CDSEmpreendimentomarca: TStringField;
    CDSEmpreendimentocodcontabancaria: TIntegerField;
    CDSEmpreendimentoidparticipante: TIntegerField;
    CDSEmpreendimentoidordem: TIntegerField;
    CDSEmpreendimentonomearquivo: TStringField;
    CDSEmpreendimentodigito_dif: TStringField;
    DS_Empreendimento: TDataSource;




    DS_RemesRec: TDataSource;
    DS_MarcaTit: TDataSource;
    CDS_MarcaTit: TClientDataSet;
    CDS_MarcaTitmarca: TStringField;
    CDS_MarcaTitIdreceb: TIntegerField;
    CDS_MarcaTitidparti: TIntegerField;
    CDS_MarcaTitDT_Entrada: TDateField;
    CDS_MarcaTitDt_Vencimento: TDateField;
    CDS_MarcaTitValor: TFloatField;
    CDS_MarcaTitordem: TStringField;
    CDS_MarcaTitnome_parte: TStringField;
    CDS_MarcaTitende_cob: TStringField;
    CDS_MarcaTitcep_cob: TStringField;
    CDS_MarcaTitdoc1: TStringField;
    CDS_MarcaTitTipDoc: TStringField;
    CDS_MarcaTitnumordem: TIntegerField;
    CDS_MarcaTitcliente: TIntegerField;
    CDS_MarcaTitidloteam: TIntegerField;
    CDS_MarcaTitquadralote: TStringField;
    CDS_MarcaTitnumboleto: TStringField;
    CDS_MarcaTitdigito_dif: TStringField;
    CDS_MarcaTitbairro: TStringField;
    CDS_MarcaTitdoc2: TStringField;
    CDS_MarcaTitcidade: TLargeintField;
    CDS_MarcaTitnomecidade: TStringField;
    CDS_MarcaTituf: TStringField;
    ZQNumero_remessa: TZQuery;
    DS_numero_remessa: TDataSource;

    RLBTitulo2: TRLBTitulo;
    CDSEmpreendimentocod_transmissao: TStringField;
    CDSEmpreendimentocomplemento: TStringField;
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    XBanner2: TXBanner;
    DBGEmpre: TDBGrid;
    Panel3: TPanel;
    XBanner3: TXBanner;
    Panel4: TPanel;
    XBanner4: TXBanner;
    DXBMarcTit: TdxButtonArround;
    DBGTit: TDBGrid;
    Panel5: TPanel;
    XBanner5: TXBanner;
    Gauge2: TGauge;
    XDBEdit1: TXDBEdit;
    XDBEdit2: TXDBEdit;
    XDBEdit3: TXDBEdit;
    XDBEdit4: TXDBEdit;
    CBIdentOcorr: TComboBox;
    CBatualizar: TCheckBox;
    Gauge1: TGauge;
    cbemail: TCheckBox;
    CBBancoemite: TCheckBox;
    Label7: TLabel;
    ComboBox3: TComboBox;
    chkemissao: TCheckBox;
    rgnosso: TCheckBox;
    CBTipoArq: TCheckBox;
    dxButton1: TdxButtonArround;
    DBGerar: TdxButtonArround;
    XBRelat: TdxButtonArround;
    DXBFechar: TdxButtonArround;
    Panel6: TPanel;
    XBanner6: TXBanner;
    DXBMarcEmp: TdxButtonArround;
    XBanner1: TXBanner;
    GroupBox2: TGroupBox;
    Label8: TLabel;
    DBGrid1: TDBGrid;
    MaskEdit1: TMaskEdit;
    Button1: TButton;
    CBParcelas: TCheckBox;
    CBbcoGera: TCheckBox;
    CDSEmpreendimentojurosemboleto: TFloatField;
    ACBrBoleto1: TACBrBoleto;
    ACBrBoletoFCFortes1: TACBrBoletoFCFortes;
    Remessa_temp: TClientDataSet;
    Remessa_tempremessa: TIntegerField;
    Remessa_tempidrec: TIntegerField;
    Remessa_tempnossonumero: TStringField;
    cb2via: TCheckBox;
    ZQRemesRec: TZQuery;
    CBCarteiraEnvio: TCheckBox;
    function SoNumero(fField : String): String;
    function SoLetra(fField : String): String;
    procedure BANCOOB_400;
    procedure SICOOB_240;
    procedure Itau_400;
    procedure bradesco_400;
    procedure brasil_400;
    procedure santander_240;
    procedure santander_400;
    procedure caixa_240;
    procedure hsbc_400;
    procedure DXBFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DBGerarClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DBETitKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBETitMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure XDBEdit1Change(Sender: TObject);
    procedure XDBEdit2Change(Sender: TObject);
    procedure XDBEdit3Change(Sender: TObject);
    procedure XDBEdit4Change(Sender: TObject);
    procedure DXBGravarClick(Sender: TObject);
    procedure XBRelatClick(Sender: TObject);
    procedure DBGEmpreColEnter(Sender: TObject);
    procedure DBGEmpreDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGEmpreKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure mudamarcaEmpreend;
    procedure DBGEmpreMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGTitDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBEEmpreChange(Sender: TObject);
    procedure DBGrid1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure mudamarcaTit;
    procedure DXBMarcTitClick(Sender: TObject);
    procedure DXBMarcEmpClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure dxButton1Click(Sender: TObject);
    procedure CBBancoemiteClick(Sender: TObject);
    procedure chkemissaoClick(Sender: TObject);
    procedure cbemailClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure MaskEdit1Change(Sender: TObject);
    procedure cb2viaClick(Sender: TObject);
    procedure MaskEdit1Exit(Sender: TObject);
    procedure DBGrid1Exit(Sender: TObject);
  private


    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmCobrancaBancaria: TFrmCobrancaBancaria;
  Varnomearq : string;
  zqaux :TZQuery;
implementation

uses funcoes, RelRecebimento, Tabelas, RelBolBanco, NumRemessa, uRuntimeFields;

{$R *.dfm}

function TFrmCobrancaBancaria.SoNumero(fField : String): String;
var
  I : Byte;
begin
   Result := '';
   for I := 1 To Length(fField) do
       if fField [I] In ['0'..'9'] Then
            Result := Result + fField [I];
end;


function TFrmCobrancaBancaria.SoLetra(fField : String): String;
var
  I : Byte;
begin
   Result := '';
   for I := 1 To Length(fField) do
       if fField [I] In ['a'..'z','A'..'Z'] Then
            Result := Result + fField [I];
end;

procedure TFrmCobrancaBancaria.BANCOOB_400;
Var
  F : TextFile;
  ArqIni : tIniFile;
  VARARQ, Vararqnome, Varmens1, Varmens2, Varmens3, Varmens4, varnossonumero,
  varnossnum, varmora, vardoc, varcpfcnpj, varendereco, varhoje, VarNomepasta,
  varhostname, vardir, varpastabanco,VarPath, ext,VarDif,ql : string;
  sq,eita,codcli,carne,tam,tam2,varregistro, varremes, varx, vary : integer;
begin
  eita:=0;
  codcli:=0;
  carne:=0;
  tam:=0;
  VarDif:='';
  ql:='';
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  VarNomepasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );

  varpastabanco := FrmRelRecebimento.CDSEmpreendimentoapelido.Value+'_'+FrmRelRecebimento.CDSEmpreendimentonomebanco.Value+FrmRelRecebimento.CDSEmpreendimentoconta.Value;
  vardir := '\\'+varhostname+'\'+VarNomePasta;
  SaveDialog1.Execute;
  tam2:=length(SaveDialog1.FileName);
  tam2:=tam2-11;
  VarPath := copy(SaveDialog1.FileName,tam2+1,11);
  if trim(VarPath)<>'remessa.rem' then
  begin
    mensagem('O nome do Arquivo tem que ser remessa.rem');
    exit;
  end;

  VarPath := copy(SaveDialog1.FileName,1,tam2);
//  if not DirectoryExists(vardir) then
    vardir := VarPath+varpastabanco;
//  else
//    vardir := vardir +'\'+varpastabanco;
  if empty(trim(VarPath)) then
  begin
    mensagem('O nome do Arquivo não pode ficar em branco, tem que ser remessa.rem');
    exit;
  end;

  DBGerar.Enabled:=false;

  CreateDir(vardir);
//  vardir := vardir +'\'+copy(datetostr(date),7,4);
//  CreateDir(vardir);
//  vardir := vardir +'\'+copy(datetostr(date),4,2);
//  CreateDir(vardir);
  vardir := vardir +'\'+copy(datetostr(date),1,2);
  CreateDir(vardir);
  Ext := '.REM';
  if CBTipoArq.Checked Then
    Ext := '.TST';
  varregistro:=1;
  for  varremes:= 1 to 99 do begin
    vararq := inttostrZero(varremes,2);
    Vararqnome := 'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+Ext;
    VARARQ :=vardir+'\'+Vararqnome;
    If not(fileexists(VARARQ)) then
      Break;
  end;
  Varnomearq := VARARQ;
  if DM_Tabelas.ZQBol_men.State in [DsEdit, DsInsert] Then
    DM_Tabelas.ZQBol_men.Edit;

  DM_Tabelas.ZQBancRemes.Close;
  DM_Tabelas.ZQBancRemes.SQL.Clear;
  DM_Tabelas.ZQBancRemes.SQL.Add(' select idbanco_remessa,remessa,idbanco,gerado,nomearq,arquivo,digito_dif,remessa2 from banco_remessa order by remessa2 ');
  DM_Tabelas.ZQBancRemes.open;
  DM_Tabelas.ZQBancRemes.Last;
  varremes := DM_Tabelas.ZQBancRemes.FieldByName('remessa2').AsInteger + 1;


  Frm_NumRemessa.XNENumRemessa.Value := varremes;
  Frm_NumRemessa.Label2.Caption := VARARQ;
  Frm_NumRemessa.Top := FrmCobrancaBancaria.Top + 30;
  Frm_NumRemessa.Left := FrmCobrancaBancaria.Left + 300;
  AbrirModal(Self, Frm_NumRemessa);
  varremes := strtoint(floattostr(Frm_NumRemessa.XNENumRemessa.value));
//  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria;n_dif_empreed',VarArrayOf([FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, DM_Tabelas.CDSEmpreendimentodigito_dif.Value]), []);

  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
//  sq:=DM_Tabelas.ZQContaBancarian_sequencial.value;


  DM_Tabelas.ZQBancRemes.Close;
  DM_Tabelas.ZQBancRemes.SQL.Clear;
  DM_Tabelas.ZQBancRemes.SQL.Add(' select idbanco_remessa,remessa,idbanco,gerado,nomearq,arquivo,digito_dif,remessa2 from banco_remessa ');
  DM_Tabelas.ZQBancRemes.open;

  DM_Tabelas.ZQBancRemes.Insert;
  DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger := FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value;
  DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger := varremes;
  DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := Vararqnome;
  DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := date;
  DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
  DM_Tabelas.ZQBancRemes.Post;


  AssignFile(f,VARARQ);
    
  varhoje := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);
  Rewrite(f); //abre o arquivo para escrita
  Write(f,'0'); // Identificação do Registro
  Write(f,'1'); // Identificação d arquivo remessa
  Write(f,'REMESSA');  // literal remessa
  Write(f,'01'); // codigo de serviço
  Write(f,'COBRANCA'); // literal do serviço
  Write(f,espacos(7));
  Write(f,Formatar(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString,4,false,' ')); // agencia
  Write(f,copy(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString,1,1)); // dig. verificador agencia
  Write(f,Formatar(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString,8,true,' ')); // codigo do cedente conta
  Write(f,copy(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString,1,1)); // dig. verificador conta
  Write(f,Formatar(DM_Tabelas.ZQContaBancaria.FieldByName('convenio').AsString,6,false,'0')); // convenio
  Write(f,copy(FrmRelRecebimento.CDSEmpreendimentonometitular.Value+espacos(30),1,30)); //  razão social da empresa

//  Write(f,inttostrZero(strtofloat(FrmRelRecebimento.CDSEmpreendimentocod_no_banco.Value),20)); //  codigo da empresa no banco (banco fornece)

  Write(f,'756'); // numero do bradesco na camara da compensação
  Write(f,'BANCOOBCED     '); //  Nome do banco por extenso//
  Write(f,varhoje); // data de hoje
  Write(f,inttostrZero(DM_Tabelas.ZQBancRemes.FieldByName('remessa2').AsInteger,7));  //  nº sequencial de remessa
  Write(f,espacos(287)); // branco
  Writeln(f,'000001');
  FrmRelRecebimento.CDS_MarcaTit.First;
  Gauge1.Visible := True;
  Gauge1.MaxValue := FrmRelRecebimento.CDS_MarcaTit.RecordCount;


  ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;


  if zqcarne.active=false then
     zqcarne.open;
  zqcarne.last;
  carne:=zqcarne.FieldByName('n_seq').AsInteger;
  inc(carne);
  //detalhe
  FrmRelRecebimento.CDS_MarcaTit.DisableControls;
  while not FrmRelRecebimento.CDS_MarcaTit.Eof do
  begin
   if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
   begin
    if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '1' Then
    else
    if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
      mensagem('A parcela '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente')
    else Begin

      if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
        varmora := inttostrZero(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,6)
      else
        varmora := inttostrZero(strtofloat(charrem(alltrim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2),'###,##0.00')))),6);
//        varmora := inttostrZero(strtofloat(charrem(floattostr(truncar(FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariataxadiaria.Value,2)))),6);
//        varmora := inttostrZero(truncar(FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariataxadiaria.Value,0),6);
//        varmora := inttostrZero(ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariataxadiaria.Value,2),13);
      varendereco := FrmRelRecebimento.CDS_MarcaTitende_cob.Value;
//      ZQRemesRec.Insert;
//      ZQRemesRecremessa.Value := DM_Tabelas.ZQBancRemesidbanco_remessa.Value;
//      ZQRemesRecidrec.Value := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
//      ZQRemesRecnossonumero.value :=varnossnum;
//      ZQRemesRec.Post;
      inc(varregistro);
      Gauge1.Progress := FrmRelRecebimento.CDS_MarcaTit.RecNo;
      eita:=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
      if eita=1 then
         codcli:=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
      if eita=2 then
         codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
      if eita=3 then
         codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
      if eita=4 then
         codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
      if eita>=5 then
         codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
      eita:=codcli;
      codcli:=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
      vardoc := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);
      varnossonumero := DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString+inttostrZero(codcli,11);
         // qdo precisar trocar o numero do nosso numero //comenta as 2 linhs anteriores e liberar as 2 seguinte
//      vardoc := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);
//      varnossonumero := DM_Tabelas.ZQContaBancariacarteira.Value+inttostrZero( FrmRelRecebimento.CDS_MarcaTitIdreceb.Value,11);

      varnossnum:='32765432765432';
      vary := 0;
      for varx := 1 to 14 do
        vary := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(varnossonumero,varx,1)));
      vary := (vary mod 11);
      vary := 11 - vary ;
      if vary = 10 Then
        varnossnum := inttostrZero(codcli,11)+'P'
      else if vary = 11
       Then
        varnossnum := inttostrZero(codcli,11)+'0'
      else
        varnossnum := inttostrZero(codcli,11)+inttostr(vary);
      codcli:=0;
      eita:=0;
      ZQRemesRec.Filtered:=false;
      ZQRemesRec.Filter:='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
      ZQRemesRec.Filtered:=true;
      if ZQRemesRec.RecordCount>0 then
      begin
        ZQRemesRec.First;
        while not ZQRemesRec.Eof do
        begin
          ZQRemesRec.delete;
        end;
        ZQRemesRec.Filtered:=false;
        ZQRemesRec.Insert;
        ZQRemesRec.FieldByName('remessa').AsInteger := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
        ZQRemesRec.FieldByName('idrec').AsInteger := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
        ZQRemesRec.FieldByName('nossonumero').AsString :=varnossnum;
        ZQRemesRec.Post;
      end
      else
      begin
        ZQRemesRec.Filtered:=false;
        ZQRemesRec.Insert;
        ZQRemesRec.FieldByName('remessa').AsInteger := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
        ZQRemesRec.FieldByName('idrec').AsInteger := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
        ZQRemesRec.FieldByName('nossonumero').AsString :=varnossnum;
        ZQRemesRec.Post;
      end;
      ZQRemesRec.Filtered:=false;
      VarDif :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;

      Write(f,'1'); // Identificação do Registro
      if length(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)=14 Then
        Write(f,'01')   // tipo cpf
      else
        Write(f,'02');  // tipo cnpj
      varcpfcnpj := inttostrZero(strtofloat(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)),14);
      Write(f,varcpfcnpj);  // cpf/cnpj
      Write(f,Formatar(FrmRelRecebimento.CDSEmpreendimenton_agencia.Value,4,false,' ')); // agencia
      Write(f,copy(FrmRelRecebimento.CDSEmpreendimenton_agencia_v.value,1,1)); // dig. verificador agencia
      Write(f,Formatar(FrmRelRecebimento.CDSEmpreendimentoconta.Value,8,true,' ')); // codigo do cedente conta
      Write(f,copy(FrmRelRecebimento.CDSEmpreendimentoconta_v.Value,1,1)); // dig. verificador conta
      Write(f,Formatar(FrmRelRecebimento.CDSEmpreendimentoconvenio.Value,6,false,'0')); // convenio
       //carne

      if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
      begin
        if empty(VarDif) then // coluna 38 até 62
        begin
          Write(f,copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,25));  // uso da empresa, nº de controle da empresa;
        end
        else
        begin
          Write(f,copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(7),1,24));  // uso da empresa, nº de controle da empresa;
          Write(f,VarDif);
        end;
      end
      else
      begin
        inc(carne);
        ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
        if empty(VarDif) then // coluna 38 até 62
        begin
          Write(f,copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,25));  // uso da empresa, nº de controle da empresa;
        end
        else
        begin
          Write(f,copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(7),1,24));  // uso da empresa, nº de controle da empresa;
          Write(f,VarDif);
        end;
      end;
      Write(f,varnossnum); //Nº bancario para cobrana com e em registro
      Write(f,'0000   ');  //  zero

      Write(f,' '); //Indicador de Sacador
      Write(f,'   '); //Prefixo do título
      Write(f,'000'); //Variação
      Write(f,'0'); //Conta Caução
      Write(f,'000000'); //Código de Responsabilidade + DV
      Write(f,'000000'); //Número do borderô
      Write(f,'00000'); //filter
      Write(f,'01'); //Carteira/Modalidade
      Write(f,'00'); //Comando/Movimento
      Write(f,vardoc); //Seu Número
      Write(f,copy(FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Text,1,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Text,4,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Text,9,2));  // Vencimento
      Write(f,inttostrZero(ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2)*100,13));   //  Valor do titulo
      Write(f,'756'); // numero do banco na camara da compensação
 //     Write(f,Formatar(FrmRelRecebimento.CDSEmpreendimenton_agencia.Value,4,false,' ')); // agencia
//      Write(f,copy(FrmRelRecebimento.CDSEmpreendimenton_agencia_v.value,1,1)); // dig. verificador agencia
      Write(f,'00000'); // Agencia depositaria
      Write(f,'01'); //Espécie do Título
      Write(f,'0'); //Aceite do Título
      Write(f,copy(FrmRelRecebimento.CDS_MarcaTitDt_Entrada.Text,1,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Entrada.Text,4,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Entrada.Text,9,2));  // emissao      
      Write(f,'00'); //Primeira instrução codificada
      Write(f,'00'); //Segunda instrução
      Write(f,varmora); // Mora por dia de atraso
      Write(f,'000000'); // Taxa de multa
      Write(f,'0'); // Filter
      Write(f,'000000'); //copy(DM_Tabelas.ZQRecebimentoDt_Vencimento.Text,1,2)+ copy(DM_Tabelas.ZQRecebimentoDt_Vencimento.Text,4,2)+ copy(DM_Tabelas.ZQRecebimentoDt_Vencimento.Text,9,2));  // Limite para desconto
      Write(f,'0000000000000');  //  Valor do desconto
      Write(f,'0000000000000');  //  Valor IOF
      Write(f,'0000000000000');  //  Valor abatimento
      if length(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)=14 Then
        Write(f,'01')   // tipo cpf
      else
        Write(f,'02');  // tipo cnpj
      varcpfcnpj := inttostrZero(strtofloat(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)),14);
      Write(f,varcpfcnpj);  // cpf/cnpj
      Write(f,copy(removeacentos(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+espacos(40),1,40));  // nome do sacado  coluna 235
      Write(f,copy(removeacentos(varendereco)+espacos(37),1,37));  // Endereco   275 a 314
      Write(f,copy(FrmRelRecebimento.CDS_MarcaTitbairro.Value+espacos(15),1,15)) ;
      Write(f,copy(charrem(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value)+espacos(8),1,8)) ;
      Write(f,copy(charrem(removeacentos(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value))+espacos(15),1,15)) ;
      Write(f,copy(charrem(FrmRelRecebimento.CDS_MarcaTituf.Value)+espacos(2),1,2)) ;
      Write(f,COPY('SR. CAIXA, NAO RECEBER APOS 30 (TRINTA) DIAS DE VENCIMENTO.'+espacos(40),1,40));  // decomposição
      Write(f,'00');  // Protesto
      Write(f,' ');  //  Filler
      Writeln(f,inttostrZero(varregistro,6)); //  nº seguencial
{      tam:=length(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text);
      tam:=37+tam;
      tam:=49-tam;

      VarDif:='';
      Write(f,'   '); // codigo do banco a ser debitado na camara de compensação
//      Write(f,'237'); // codigo do banco a ser debitado na camara de compensação
      Write(f,'2'); // cobrar multa
      Write(f,inttostrZero(truncar(FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Value*100,0),4));  // % de multa a ser considerado

      Write(f,inttostrZero(0,10));  // Valor do desconto bonif. / dia
      if CBBancoemite.Checked Then
        Write(f,'1')  //   banco emite e registra o beleto
      else
        Write(f,'2');  //   nós emitimos o boleto e o banco registra o beleto
      Write(f,'N');  //  Não registra na cobrança. Diferente de N registra e emite boleto caso os dados estejam errados  coluna 84
//      Write(f,inttostrZero(0,10));  // Identificação da operação do banco
      Write(f,espacos(10));  //  Identificação da operação do banco
      Write(f,' ');  //  Indicador rateio credito
      Write(f,' ');  //  cod. endereco de aviso do sacado
      Write(f,'  ');  //  Brancos              coluna 107
      Write(f,copy(CBIdentOcorr.Items.Strings[CBIdentOcorr.ItemIndex],1,2));  //  codigo de ocorrencia
      Write(f,vardoc);  //  nº do documento      coluna 111

//      Write(f,inttostrZero(truncar(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2)*100,13));   //  Valor do titulo
      Write(f,inttostrZero(ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2)*100,13));   //  Valor do titulo
      Write(f,'000');  // Banco encarregado da cobrança
      Write(f,'00000'); // Agencia depositaria
      Write(f,'01');  // Especie de Titulo  01 = duplicata
      Write(f,'N');  //  identificação - sempre N             coluna 150
      Write(f,copy(FrmRelRecebimento.CDS_MarcaTitDt_Entrada.Text,1,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Entrada.Text,4,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Entrada.Text,9,2));  // emissao
      Write(f,'0000'); //  1ª e 2ª instrução
      Write(f,varmora); // Mora por dia de atraso
      Write(f,'000000'); //copy(DM_Tabelas.ZQRecebimentoDt_Vencimento.Text,1,2)+ copy(DM_Tabelas.ZQRecebimentoDt_Vencimento.Text,4,2)+ copy(DM_Tabelas.ZQRecebimentoDt_Vencimento.Text,9,2));  // Limite para desconto
      Write(f,'0000000000000');  //  Valor do desconto
      Write(f,'0000000000000');  //  Valor IOF
      Write(f,'0000000000000');  //  Valor abatimento
      if length(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)=14 Then
        Write(f,'01')   // tipo cpf
      else
        Write(f,'02');  // tipo cnpj
      varcpfcnpj := inttostrZero(strtofloat(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)),14);
      Write(f,varcpfcnpj);  // cpf/cnpj
      Write(f,copy(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+espacos(40),1,40));  // nome do sacado  coluna 235
      Write(f,copy(varendereco+espacos(40),1,40));  // Endereco   275 a 314
      Write(f,espacos(12)); // 1ª mensagem 315 a 326
//      Write(f,copy(FrmRelRecebimento.CDS_MarcaTitquadralote.Value,1,12)); // 1ª mensagem 315 a 326
      Write(f,copy(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value)+espacos(8),1,8));  // cep
      Write(f,COPY('SR. CAIXA, NºO RECEBER APÓS 30 (TRINTA) DIAS DE VENCIMENTO.'+espacos(60),1,60));  // decomposição}



    {  varmens1 := copy(DM_Tabelas.ZQBol_menlinha1.Value+espacos(80),1,80);
      varmens2 := copy(DM_Tabelas.ZQBol_menlinha2.Value+espacos(80),1,80);
      varmens3 := copy(DM_Tabelas.ZQBol_menlinha3.Value+espacos(80),1,80);
      varmens4 := copy(DM_Tabelas.ZQBol_menlinha4.Value+espacos(80),1,80);
      if not empty(varmens1+varmens2+varmens3+varmens4) Then
      Begin
        inc(varregistro);
        Write(f,'2');  //  Tipo de Registro
        Write(f,varmens1+varmens2+varmens3+varmens4);  //  mensagens 80 cacacteres cada
        Write(f,espacos(45));  //  reserva   -  filler
        Write(f,DM_Tabelas.ZQContaBancariacarteira.Value);  // Código da carteira
        Write(f,inttostrZero(strtofloat(FrmRelRecebimento.CDSEmpreendimenton_agencia.Value),5));  // codigo da agencia cedente, sem o digito
        Write(f,inttostrZero(strtofloat(FrmRelRecebimento.CDSEmpreendimentoconta.Value),7)); // conta corrente
        Write(f,copy(FrmRelRecebimento.CDSEmpreendimentoconta_v.Value,1,1)); // digito da conta
        Write(f,varnossnum); //Nº bancario para cobrana com e em registro
        Writeln(f,inttostrZero(varregistro,6)); //  nº seguencial
      end;}
    end;
    FrmRelRecebimento.CDS_MarcaTit.Next;
   end
   else
   begin
     FrmRelRecebimento.CDS_MarcaTit.Next;
   end;
  end;
  FrmRelRecebimento.CDS_MarcaTit.EnableControls;
  zqcarne.Edit;
  zqcarne.FieldByName('n_seq').AsInteger:=carne;
  zqcarne.post;
  zqcarne.Close;
  //TRAILLER
  inc(varregistro);
  Write(f,'9');  // identificação do registro
  Write(f,espacos(393));   // branco
  Writeln(f,inttostrZero(varregistro,6)); //  nº seguencial
  Closefile(f); //fecha o handle de arquivo
  Gauge1.Visible := False;

  Gauge1.Visible := True;
  Gauge1.Progress:=0;
  ZQRemesRec.First;
  Gauge1.MaxValue := ZQRemesRec.RecordCount;
//  DM_Tabelas.ZQRemes_Receb_atualiza.open;
//  DM_Tabelas.ZQRemes_Receb_atualiza.first;

  ZQRemesRec.DisableControls;
  while not ZQRemesRec.Eof do
  begin
    Gauge1.Progress := ZQRemesRec.RecNo;
    DM_Tabelas.ZQRemes_Receb_atualiza.Close;
    DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
    DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text));
    DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;


{    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRemes_Receb_atualiza.sql.clear;
    DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('SELECT * FROM recebimento where idrecebimento='+quotedstr(ZQRemesRecidrec.Text));
    DM_Tabelas.ZQRemes_Receb_atualiza.open;
    if DM_Tabelas.ZQRemes_Receb_atualiza.recordcount>0 then
    begin
      DM_Tabelas.ZQRemes_Receb_atualiza.Edit;
      DM_Tabelas.ZQRemes_Receb_atualizanumboleto.Value:=ZQRemesRecnossonumero.Value;
      DM_Tabelas.ZQRemes_Receb_atualiza.post;
    end;}
    ZQRemesRec.Next;
  end;
  ZQRemesRec.EnableControls;
//  DM_Tabelas.ZQRemes_Receb_atualiza.refresh;
//  DM_Tabelas.ZQRecebimento.refresh;
  DM_Tabelas.ZQRemes_Receb_atualiza.close;
  Gauge1.Progress:=0;
  Gauge1.Visible := false;
  mensagem('Relatório de Boletos gerado com sucesso!!!'+chr(13)+'O relatório foi gravado em '+VARARQ);
end;

procedure TFrmCobrancaBancaria.SICOOB_240;
Var
  F : TextFile;
  ArqIni : tIniFile;
  VARARQ, Vararqnome, Varmens1, Varmens2, Varmens3, Varmens4, varnossonumero,
  varnossnum, varmora, vardoc, varcpfcnpj, varendereco, varhoje, VarNomepasta,
  varhostname, vardir, varpastabanco,VarPath, ext,VarDif,ql : string;
  vpar,NN,sq,eita,codcli,carne,tam,tam2,varregistro, varremes, varx, vary : integer;
  Titulo : TACBrTitulo;
  I : TACBrBolLayOut ;
begin
  eita:=0;
  codcli:=0;
  carne:=0;
  tam:=0;
  VarDif:='';
  ql:='';
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  VarNomepasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );

  VARARQ:=ExtractFilePath( Application.ExeName );

  varpastabanco := FrmRelRecebimento.CDSEmpreendimentoapelido.Value+'_'+FrmRelRecebimento.CDSEmpreendimentonomebanco.Value+FrmRelRecebimento.CDSEmpreendimentoconta.Value;
  vardir := '\\'+varhostname+'\'+VarNomePasta;
  if cb2via.Checked=false then
  begin

      SaveDialog1.Execute;
      tam2:=length(SaveDialog1.FileName);
      tam2:=tam2-11;
      VarPath := copy(SaveDialog1.FileName,tam2+1,11);
      if trim(VarPath)<>'remessa.rem' then
      begin
        mensagem('O nome do Arquivo tem que ser remessa.rem');
        exit;
      end;

      VarPath := copy(SaveDialog1.FileName,1,tam2);
    //  if not DirectoryExists(vardir) then
        vardir := VarPath+varpastabanco;
    //  else
    //    vardir := vardir +'\'+varpastabanco;
      if empty(trim(VarPath)) then
      begin
        mensagem('O nome do Arquivo não pode ficar em branco, tem que ser remessa.rem');
        exit;
      end;

      DBGerar.Enabled:=false;

      CreateDir(vardir);
      vardir := vardir +'\'+copy(datetostr(date),1,2);
      CreateDir(vardir);
      Ext := '.REM';
      if CBTipoArq.Checked Then
        Ext := '.TST';
      varregistro:=1;
      for  varremes:= 1 to 99 do begin
        vararq := inttostrZero(varremes,2);
        Vararqnome := 'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+Ext;
        VARARQ :=vardir+'\'+Vararqnome;
        If not(fileexists(VARARQ)) then
          Break;
      end;
      Varnomearq := VARARQ;
      if DM_Tabelas.ZQBol_men.State in [DsEdit, DsInsert] Then
        DM_Tabelas.ZQBol_men.Edit;

      //pegar o ultimo registro auto incremento criado

     Remessa_temp.Close;
     Remessa_temp.CreateDataSet;
     Remessa_temp.Open;

      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

      DM_tabelas.ZQsequencia_remessa.close;
      DM_tabelas.ZQsequencia_remessa.SQL.Clear;
      DM_tabelas.ZQsequencia_remessa.SQL.Add('select * from remessa_sequencia where idconta_bancaria='+quotedstr(FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Text));
      DM_tabelas.ZQsequencia_remessa.Open;
      if DM_tabelas.ZQsequencia_remessa.RecordCount>0 then
         DM_tabelas.ZQsequencia_remessa.edit
      else
          DM_tabelas.ZQsequencia_remessa.insert;
      DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt:=DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt+1;
      DM_tabelas.ZQsequencia_remessa.Post;
      DM_tabelas.ZQsequencia_remessa.Refresh;
      varremes := DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt;//DM_Tabelas.ZQContaBancarian_sequencial.Value+1;
      DM_tabelas.ZQsequencia_remessa.close;


      Frm_NumRemessa.XNENumRemessa.Value := varremes;
      Frm_NumRemessa.Label2.Caption := VARARQ;
      Frm_NumRemessa.Top := FrmCobrancaBancaria.Top + 30;
      Frm_NumRemessa.Left := FrmCobrancaBancaria.Left + 300;
      AbrirModal(Self, Frm_NumRemessa);
      //numero da remessa

      varremes := strtoint(floattostr(Frm_NumRemessa.XNENumRemessa.value));
  end;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  if cb2via.Checked=false then
  begin
    DM_Tabelas.ZQBancRemes.Close;
    DM_Tabelas.ZQBancRemes.SQL.Clear;
    DM_Tabelas.ZQBancRemes.SQL.Add(' select idbanco_remessa,remessa,idbanco,gerado,nomearq,arquivo,digito_dif,remessa2 from banco_remessa ');
    DM_Tabelas.ZQBancRemes.open;

    DM_Tabelas.ZQBancRemes.Insert;
    DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger := FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value;
    DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger := varremes;
    DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := Vararqnome;
    DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := date;
    DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
    DM_Tabelas.ZQBancRemes.Post;
  end;

//  if not DirectoryExists(ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB') then
  //       ForceDirectories(ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB');
  ACBrBoleto1.DirArqRemessa := VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
  ACBrBoleto1.LayoutRemessa:=c240;
  case CBIdentOcorr.ItemIndex of
      1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
      2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
      3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
      4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
      5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
      6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
      7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
      8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
      9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
      10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
      11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
      12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
      13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
      14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
      15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
      16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
      17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
  end;
  ACBrBoleto1.Banco.TipoCobranca:=cobBancoob;
  ACBrBoleto1.Banco.TamanhoMaximoNossoNum:=7;
  if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
  begin
     ACBrBoleto1.Banco.LayoutVersaoArquivo:=81;
     ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=81;
  end
  else
  begin
     ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
     ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
  end;
  if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
  begin
     ACBrBoleto1.Banco.LayoutVersaoLote:=40;
     ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=40;
  end
  else
  begin
    ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
    ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
  end;

  ACBrBoletoFCFortes1.DirLogo:=ExtractFilePath( Application.ExeName)+'Logos_bco';
  FrmRelRecebimento.CDS_MarcaTit.First;
  Gauge1.Visible := True;
  Gauge1.MaxValue := FrmRelRecebimento.CDS_MarcaTit.RecordCount;


  ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

  carne:=0;

  if cb2via.Checked=false then
  begin

    if zqcarne.active=false then
       zqcarne.open;
    zqcarne.last;
    carne:=zqcarne.FieldByName('n_seq').AsInteger;
    inc(carne);
  end;


  //detalhe
  FrmRelRecebimento.CDS_MarcaTit.DisableControls;
  ACBrBoleto1.ListadeBoletos.Clear;

  ACBrBoleto1.Cedente.Agencia           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
  ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
  ACBrBoleto1.Cedente.AgenciaDigito     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
  ACBrBoleto1.Cedente.Conta             := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
  ACBrBoleto1.Cedente.ContaDigito       := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
  if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
     ACBrBoleto1.Cedente.Modalidade     :='01'
  else
     ACBrBoleto1.Cedente.Modalidade     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);
  ACBrBoleto1.Cedente.CodigoCedente     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
  ACBrBoleto1.Cedente.CodigoTransmissao := trim(DM_Tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString);

  case ComboBox3.ItemIndex of
       0: ACBrBoleto1.ACBrBoletoFC.LayOut := TACBrBolLayOut(lPadrao);
       1: ACBrBoleto1.ACBrBoletoFC.LayOut := TACBrBolLayOut(lCarne);
  end;

//          if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
  if length(charrem(trim(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString)))=14 then
  begin
      ACBrBoleto1.Cedente.CNPJCPF       := trim(charrem(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString));
      ACBrBoleto1.Cedente.TipoInscricao := pJuridica;
  end
  else
  begin
      ACBrBoleto1.Cedente.CNPJCPF       := trim(charrem(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString));
      ACBrBoleto1.Cedente.TipoInscricao := pFisica
  end;

  ACBrBoleto1.Cedente.Nome             := trim(FrmRelRecebimento.CDSEmpreendimentonometitular.Value);
 // ACBrBoleto1.Cedente.Logradouro       := trim(DM_Tabelas.zqempresaendereco.Value);
 // tam:=pos(',',DM_Tabelas.zqempresaendereco.Value);
//          inc(tam);
//        if tam>1 then
//         ACBrBoleto1.Cedente.NumeroRes     := SoNumero(Trim(DM_Tabelas.zqempresaendereco.Value))
//    else
//     ACBrBoleto1.Cedente.NumeroRes     := 'SN';
  try
    zqaux:= tzquery.Create(nil);
    zqaux.Connection := DM_Tabelas.zconeccao;
    zqaux.Close;
    zqaux.sql.clear;
    zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
    zqaux.Open;
    if zqaux.RecordCount>0 then
    begin
        ACBrBoleto1.Cedente.Logradouro        := RemoveAcentos(uppercase(trim(zqaux.FieldByName('ende_cob').AsString)));
      //  ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
        ACBrBoleto1.Cedente.Complemento       := '';
        ACBrBoleto1.Cedente.Bairro            := RemoveAcentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
        ACBrBoleto1.Cedente.Cidade            := RemoveAcentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
        ACBrBoleto1.Cedente.UF                := RemoveAcentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
        ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
    end;
  finally
    zqaux.Close;
    zqaux.free;
  end;

  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  NN:=DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;
  while not FrmRelRecebimento.CDS_MarcaTit.Eof do
  begin

   if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
   begin
    if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '1' Then
    else
    if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
      mensagem('A parcela '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente')
    else Begin

        DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
        Titulo := ACBrBoleto1.CriarTituloNaLista;
        with Titulo do
        begin
                {TACBrTipoCobranca = (cobNenhum, cobBancoDoBrasil, cobSantander, cobCaixaEconomica, cobCaixaSicob, cobBradesco, cobItau, cobBancoMercantil, cobSicred,
                                      cobBancoob, cobBanrisul, cobBanestes, cobHSBC, cobBancoDoNordeste, cobBRB, cobBicBanco, cobBradescoSICOOB, cobBancoSafra,
                                      cobSafraBradesco, cobBancoCECRED );}

          EspecieDoc        := 'DM';
          Vencimento        := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
          DataDocumento     := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;

          eita:=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
          if eita=1 then
             codcli:=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
          if eita=2 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
          if eita=3 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
          if eita=4 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
          if eita>=5 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
          eita:=codcli;
          codcli:=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
          vardoc := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);

          NumeroDocumento   := vardoc;
          SeuNumero         := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
          EspecieMod        := '$';
          Aceite            := atNao; //atSim
          DataProcessamento := Now;
          if CBCarteiraEnvio.Checked then
             CarteiraEnvio     := tceBanco   // vai retornar '1'   tceCedente; // vai retornar '2'
          else
             CarteiraEnvio     := tceCedente;
          Carteira          := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);
          if (rgnosso.Checked) and (CBbcoGera.Checked=false) then
          begin
            inc(NN);
            varnossnum:=Trim(strzero(NN,7));
          end;
          if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
          begin
           if CBbcoGera.Checked=false then
           begin
              varnossnum :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;
              NN := varnossnum.ToInteger;
           end;
          end;

         if cb2via.Checked=false then
         begin

          ZQRemesRec.Filtered:=false;
          ZQRemesRec.Filter:='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
          ZQRemesRec.Filtered:=true;
          if ZQRemesRec.RecordCount>0 then
          begin
            ZQRemesRec.First;
            while not ZQRemesRec.Eof do
            begin
              ZQRemesRec.delete;
            end;
            ZQRemesRec.Filtered:=false;
            ZQRemesRec.Insert;
            ZQRemesRec.FieldByName('remessa').AsInteger     := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
            ZQRemesRec.FieldByName('idrec').AsInteger       := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
            ZQRemesRec.FieldByName('nossonumero').AsString := varnossnum;
            ZQRemesRec.Post;
          end
          else
          begin
            ZQRemesRec.Filtered:=false;
            ZQRemesRec.Insert;
            ZQRemesRec.FieldByName('remessa').AsInteger     := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
            ZQRemesRec.FieldByName('idrec').AsInteger       := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
            ZQRemesRec.FieldByName('nossonumero').AsString := varnossnum;
            ZQRemesRec.Post;
          end;
          ZQRemesRec.Filtered:=false;

         end;


         if cb2via.Checked=false then
         begin

            Remessa_temp.Filtered:=false;
            Remessa_temp.Filter:='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
            Remessa_temp.Filtered:=true;
            if Remessa_temp.RecordCount>0 then
            begin
              Remessa_temp.First;
              while not Remessa_temp.Eof do
              begin
                Remessa_temp.delete;
              end;
              Remessa_temp.Filtered:=false;
              Remessa_temp.Insert;
              Remessa_tempremessa.Value     := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
              Remessa_tempidrec.Value       := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
              Remessa_tempnossonumero.value := varnossnum;
              Remessa_temp.Post;
            end
            else
            begin
              Remessa_temp.Filtered:=false;
              Remessa_temp.Insert;
              Remessa_tempremessa.Value     := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
              Remessa_tempidrec.Value       := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
              Remessa_tempnossonumero.value := varnossnum;
              Remessa_temp.Post;
            end;
            Remessa_temp.Filtered:=false;

         end;

          VarDif :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;

          NossoNumero       := varnossnum;
          ValorDocumento    := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);

          if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
          begin
             vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
             inc(vpar);
             TotalParcelas     := strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
             vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
             dec(vpar);
             parcela           :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
          end
          else
          begin
             TotalParcelas     :=  10;
             parcela           :=  1;
          end;


        //dados do sacado
          Sacado.NomeSacado:=Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
         if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
          begin
              Sacado.CNPJCPF:= trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
              Sacado.Pessoa := pJuridica;
          end
          else
          begin
              Sacado.CNPJCPF:= trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
              Sacado.Pessoa := pFisica;
          end;
          Sacado.Logradouro := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
//          Sacado.Numero     := SoNumero(Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value));
          Sacado.Bairro     := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);
          Sacado.Cidade     := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
          Sacado.UF         := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
          Sacado.CEP        := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
          ValorAbatimento   := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
          LocalPagamento    := 'PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO';

          CodigoMoraJuros   := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
          CodigoMora        := '1'; //1-por dia 2-mensal 3-isento
          CodigoMulta       := cmPercentual;//(cmValorFixo, cmPercentual)
          DataMulta         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

          if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
             PercentualMulta   := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
          else
             PercentualMulta   := 0;

      //            MultaValorFixo    := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
          MultaValorFixo    := False;

         if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
            ValorMoraJuros    := DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat
         else
            ValorMoraJuros    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);


         DataMoraJuros     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

//            DataProtesto      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;
         DataProtesto      := 0;


         if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
         begin
            {  TACBrTipoDesconto = (
            tdNaoConcederDesconto,
            tdValorFixoAteDataInformada,
            tdPercentualAteDataInformada,
            tdValorAntecipacaoDiaCorrido,
            tdValorAntecipacaoDiaUtil,
            tdPercentualSobreValorNominalDiaCorrido,
            tdPercentualSobreValorNominalDiaUtil,
            tdCancelamentoDesconto);}


           CodigoDesconto    := cdValorFixo; //(cdSemDesconto, cdValorFixo);
           TipoDesconto      := tdValorFixoAteDataInformada;
           DataDesconto      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
           ValorDesconto     := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);

         end;

         if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
         begin
              if empty(VarDif) then // coluna 38 até 62
              begin
                Mensagem.Text     := 'SR. CAIXA, NAO RECEBER APOS 30 (TRINTA) DIAS DE VENCIMENTO. N.Tit.'+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+'  N.Carnê:'+strzero(carne,5);
              end
              else
              begin
                Mensagem.Text     := 'SR. CAIXA, NAO RECEBER APOS 30 (TRINTA) DIAS DE VENCIMENTO. N.Tit.'+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+' N.Carnê:'+strzero(carne,5)+'-'+Vardif;
              end;
         end
         else
         begin
              inc(carne);
              ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
              if empty(VarDif) then // coluna 38 até 62
              begin
                Mensagem.Text     := 'SR. CAIXA, NAO RECEBER APOS 30 (TRINTA) DIAS DE VENCIMENTO. N.Tit.'+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+' N.Carnê:'+strzero(carne,5);
              end
              else
              begin
                Mensagem.Text     := 'SR. CAIXA, NAO RECEBER APOS 30 (TRINTA) DIAS DE VENCIMENTO. N.Tit.'+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+' N.Carnê:'+strzero(carne,5)+'-'+Vardif;
              end;
         end;
//         if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2)>0 then
//         begin
//            Mensagem.Text := Mensagem.Text +#13+' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'));
//         end;
            case CBIdentOcorr.ItemIndex of
                1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
            end;
        end;
    end;
    FrmRelRecebimento.CDS_MarcaTit.Next;
   end
   else
   begin
     FrmRelRecebimento.CDS_MarcaTit.Next;
   end;
  end;
  try
    if cb2via.Checked=false then
       ACBrBoleto1.GerarRemessa( varremes );
    if CBBancoemite.State=cbUnChecked then
    begin
       ACBrBoleto1.Imprimir;
 //      tam:=pos('.',vardir+'\'+Vararqnome);
 //      dec(tam);
 //      ACBrBoletoFCFortes1.NomeArquivo:=copy(vardir+'\'+Vararqnome,1,tam)+'.pdf';
 //      ACBrBoleto1.GerarPDF;


    end;
  finally
    FrmRelRecebimento.CDS_MarcaTit.EnableControls;
    if cb2via.Checked=false then
    begin
      zqcarne.Edit;
      zqcarne.FieldByName('n_seq').AsInteger:=carne;
      zqcarne.post;
    end;
    zqcarne.Close;
    Gauge1.Visible := False;
    if cb2via.Checked=false then
    begin
      Gauge1.Visible := True;
      Gauge1.Progress:=0;
      Remessa_temp.First;
      Gauge1.MaxValue := Remessa_temp.RecordCount;
      Remessa_temp.DisableControls;
      while not Remessa_temp.Eof do
      begin
        Gauge1.Progress := Remessa_temp.RecNo;

        DM_Tabelas.ZQRemes_Receb_atualiza.Close;
        DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
        DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(Remessa_tempnossonumero.Value)+') where idrecebimento='+quotedstr(Remessa_tempidrec.Text));
        DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;
        Remessa_temp.Next;
      end;
    end;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRecebimento.refresh;

    DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
    if cb2via.Checked=false then
    begin

      DM_Tabelas.ZQContaBancaria.Edit;
      DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt:=NN;
      DM_Tabelas.ZQContaBancaria.Post;
    end;
    if cb2via.Checked=false then
       Remessa_temp.EnableControls;
    Remessa_temp.Close;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    Gauge1.Progress:=0;
    Gauge1.Visible := false;
    if cb2via.Checked=false then
       mensagem('Relatório de Boletos gerado com sucesso!!!'+chr(13)+'O relatório foi gravado em '+VARARQ);
  end;


end;




procedure TFrmCobrancaBancaria.Itau_400;
Var
  F : TextFile;
  ArqIni : tIniFile;
  VARARQ, Vararqnome, Varmens1, Varmens2, Varmens3, Varmens4, varnossonumero,
  varnossnum, varmora, vardoc, varcpfcnpj, varendereco, varhoje, VarNomepasta,
  varhostname, vardir, varpastabanco,VarPath, ext,VarDif,ql : string;
  vpar,NN,sq,eita,codcli,carne,tam,tam2,varregistro, varremes, varx, vary : integer;
  Titulo : TACBrTitulo;
  I : TACBrBolLayOut ;
begin
  eita:=0;
  codcli:=0;
  carne:=0;
  tam:=0;
  VarDif:='';
  ql:='';
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  VarNomepasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );

  VARARQ:=ExtractFilePath( Application.ExeName );

  varpastabanco := FrmRelRecebimento.CDSEmpreendimentoapelido.Value+'_'+FrmRelRecebimento.CDSEmpreendimentonomebanco.Value+FrmRelRecebimento.CDSEmpreendimentoconta.Value;
  vardir := '\\'+varhostname+'\'+VarNomePasta;
  if cb2via.Checked=false then
  begin

      SaveDialog1.Execute;
      tam2:=length(SaveDialog1.FileName);
      tam2:=tam2-11;
      VarPath := copy(SaveDialog1.FileName,tam2+1,11);
      if trim(VarPath)<>'remessa.rem' then
      begin
        mensagem('O nome do Arquivo tem que ser remessa.rem');
        exit;
      end;

      VarPath := copy(SaveDialog1.FileName,1,tam2);
    //  if not DirectoryExists(vardir) then
        vardir := VarPath+varpastabanco;
    //  else
    //    vardir := vardir +'\'+varpastabanco;
      if empty(trim(VarPath)) then
      begin
        mensagem('O nome do Arquivo não pode ficar em branco, tem que ser remessa.rem');
        exit;
      end;

      DBGerar.Enabled:=false;

      CreateDir(vardir);
      vardir := vardir +'\'+copy(datetostr(date),1,2);
      CreateDir(vardir);
      Ext := '.REM';
      if CBTipoArq.Checked Then
        Ext := '.TST';
      varregistro:=1;
      for  varremes:= 1 to 99 do begin
        vararq := inttostrZero(varremes,2);
        Vararqnome := 'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+Ext;
        VARARQ :=vardir+'\'+Vararqnome;
        If not(fileexists(VARARQ)) then
          Break;
      end;
      Varnomearq := VARARQ;
      if DM_Tabelas.ZQBol_men.State in [DsEdit, DsInsert] Then
        DM_Tabelas.ZQBol_men.Edit;

      //pegar o ultimo registro auto incremento criado

     Remessa_temp.Close;
     Remessa_temp.CreateDataSet;
     Remessa_temp.Open;

      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

      DM_tabelas.ZQsequencia_remessa.close;
      DM_tabelas.ZQsequencia_remessa.SQL.Clear;
      DM_tabelas.ZQsequencia_remessa.SQL.Add('select * from remessa_sequencia where idconta_bancaria='+quotedstr(FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Text));
      DM_tabelas.ZQsequencia_remessa.Open;
      if DM_tabelas.ZQsequencia_remessa.RecordCount>0 then
         DM_tabelas.ZQsequencia_remessa.edit
      else
          DM_tabelas.ZQsequencia_remessa.insert;
      DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt:=DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt+1;
      DM_tabelas.ZQsequencia_remessa.FieldByName('idconta_bancaria').AsLargeInt:=FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.value;
      DM_tabelas.ZQsequencia_remessa.Post;
      DM_tabelas.ZQsequencia_remessa.Refresh;
      varremes := DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt;//DM_Tabelas.ZQContaBancarian_sequencial.Value+1;
      DM_tabelas.ZQsequencia_remessa.close;


      Frm_NumRemessa.XNENumRemessa.Value := varremes;
      Frm_NumRemessa.Label2.Caption := VARARQ;
      Frm_NumRemessa.Top := FrmCobrancaBancaria.Top + 30;
      Frm_NumRemessa.Left := FrmCobrancaBancaria.Left + 300;
      AbrirModal(Self, Frm_NumRemessa);
      //numero da remessa

      varremes := strtoint(floattostr(Frm_NumRemessa.XNENumRemessa.value));
  end;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  if cb2via.Checked=false then
  begin
    DM_Tabelas.ZQBancRemes.Close;
    DM_Tabelas.ZQBancRemes.SQL.Clear;
    DM_Tabelas.ZQBancRemes.SQL.Add(' select idbanco_remessa,remessa,idbanco,gerado,nomearq,arquivo,digito_dif,remessa2 from banco_remessa ');
    DM_Tabelas.ZQBancRemes.open;

    DM_Tabelas.ZQBancRemes.Insert;
    DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger := FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value;
    DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger := varremes;
    DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := Vararqnome;
    DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := date;
    DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
    DM_Tabelas.ZQBancRemes.Post;
  end;

//  if not DirectoryExists(ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB') then
  //       ForceDirectories(ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB');
  ACBrBoleto1.DirArqRemessa := VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
  ACBrBoleto1.LayoutRemessa:=c400;
  case CBIdentOcorr.ItemIndex of
      1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
      2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
      3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
      4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
      5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
      6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
      7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
      8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
      9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
      10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
      11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
      12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
      13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
      14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
      15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
      16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
      17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
  end;
  ACBrBoleto1.Banco.TipoCobranca:=cobitau;
  ACBrBoleto1.Banco.TamanhoMaximoNossoNum:=8;
  if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
  begin
     ACBrBoleto1.Banco.LayoutVersaoArquivo:=40;
     ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=40;
  end
  else
  begin
     ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
     ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
  end;
  if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
  begin
     ACBrBoleto1.Banco.LayoutVersaoLote:=30;
     ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=30;
  end
  else
  begin
    ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
    ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
  end;


  ACBrBoletoFCFortes1.DirLogo:=ExtractFilePath( Application.ExeName)+'Logos_bco';
  FrmRelRecebimento.CDS_MarcaTit.First;
  Gauge1.Visible := True;
  Gauge1.MaxValue := FrmRelRecebimento.CDS_MarcaTit.RecordCount;


  ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

  carne:=0;

  if cb2via.Checked=false then
  begin

    if zqcarne.active=false then
       zqcarne.open;
    zqcarne.last;
    carne:=zqcarne.FieldByName('n_seq').AsInteger;
    inc(carne);
  end;


  //detalhe
  FrmRelRecebimento.CDS_MarcaTit.DisableControls;
  ACBrBoleto1.ListadeBoletos.Clear;

  ACBrBoleto1.Cedente.Agencia           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
  ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
  ACBrBoleto1.Cedente.AgenciaDigito     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
  ACBrBoleto1.Cedente.Conta             := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
  ACBrBoleto1.Cedente.ContaDigito       := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
//  if trim(DM_Tabelas.ZQContaBancariavariacao.Value)=Emptystr then
//     ACBrBoleto1.Cedente.Modalidade     :='01'
//  else
//     ACBrBoleto1.Cedente.Modalidade     := trim(DM_Tabelas.ZQContaBancariavariacao.Value);
//  ACBrBoleto1.Cedente.CodigoCedente     := trim(DM_Tabelas.ZQContaBancariaconvenio.Value);
//  ACBrBoleto1.Cedente.CodigoTransmissao := '';

  case ComboBox3.ItemIndex of
       0: ACBrBoleto1.ACBrBoletoFC.LayOut := TACBrBolLayOut(lPadrao);
       1: ACBrBoleto1.ACBrBoletoFC.LayOut := TACBrBolLayOut(lCarne);
  end;

//          if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
  if length(charrem(trim(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString)))=14 then
  begin
      ACBrBoleto1.Cedente.CNPJCPF       := trim(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString));
      ACBrBoleto1.Cedente.TipoInscricao := pJuridica;
  end
  else
  begin
      ACBrBoleto1.Cedente.CNPJCPF       := trim(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString));
      ACBrBoleto1.Cedente.TipoInscricao := pFisica
  end;
  ACBrBoleto1.Cedente.Nome             := trim(FrmRelRecebimento.CDSEmpreendimentonometitular.Value);
//  ACBrBoleto1.Cedente.Logradouro       := trim(DM_Tabelas.zqempresaendereco.Value);
//  tam:=pos(',',DM_Tabelas.zqempresaendereco.Value);
//          inc(tam);
//        if tam>1 then
//         ACBrBoleto1.Cedente.NumeroRes     := SoNumero(Trim(DM_Tabelas.zqempresaendereco.Value))
//    else
//     ACBrBoleto1.Cedente.NumeroRes     := 'SN';
  try
    zqaux:= tzquery.Create(nil);
    zqaux.Connection := DM_Tabelas.zconeccao;
    zqaux.Close;
    zqaux.sql.clear;
    zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
    zqaux.Open;
    if zqaux.RecordCount>0 then
    begin
        ACBrBoleto1.Cedente.Logradouro        := RemoveAcentos(uppercase(trim(zqaux.FieldByName('ende_cob').AsString)));
      //  ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
        ACBrBoleto1.Cedente.Complemento       := '';
        ACBrBoleto1.Cedente.Bairro            := RemoveAcentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
        ACBrBoleto1.Cedente.Cidade            := RemoveAcentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
        ACBrBoleto1.Cedente.UF                := RemoveAcentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
        ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
    end;
  finally
    zqaux.Close;
    zqaux.free;
  end;

  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

  NN:=DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;

  while not FrmRelRecebimento.CDS_MarcaTit.Eof do
  begin

   if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
   begin
    if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '1' Then
    else
    if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
      mensagem('A parcela '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente')
    else Begin

        DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
        Titulo := ACBrBoleto1.CriarTituloNaLista;
        with Titulo do
        begin
                {TACBrTipoCobranca = (cobNenhum, cobBancoDoBrasil, cobSantander, cobCaixaEconomica, cobCaixaSicob, cobBradesco, cobItau, cobBancoMercantil, cobSicred,
                                      cobBancoob, cobBanrisul, cobBanestes, cobHSBC, cobBancoDoNordeste, cobBRB, cobBicBanco, cobBradescoSICOOB, cobBancoSafra,
                                      cobSafraBradesco, cobBancoCECRED );}

          EspecieDoc        := 'DM';
          Vencimento        := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
          DataDocumento     := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;

          eita:=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
          if eita=1 then
             codcli:=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
          if eita=2 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
          if eita=3 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
          if eita=4 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
          if eita>=5 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
          eita:=codcli;
          codcli:=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
          vardoc := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);

          NumeroDocumento   := vardoc;
          SeuNumero         := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
          EspecieMod        := '$';
          Aceite            := atNao; //atSim
          DataProcessamento := Now;
          Carteira          := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);
          if (rgnosso.Checked) and (CBbcoGera.Checked=false) and (cb2via.Checked=false) then
          begin
            inc(NN);
            varnossnum:=inttostr(NN);
          end;
          if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
          begin
           if (CBbcoGera.Checked=false) and (cb2via.Checked) then
           begin
              varnossnum :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;
              NN := varnossnum.ToInteger;
           end;
          end;

         if cb2via.Checked=false then
         begin

          ZQRemesRec.Filtered:=false;
          ZQRemesRec.Filter:='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
          ZQRemesRec.Filtered:=true;
          if ZQRemesRec.RecordCount>0 then
          begin
            ZQRemesRec.First;
            while not ZQRemesRec.Eof do
            begin
              ZQRemesRec.delete;
            end;
            ZQRemesRec.Filtered:=false;
            ZQRemesRec.Insert;
            ZQRemesRec.FieldByName('remessa').AsInteger     := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
            ZQRemesRec.FieldByName('idrec').AsInteger       := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
            ZQRemesRec.FieldByName('nossonumero').AsString := varnossnum;
            ZQRemesRec.Post;
          end
          else
          begin
            ZQRemesRec.Filtered:=false;
            ZQRemesRec.Insert;
            ZQRemesRec.FieldByName('remessa').AsInteger     := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
            ZQRemesRec.FieldByName('idrec').AsInteger       := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
            ZQRemesRec.FieldByName('nossonumero').AsString := varnossnum;
            ZQRemesRec.Post;
          end;
          ZQRemesRec.Filtered:=false;

         end;


         if cb2via.Checked=false then
         begin

            Remessa_temp.Filtered:=false;
            Remessa_temp.Filter:='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
            Remessa_temp.Filtered:=true;
            if Remessa_temp.RecordCount>0 then
            begin
              Remessa_temp.First;
              while not Remessa_temp.Eof do
              begin
                Remessa_temp.delete;
              end;
              Remessa_temp.Filtered:=false;
              Remessa_temp.Insert;
              Remessa_tempremessa.Value     := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
              Remessa_tempidrec.Value       := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
              Remessa_tempnossonumero.value := varnossnum;
              Remessa_temp.Post;
            end
            else
            begin
              Remessa_temp.Filtered:=false;
              Remessa_temp.Insert;
              Remessa_tempremessa.Value     := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
              Remessa_tempidrec.Value       := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
              Remessa_tempnossonumero.value := varnossnum;
              Remessa_temp.Post;
            end;
            Remessa_temp.Filtered:=false;

         end;

          VarDif :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;

          NossoNumero       := varnossnum;
          ValorDocumento    := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);



          if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
          begin
             vpar           :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
             inc(vpar);
             TotalParcelas  :=  strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
          end
          else
             TotalParcelas  :=  1;

//          TotalParcelas     := 36000;


          if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
          begin
             vpar           :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
             dec(vpar);
             parcela        :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
          end
          else
             parcela        :=  1;


        //dados do sacado
          Sacado.NomeSacado:=Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
         if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
          begin
              Sacado.CNPJCPF:= trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
              Sacado.Pessoa := pJuridica;
          end
          else
          begin
              Sacado.CNPJCPF:= trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
              Sacado.Pessoa := pFisica;
          end;
          Sacado.Logradouro := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
//          Sacado.Numero     := SoNumero(Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value));
          Sacado.Bairro     := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);
          Sacado.Cidade     := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
          Sacado.UF         := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
          Sacado.CEP        := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
          ValorAbatimento   := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
          LocalPagamento    := 'PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO';

          CodigoMoraJuros   := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
          CodigoMora        := '1'; //1-por dia 2-mensal 3-isento
          CodigoMulta       := cmPercentual;//(cmValorFixo, cmPercentual)
          DataMulta         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

          if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
             PercentualMulta   := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
          else
             PercentualMulta   := 0;

      //            MultaValorFixo    := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
          MultaValorFixo    := False;

         if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
            ValorMoraJuros    := DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat
         else
            ValorMoraJuros    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);


         DataMoraJuros     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

//            DataProtesto      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;
         DataProtesto      := 0;


         if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
         begin
            {  TACBrTipoDesconto = (
            tdNaoConcederDesconto,
            tdValorFixoAteDataInformada,
            tdPercentualAteDataInformada,
            tdValorAntecipacaoDiaCorrido,
            tdValorAntecipacaoDiaUtil,
            tdPercentualSobreValorNominalDiaCorrido,
            tdPercentualSobreValorNominalDiaUtil,
            tdCancelamentoDesconto);}


           CodigoDesconto    := cdValorFixo; //(cdSemDesconto, cdValorFixo);
           TipoDesconto      := tdValorFixoAteDataInformada;
           DataDesconto      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
           ValorDesconto     := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);

         end;

         if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
         begin
              if empty(VarDif) then // coluna 38 até 62
              begin
                Mensagem.Text     := 'SR. CAIXA, NAO RECEBER APOS 30 (TRINTA) DIAS DE VENCIMENTO. N.Tit.'+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+'  N.Carnê:'+strzero(carne,5);
              end
              else
              begin
                Mensagem.Text     := 'SR. CAIXA, NAO RECEBER APOS 30 (TRINTA) DIAS DE VENCIMENTO. N.Tit.'+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+' N.Carnê:'+strzero(carne,5)+'-'+Vardif;
              end;
         end
         else
         begin
              inc(carne);
              ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
              if empty(VarDif) then // coluna 38 até 62
              begin
                Mensagem.Text     := 'SR. CAIXA, NAO RECEBER APOS 30 (TRINTA) DIAS DE VENCIMENTO. N.Tit.'+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+' N.Carnê:'+strzero(carne,5);
              end
              else
              begin
                Mensagem.Text     := 'SR. CAIXA, NAO RECEBER APOS 30 (TRINTA) DIAS DE VENCIMENTO. N.Tit.'+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+' N.Carnê:'+strzero(carne,5)+'-'+Vardif;
              end;
         end;
//         if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2)>0 then
//         begin
//            Mensagem.Text := Mensagem.Text +#13+' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'));
//         end;
            case CBIdentOcorr.ItemIndex of
                1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
            end;
        end;
    end;
    FrmRelRecebimento.CDS_MarcaTit.Next;
   end
   else
   begin
     FrmRelRecebimento.CDS_MarcaTit.Next;
   end;
  end;
  try
    if cb2via.Checked=false then
       ACBrBoleto1.GerarRemessa( varremes );
    if CBBancoemite.State=cbUnChecked then
    begin
       ACBrBoleto1.Imprimir;
 //      tam:=pos('.',vardir+'\'+Vararqnome);
 //      dec(tam);
 //      ACBrBoletoFCFortes1.NomeArquivo:=copy(vardir+'\'+Vararqnome,1,tam)+'.pdf';
 //      ACBrBoleto1.GerarPDF;


    end;
  finally
    FrmRelRecebimento.CDS_MarcaTit.EnableControls;
    if cb2via.Checked=false then
    begin
      zqcarne.Edit;
      zqcarne.FieldByName('n_seq').AsInteger:=carne;
      zqcarne.post;
    end;
    zqcarne.Close;
    Gauge1.Visible := False;
    if cb2via.Checked=false then
    begin
      Gauge1.Visible := True;
      Gauge1.Progress:=0;
      Remessa_temp.First;
      Gauge1.MaxValue := Remessa_temp.RecordCount;
      Remessa_temp.DisableControls;
      while not Remessa_temp.Eof do
      begin
        Gauge1.Progress := Remessa_temp.RecNo;

        DM_Tabelas.ZQRemes_Receb_atualiza.Close;
        DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
        DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(Remessa_tempnossonumero.Value)+') where idrecebimento='+quotedstr(Remessa_tempidrec.Text));
        DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;
        Remessa_temp.Next;
      end;
    end;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRecebimento.refresh;

    DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
    if (cb2via.Checked=false) and (rgnosso.Checked) then
    begin

      DM_Tabelas.ZQContaBancaria.Edit;
      DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt:=NN;
      DM_Tabelas.ZQContaBancaria.Post;
    end;
    if cb2via.Checked=false then
       Remessa_temp.EnableControls;
    Remessa_temp.Close;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    Gauge1.Progress:=0;
    Gauge1.Visible := false;
    DM_Tabelas.ZQContaBancaria.refresh;
    DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
    if cb2via.Checked=false then
       mensagem('Relatório de Boletos gerado com sucesso!!!'+chr(13)+'O relatório foi gravado em '+VARARQ);
  end;

End;


procedure TFrmCobrancaBancaria.MaskEdit1Change(Sender: TObject);
begin
  DM_Tabelas.ZQEmpresa.Filtered:=false;
  DM_Tabelas.ZQEmpresa.Close;
  DM_Tabelas.ZQEmpresa.SQL.Clear;
  DM_Tabelas.ZQEmpresa.SQL.Add('select * from empresa where ativa = '+quotedstr('S')+' and razao like '+quotedstr('%'+MaskEdit1.Text+'%')+' or fantasia like '+quotedstr('%'+MaskEdit1.Text+'%')+' order by razao');
  DM_Tabelas.ZQEmpresa.Open;
//  dbgrid1.SetFocus;

end;

procedure TFrmCobrancaBancaria.MaskEdit1Exit(Sender: TObject);
begin
  Dbgrid1.SetFocus;
end;

procedure TFrmCobrancaBancaria.bradesco_400;
Var
  F : TextFile;
  ArqIni : tIniFile;
  VARARQ, Vararqnome, Varmens1, Varmens2, Varmens3, Varmens4, varnossonumero,
  varnossnum, varmora, vardoc, varcpfcnpj, varendereco, varhoje, VarNomepasta,
  hora,varhostname, vardir, varpastabanco,VarPath, ext,VarDif,ql,codnosso,pega : string;
  tmnosso,num,numeroremessa,codcli,eita,carne,tam,tam2,varregistro, varremes, varx, vary,ps : integer;
  Titulo : TACBrTitulo;
  I : TACBrBolLayOut;

begin
  carne:=0;
  codcli:=0;
  tam:=0;
  tmnosso:=0;
  VarDif:='';
  ql:='';
  hora:='';
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  VarNomepasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );

//  varpastabanco := FrmRelRecebimento.CDSEmpreendimentoapelido.Value+'_'+FrmRelRecebimento.CDSEmpreendimentonomebanco.Value+FrmRelRecebimento.CDSEmpreendimentoconta.Value;

  vardir := '\\'+varhostname+'\'+VarNomePasta;
  if cb2via.Checked=false then
  begin
    // aqui gera a remessa
    if copy(CBIdentOcorr.Text,1,2)<>'00' then
    begin
     SaveDialog1.Execute;
      tam2:=length(SaveDialog1.FileName);
      tam2:=tam2-11;
      VarPath := copy(SaveDialog1.FileName,tam2+1,11);
      if trim(VarPath)<>'remessa.rem' then
      begin
        mensagem('O nome do Arquivo tem que ser remessa.rem');
        exit;
      end;

      VarPath := copy(SaveDialog1.FileName,1,tam2);

      vardir := VarPath;

      if empty(trim(VarPath)) then
      begin
        mensagem('O nome do Arquivo não pode ficar em branco, tem que ser remessa.rem');
        exit;
      end;


      DBGerar.Enabled:=false;

      vardir := vardir +'\'+copy(datetostr(date),1,2);
      CreateDir(vardir);
      Ext := '.REM';
      if CBTipoArq.Checked Then
        Ext := '.TST';

      varregistro:=2;

      for  varremes:= 1 to 99 do begin
        vararq := inttostrZero(varremes,2);
        Vararqnome := 'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+Ext;
        VARARQ :=vardir+'\'+Vararqnome;
        If not(fileexists(VARARQ)) then
          Break;
      end;
      Varnomearq := VARARQ;


      if DM_Tabelas.ZQBol_men.State in [DsEdit, DsInsert] Then
        DM_Tabelas.ZQBol_men.Edit;

      //pegar o ultimo registro auto incremento criado
      ZQNumero_remessa.Close;
      ZQNumero_remessa.SQL.Clear;
      ZQNumero_remessa.SQL.Add('Select idnumero_remessa from numero_remessa');
      ZQNumero_remessa.open;
      ZQNumero_remessa.last;
      varremes := ZQNumero_remessa.FieldByName('idnumero_remessa').AsInteger+1;


      ZQNumero_remessa.Close;
      ZQNumero_remessa.SQL.Clear;
      ZQNumero_remessa.SQL.Add('INSERT INTO numero_remessa (idnumero_remessa) VALUES ( '+quotedstr(inttostr(varremes))+')');
      ZQNumero_remessa.ExecSQL;

      //pegar o ultimo registro auto incremento criado
      ZQNumero_remessa.Close;
      ZQNumero_remessa.SQL.Clear;
      ZQNumero_remessa.SQL.Add('Select idnumero_remessa from numero_remessa');
      ZQNumero_remessa.open;
      ZQNumero_remessa.last;
      varremes := ZQNumero_remessa.FieldByName('idnumero_remessa').AsInteger;
      ZQNumero_remessa.Close;
      DM_Tabelas.ZQBancRemes.Last;
  //    varremes := DM_Tabelas.ZQBancRemesremessa.Value + 1;

      Frm_NumRemessa.XNENumRemessa.Value := varremes;
      Frm_NumRemessa.Label2.Caption      := VARARQ;
      Frm_NumRemessa.Top                 := FrmCobrancaBancaria.Top + 30;
      Frm_NumRemessa.Left                := FrmCobrancaBancaria.Left + 300;
      AbrirModal(Self, Frm_NumRemessa);
      //numero da remessa

      varremes := strtoint(floattostr(Frm_NumRemessa.XNENumRemessa.value));
  //  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria;n_dif_empreed',VarArrayOf([FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value]), []);
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

      DM_Tabelas.ZQBancRemes.Insert;
      DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger    := FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value;
      DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger    := varremes;
      DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString    := Vararqnome;
      DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime     := date;
      DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
      DM_Tabelas.ZQBancRemes.Post;

      varhoje := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);


      ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

      if zqcarne.active=false then
         zqcarne.open;
      zqcarne.last;
      carne:=zqcarne.FieldByName('n_seq').AsInteger;
      inc(carne);

      FrmRelRecebimento.CDS_MarcaTit.DisableControls;

      Gauge2.Progress:=0;

      FrmRelRecebimento.CDS_MarcaTit.DisableControls;

      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible := True;
      Gauge1.MaxValue := FrmRelRecebimento.CDS_MarcaTit.RecordCount;


      ACBrBoleto1.ListadeBoletos.Clear;
      ACBrBoleto1.DirArqRemessa:=vardir;
      ACBrBoleto1.DirArqRetorno:=ExtractFilePath( Application.ExeName ) +'Retornos';
      ACBrBoleto1.ACBrBoletoFC.DirLogo:=ExtractFilePath( Application.ExeName ) +'Logos_bco';
      ACBrBoleto1.ACBrBoletoFC.NomeArquivo:=vardir+'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+'\'+Vararqnome+'.pdf';

      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.Banco.TipoCobranca          := cobBradesco;
      ACBrBoleto1.Banco.Digito                := 2;
      ACBrBoleto1.LayoutRemessa               := c400;
      ACBrBoleto1.Banco.Nome                  := trim(DM_Tabelas.ZQContaBancaria.FieldByName('nomebanco').AsString);
      ACBrBoleto1.Banco.Numero                := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString).ToInteger;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum :=11;
      case ComboBox3.ItemIndex of
         0: ACBrBoleto1.ACBrBoletoFC.LayOut := TACBrBolLayOut(lPadrao);
         1: ACBrBoleto1.ACBrBoletoFC.LayOut := TACBrBolLayOut(lCarne);
      end;

      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=84;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=84;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=42;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=42;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;

      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin

       if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
       begin
        if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '1' Then
        else
         if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            mensagem('A parcela '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente')
        else
        Begin
          if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
            varmora := inttostrZero(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,13)
          else
            varmora := inttostrZero(strtofloat(charrem(alltrim(transform(ExRound((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2),'###,##0.00')))),13);
          varendereco := removeacentos(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
          inc(varregistro);
          Gauge1.Progress := FrmRelRecebimento.CDS_MarcaTit.RecNo;
          eita:=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
          if eita=1 then
             codcli:=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
          if eita=2 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
          if eita=3 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
          if eita=4 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
          if eita>=5 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
          eita:=codcli;

          hora:=timetostr(time);
          hora:=copy(hora,1,2)+copy(hora,4,2)+copy(hora,7,2);

          if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=true) then
          begin
            randomize;
            num := Round(random(999)); // 999 é o limite do numero randomico
            codcli:=num+strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
            codnosso:= copy(trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.text+inttostr(num+eita)),1,11);

          end
          else
          begin
          //mudei aqui 09/06/2021 para ver se para de dar nosso numero duplicado
            randomize;
            num := Round(random(999)); // 999 é o limite do numero randomico
            codcli:=num+strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
            codnosso:= copy(trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.text+inttostr(num+eita)),1,11);
          end;

            //codcli:=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
       //   codcli:=strtoint(inttostr(eita)+hora);
          vardoc := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);

   //       varnossonumero := DM_Tabelas.ZQContaBancariacarteira.Value+inttostrZero(codcli,11);
          tmnosso:=length(codnosso);
          tmnosso:=11-tmnosso;
          if tmnosso<=0 then
             tmnosso:=0;
          if tmnosso>11 then
             tmnosso:=11;
          varnossonumero := DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString+repl('0',tmnosso)+codnosso;
           // qdo precisar trocar o numero do nosso numero //comenta as 2 linhs anteriores e liberar as 2 seguinte
  //      vardoc := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);
  //      varnossonumero := DM_Tabelas.ZQContaBancariacarteira.Value+inttostrZero( FrmRelRecebimento.CDS_MarcaTitIdreceb.Value,11);

          varnossnum:='32765432765432';
          vary := 0;
          for varx := 1 to 14 do
            vary := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(varnossonumero,varx,1)));
          vary := (vary mod 11);
          vary := 11 - vary ;
          if vary = 10 Then
            varnossnum := repl('0',tmnosso)+codnosso+'P'
          else if vary = 11
           Then
            varnossnum := repl('0',tmnosso)+codnosso+'0'
          else
            varnossnum := repl('0',tmnosso)+codnosso+inttostr(vary);

          // 14/12/2012
          if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
             varnossnum :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;

          if FrmRelRecebimento.CDS_MarcaTitnumboleto.Value='' then
          begin
            FrmRelRecebimento.CDS_MarcaTit.edit;
            FrmRelRecebimento.CDS_MarcaTitnumboleto.Value:=varnossnum;
            FrmRelRecebimento.CDS_MarcaTit.post;
          end
          else
          begin
            if rgnosso.Checked then
            begin
              FrmRelRecebimento.CDS_MarcaTit.edit;
              FrmRelRecebimento.CDS_MarcaTitnumboleto.Value:=varnossnum;
              FrmRelRecebimento.CDS_MarcaTit.post;
            end;

          end;
          codcli:=0;
          eita:=0;
          ZQRemesRec.close;
          ZQRemesRec.SQL.Clear;
          ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
          ZQRemesRec.open;
  {        ZQRemesRec.Filtered:=false;
          ZQRemesRec.Filter:='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
          ZQRemesRec.Filtered:=true;}
          if ZQRemesRec.RecordCount>0 then
          begin
            ZQRemesRec.First;
            while not ZQRemesRec.Eof do
            begin
              ZQRemesRec.delete;
            end;
            ZQRemesRec.Filtered:=false;
            ZQRemesRec.Insert;
            ZQRemesRec.FieldByName('remessa').AsInteger := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
            ZQRemesRec.FieldByName('idrec').AsInteger := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
            ZQRemesRec.FieldByName('nossonumero').AsString :=varnossnum;
            ZQRemesRec.Post;
          end
          else
          begin
            ZQRemesRec.Filtered:=false;
            ZQRemesRec.Insert;
            ZQRemesRec.FieldByName('remessa').AsInteger := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
            ZQRemesRec.FieldByName('idrec').AsInteger := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
            ZQRemesRec.FieldByName('nossonumero').AsString :=varnossnum;
            ZQRemesRec.Post;
          end;
          ZQRemesRec.Filtered:=false;
          VarDif :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


          // 17/09/2013

          numeroremessa:=ZQRemesRec.FieldByName('remessa').AsInteger;


          Titulo := ACBrBoleto1.CriarTituloNaLista;
          with Titulo do
          begin
            Gauge1.MaxValue   := FrmRelRecebimento.CDS_MarcaTit.RecNo;
            EspecieDoc        := '01';
            Vencimento        := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.ValUE;
            DataLimitePagto   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.ValUE+30;

            tam:=length(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text);
            tam:=37+tam;
            tam:=49-tam;

            if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
            begin
              if empty(VarDif) then // coluna 38 até 62
                 SeuNumero         := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,25)  // uso da empresa, nº de controle da empresa;
              else
                 SeuNumero         := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,24)+VarDif;  // uso da empresa, nº de controle da empresa;
            end
            else
            begin
              inc(carne);
              ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
              if empty(VarDif) then // coluna 38 até 62
                  SeuNumero         := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,25) // uso da empresa, nº de controle da empresa;
              else
                  SeuNumero         :=copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(7),1,24)+VarDif;  // uso da empresa, nº de controle da empresa;
            end;
            VarDif:='';


            NumeroDocumento   := inttostrZero(FrmRelRecebimento.CDS_MarcaTitidreceb.Value,10);//copy(DM_Comercio.CDSMarcaTitordem.Value+space(10),1,10);
            DataDocumento     := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.VALUE;
            EspecieMod        := 'R$';
            Aceite            := atNao; //atSim
            DataProcessamento := Now;
            Carteira          := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);
            if CBCarteiraEnvio.Checked then
               CarteiraEnvio     := tceBanco   // vai retornar '1'   tceCedente; // vai retornar '2'    coluna 93
            else
              CarteiraEnvio     := tceCedente;       // coluna 93 cnab400

            ValorDocumento    := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);

            //-------------------------------------------------------
            ps:=pos('-',FrmRelRecebimento.CDS_MarcaTitordem.Value);
            inc(ps);
            pega:=copy(Trim(FrmRelRecebimento.CDS_MarcaTitordem.Value),ps,10);
            ps:=pos('/',pega);
            inc(ps);
            pega:=trim(copy(pega,ps,3));
            if not empty(pega) then
               ps:=pega.ToInteger
            else
              ps:=1;

            TotalParcelas     := ps;
            ps:=pos('-',FrmRelRecebimento.CDS_MarcaTitordem.Value);
            inc(ps);
            pega:=copy(Trim(FrmRelRecebimento.CDS_MarcaTitordem.Value),ps,10);
            ps:=pos('/',pega);
            dec(ps);
            pega:=trim(copy(pega,1,ps));
            if not empty(pega) then
               ps:=pega.ToInteger
            else
              ps:=1;

            parcela           := ps;

            //Dados do Cedente

            ACBrBoleto.Cedente.Agencia           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
            ACBrBoleto.Cedente.AgenciaDigito     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
            ACBrBoleto.Cedente.Conta             := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
            ACBrBoleto.Cedente.ContaDigito       := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
            ACBrBoleto.Cedente.CodigoTransmissao := trim(DM_Tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString); //trim(DM_comercio.ZQConta_BancariaCodigo_cedente.Value);
            if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
               ACBrBoleto.Cedente.Modalidade     :='01'
            else
               ACBrBoleto.Cedente.Modalidade     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);
            ACBrBoleto.Cedente.CodigoCedente     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_no_banco').AsString);
            ACBrBoleto.Cedente.Convenio          := trim(DM_Tabelas.ZQContaBancaria.FieldByName('convenio').AsString);


            if CBBancoemite.checked then
               ACBrBoleto.Cedente.ResponEmissao := tbBancoEmite
            else
              ACBrBoleto.Cedente.ResponEmissao := tbCliEmite;

            ACBrBoleto.Cedente.CaracTitulo                      := tcSimplesRapComReg; //(tcSimples, tcVinculada, tcCaucionada, tcDescontada, tcVendor, tcDireta, tcSimplesRapComReg, tcCaucionadaRapComReg, tcDiretaEspecial)
            ACBrBoleto.Cedente.TipoCarteira                     := tctRegistrada;//tctSimples

            if length(tiramascara(trim(DM_Tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString)))<14 then
            begin
               ACBrBoleto.Cedente.CNPJCPF :=tiramascara(trim(DM_Tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString));
               ACBrBoleto.Cedente.TipoInscricao := pFisica;
            end
            else
            begin
               ACBrBoleto.Cedente.CNPJCPF := copy(tiramascara(trim(DM_Tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString)),1,14);
               ACBrBoleto.Cedente.TipoInscricao := pJuridica;
            end;
            ACBrBoleto.Cedente.Nome              := charrem(trim(DM_Tabelas.ZQContaBancaria.FieldByName('nometitular').AsString));
            try
              zqaux:= tzquery.Create(nil);
              zqaux.Connection := DM_Tabelas.zconeccao;
              zqaux.Close;
              zqaux.sql.clear;
              zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
              zqaux.Open;
              if zqaux.RecordCount>0 then
              begin
                  ACBrBoleto.Cedente.Logradouro        := removeacentos(uppercase(trim(copy(zqaux.FieldByName('ende_cob').AsString,1,pos(',',zqaux.FieldByName('ende_cob').AsString)-1))));
                  ACBrBoleto.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                  ACBrBoleto.Cedente.Complemento       := '';
                  ACBrBoleto.Cedente.Bairro            := Removeacentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                  ACBrBoleto.Cedente.Cidade            := Removeacentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                  ACBrBoleto.Cedente.UF                := Removeacentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                  ACBrBoleto.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
              end;
            finally
              zqaux.Close;
              zqaux.free;
            end;

            if (CBBancoemite.Checked) then
              ACBrBoleto.Cedente.ResponEmissao:=tbBancoEmite     // 2 Cliente Emite - 1 Banco Emite
            else
              ACBrBoleto.Cedente.ResponEmissao:=tbCliEmite;


            //dados do sacado
            Sacado.NomeSacado   := charrem(Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value));
            if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
            begin
              Sacado.CNPJCPF  := Charrem(trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)));
              Sacado.Pessoa   := pJuridica;
            end
            else
            begin
              Sacado.CNPJCPF  := Charrem(trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)));
              Sacado.Pessoa   := pFisica;
            end;
            Sacado.Logradouro := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
            Sacado.CEP        := charrem(Trim(Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value))));
            Sacado.Cidade     := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
            Sacado.UF         := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
            Sacado.Bairro     := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);
//            if CBParcelas.Checked=true then
//               Sacado.Avalista:=COPY('N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+espacos(60),1,60)  // decomposição
//            else
//               Sacado.Avalista:=COPY('Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+espacos(60),1,60);  // decomposição


          //dados da cobranca

            codcli:=0;
            eita:=0;
//
            NossoNumero                  := copy(varnossnum,1,11);
            ValorDocumento               := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
            DataDocumento                := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
            Vencimento                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;
            NumeroDocumento              := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);
            if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
            begin
           //  RLBTitulo1.TipoOcorrencia := toRemessaConcederDesconto;
              DataDesconto               := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
              ValorDesconto              := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
            end;
            ValorAbatimento   := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
            LocalPagamento := 'PAGÁVEL PREFERENCIALMENTE NA REDE BRADESCO OU BRADESCO EXPRESSO';//trim(DM_comercio.ZQConta_BancariaInstrucoes.Value);

             DataProtesto   := 0; //FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
            // dias de protestos e carteira de envio
            TipoDiasProtesto := diCorridos;

            if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                 PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
            else
                 PercentualMulta                                 := 0;
            CodigoMora        :='1'; // 1 = Valor 2 = Percentual 9 = Dispensar Multa
            CodigoMulta       := cmPercentual;//(cmValorFixo, cmPercentual)
            DataMulta         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;
   //       MultaValorFixo    := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
            MultaValorFixo    := False;
            if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
               ValorMoraJuros            := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
            else
              ValorMoraJuros             := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
            DataMoraJuros                := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
            CodigoMoraJuros              := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
            CodigoMora                   := '1'; //1-por dia 2-mensal 3-isento
            CodigoMulta                  := cmPercentual;//(cmValorFixo, cmPercentual)

//           if CBParcelas.Checked=true then
//               mensagem.text := 'COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+'  Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
//           else
//               mensagem.text := 'COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

            if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2) > 0 then
            begin
              if CBParcelas.Checked=false then
              begin
               // mensagem.Add('NAO RECEBER APOS 30 DIAS DO VENCIMENTO.');
                if DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat>0 then
                   mensagem.Add(' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2),'###,###,##0.00')));
                mensagem.Add(' N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);
              end
              else
              begin
                if DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat>0 then
                   mensagem.Add(' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2),'###,###,##0.00')));
                mensagem.Add(' N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);
              end;
            end
            else
            begin
              if CBParcelas.Checked=false then
              begin
          //      mensagem.Add('NAO RECEBER APOS 30 DIAS DO VENCIMENTO.');
                if DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat>0 then
                   mensagem.Add(' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2),'###,###,##0.00')));
                mensagem.Add(' N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);
              end
              else
              begin
                if DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat>0 then
                   mensagem.Add(' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2),'###,###,##0.00')));
                mensagem.Add(' N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);
              end;
            end;
            case CBIdentOcorr.ItemIndex of
                1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
            end;
          end;

        end;
        FrmRelRecebimento.CDS_MarcaTit.Next;
       end
       else
       begin
         FrmRelRecebimento.CDS_MarcaTit.Next;
       end;
      end;
      FrmRelRecebimento.CDS_MarcaTit.EnableControls;
      //gauge2.Progress:=0;
      Gauge2.Visible:=false;
      FrmRelRecebimento.CDS_MarcaTit.EnableControls;
      ACBrBoleto1.GerarRemessa( varremes );
      zqcarne.Edit;
      zqcarne.FieldByName('n_seq').AsInteger:=carne;
      zqcarne.post;
      zqcarne.Close;

      inc(varregistro);
      mensagem('Remessa de Boletos gerado com sucesso!!!'+chr(13)+'A remessa foi gravado em '+VARARQ);
    end;
  end
  else
  begin
      varhoje := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);

      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible                                               := True;
      Gauge1.Progress                                              :=0;
      Gauge1.MaxValue                                              := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

      ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;



      FrmRelRecebimento.CDS_MarcaTit.DisableControls;

      Gauge2.Progress:=0;

      FrmRelRecebimento.CDS_MarcaTit.DisableControls;

      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible := True;
      Gauge1.MaxValue := FrmRelRecebimento.CDS_MarcaTit.RecordCount;



      ACBrBoleto1.ListadeBoletos.Clear;
      ACBrBoleto1.DirArqRemessa:=VARARQ;
      ACBrBoleto1.DirArqRetorno:=ExtractFilePath( Application.ExeName ) +'Retornos';
      ACBrBoleto1.ACBrBoletoFC.DirLogo:=ExtractFilePath( Application.ExeName ) +'Logos_bco';
      ACBrBoleto1.ACBrBoletoFC.NomeArquivo:=VARARQ+'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+'.pdf';
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;


      ACBrBoleto1.Banco.TipoCobranca          := cobBradesco;
      ACBrBoleto1.Banco.Digito                := 2;
      ACBrBoleto1.LayoutRemessa               := c400;
      ACBrBoleto1.Banco.Nome                  := trim(DM_Tabelas.ZQContaBancaria.FieldByName('nomebanco').AsString);
      ACBrBoleto1.Banco.Numero                := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString).ToInteger;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum :=11;
      case ComboBox3.ItemIndex of
         0: ACBrBoleto1.ACBrBoletoFC.LayOut := TACBrBolLayOut(lPadrao);
         1: ACBrBoleto1.ACBrBoletoFC.LayOut := TACBrBolLayOut(lCarne);
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=84;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=84;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=42;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=42;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;
      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin

       if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
       begin
        if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '1' Then
        else
         if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            mensagem('A parcela '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente')
        else
        Begin
          if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
            varmora := inttostrZero(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,13)
          else
            varmora := inttostrZero(strtofloat(charrem(alltrim(transform(ExRound((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2),'###,##0.00')))),13);
          varendereco := removeacentos(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
          inc(varregistro);
          Gauge1.Progress := FrmRelRecebimento.CDS_MarcaTit.RecNo;
          eita:=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
          if eita=1 then
             codcli:=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
          if eita=2 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
          if eita=3 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
          if eita=4 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
          if eita>=5 then
             codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
          eita:=codcli;

          hora:=timetostr(time);
          hora:=copy(hora,1,2)+copy(hora,4,2)+copy(hora,7,2);
          eita                                               :=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
          if eita=1 then
             codcli                                          :=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
          if eita=2 then
             codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
          if eita=3 then
             codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
          if eita=4 then
             codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
          if eita>=5 then
             codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
          eita                                               :=codcli;
          codcli                                             :=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);

          vardoc := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);

          varnossnum :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;

          if FrmRelRecebimento.CDS_MarcaTitnumboleto.Value='' then
          begin
            FrmRelRecebimento.CDS_MarcaTit.edit;
            FrmRelRecebimento.CDS_MarcaTitnumboleto.Value:=varnossnum;
            FrmRelRecebimento.CDS_MarcaTit.post;
          end
          else
          begin
            if rgnosso.Checked then
            begin
              FrmRelRecebimento.CDS_MarcaTit.edit;
              FrmRelRecebimento.CDS_MarcaTitnumboleto.Value:=varnossnum;
              FrmRelRecebimento.CDS_MarcaTit.post;
            end;

          end;
          codcli:=0;
          eita:=0;
          ZQRemesRec.close;
          ZQRemesRec.SQL.Clear;
          ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
          ZQRemesRec.open;
  {        ZQRemesRec.Filtered:=false;
          ZQRemesRec.Filter:='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
          ZQRemesRec.Filtered:=true;}
          if ZQRemesRec.RecordCount>0 then
          begin
            ZQRemesRec.First;
            while not ZQRemesRec.Eof do
            begin
              ZQRemesRec.delete;
            end;
            ZQRemesRec.Filtered:=false;
            ZQRemesRec.Insert;
            ZQRemesRec.FieldByName('remessa').AsInteger := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
            ZQRemesRec.FieldByName('idrec').AsInteger := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
            ZQRemesRec.FieldByName('nossonumero').AsString :=varnossnum;
            ZQRemesRec.Post;
          end
          else
          begin
            ZQRemesRec.Filtered:=false;
            ZQRemesRec.Insert;
            ZQRemesRec.FieldByName('remessa').AsInteger := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
            ZQRemesRec.FieldByName('idrec').AsInteger := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
            ZQRemesRec.FieldByName('nossonumero').AsString :=varnossnum;
            ZQRemesRec.Post;
          end;
          ZQRemesRec.Filtered:=false;
          VarDif :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


          // 17/09/2013

          numeroremessa:=ZQRemesRec.FieldByName('remessa').AsInteger;


          Titulo := ACBrBoleto1.CriarTituloNaLista;
          with Titulo do
          begin
            Gauge1.MaxValue   := FrmRelRecebimento.CDS_MarcaTit.RecNo;
            EspecieDoc        := '01';
            Vencimento        := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.ValUE;
            DataLimitePagto   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.ValUE+30;

            tam:=length(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text);
            tam:=37+tam;
            tam:=49-tam;

            if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
            begin
              if empty(VarDif) then // coluna 38 até 62
                 SeuNumero         := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,25)  // uso da empresa, nº de controle da empresa;
              else
                 SeuNumero         := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,24)+VarDif;  // uso da empresa, nº de controle da empresa;
            end
            else
            begin
              inc(carne);
              ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
              if empty(VarDif) then // coluna 38 até 62
                  SeuNumero         := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,25) // uso da empresa, nº de controle da empresa;
              else
                  SeuNumero         :=copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(7),1,24)+VarDif;  // uso da empresa, nº de controle da empresa;
            end;
            VarDif:='';


            NumeroDocumento   := inttostrZero(FrmRelRecebimento.CDS_MarcaTitidreceb.Value,10);//copy(DM_Comercio.CDSMarcaTitordem.Value+space(10),1,10);
            DataDocumento     := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.VALUE;
            EspecieMod        := 'R$';
            Aceite            := atNao; //atSim
            DataProcessamento := Now;
            Carteira          := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);
            if CBCarteiraEnvio.Checked then
               CarteiraEnvio     := tceBanco   // vai retornar '1'   tceCedente; // vai retornar '2'    coluna 93
            else
               CarteiraEnvio     := tceCedente;       // coluna 93 cnab400
            ValorDocumento    := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);

            //-------------------------------------------------------
            ps:=pos('-',FrmRelRecebimento.CDS_MarcaTitordem.Value);
            inc(ps);
            pega:=copy(Trim(FrmRelRecebimento.CDS_MarcaTitordem.Value),ps,10);
            ps:=pos('/',pega);
            inc(ps);
            pega:=trim(copy(pega,ps,3));
            if not empty(pega) then
               ps:=pega.ToInteger
            else
              ps:=1;

            TotalParcelas     := ps;
            ps:=pos('-',FrmRelRecebimento.CDS_MarcaTitordem.Value);
            inc(ps);
            pega:=copy(Trim(FrmRelRecebimento.CDS_MarcaTitordem.Value),ps,10);
            ps:=pos('/',pega);
            dec(ps);
            pega:=trim(copy(pega,1,ps));
            if not empty(pega) then
               ps:=pega.ToInteger
            else
              ps:=1;

            parcela           := ps;

            //Dados do Cedente

            ACBrBoleto.Cedente.Agencia           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
            ACBrBoleto.Cedente.AgenciaDigito     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
            ACBrBoleto.Cedente.Conta             := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
            ACBrBoleto.Cedente.ContaDigito       := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
            ACBrBoleto.Cedente.CodigoTransmissao := trim(DM_Tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString); //trim(DM_comercio.ZQConta_BancariaCodigo_cedente.Value);
            if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
               ACBrBoleto.Cedente.Modalidade     :='01'
            else
               ACBrBoleto.Cedente.Modalidade     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);
            ACBrBoleto.Cedente.CodigoCedente     := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_no_banco').AsString);
            ACBrBoleto.Cedente.Convenio          := trim(DM_Tabelas.ZQContaBancaria.FieldByName('convenio').AsString);

            if length(tiramascara(trim(DM_Tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString)))<14 then
            begin
               ACBrBoleto.Cedente.CNPJCPF := tiramascara(trim(DM_Tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString));
               ACBrBoleto.Cedente.TipoInscricao := pFisica;
            end
            else
            begin
               ACBrBoleto.Cedente.CNPJCPF := copy(tiramascara(trim(DM_Tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString)),1,14);
               ACBrBoleto.Cedente.TipoInscricao := pJuridica;
            end;
            ACBrBoleto.Cedente.Nome              := charrem(trim(DM_Tabelas.ZQContaBancaria.FieldByName('nometitular').AsString));
            try
              zqaux:= tzquery.Create(nil);
              zqaux.Connection := DM_Tabelas.zconeccao;
              zqaux.Close;
              zqaux.sql.clear;
              zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
              zqaux.Open;
              if zqaux.RecordCount>0 then
              begin
                  ACBrBoleto.Cedente.Logradouro        := Removeacentos(uppercase(trim(copy(zqaux.FieldByName('ende_cob').AsString,1,pos(',',zqaux.FieldByName('ende_cob').AsString)-1))));
                  ACBrBoleto.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                  ACBrBoleto.Cedente.Complemento       := '';
                  ACBrBoleto.Cedente.Bairro            := Removeacentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                  ACBrBoleto.Cedente.Cidade            := Removeacentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                  ACBrBoleto.Cedente.UF                := Removeacentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                  ACBrBoleto.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
              end;
            finally
              zqaux.Close;
              zqaux.free;
            end;
            if (CBBancoemite.Checked=false) then
              ACBrBoleto.Cedente.ResponEmissao:=tbCliEmite
            else
              ACBrBoleto.Cedente.ResponEmissao:=tbBancoEmite;


            //dados do sacado
            Sacado.NomeSacado   := charrem(Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value));
            if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
            begin
              Sacado.CNPJCPF  := Charrem(trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)));
              Sacado.Pessoa   := pJuridica;
            end
            else
            begin
              Sacado.CNPJCPF  := Charrem(trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)));
              Sacado.Pessoa   := pFisica;
            end;
            Sacado.Logradouro := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
            Sacado.CEP        := charrem(Trim(Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value))));
            Sacado.Cidade     := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
            Sacado.UF         := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
            Sacado.Bairro     := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);
//            if CBParcelas.Checked=true then
//               Sacado.Avalista:=COPY('N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+espacos(60),1,60)  // decomposição
//            else
//               Sacado.Avalista:=COPY('Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+espacos(60),1,60);  // decomposição


          //dados da cobranca

            codcli:=0;
            eita:=0;
//
            NossoNumero                  := copy(varnossnum,1,11);
            ValorDocumento               := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
            DataDocumento                := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
            Vencimento                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;
            NumeroDocumento              := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);
            if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
            begin
           //  RLBTitulo1.TipoOcorrencia := toRemessaConcederDesconto;
              DataDesconto               := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
              ValorDesconto              := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
            end;
            ValorAbatimento   := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
            LocalPagamento := 'PAGÁVEL PREFERENCIALMENTE NA REDE BRADESCO OU BRADESCO EXPRESSO';//trim(DM_comercio.ZQConta_BancariaInstrucoes.Value);

             DataProtesto   := 0; //FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
            // dias de protestos e carteira de envio
            TipoDiasProtesto := diCorridos;

            if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                 PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
            else
                 PercentualMulta                                 := 0;
            CodigoMora        :='1'; // 1 = Valor 2 = Percentual 9 = Dispensar Multa
            CodigoMulta       := cmPercentual;//(cmValorFixo, cmPercentual)
            DataMulta         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;
   //       MultaValorFixo    := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
            MultaValorFixo    := False;
            if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
               ValorMoraJuros            := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
            else
              ValorMoraJuros             := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
            DataMoraJuros                := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
            CodigoMoraJuros              := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
            CodigoMora                   := '1'; //1-por dia 2-mensal 3-isento
            CodigoMulta                  := cmPercentual;//(cmValorFixo, cmPercentual)



//           if CBParcelas.Checked=true then
//               mensagem.text := 'COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+'  Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
//           else
//               mensagem.text := 'COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

            if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2) > 0 then
            begin
              if CBParcelas.Checked=false then
              begin
            //    mensagem.Add('NAO RECEBER APOS 30 DIAS DO VENCIMENTO.');
                if DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat>0 then
                   mensagem.Add(' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2),'###,###,##0.00')));
                mensagem.Add(' N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);
              end
              else
              begin
                if DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat>0 then
                   mensagem.Add(' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2),'###,###,##0.00')));
                mensagem.Add(' N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);
              end;
            end
            else
            begin
              if CBParcelas.Checked=false then
              begin
      //          mensagem.Add('NAO RECEBER APOS 30 DIAS DO VENCIMENTO.');
                if DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat>0 then
                   mensagem.Add(' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2),'###,###,##0.00')));
                mensagem.Add(' N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);
              end
              else
              begin
                if DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat>0 then
                   mensagem.Add(' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2),'###,###,##0.00')));
                mensagem.Add(' N.Parc.:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);
              end;
            end;
            case CBIdentOcorr.ItemIndex of
                1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
            end;
          end;

        end;
        FrmRelRecebimento.CDS_MarcaTit.Next;
       end
       else
       begin
         FrmRelRecebimento.CDS_MarcaTit.Next;
       end;
      end;
      FrmRelRecebimento.CDS_MarcaTit.EnableControls;
      //gauge2.Progress:=0;
      Gauge2.Visible:=false;
      FrmRelRecebimento.CDS_MarcaTit.EnableControls;
    //  ACBrBoleto1.GerarRemessa( varremes );
      inc(varregistro);
  //    mensagem('Remessa de Boletos gerado com sucesso!!!'+chr(13)+'A remessa foi gravado em '+VARARQ);
  end;


    //Emissão dos Boletos
  if CBBancoemite.Checked=false then
  begin
    ACBrBoleto1.Imprimir;
  end;

  FrmRelRecebimento.CDS_MarcaTit.EnableControls;
//  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
//  DM_Tabelas.ZQContaBancaria.edit;
//  DM_Tabelas.ZQContaBancarian_sequencial.Value:=varremes;
//  if rgnosso.Checked then
//     DM_Tabelas.ZQContaBancarianosso_numero.Value:=varnossnum.ToInteger;
//  DM_Tabelas.ZQContaBancaria.post;
//  DM_Tabelas.ZQContaBancaria.ApplyUpdates;
//  DM_Tabelas.ZQContaBancaria.Refresh;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

  Gauge1.Visible      := False;
  // 14/12/2012
  if  rgnosso.Checked then
  begin
    Gauge1.Visible := True;
    Gauge1.Progress:=0;

    //17/09/2013
    ZQRemesRec.close;
    ZQRemesRec.SQL.Clear;
    ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where remessa='+quotedstr(inttostr(numeroremessa)));
    ZQRemesRec.open;



    ZQRemesRec.First;
    Gauge1.MaxValue := ZQRemesRec.RecordCount;
//    DM_Tabelas.ZQRemes_Receb_atualiza.open;
  //  DM_Tabelas.ZQRemes_Receb_atualiza.first;
    ZQRemesRec.DisableControls;
    while not ZQRemesRec.Eof do
    begin
      Gauge1.Progress := ZQRemesRec.RecNo;

      DM_Tabelas.ZQRemes_Receb_atualiza.Close;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text));
      DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;
      ZQRemesRec.Next;
    end;
    ZQRemesRec.EnableControls;
//    DM_Tabelas.ZQRemes_Receb_atualiza.refresh;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
//    DM_Tabelas.ZQRecebimento.refresh;
  end;
  //16/07/2021
  if cb2via.Checked=false then
  begin
    Gauge1.Visible := True;
    Gauge1.Progress:=0;


    ZQRemesRec.close;
    ZQRemesRec.SQL.Clear;
    ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where remessa='+quotedstr(inttostr(numeroremessa)));
    ZQRemesRec.open;



    ZQRemesRec.First;
    Gauge1.MaxValue := ZQRemesRec.RecordCount;
//    DM_Tabelas.ZQRemes_Receb_atualiza.open;
  //  DM_Tabelas.ZQRemes_Receb_atualiza.first;
    ZQRemesRec.DisableControls;
    while not ZQRemesRec.Eof do
    begin
      Gauge1.Progress := ZQRemesRec.RecNo;

      DM_Tabelas.ZQRemes_Receb_atualiza.Close;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text));
      DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;

      ZQRemesRec.Next;
    end;
    ZQRemesRec.EnableControls;
//    DM_Tabelas.ZQRemes_Receb_atualiza.refresh;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
//    DM_Tabelas.ZQRecebimento.refresh;
  end;
  FrmRelRecebimento.CDS_MarcaTit.EnableControls;
  Gauge1.Progress:=0;
  Gauge1.Visible := false;
end;

procedure TFrmCobrancaBancaria.brasil_400;
Var
  F : TextFile;
  ArqIni : tIniFile;
  VARARQ, Vararqnome, Varmens1, Varmens2, Varmens3, Varmens4, varnossonumero,
  varnossnum, varmora, vardoc, varcpfcnpj, varendereco, varhoje, VarNomepasta,
  varhostname, vardir, varpastabanco,VarPath, ext,VarDif,ql : string;
  xnossonumero,vpar,NN,sq,totlin,xnosso,tm,numeroremessa,codcli,eita,carne,tam,tam2,varregistro, varremes, varx, vary : integer;
  data:Tdatetime;
  totvr:double;
  Titulo : TACBrTitulo;
  I : TACBrBolLayOut ;
begin
  totlin:=0;
  toTvr:=0;
  carne:=0;
  codcli:=0;
  tam:=0;
  VarDif:='';
  ql:='';
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  VarNomepasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );

//  varpastabanco := FrmRelRecebimento.CDSEmpreendimentoapelido.Value+'_'+FrmRelRecebimento.CDSEmpreendimentonomebanco.Value+FrmRelRecebimento.CDSEmpreendimentoconta.Value;
  if cb2via.Checked=false then
  begin

    vardir := '\\'+varhostname+'\'+VarNomePasta;
    // aqui gera a remessa
    if copy(CBIdentOcorr.Text,1,2)<>'00' then
    begin
      SaveDialog1.Execute;
      tam2:=length(SaveDialog1.FileName);
      tam2:=tam2-11;
      VarPath := copy(SaveDialog1.FileName,tam2+1,11);
      if trim(VarPath)<>'remessa.rem' then
      begin
        mensagem('O nome do Arquivo tem que ser remessa.rem');
        exit;
      end;

      VarPath := copy(SaveDialog1.FileName,1,tam2);
  //  if not DirectoryExists(vardir) then
  //    vardir := VarPath+varpastabanco;
      vardir := VarPath;
  //  else
  //    vardir := vardir +'\'+varpastabanco;
      if empty(trim(VarPath)) then
      begin
        mensagem('O nome do Arquivo não pode ficar em branco, tem que ser remessa.rem');
        exit;
      end;


      DBGerar.Enabled:=false;

      //CreateDir(vardir);

  //  vardir := vardir +'\'+copy(datetostr(date),7,4);
  //  CreateDir(vardir);
  //  vardir := vardir +'\'+copy(datetostr(date),4,2);
  //  CreateDir(vardir);
      vardir := vardir +'\'+copy(datetostr(date),1,2);
      CreateDir(vardir);
      Ext := '.REM';
      if CBTipoArq.Checked Then
        Ext := '.TST';

      varregistro:=2;

      for  varremes:= 1 to 99 do begin
        vararq := inttostrZero(varremes,2);
        Vararqnome := 'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+Ext;
        VARARQ :=vardir+'\'+Vararqnome;
        If not(fileexists(VARARQ)) then
          Break;
      end;
      Varnomearq := VARARQ;
      if DM_Tabelas.ZQBol_men.State in [DsEdit, DsInsert] Then
        DM_Tabelas.ZQBol_men.Edit;


      //pegar o ultimo registro auto incremento criado

      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

      DM_tabelas.ZQsequencia_remessa.close;
      DM_tabelas.ZQsequencia_remessa.SQL.Clear;
      DM_tabelas.ZQsequencia_remessa.SQL.Add('select * from remessa_sequencia where idconta_bancaria='+quotedstr(FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Text));
      DM_tabelas.ZQsequencia_remessa.Open;

      if DM_tabelas.ZQsequencia_remessa.RecordCount>0 then
      begin
        DM_tabelas.ZQsequencia_remessa.edit;
        DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt:=DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt+1;
        DM_tabelas.ZQsequencia_remessa.Post;
        DM_tabelas.ZQsequencia_remessa.Refresh;
      end;
      varremes                                                     := DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt;//DM_Tabelas.ZQContaBancarian_sequencial.Value+1;
      DM_tabelas.ZQsequencia_remessa.close;

      xnosso                                                       := DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;

      Frm_NumRemessa.XNENumRemessa.Value                           := varremes;
      Frm_NumRemessa.Label2.Caption                                := VARARQ;
      Frm_NumRemessa.Top                                           := FrmCobrancaBancaria.Top + 30;
      Frm_NumRemessa.Left                                          := FrmCobrancaBancaria.Left + 300;
      AbrirModal(Self, Frm_NumRemessa);
      //numero da remessa
      varremes                                                     := strtoint(floattostr(Frm_NumRemessa.XNENumRemessa.value));


      DM_Tabelas.ZQBancRemes.Insert;
      DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger                          := FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value;
      DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger                          := varremes;
      DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString                          := Vararqnome;
      DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime                           := date;
      DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString                       := FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
      DM_Tabelas.ZQBancRemes.Post;

      varhoje                                                      := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);



      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible                                               := True;
      Gauge1.Progress                                              :=0;
      Gauge1.MaxValue                                              := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  //    xnosso                                                     :=DM_Tabelas.ZQContaBancarianosso_numero.Value;

      ACBrBoleto1.DirArqRemessa                                    := vardir; //VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.LayoutRemessa                                    :=c400;
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.Banco.Digito                                     := 9;
      ACBrBoleto1.Banco.TipoCobranca                               :=cobBancoDoBrasil;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum                      :=10;

      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=0;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=0;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=0;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=0;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;


      ACBrBoletoFCFortes1.DirLogo                                  :=ExtractFilePath( Application.ExeName)+'Logos_bco';

      ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

      if zqcarne.active=false then
         zqcarne.open;
      zqcarne.last;
      carne                                                        :=zqcarne.FieldByName('n_seq').AsInteger;
      inc(carne);

      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge1.Progress                                            :=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
        begin
          if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0' Then
          begin
            if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            begin
               mensagem('O boleto '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente');
            end
            else
            begin

              //Dados do Cedente
              DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

              ACBrBoleto1.DataArquivo                              := Date;
              ACBrBoleto1.Banco.Numero                             := strtoint(trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString));
              ACBrBoleto1.Cedente.Agencia                          :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString)<>emptystr then
                 ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
              //ACBrBoleto1.Cedente.AgenciaDigito                  :=trim(DM_Tabelas.ZQContaBancarian_agencia_v.value);
              ACBrBoleto1.Cedente.Conta                            := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
              ACBrBoleto1.Cedente.ContaDigito                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
              ACBrBoleto1.Cedente.CodigoCedente                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString)+'-'+trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString); //trim(DM_tabelas.ZQContaBancariaconvenio.Value);
              ACBrBoleto1.Cedente.CodigoTransmissao                := trim(DM_tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString);
              case ComboBox3.ItemIndex of
                   0: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lPadrao);
                   1: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lCarne);
              end;

              ACBrBoleto1.Cedente.Convenio                         := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.TipoDocumento                    := escritural;
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
                 ACBrBoleto1.Cedente.Modalidade                    :='01'
              else
                 ACBrBoleto1.Cedente.Modalidade                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);  {Tipo de Modalidade Carteira
                                                                                                      101-Cobrança Rápida COM Registro  cliente emite os boletos
                                                                                                      104-Cobrança Eletrônica COM Registro bco emite os boletos}

              if CBBancoemite.checked then
                 ACBrBoleto1.Cedente.ResponEmissao := tbBancoEmite
              else
                ACBrBoleto1.Cedente.ResponEmissao := tbCliEmite;
              ACBrBoleto1.Cedente.CaracTitulo                      := tcSimplesRapComReg; //(tcSimples, tcVinculada, tcCaucionada, tcDescontada, tcVendor, tcDireta, tcSimplesRapComReg, tcCaucionadaRapComReg, tcDiretaEspecial)
              ACBrBoleto1.Cedente.TipoCarteira                     := tctRegistrada;//tctSimples
              ACBrBoleto1.Cedente.Nome                             := trim(DM_tabelas.ZQContaBancaria.FieldByName('nometitular').AsString);
              if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pJuridica
              else  if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))=11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pFisica
              else
              begin
                mensagem('Falta o documento (CNPJ/CPF) do cedente');
                exit;
              end;
              ACBrBoleto1.Cedente.CNPJCPF                          := tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString);

              try
                zqaux:= tzquery.Create(nil);
                zqaux.Connection := DM_Tabelas.zconeccao;
                zqaux.Close;
                zqaux.sql.clear;
                zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
                zqaux.Open;
                if zqaux.RecordCount>0 then
                begin
                    ACBrBoleto1.Cedente.Logradouro        := Removeacentos(uppercase(trim(copy(zqaux.FieldByName('ende_cob').AsString,1,pos(',',zqaux.FieldByName('ende_cob').AsString)-1))));
                    ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                    ACBrBoleto1.Cedente.Complemento       := '';
                    ACBrBoleto1.Cedente.Bairro            := Removeacentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                    ACBrBoleto1.Cedente.Cidade            := Removeacentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                    ACBrBoleto1.Cedente.UF                := Removeacentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                    ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
                end;
              finally
                zqaux.Close;
                zqaux.free;
              end;



              Titulo                                               := ACBrBoleto1.CriarTituloNaLista;
              with Titulo do
              begin
                //dados do sacado
                Sacado.NomeSacado                                  := Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
                if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pJuridica;;
                end
                else
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pFisica;
                end;
                Sacado.Logradouro                                  := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
                Sacado.CEP                                         := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
                Sacado.Cidade                                      := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
                Sacado.UF                                          := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
                Sacado.Bairro                                      := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);

                //dados da cobranca

                ValorAbatimento                                    := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
                LocalPagamento                                     := 'PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO';

                CodigoMoraJuros                                    := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
                CodigoMora                                         := '1'; //1-por dia 2-mensal 3-isento
                CodigoMulta                                        := cmPercentual;//(cmValorFixo, cmPercentual)
                DataMulta                                          := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

                if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                   PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
                else
                   PercentualMulta                                 := 0;

       //            MultaValorFixo                                := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
                MultaValorFixo                                     := False;

      //            DataProtesto                                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;
               DataProtesto                                        := 0;


                eita                                               :=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
                if eita=1 then
                   codcli                                          :=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
                if eita=2 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
                if eita=3 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
                if eita=4 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
                if eita>=5 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
                eita                                               :=codcli;
                codcli                                             :=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                vardoc                                             := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(25),1,25);

                inc(xnosso);

                varnossonumero                                     := inttostrZero(xnosso,10);
                xnossonumero                                       := DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;
                // 14/12/2012
                if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
                begin
                 xnossonumero   := strtoint(copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,1,17));
                 varnossonumero :=copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,8,10);
                end;

                codcli                                             :=0;
                eita                                               :=0;


                 ///////////////  28/11/2022
                ZQRemesRec.close;
                ZQRemesRec.SQL.Clear;
                ZQRemesRec.SQL.Add('Select  idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
                ZQRemesRec.open;
        {        ZQRemesRec.Filtered                               :=false;
                ZQRemesRec.Filter                                  :='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                ZQRemesRec.Filtered                                :=true;}
                if ZQRemesRec.RecordCount>0 then
                begin
                  ZQRemesRec.First;
                  ZQRemesRec.DisableControls;
                  while not ZQRemesRec.Eof do
                  begin
                    ZQRemesRec.delete;
                  end;
                  ZQRemesRec.EnableControls;
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('convenio').AsString+varnossonumero);
                  ZQRemesRec.Post;
                end
                else
                begin
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('convenio').AsString+varnossonumero);
                  ZQRemesRec.Post;
                end;
                ZQRemesRec.Filtered                                :=false;

                //02/08/2022
                FrmRelRecebimento.CDS_MarcaTit.edit;
                FrmRelRecebimento.CDS_MarcaTitnumboleto.Value      :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('convenio').AsString+varnossonumero);
                FrmRelRecebimento.CDS_MarcaTit.Post;
                //
                VarDif                                             :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


                // 17/09/2013

                numeroremessa                                      :=ZQRemesRec.FieldByName('remessa').AsInteger;

            /////////////////////////////////////////////////////////////////////

                ValorDocumento                                     := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
                if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
                begin
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   inc(vpar);
                   TotalParcelas     := strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   dec(vpar);
                   parcela           :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
                end
                else
                begin
                   TotalParcelas     :=  10;
                   parcela           :=  1;
                end;


                DataDocumento                                      := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
                if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
                   ValorMoraJuros                                  := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
                else
                  ValorMoraJuros                                   := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
                DataMoraJuros                                      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
                Vencimento                                         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;



                NumeroDocumento                                    := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);

                SeuNumero                                          := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
                EspecieMod                                         := '$';
                Aceite                                             := atNao; //atSim
                DataProcessamento                                  := Now;
                Carteira                                           := copy(trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString),2,2);
                if CBCarteiraEnvio.Checked then
                   CarteiraEnvio     := tceBanco   // vai retornar '1'   tceCedente; // vai retornar '2'    coluna 93
                else
                  CarteiraEnvio     := tceCedente;       // coluna 93 cnab400
                NossoNumero                                        := copy(varnossonumero,1,10);
                if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
                begin
    //              RLBTitulo2.TipoOcorrencia                      := toRemessaConcederDesconto;
                  CodigoDesconto                                   := cdValorFixo; //(cdSemDesconto, cdValorFixo);
                  TipoDesconto                                     := tdValorFixoAteDataInformada;
                  DataDesconto                                     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
                  ValorDesconto                                    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
                end;
               // tony pediu para trocar a mensagem no dia 28/11/2012
    //            RLBTitulo2.Instrucoes.Text                       := trim('NºO RECEBER APÓS 30 DIAS DE VENCIMENTO.Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            RLBTitulo2.Instrucoes.Text                       := trim('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100),2) >0 then
    //               RLBTitulo2.Instrucoes.Text                    := 'Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'))+#13+
    //                                             ' COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
    //            else


                if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
                begin
                  ql                                               :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                  Mensagem.Text                                    := 'NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+' N.Carnê:'+strzero(carne,5);
                  inc(carne);
                end;

                case CBIdentOcorr.ItemIndex of
                    1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                    2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                    3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                    4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                    5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                    6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                    7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                    8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                    9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                    10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                    11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                    12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                    13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                    14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                    15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                    16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                    17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
                end;
              end;

              if cbemail.Checked=true then
              begin
                 ZQAniversariante.close;
                 ZQAniversariante.SQL.clear;
                 ZQAniversariante.SQL.add('select * from participante where nome_parte='+quotedstr(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+' order by nome_parte');
                 ZQAniversariante.open;
                 if not empty(ZQAniversariante.FieldByName('email').AsString) then
                 begin
                   CDParticipante.Insert;
                   CDParticipanteidpaticipante.Value               :=ZQAniversariante.FieldByName('idpaticipante').AsLargeInt;
                   CDParticipantenome_parte.Value                  :=ZQAniversariante.FieldByName('nome_parte').AsString;
                   CDParticipanteemail.Value                       :=ZQAniversariante.FieldByName('email').AsString;
                   CDParticipanteinativo.Value                     :=ZQAniversariante.FieldByName('Inativo').AsString;
                   CDParticipantenascimento.Value                  :=ZQAniversariante.FieldByName('aniversario').AsDateTime;
                   CDParticipanteProfissao.Value                   :=ZQAniversariante.FieldByName('profissao').AsString;
                   CDParticipante.post;
                end;
                ZQAniversariante.close;
              end;
            end;

          end;
        end;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;
      ACBrBoleto1.GerarRemessa( varremes );
      zqcarne.Edit;
      zqcarne.FieldByName('n_seq').AsInteger                                           :=carne;
      zqcarne.post;
      zqcarne.Close;
      mensagem('Relatório de Boletos gerado com sucesso!!!'+chr(13)+'O relatório foi gravado em '+VARARQ);
    end;
  end
  else
  begin
      varhoje                                                      := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);

      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible                                               := True;
      Gauge1.Progress                                              :=0;
      Gauge1.MaxValue                                              := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  //    xnosso                                                     :=DM_Tabelas.ZQContaBancarianosso_numero.Value;

      ACBrBoleto1.DirArqRemessa                                    := VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
      ACBrBoleto1.Banco.Digito                                     := 9;
      ACBrBoleto1.LayoutRemessa                                    :=c400;
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.Banco.TipoCobranca                               :=cobBancoDoBrasil;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum                      :=10;

      ACBrBoletoFCFortes1.DirLogo                                  :=ExtractFilePath( Application.ExeName)+'Logos_bco';

      ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=0;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=0;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=0;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=0;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;

//      if zqcarne.active=false then
//         zqcarne.open;
//      zqcarne.last;
//      carne                                                        :=zqcarnen_seq.Value;
//      inc(carne);

      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge1.Progress                                            :=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
        begin
          if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0' Then
          begin
            if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            begin
               mensagem('O boleto '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente');
            end
            else
            begin

              //Dados do Cedente
              DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

              ACBrBoleto1.DataArquivo                              := Date;
              ACBrBoleto1.Banco.Numero                             := strtoint(trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString));
              ACBrBoleto1.Cedente.Agencia                          :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString)<>emptystr then
                 ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
              //ACBrBoleto1.Cedente.AgenciaDigito                  :=trim(DM_Tabelas.ZQContaBancarian_agencia_v.value);
              ACBrBoleto1.Cedente.Conta                            := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
              ACBrBoleto1.Cedente.ContaDigito                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
              ACBrBoleto1.Cedente.CodigoCedente                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString)+'-'+trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);//trim(DM_tabelas.ZQContaBancariaconvenio.Value);
              ACBrBoleto1.Cedente.CodigoTransmissao                := trim(DM_tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString);
              case ComboBox3.ItemIndex of
                   0: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lPadrao);
                   1: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lCarne);
              end;

              ACBrBoleto1.Cedente.Convenio                         := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.TipoDocumento                    := escritural;
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
                 ACBrBoleto1.Cedente.Modalidade                    :='01'
              else
                 ACBrBoleto1.Cedente.Modalidade                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);  {Tipo de Modalidade Carteira
                                                                                                      101-Cobrança Rápida COM Registro  cliente emite os boletos
                                                                                                      104-Cobrança Eletrônica COM Registro bco emite os boletos}

              if CBBancoemite.checked then
                 ACBrBoleto1.Cedente.ResponEmissao := tbBancoEmite
              else
                ACBrBoleto1.Cedente.ResponEmissao := tbCliEmite;
              ACBrBoleto1.Cedente.CaracTitulo                      := tcSimplesRapComReg; //(tcSimples, tcVinculada, tcCaucionada, tcDescontada, tcVendor, tcDireta, tcSimplesRapComReg, tcCaucionadaRapComReg, tcDiretaEspecial)
              ACBrBoleto1.Cedente.TipoCarteira                     := tctRegistrada;//tctSimples
              ACBrBoleto1.Cedente.Nome                             := trim(DM_tabelas.ZQContaBancaria.FieldByName('nometitular').AsString);
              if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pJuridica
              else  if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))=11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pFisica
              else
              begin
                mensagem('Falta o documento (CNPJ/CPF) do cedente');
                exit;
              end;
              ACBrBoleto1.Cedente.CNPJCPF                          := tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString);

             try
              zqaux:= tzquery.Create(nil);
              zqaux.Connection := DM_Tabelas.zconeccao;
              zqaux.Close;
              zqaux.sql.clear;
              zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
              zqaux.Open;
              if zqaux.RecordCount>0 then
              begin
                  ACBrBoleto1.Cedente.Logradouro        := Removeacentos(uppercase(trim(zqaux.FieldByName('ende_cob').AsString)));
                //  ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                  ACBrBoleto1.Cedente.Complemento       := '';
                  ACBrBoleto1.Cedente.Bairro            := Removeacentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                  ACBrBoleto1.Cedente.Cidade            := Removeacentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                  ACBrBoleto1.Cedente.UF                := Removeacentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                  ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
              end;
            finally
              zqaux.Close;
              zqaux.free;
            end;



              Titulo                                               := ACBrBoleto1.CriarTituloNaLista;
              with Titulo do
              begin
                //dados do sacado
                Sacado.NomeSacado                                  := Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
                if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pJuridica;;
                end
                else
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pFisica;
                end;
                Sacado.Logradouro                                  := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
                Sacado.CEP                                         := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
                Sacado.Cidade                                      := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
                Sacado.UF                                          := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
                Sacado.Bairro                                      := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);

                //dados da cobranca

                ValorAbatimento                                    := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
                LocalPagamento                                     := 'PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO';

                CodigoMoraJuros                                    := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
                CodigoMora                                         := '1'; //1-por dia 2-mensal 3-isento
                CodigoMulta                                        := cmPercentual;//(cmValorFixo, cmPercentual)
                DataMulta                                          := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

                if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                   PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
                else
                   PercentualMulta                                 := 0;

       //            MultaValorFixo                                := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
                MultaValorFixo                                     := False;

      //            DataProtesto                                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;
               DataProtesto                                        := 0;


                eita                                               :=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
                if eita=1 then
                   codcli                                          :=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
                if eita=2 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
                if eita=3 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
                if eita=4 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
                if eita>=5 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
                eita                                               :=codcli;
                codcli                                             :=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                vardoc                                             := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(25),1,25);

                xnosso                                             := strtoint(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value);
                varnossonumero                                     := inttostrZero(xnosso,10);
                xnossonumero                                       := DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;

                // 14/12/2012
                if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
                begin
                 xnossonumero   := strtoint(copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,1,17));
                 varnossonumero :=copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,8,10);
                end;

                codcli                                             :=0;
                eita                                               :=0;


                 ///////////////  28/11/2022
                ZQRemesRec.close;
                ZQRemesRec.SQL.Clear;
                ZQRemesRec.SQL.Add('Select  idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
                ZQRemesRec.open;
                ZQRemesRec.First;
        {        ZQRemesRec.Filtered                               :=false;
                ZQRemesRec.Filter                                  :='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                ZQRemesRec.Filtered                                :=true;}
                if ZQRemesRec.RecordCount=0 then
                begin
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('convenio').AsString+varnossonumero);
                  ZQRemesRec.Post;
                end;

                numeroremessa                                      :=ZQRemesRec.FieldByName('remessa').AsInteger;
                ZQRemesRec.Filtered                                :=false;
                //02/08/2022
                FrmRelRecebimento.CDS_MarcaTit.edit;
                FrmRelRecebimento.CDS_MarcaTitnumboleto.Value      :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('convenio').AsString+varnossonumero);
                FrmRelRecebimento.CDS_MarcaTit.Post;
                //
                VarDif                                             :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


            /////////////////////////////////////////////////////////////////////

                ValorDocumento                                     := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
                if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
                begin
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   inc(vpar);
                   TotalParcelas     := strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   dec(vpar);
                   parcela           :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
                end
                else
                begin
                   TotalParcelas     :=  10;
                   parcela           :=  1;
                end;


                DataDocumento                                      := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
                if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
                   ValorMoraJuros                                  := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
                else
                  ValorMoraJuros                                   := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
                DataMoraJuros                                      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
                Vencimento                                         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;



                NumeroDocumento                                    := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);

                SeuNumero                                          := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
                EspecieMod                                         := '$';
                Aceite                                             := atNao; //atSim
                DataProcessamento                                  := Now;
                Carteira                                           := copy(trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString),2,2);
                NossoNumero                                        := copy(varnossonumero,1,10);
                if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
                begin
    //              RLBTitulo2.TipoOcorrencia                      := toRemessaConcederDesconto;
                  CodigoDesconto                                   := cdValorFixo; //(cdSemDesconto, cdValorFixo);
                  TipoDesconto                                     := tdValorFixoAteDataInformada;
                  DataDesconto                                     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
                  ValorDesconto                                    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
                end;
               // tony pediu para trocar a mensagem no dia 28/11/2012
    //            RLBTitulo2.Instrucoes.Text                       := trim('NºO RECEBER APÓS 30 DIAS DE VENCIMENTO.Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            RLBTitulo2.Instrucoes.Text                       := trim('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100),2) >0 then
    //               RLBTitulo2.Instrucoes.Text                    := 'Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'))+#13+
    //                                             ' COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
    //            else
                if zqcarne.active=false then
                   zqcarne.open;
                zqcarne.last;
                carne                                              :=zqcarne.FieldByName('n_seq').AsInteger;
                ql                                                 :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
                begin
                  ql                                               :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                  Mensagem.Text                                    := 'NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+' N.Carnê:'+strzero(carne,5);
                end;
                case CBIdentOcorr.ItemIndex of
                    1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                    2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                    3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                    4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                    5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                    6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                    7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                    8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                    9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                    10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                    11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                    12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                    13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                    14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                    15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                    16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                    17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
                end;
              end;

              if cbemail.Checked=true then
              begin
                 ZQAniversariante.close;
                 ZQAniversariante.SQL.clear;
                 ZQAniversariante.SQL.add('select * from participante where nome_parte='+quotedstr(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+' order by nome_parte');
                 ZQAniversariante.open;
                 if not empty(ZQAniversariante.FieldByName('email').AsString) then
                 begin
                   CDParticipante.Insert;
                   CDParticipanteidpaticipante.Value               :=ZQAniversariante.FieldByName('idpaticipante').AsLargeInt;
                   CDParticipantenome_parte.Value                  :=ZQAniversariante.FieldByName('nome_parte').AsString;
                   CDParticipanteemail.Value                       :=ZQAniversariante.FieldByName('email').AsString;
                   CDParticipanteinativo.Value                     :=ZQAniversariante.FieldByName('Inativo').AsString;
                   CDParticipantenascimento.Value                  :=ZQAniversariante.FieldByName('aniversario').AsDateTime;
                   CDParticipanteProfissao.Value                   :=ZQAniversariante.FieldByName('profissao').AsString;
                   CDParticipante.post;
                end;
                ZQAniversariante.close;
              end;
            end;
          end;
        end;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;

  end;


    //Emissão dos Boletos
  if CBBancoemite.Checked=false then
  begin
    ACBrBoleto1.Imprimir;
  end;

  FrmRelRecebimento.CDS_MarcaTit.EnableControls;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  DM_Tabelas.ZQContaBancaria.edit;
  DM_Tabelas.ZQContaBancaria.FieldByName('n_sequencial').AsLargeInt:=varremes;
  if rgnosso.Checked then
     DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt:=xnosso;
  DM_Tabelas.ZQContaBancaria.post;
  DM_Tabelas.ZQContaBancaria.ApplyUpdates;
  DM_Tabelas.ZQContaBancaria.Refresh;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

  Gauge1.Visible      := False;
//  if cbemail.Checked=false then
  //   RLBTitulo2.PreviewModal;

  // 14/12/2012
  if  rgnosso.Checked then
  begin
    Gauge1.Visible    := True;
    Gauge1.Progress   :=0;

    //17/09/2013
    ZQRemesRec.close;
    ZQRemesRec.SQL.Clear;
    ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where remessa='+quotedstr(inttostr(numeroremessa)));
    ZQRemesRec.open;



    ZQRemesRec.First;
    Gauge1.MaxValue   := ZQRemesRec.RecordCount;
//    DM_Tabelas.ZQRemes_Receb_atualiza.open;
  //  DM_Tabelas.ZQRemes_Receb_atualiza.first;
    ZQRemesRec.DisableControls;
    while not ZQRemesRec.Eof do
    begin
      Gauge1.Progress := ZQRemesRec.RecNo;

      DM_Tabelas.ZQRemes_Receb_atualiza.Close;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text));
      DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;



      ZQRemesRec.Next;
    end;
    ZQRemesRec.EnableControls;
//    DM_Tabelas.ZQRemes_Receb_atualiza.refresh;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRecebimento.refresh;
  end;
  Gauge1.Progress     :=0;
  Gauge1.Visible      := false;
  varremes            :=0;
  xnosso              :=0;

end;




procedure TFrmCobrancaBancaria.Button1Click(Sender: TObject);
var
vazio:string;
begin
  GroupBox2.Visible:=false;
  vazio:='';
  if  CBatualizar.Checked then
  begin
    Gauge1.Visible := True;
    Gauge1.Progress:=0;

    //17/09/2013
    ZQRemesRec.close;
    ZQRemesRec.SQL.Clear;
    ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb ');
    ZQRemesRec.open;

    ZQRemesRec.First;
    Gauge1.MaxValue := ZQRemesRec.RecordCount;
//    DM_Tabelas.ZQRemes_Receb_atualiza.open;
  //  DM_Tabelas.ZQRemes_Receb_atualiza.first;
    ZQRemesRec.DisableControls;
    while not ZQRemesRec.Eof do
    begin
      Gauge1.Progress := ZQRemesRec.RecNo;

      DM_Tabelas.ZQRemes_Receb_atualiza.Close;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where ( idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text)+') and ((numboleto='+quotedstr(vazio)+') or (numboleto is null))');
      DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;
      ZQRemesRec.Next;
    end;
    ZQRemesRec.EnableControls;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRecebimento.refresh;
  end;
  Gauge1.Progress:=0;
  Gauge1.Visible := false;


   FrmRelRecebimento.CDSEmpreendimento.First;
   FrmRelRecebimento.CDSEmpreendimento.DisableControls;
   while not FrmRelRecebimento.CDSEmpreendimento.Eof do
   begin
      FrmRelRecebimento.CDS_MarcaTit.Filtered:=false;
      FrmRelRecebimento.CDS_MarcaTit.Filter:='idloteam='+quotedstr(FrmRelRecebimento.CDSEmpreendimentoidloteamento.Text);
      FrmRelRecebimento.CDS_MarcaTit.Filtered:=true;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge2.MaxValue:=FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      Gauge2.Progress:=0;
      Gauge2.Visible:=true;
      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge2.Progress:=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        FrmRelRecebimento.CDS_MarcaTit.Edit;
        FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value:=FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
        FrmRelRecebimento.CDS_MarcaTit.post;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;
      FrmRelRecebimento.CDS_MarcaTit.EnableControls;
      Gauge2.Progress:=0;

      Gauge2.Visible:=false;
      FrmRelRecebimento.CDSEmpreendimento.Next;
   end;
   FrmRelRecebimento.CDSEmpreendimento.EnableControls;
   FrmRelRecebimento.CDSEmpreendimento.First;
   FrmRelRecebimento.CDS_MarcaTit.Filtered := False;
   FrmRelRecebimento.CDS_MarcaTit.Filter := 'idloteam='+FrmRelRecebimento.CDSEmpreendimentoidloteamento.Text+' and cliente='+FrmRelRecebimento.CDSEmpreendimentoidparticipante.Text;
   FrmRelRecebimento.CDS_MarcaTit.Filtered := True;
   DBGerar.Enabled:=true;
end;

procedure TFrmCobrancaBancaria.cb2viaClick(Sender: TObject);
begin
   rgnosso.Checked:=false;
   CBBancoemite.Checked:=false;
end;




procedure TFrmCobrancaBancaria.santander_400;
Var
  F : TextFile;
  ArqIni : tIniFile;
  VARARQ, Vararqnome, Varmens1, Varmens2, Varmens3, Varmens4, varnossonumero,
  varnossnum, varmora, vardoc, varcpfcnpj, varendereco, varhoje, VarNomepasta,
  varhostname, vardir, varpastabanco,VarPath, ext,VarDif,ql : string;
  vpar,NN,sq,totlin,xnosso,tm,numeroremessa,codcli,eita,carne,tam,tam2,varregistro, varremes, varx, vary : integer;
  data:Tdatetime;
  totvr:double;
  Titulo : TACBrTitulo;
  I : TACBrBolLayOut ;
begin
  totlin:=0;
  toTvr:=0;
  carne:=0;
  codcli:=0;
  tam:=0;
  VarDif:='';
  ql:='';
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  VarNomepasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );

//  varpastabanco := FrmRelRecebimento.CDSEmpreendimentoapelido.Value+'_'+FrmRelRecebimento.CDSEmpreendimentonomebanco.Value+FrmRelRecebimento.CDSEmpreendimentoconta.Value;
  if cb2via.Checked=false then
  begin

    vardir := '\\'+varhostname+'\'+VarNomePasta;
    // aqui gera a remessa
    if copy(CBIdentOcorr.Text,1,2)<>'00' then
    begin
      SaveDialog1.Execute;
      tam2:=length(SaveDialog1.FileName);
      tam2:=tam2-11;
      VarPath := copy(SaveDialog1.FileName,tam2+1,11);
      if trim(VarPath)<>'remessa.rem' then
      begin
        mensagem('O nome do Arquivo tem que ser remessa.rem');
        exit;
      end;

      VarPath := copy(SaveDialog1.FileName,1,tam2);
  //  if not DirectoryExists(vardir) then
  //    vardir := VarPath+varpastabanco;
      vardir := VarPath;
  //  else
  //    vardir := vardir +'\'+varpastabanco;
      if empty(trim(VarPath)) then
      begin
        mensagem('O nome do Arquivo não pode ficar em branco, tem que ser remessa.rem');
        exit;
      end;


      DBGerar.Enabled:=false;

      //CreateDir(vardir);

  //  vardir := vardir +'\'+copy(datetostr(date),7,4);
  //  CreateDir(vardir);
  //  vardir := vardir +'\'+copy(datetostr(date),4,2);
  //  CreateDir(vardir);
      vardir := vardir +'\'+copy(datetostr(date),1,2);
      CreateDir(vardir);
      Ext := '.REM';
      if CBTipoArq.Checked Then
        Ext := '.TST';

      varregistro:=2;

      for  varremes:= 1 to 99 do begin
        vararq := inttostrZero(varremes,2);
        Vararqnome := 'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+Ext;
        VARARQ :=vardir+'\'+Vararqnome;
        If not(fileexists(VARARQ)) then
          Break;
      end;
      Varnomearq := VARARQ;
      if DM_Tabelas.ZQBol_men.State in [DsEdit, DsInsert] Then
        DM_Tabelas.ZQBol_men.Edit;


      //pegar o ultimo registro auto incremento criado

      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

      DM_tabelas.ZQsequencia_remessa.close;
      DM_tabelas.ZQsequencia_remessa.SQL.Clear;
      DM_tabelas.ZQsequencia_remessa.SQL.Add('select * from remessa_sequencia where idconta_bancaria='+quotedstr(FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Text));
      DM_tabelas.ZQsequencia_remessa.Open;

      if DM_tabelas.ZQsequencia_remessa.RecordCount>0 then
      begin
        DM_tabelas.ZQsequencia_remessa.edit;
        DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt:=DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt+1;
        DM_tabelas.ZQsequencia_remessa.Post;
        DM_tabelas.ZQsequencia_remessa.Refresh;
      end;
      varremes                                                     := DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt;//DM_Tabelas.ZQContaBancarian_sequencial.Value+1;
      DM_tabelas.ZQsequencia_remessa.close;

      xnosso                                                       := DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;

      Frm_NumRemessa.XNENumRemessa.Value                           := varremes;
      Frm_NumRemessa.Label2.Caption                                := VARARQ;
      Frm_NumRemessa.Top                                           := FrmCobrancaBancaria.Top + 30;
      Frm_NumRemessa.Left                                          := FrmCobrancaBancaria.Left + 300;
      AbrirModal(Self, Frm_NumRemessa);
      //numero da remessa
      varremes                                                     := strtoint(floattostr(Frm_NumRemessa.XNENumRemessa.value));


      DM_Tabelas.ZQBancRemes.Insert;
      DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger                          := FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value;
      DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger                          := varremes;
      DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString                          := Vararqnome;
      DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime                           := date;
      DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString                       := FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
      DM_Tabelas.ZQBancRemes.Post;

      varhoje                                                      := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);



      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible                                               := True;
      Gauge1.Progress                                              :=0;
      Gauge1.MaxValue                                              := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  //    xnosso                                                     :=DM_Tabelas.ZQContaBancarianosso_numero.Value;

      ACBrBoleto1.DirArqRemessa                                    := VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
      ACBrBoleto1.LayoutRemessa                                    :=c400;
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.Banco.TipoCobranca                               :=cobSantander;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum                      :=12;

      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=40;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=40;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=30;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=30;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;

      ACBrBoletoFCFortes1.DirLogo                                  :=ExtractFilePath( Application.ExeName)+'Logos_bco';

      ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

      if zqcarne.active=false then
         zqcarne.open;
      zqcarne.last;
      carne                                                        :=zqcarne.FieldByName('n_seq').AsInteger;
      inc(carne);

      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge1.Progress                                            :=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
        begin
          if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0' Then
          begin
            if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            begin
               mensagem('O boleto '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente');
            end
            else
            begin

              //Dados do Cedente
              DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

              ACBrBoleto1.DataArquivo                              := Date;
              ACBrBoleto1.Banco.Numero                             := strtoint(trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString));
              ACBrBoleto1.Cedente.Agencia                          :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString)<>emptystr then
                 ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
              //ACBrBoleto1.Cedente.AgenciaDigito                  :=trim(DM_Tabelas.ZQContaBancarian_agencia_v.value);
              ACBrBoleto1.Cedente.Conta                            := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
              ACBrBoleto1.Cedente.ContaDigito                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
              ACBrBoleto1.Cedente.CodigoCedente                    := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.CodigoTransmissao                := trim(DM_tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString);
              case ComboBox3.ItemIndex of
                   0: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lPadrao);
                   1: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lCarne);
              end;

              ACBrBoleto1.Cedente.Convenio                         := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.TipoDocumento                    := escritural;
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
                 ACBrBoleto1.Cedente.Modalidade                    :='101'
              else
                 ACBrBoleto1.Cedente.Modalidade                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);  {Tipo de Modalidade Carteira
                                                                                                      101-Cobrança Rápida COM Registro  cliente emite os boletos
                                                                                                      104-Cobrança Eletrônica COM Registro bco emite os boletos}
              if CBBancoemite.checked then
                 ACBrBoleto1.Cedente.ResponEmissao := tbBancoEmite
              else
                ACBrBoleto1.Cedente.ResponEmissao := tbCliEmite;
              ACBrBoleto1.Cedente.CaracTitulo                      := tcSimplesRapComReg; //(tcSimples, tcVinculada, tcCaucionada, tcDescontada, tcVendor, tcDireta, tcSimplesRapComReg, tcCaucionadaRapComReg, tcDiretaEspecial)
              ACBrBoleto1.Cedente.TipoCarteira                     := tctRegistrada;//tctSimples
              ACBrBoleto1.Cedente.Nome                             := trim(DM_tabelas.ZQContaBancaria.FieldByName('nometitular').AsString);
              if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pJuridica
              else  if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))=11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pFisica
              else
              begin
                mensagem('Falta o documento (CNPJ/CPF) do cedente');
                exit;
              end;
              ACBrBoleto1.Cedente.CNPJCPF                          := tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString);

              try
                zqaux:= tzquery.Create(nil);
                zqaux.Connection := DM_Tabelas.zconeccao;
                zqaux.Close;
                zqaux.sql.clear;
                zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
                zqaux.Open;
                if zqaux.RecordCount>0 then
                begin
                    ACBrBoleto1.Cedente.Logradouro        := Removeacentos(uppercase(trim(zqaux.FieldByName('ende_cob').AsString)));
                  //  ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                    ACBrBoleto1.Cedente.Complemento       := '';
                    ACBrBoleto1.Cedente.Bairro            := Removeacentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                    ACBrBoleto1.Cedente.Cidade            := Removeacentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                    ACBrBoleto1.Cedente.UF                := Removeacentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                    ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
                end;
              finally
                zqaux.Close;
                zqaux.free;
              end;



              Titulo                                               := ACBrBoleto1.CriarTituloNaLista;
              with Titulo do
              begin
                //dados do sacado
                Sacado.NomeSacado                                  := Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
                if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pJuridica;;
                end
                else
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pFisica;
                end;
                Sacado.Logradouro                                  := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
                Sacado.CEP                                         := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
                Sacado.Cidade                                      := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
                Sacado.UF                                          := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
                Sacado.Bairro                                      := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);

                //dados da cobranca

                ValorAbatimento                                    := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
                LocalPagamento                                     := 'PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO';

                CodigoMoraJuros                                    := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
                CodigoMora                                         := '1'; //1-por dia 2-mensal 3-isento
                CodigoMulta                                        := cmPercentual;//(cmValorFixo, cmPercentual)
                DataMulta                                          := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

                if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                   PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
                else
                   PercentualMulta                                 := 0;

       //            MultaValorFixo                                := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
                MultaValorFixo                                     := False;

      //            DataProtesto                                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;
               DataProtesto                                        := 0;


                eita                                               :=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
                if eita=1 then
                   codcli                                          :=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
                if eita=2 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
                if eita=3 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
                if eita=4 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
                if eita>=5 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
                eita                                               :=codcli;
                codcli                                             :=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                vardoc                                             := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(25),1,25);

                inc(xnosso);

                varnossonumero                                     := inttostrZero(xnosso,12);


               // 14/12/2012
               if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
                  varnossonumero                                   :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;
               //     varnossonumero                               :=inttostr(strtoint(copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,1,7)));


                codcli                                             :=0;
                eita                                               :=0;


                 ///////////////  28/11/2022
                ZQRemesRec.close;
                ZQRemesRec.SQL.Clear;
                ZQRemesRec.SQL.Add('Select  idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
                ZQRemesRec.open;
        {        ZQRemesRec.Filtered                               :=false;
                ZQRemesRec.Filter                                  :='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                ZQRemesRec.Filtered                                :=true;}
                if ZQRemesRec.RecordCount>0 then
                begin
                  ZQRemesRec.First;
                  ZQRemesRec.DisableControls;
                  while not ZQRemesRec.Eof do
                  begin
                    ZQRemesRec.delete;
                  end;
                  ZQRemesRec.EnableControls;
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      :=varnossonumero;
                  ZQRemesRec.Post;
                end
                else
                begin
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      :=varnossonumero;
                  ZQRemesRec.Post;
                end;
                ZQRemesRec.Filtered                                :=false;

                //02/08/2022
                FrmRelRecebimento.CDS_MarcaTit.edit;
                FrmRelRecebimento.CDS_MarcaTitnumboleto.Value      :=varnossonumero;
                FrmRelRecebimento.CDS_MarcaTit.Post;
                //
                VarDif                                             :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


                // 17/09/2013

                numeroremessa                                      :=ZQRemesRec.FieldByName('remessa').AsInteger;

            /////////////////////////////////////////////////////////////////////


                NossoNumero                                        := copy(varnossonumero,1,12);

                ValorDocumento                                     := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
                if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
                begin
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   inc(vpar);
                   TotalParcelas     := strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   dec(vpar);
                   parcela           :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
                end
                else
                begin
                   TotalParcelas     :=  10;
                   parcela           :=  1;
                end;


                DataDocumento                                      := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
                if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
                   ValorMoraJuros                                  := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
                else
                  ValorMoraJuros                                   := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
                DataMoraJuros                                      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
                Vencimento                                         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;



                NumeroDocumento                                    := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);

                SeuNumero                                          := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
                EspecieMod                                         := '$';
                Aceite                                             := atNao; //atSim
                DataProcessamento                                  := Now;
                Carteira                                           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);//copy(trim(DM_Tabelas.ZQContaBancariacarteira.Value),3,1);  // verificar se é só um digito mesmo

                if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
                begin
    //              RLBTitulo2.TipoOcorrencia                      := toRemessaConcederDesconto;
                  CodigoDesconto                                   := cdValorFixo; //(cdSemDesconto, cdValorFixo);
                  TipoDesconto                                     := tdValorFixoAteDataInformada;
                  DataDesconto                                     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
                  ValorDesconto                                    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
                end;
               // tony pediu para trocar a mensagem no dia 28/11/2012
    //            RLBTitulo2.Instrucoes.Text                       := trim('NºO RECEBER APÓS 30 DIAS DE VENCIMENTO.Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            RLBTitulo2.Instrucoes.Text                       := trim('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100),2) >0 then
    //               RLBTitulo2.Instrucoes.Text                    := 'Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'))+#13+
    //                                             ' COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
    //            else

                ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
                begin
                  ql                                               := FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                  Mensagem.Text                                    := 'NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+' N.Carnê:'+strzero(carne,5);
                  inc(carne);
                end;
                case CBIdentOcorr.ItemIndex of
                    1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                    2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                    3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                    4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                    5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                    6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                    7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                    8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                    9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                    10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                    11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                    12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                    13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                    14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                    15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                    16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                    17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
                end;
              end;

              if cbemail.Checked=true then
              begin
                 ZQAniversariante.close;
                 ZQAniversariante.SQL.clear;
                 ZQAniversariante.SQL.add('select * from participante where nome_parte='+quotedstr(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+' order by nome_parte');
                 ZQAniversariante.open;
                 if not empty(ZQAniversariante.FieldByName('email').AsString) then
                 begin
                   CDParticipante.Insert;
                   CDParticipanteidpaticipante.Value               :=ZQAniversariante.FieldByName('idpaticipante').AsLargeInt;
                   CDParticipantenome_parte.Value                  :=ZQAniversariante.FieldByName('nome_parte').AsString;
                   CDParticipanteemail.Value                       :=ZQAniversariante.FieldByName('email').AsString;
                   CDParticipanteinativo.Value                     :=ZQAniversariante.FieldByName('Inativo').AsString;
                   CDParticipantenascimento.Value                  :=ZQAniversariante.FieldByName('aniversario').AsDateTime;
                   CDParticipanteProfissao.Value                   :=ZQAniversariante.FieldByName('profissao').AsString;
                   CDParticipante.post;
                end;
                ZQAniversariante.close;
              end;
            end;
          end;
        end;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;
      ACBrBoleto1.GerarRemessa( varremes );
      zqcarne.Edit;
      zqcarne.FieldByName('n_seq').AsInteger                                           :=carne;
      zqcarne.post;
      zqcarne.Close;
      mensagem('Relatório de Boletos gerado com sucesso!!!'+chr(13)+'O relatório foi gravado em '+VARARQ);
    end;
  end
  else
  begin
      varhoje                                                      := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);



      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible                                               := True;
      Gauge1.Progress                                              :=0;
      Gauge1.MaxValue                                              := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  //    xnosso                                                     :=DM_Tabelas.ZQContaBancarianosso_numero.Value;

      ACBrBoleto1.DirArqRemessa                                    := VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
      ACBrBoleto1.LayoutRemessa                                    :=c400;
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.Banco.TipoCobranca                               :=cobSantander;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum                      :=12;

      ACBrBoletoFCFortes1.DirLogo                                  :=ExtractFilePath( Application.ExeName)+'Logos_bco';

      ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=40;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=40;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=30;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=30;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;

//      if zqcarne.active=false then
//         zqcarne.open;
//      zqcarne.last;
//      carne                                                        :=zqcarnen_seq.Value;
//      inc(carne);
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=40;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=40;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=30;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=30;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;
      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge1.Progress                                            :=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
        begin
          if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0' Then
          begin
            if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            begin
               mensagem('O boleto '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente');
            end
            else
            begin

              //Dados do Cedente
              DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

              ACBrBoleto1.DataArquivo                              := Date;
              ACBrBoleto1.Banco.Numero                             := strtoint(trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString));
              ACBrBoleto1.Cedente.Agencia                          :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString)<>emptystr then
                 ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
              //ACBrBoleto1.Cedente.AgenciaDigito                  :=trim(DM_Tabelas.ZQContaBancarian_agencia_v.value);
              ACBrBoleto1.Cedente.Conta                            := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
              ACBrBoleto1.Cedente.ContaDigito                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
              ACBrBoleto1.Cedente.CodigoCedente                    := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.CodigoTransmissao                := trim(DM_tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString);
              case ComboBox3.ItemIndex of
                   0: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lPadrao);
                   1: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lCarne);
              end;

              ACBrBoleto1.Cedente.Convenio                         := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.TipoDocumento                    := escritural;
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
                 ACBrBoleto1.Cedente.Modalidade                    :='101'
              else
                 ACBrBoleto1.Cedente.Modalidade                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);  {Tipo de Modalidade Carteira
                                                                                                      101-Cobrança Rápida COM Registro  cliente emite os boletos
                                                                                                      104-Cobrança Eletrônica COM Registro bco emite os boletos}

              if CBBancoemite.checked then
                 ACBrBoleto1.Cedente.ResponEmissao := tbBancoEmite
              else
                ACBrBoleto1.Cedente.ResponEmissao := tbCliEmite;
              ACBrBoleto1.Cedente.CaracTitulo                      := tcSimplesRapComReg; //(tcSimples, tcVinculada, tcCaucionada, tcDescontada, tcVendor, tcDireta, tcSimplesRapComReg, tcCaucionadaRapComReg, tcDiretaEspecial)
              ACBrBoleto1.Cedente.TipoCarteira                     := tctRegistrada;//tctSimples
              ACBrBoleto1.Cedente.Nome                             := trim(DM_tabelas.ZQContaBancaria.FieldByName('nometitular').AsString);
              if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pJuridica
              else  if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))=11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pFisica
              else
              begin
                mensagem('Falta o documento (CNPJ/CPF) do cedente');
                exit;
              end;
              ACBrBoleto1.Cedente.CNPJCPF                          := tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString);

              try
                zqaux:= tzquery.Create(nil);
                zqaux.Connection := DM_Tabelas.zconeccao;
                zqaux.Close;
                zqaux.sql.clear;
                zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
                zqaux.Open;
                if zqaux.RecordCount>0 then
                begin
                    ACBrBoleto1.Cedente.Logradouro        := Removeacentos(uppercase(trim(zqaux.FieldByName('ende_cob').AsString)));
                  //  ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                    ACBrBoleto1.Cedente.Complemento       := '';
                    ACBrBoleto1.Cedente.Bairro            := Removeacentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                    ACBrBoleto1.Cedente.Cidade            := Removeacentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                    ACBrBoleto1.Cedente.UF                := Removeacentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                    ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
                end;
              finally
                zqaux.Close;
                zqaux.free;
              end;



              Titulo                                               := ACBrBoleto1.CriarTituloNaLista;
              with Titulo do
              begin
                //dados do sacado
                Sacado.NomeSacado                                  := Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
                if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pJuridica;;
                end
                else
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pFisica;
                end;
                Sacado.Logradouro                                  := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
                Sacado.CEP                                         := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
                Sacado.Cidade                                      := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
                Sacado.UF                                          := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
                Sacado.Bairro                                      := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);

                //dados da cobranca

                ValorAbatimento                                    := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
                LocalPagamento                                     := 'PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO';

                CodigoMoraJuros                                    := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
                CodigoMora                                         := '1'; //1-por dia 2-mensal 3-isento
                CodigoMulta                                        := cmPercentual;//(cmValorFixo, cmPercentual)
                DataMulta                                          := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

                if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                   PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
                else
                   PercentualMulta                                 := 0;

       //            MultaValorFixo                                := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
                MultaValorFixo                                     := False;

      //            DataProtesto                                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;
               DataProtesto                                        := 0;


                eita                                               :=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
                if eita=1 then
                   codcli                                          :=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
                if eita=2 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
                if eita=3 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
                if eita=4 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
                if eita>=5 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
                eita                                               :=codcli;
                codcli                                             :=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                vardoc                                             := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(25),1,25);

                xnosso                                             := strtoint(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value);

                varnossonumero                                     := inttostrZero(xnosso,12);


               // 14/12/2012
               if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
                  varnossonumero                                   :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;
               //     varnossonumero                               :=inttostr(strtoint(copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,1,7)));


                codcli                                             :=0;
                eita                                               :=0;


                 ///////////////  28/11/2022
                ZQRemesRec.close;
                ZQRemesRec.SQL.Clear;
                ZQRemesRec.SQL.Add('Select  idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
                ZQRemesRec.open;
                ZQRemesRec.First;
        {        ZQRemesRec.Filtered                               :=false;
                ZQRemesRec.Filter                                  :='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                ZQRemesRec.Filtered                                :=true;}
                if ZQRemesRec.RecordCount=0 then
                begin
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      := varnossonumero;
                  ZQRemesRec.Post;
                end;

                numeroremessa                                      :=ZQRemesRec.FieldByName('remessa').AsInteger;
                ZQRemesRec.Filtered                                :=false;
                //02/08/2022
                FrmRelRecebimento.CDS_MarcaTit.edit;
                FrmRelRecebimento.CDS_MarcaTitnumboleto.Value      :=varnossonumero;
                FrmRelRecebimento.CDS_MarcaTit.Post;
                //
                VarDif                                             :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


            /////////////////////////////////////////////////////////////////////


                NossoNumero                                        := copy(varnossonumero,1,12);

                ValorDocumento                                     := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
                if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
                begin
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   inc(vpar);
                   TotalParcelas     := strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   dec(vpar);
                   parcela           :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
                end
                else
                begin
                   TotalParcelas     :=  10;
                   parcela           :=  1;
                end;


                DataDocumento                                      := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
                if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
                   ValorMoraJuros                                  := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
                else
                  ValorMoraJuros                                   := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
                DataMoraJuros                                      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
                Vencimento                                         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;



                NumeroDocumento                                    := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);

                SeuNumero                                          := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
                EspecieMod                                         := '$';
                Aceite                                             := atNao; //atSim
                DataProcessamento                                  := Now;
                Carteira                                           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);//copy(trim(DM_Tabelas.ZQContaBancariacarteira.Value),3,1);  // verificar se é só um digito mesmo

                if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
                begin
    //              RLBTitulo2.TipoOcorrencia                      := toRemessaConcederDesconto;
                  CodigoDesconto                                   := cdValorFixo; //(cdSemDesconto, cdValorFixo);
                  TipoDesconto                                     := tdValorFixoAteDataInformada;
                  DataDesconto                                     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
                  ValorDesconto                                    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
                end;
               // tony pediu para trocar a mensagem no dia 28/11/2012
    //            RLBTitulo2.Instrucoes.Text                       := trim('NºO RECEBER APÓS 30 DIAS DE VENCIMENTO.Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            RLBTitulo2.Instrucoes.Text                       := trim('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100),2) >0 then
    //               RLBTitulo2.Instrucoes.Text                    := 'Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'))+#13+
    //                                             ' COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
    //            else
                if zqcarne.active=false then
                   zqcarne.open;
                zqcarne.last;
                carne                                                        :=zqcarne.FieldByName('n_seq').AsInteger;
                ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
                begin
                  ql                                               := FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                  Mensagem.Text                                    := 'NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+' N.Carnê:'+strzero(carne,5);
                end;
                case CBIdentOcorr.ItemIndex of
                    1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                    2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                    3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                    4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                    5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                    6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                    7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                    8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                    9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                    10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                    11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                    12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                    13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                    14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                    15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                    16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                    17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
                end;
              end;

              if cbemail.Checked=true then
              begin
                 ZQAniversariante.close;
                 ZQAniversariante.SQL.clear;
                 ZQAniversariante.SQL.add('select * from participante where nome_parte='+quotedstr(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+' order by nome_parte');
                 ZQAniversariante.open;
                 if not empty(ZQAniversariante.FieldByName('email').AsString) then
                 begin
                   CDParticipante.Insert;
                   CDParticipanteidpaticipante.Value               :=ZQAniversariante.FieldByName('idpaticipante').AsLargeInt;
                   CDParticipantenome_parte.Value                  :=ZQAniversariante.FieldByName('nome_parte').AsString;
                   CDParticipanteemail.Value                       :=ZQAniversariante.FieldByName('email').AsString;
                   CDParticipanteinativo.Value                     :=ZQAniversariante.FieldByName('Inativo').AsString;
                   CDParticipantenascimento.Value                  :=ZQAniversariante.FieldByName('aniversario').AsDateTime;
                   CDParticipanteProfissao.Value                   :=ZQAniversariante.FieldByName('profissao').AsString;
                   CDParticipante.post;
                end;
                ZQAniversariante.close;
              end;
            end;
          end;
        end;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;

  end;


    //Emissão dos Boletos
  if CBBancoemite.Checked=false then
  begin
    ACBrBoleto1.Imprimir;
  end;

  FrmRelRecebimento.CDS_MarcaTit.EnableControls;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  DM_Tabelas.ZQContaBancaria.edit;
  DM_Tabelas.ZQContaBancaria.FieldByName('n_sequencial').AsLargeInt:=varremes;
  if rgnosso.Checked then
     DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt:=xnosso;
  DM_Tabelas.ZQContaBancaria.post;
  DM_Tabelas.ZQContaBancaria.ApplyUpdates;
  DM_Tabelas.ZQContaBancaria.Refresh;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

  Gauge1.Visible      := False;
//  if cbemail.Checked=false then
  //   RLBTitulo2.PreviewModal;

  // 14/12/2012
  if  rgnosso.Checked then
  begin
    Gauge1.Visible    := True;
    Gauge1.Progress   :=0;

    //17/09/2013
    ZQRemesRec.close;
    ZQRemesRec.SQL.Clear;
    ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where remessa='+quotedstr(inttostr(numeroremessa)));
    ZQRemesRec.open;



    ZQRemesRec.First;
    Gauge1.MaxValue   := ZQRemesRec.RecordCount;
//    DM_Tabelas.ZQRemes_Receb_atualiza.open;
  //  DM_Tabelas.ZQRemes_Receb_atualiza.first;
    ZQRemesRec.DisableControls;
    while not ZQRemesRec.Eof do
    begin
      Gauge1.Progress := ZQRemesRec.RecNo;

      DM_Tabelas.ZQRemes_Receb_atualiza.Close;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text));
      DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;



      ZQRemesRec.Next;
    end;
    ZQRemesRec.EnableControls;
//    DM_Tabelas.ZQRemes_Receb_atualiza.refresh;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRecebimento.refresh;
  end;
  Gauge1.Progress     :=0;
  Gauge1.Visible      := false;
  varremes            :=0;
  xnosso              :=0;

end;


procedure TFrmCobrancaBancaria.santander_240;
Var
  F : TextFile;
  ArqIni : tIniFile;
  VARARQ, Vararqnome, Varmens1, Varmens2, Varmens3, Varmens4, varnossonumero,
  varnossnum, varmora, vardoc, varcpfcnpj, varendereco, varhoje, VarNomepasta,
  varhostname, vardir, varpastabanco,VarPath, ext,VarDif,ql : string;
  vpar,NN,sq,totlin,xnosso,tm,numeroremessa,codcli,eita,carne,tam,tam2,varregistro, varremes, varx, vary : integer;
  data:Tdatetime;
  totvr:double;
  Titulo : TACBrTitulo;
  I : TACBrBolLayOut ;
begin
  totlin:=0;
  toTvr:=0;
  carne:=0;
  codcli:=0;
  tam:=0;
  VarDif:='';
  ql:='';
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  VarNomepasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );

//  varpastabanco := FrmRelRecebimento.CDSEmpreendimentoapelido.Value+'_'+FrmRelRecebimento.CDSEmpreendimentonomebanco.Value+FrmRelRecebimento.CDSEmpreendimentoconta.Value;
  if cb2via.Checked=false then
  begin

    vardir := '\\'+varhostname+'\'+VarNomePasta;
    // aqui gera a remessa
    if copy(CBIdentOcorr.Text,1,2)<>'00' then
    begin
      SaveDialog1.Execute;
      tam2:=length(SaveDialog1.FileName);
      tam2:=tam2-11;
      VarPath := copy(SaveDialog1.FileName,tam2+1,11);
      if trim(VarPath)<>'remessa.rem' then
      begin
        mensagem('O nome do Arquivo tem que ser remessa.rem');
        exit;
      end;

      VarPath := copy(SaveDialog1.FileName,1,tam2);
  //  if not DirectoryExists(vardir) then
  //    vardir := VarPath+varpastabanco;
      vardir := VarPath;
  //  else
  //    vardir := vardir +'\'+varpastabanco;
      if empty(trim(VarPath)) then
      begin
        mensagem('O nome do Arquivo não pode ficar em branco, tem que ser remessa.rem');
        exit;
      end;


      DBGerar.Enabled:=false;

      //CreateDir(vardir);

  //  vardir := vardir +'\'+copy(datetostr(date),7,4);
  //  CreateDir(vardir);
  //  vardir := vardir +'\'+copy(datetostr(date),4,2);
  //  CreateDir(vardir);
      vardir := vardir +'\'+copy(datetostr(date),1,2);
      CreateDir(vardir);
      Ext := '.REM';
      if CBTipoArq.Checked Then
        Ext := '.TST';

      varregistro:=2;

      for  varremes:= 1 to 99 do begin
        vararq := inttostrZero(varremes,2);
        Vararqnome := 'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+Ext;
        VARARQ :=vardir+'\'+Vararqnome;
        If not(fileexists(VARARQ)) then
          Break;
      end;
      Varnomearq := VARARQ;
      if DM_Tabelas.ZQBol_men.State in [DsEdit, DsInsert] Then
        DM_Tabelas.ZQBol_men.Edit;


      //pegar o ultimo registro auto incremento criado

      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

      DM_tabelas.ZQsequencia_remessa.close;
      DM_tabelas.ZQsequencia_remessa.SQL.Clear;
      DM_tabelas.ZQsequencia_remessa.SQL.Add('select * from remessa_sequencia where idconta_bancaria='+quotedstr(FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Text));
      DM_tabelas.ZQsequencia_remessa.Open;

      if DM_tabelas.ZQsequencia_remessa.RecordCount>0 then
      begin
        DM_tabelas.ZQsequencia_remessa.edit;
        DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt:=DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt+1;
        DM_tabelas.ZQsequencia_remessa.Post;
        DM_tabelas.ZQsequencia_remessa.Refresh;
      end;
      varremes                                                     := DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt;//DM_Tabelas.ZQContaBancarian_sequencial.Value+1;
      DM_tabelas.ZQsequencia_remessa.close;

      xnosso                                                       := DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;

      Frm_NumRemessa.XNENumRemessa.Value                           := varremes;
      Frm_NumRemessa.Label2.Caption                                := VARARQ;
      Frm_NumRemessa.Top                                           := FrmCobrancaBancaria.Top + 30;
      Frm_NumRemessa.Left                                          := FrmCobrancaBancaria.Left + 300;
      AbrirModal(Self, Frm_NumRemessa);
      //numero da remessa
      varremes                                                     := strtoint(floattostr(Frm_NumRemessa.XNENumRemessa.value));


      DM_Tabelas.ZQBancRemes.Insert;
      DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger                          := FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value;
      DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger                          := varremes;
      DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString                          := Vararqnome;
      DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime                           := date;
      DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString                       := FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
      DM_Tabelas.ZQBancRemes.Post;

      varhoje                                                      := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);



      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible                                               := True;
      Gauge1.Progress                                              :=0;
      Gauge1.MaxValue                                              := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  //    xnosso                                                     :=DM_Tabelas.ZQContaBancarianosso_numero.Value;

      ACBrBoleto1.DirArqRemessa                                    := VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
      ACBrBoleto1.LayoutRemessa                                    :=c240;
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.Banco.TipoCobranca                               :=cobSantander;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum                      :=12;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=40;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=40;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=30;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=30;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;

      ACBrBoletoFCFortes1.DirLogo                                  :=ExtractFilePath( Application.ExeName)+'Logos_bco';

      ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

      if zqcarne.active=false then
         zqcarne.open;
      zqcarne.last;
      carne                                                        :=zqcarne.FieldByName('n_seq').AsInteger;
      inc(carne);

      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge1.Progress                                            :=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
        begin
          if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0' Then
          begin
            if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            begin
               mensagem('O boleto '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente');
            end
            else
            begin

              //Dados do Cedente
              DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

              ACBrBoleto1.DataArquivo                              := Date;
              ACBrBoleto1.Banco.Numero                             := strtoint(trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString));
              ACBrBoleto1.Cedente.Agencia                          :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString)<>emptystr then
                 ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
              //ACBrBoleto1.Cedente.AgenciaDigito                  :=trim(DM_Tabelas.ZQContaBancarian_agencia_v.value);
              ACBrBoleto1.Cedente.Conta                            := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
              ACBrBoleto1.Cedente.ContaDigito                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
              ACBrBoleto1.Cedente.CodigoCedente                    := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.CodigoTransmissao                := trim(DM_tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString);
              case ComboBox3.ItemIndex of
                   0: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lPadrao);
                   1: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lCarne);
              end;

              ACBrBoleto1.Cedente.Convenio                         := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.TipoDocumento                    := escritural;
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
                 ACBrBoleto1.Cedente.Modalidade                    :='101'
              else
                 ACBrBoleto1.Cedente.Modalidade                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);  {Tipo de Modalidade Carteira
                                                                                                      101-Cobrança Rápida COM Registro  cliente emite os boletos
                                                                                                      104-Cobrança Eletrônica COM Registro bco emite os boletos}
              if CBBancoemite.checked then
                 ACBrBoleto1.Cedente.ResponEmissao := tbBancoEmite
              else
                ACBrBoleto1.Cedente.ResponEmissao := tbCliEmite;
              ACBrBoleto1.Cedente.CaracTitulo                      := tcSimplesRapComReg; //(tcSimples, tcVinculada, tcCaucionada, tcDescontada, tcVendor, tcDireta, tcSimplesRapComReg, tcCaucionadaRapComReg, tcDiretaEspecial)
              ACBrBoleto1.Cedente.TipoCarteira                     := tctRegistrada;//tctSimples
              ACBrBoleto1.Cedente.Nome                             := trim(DM_tabelas.ZQContaBancaria.FieldByName('nometitular').AsString);
              if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pJuridica
              else  if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))=11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pFisica
              else
              begin
                mensagem('Falta o documento (CNPJ/CPF) do cedente');
                exit;
              end;
              ACBrBoleto1.Cedente.CNPJCPF                          := tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString);

              try
                zqaux:= tzquery.Create(nil);
                zqaux.Connection := DM_Tabelas.zconeccao;
                zqaux.Close;
                zqaux.sql.clear;
                zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
                zqaux.Open;
                if zqaux.RecordCount>0 then
                begin
                    ACBrBoleto1.Cedente.Logradouro        := Removeacentos(uppercase(trim(zqaux.FieldByName('ende_cob').AsString)));
                  //  ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                    ACBrBoleto1.Cedente.Complemento       := '';
                    ACBrBoleto1.Cedente.Bairro            := Removeacentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                    ACBrBoleto1.Cedente.Cidade            := Removeacentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                    ACBrBoleto1.Cedente.UF                := Removeacentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                    ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
                end;
              finally
                zqaux.Close;
                zqaux.free;
              end;



              Titulo                                               := ACBrBoleto1.CriarTituloNaLista;
              with Titulo do
              begin
                //dados do sacado
                Sacado.NomeSacado                                  := Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
                if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pJuridica;;
                end
                else
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pFisica;
                end;
                Sacado.Logradouro                                  := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
                Sacado.CEP                                         := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
                Sacado.Cidade                                      := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
                Sacado.UF                                          := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
                Sacado.Bairro                                      := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);

                //dados da cobranca

                ValorAbatimento                                    := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
                LocalPagamento                                     := 'PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO';

                CodigoMoraJuros                                    := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
                CodigoMora                                         := '1'; //1-por dia 2-mensal 3-isento
                CodigoMulta                                        := cmPercentual;//(cmValorFixo, cmPercentual)
                DataMulta                                          := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

                if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                   PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
                else
                   PercentualMulta                                 := 0;

       //            MultaValorFixo                                := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
                MultaValorFixo                                     := False;

      //            DataProtesto                                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;
               DataProtesto                                        := 0;


                eita                                               :=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
                if eita=1 then
                   codcli                                          :=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
                if eita=2 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
                if eita=3 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
                if eita=4 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
                if eita>=5 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
                eita                                               :=codcli;
                codcli                                             :=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                vardoc                                             := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(25),1,25);

                inc(xnosso);

                varnossonumero                                     := inttostrZero(xnosso,12);

               //       vary                                       := strtoint(Modulo11(varnossonumero,9));

//                varnossnum                                       :='234567892345';
//                vary                                             := 0;
//                for varx                                         := 1 to 12 do
//                    vary                                         := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(varnossonumero,13-varx,1)));
//                vary                                             := (vary mod 11);
//                if vary = 10 Then
//                begin
//                  varnossnum                                     := inttostrZero(xnosso,7)+'1';
//                end
//                else if vary = 1 Then
//                begin
//                 varnossnum                                      := inttostrZero(xnosso,7)+'0';
//                end
//                else if vary = 0 Then
//                begin
//                  varnossnum                                     := inttostrZero(xnosso,7)+'0';
//                end
//                else
//                begin
//                  vary                                           := 11 - vary ;
//                  varnossnum                                     := inttostrZero(xnosso,7)+inttostr(vary);
//                end;
//
//                varnossonumero                                   := varnossnum;

               // 14/12/2012
               if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
                  varnossonumero                                   :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;
               //     varnossonumero                               :=inttostr(strtoint(copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,1,7)));


                codcli                                             :=0;
                eita                                               :=0;


                 ///////////////  28/11/2022
                ZQRemesRec.close;
                ZQRemesRec.SQL.Clear;
                ZQRemesRec.SQL.Add('Select  idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
                ZQRemesRec.open;
        {        ZQRemesRec.Filtered                               :=false;
                ZQRemesRec.Filter                                  :='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                ZQRemesRec.Filtered                                :=true;}
                if ZQRemesRec.RecordCount>0 then
                begin
                  ZQRemesRec.First;
                  ZQRemesRec.DisableControls;
                  while not ZQRemesRec.Eof do
                  begin
                    ZQRemesRec.delete;
                  end;
                  ZQRemesRec.EnableControls;
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      := varnossonumero;
                  ZQRemesRec.Post;
                end
                else
                begin
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      := varnossonumero;
                  ZQRemesRec.Post;
                end;
                ZQRemesRec.Filtered                                :=false;

                //02/08/2022
                FrmRelRecebimento.CDS_MarcaTit.edit;
                FrmRelRecebimento.CDS_MarcaTitnumboleto.Value      :=varnossonumero;
                FrmRelRecebimento.CDS_MarcaTit.Post;
                //
                VarDif                                             :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


                // 17/09/2013

                numeroremessa                                      :=ZQRemesRec.FieldByName('remessa').AsInteger;

            /////////////////////////////////////////////////////////////////////


                NossoNumero                                        := copy(varnossonumero,1,12);

                ValorDocumento                                     := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
                if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
                begin
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   inc(vpar);
                   TotalParcelas     := strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   dec(vpar);
                   parcela           :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
                end
                else
                begin
                   TotalParcelas     :=  10;
                   parcela           :=  1;
                end;


                DataDocumento                                      := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
                if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
                   ValorMoraJuros                                  := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
                else
                  ValorMoraJuros                                   := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
                DataMoraJuros                                      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
                Vencimento                                         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;



                NumeroDocumento                                    := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);

                SeuNumero                                          := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
                EspecieMod                                         := '$';
                Aceite                                             := atNao; //atSim
                DataProcessamento                                  := Now;
                Carteira                                           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);//copy(trim(DM_Tabelas.ZQContaBancariacarteira.Value),3,1);  // verificar se é só um digito mesmo

                if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
                begin
    //              RLBTitulo2.TipoOcorrencia                      := toRemessaConcederDesconto;
                  CodigoDesconto                                   := cdValorFixo; //(cdSemDesconto, cdValorFixo);
                  TipoDesconto                                     := tdValorFixoAteDataInformada;
                  DataDesconto                                     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
                  ValorDesconto                                    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
                end;
               // tony pediu para trocar a mensagem no dia 28/11/2012
    //            RLBTitulo2.Instrucoes.Text                       := trim('NºO RECEBER APÓS 30 DIAS DE VENCIMENTO.Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            RLBTitulo2.Instrucoes.Text                       := trim('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100),2) >0 then
    //               RLBTitulo2.Instrucoes.Text                    := 'Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'))+#13+
    //                                             ' COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
    //            else

                ql                                                 := FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
                begin
                  ql                                               := FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                  Mensagem.Text                                    := 'NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+' N.Carnê:'+strzero(carne,5);
                  inc(carne);
                end;
                case CBIdentOcorr.ItemIndex of
                    1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                    2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                    3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                    4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                    5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                    6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                    7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                    8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                    9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                    10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                    11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                    12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                    13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                    14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                    15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                    16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                    17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
                end;
              end;

              if cbemail.Checked=true then
              begin
                 ZQAniversariante.close;
                 ZQAniversariante.SQL.clear;
                 ZQAniversariante.SQL.add('select * from participante where nome_parte='+quotedstr(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+' order by nome_parte');
                 ZQAniversariante.open;
                 if not empty(ZQAniversariante.FieldByName('email').AsString) then
                 begin
                   CDParticipante.Insert;
                   CDParticipanteidpaticipante.Value               :=ZQAniversariante.FieldByName('idpaticipante').AsLargeInt;
                   CDParticipantenome_parte.Value                  :=ZQAniversariante.FieldByName('nome_parte').AsString;
                   CDParticipanteemail.Value                       :=ZQAniversariante.FieldByName('email').AsString;
                   CDParticipanteinativo.Value                     :=ZQAniversariante.FieldByName('Inativo').AsString;
                   CDParticipantenascimento.Value                  :=ZQAniversariante.FieldByName('aniversario').AsDateTime;
                   CDParticipanteProfissao.Value                   :=ZQAniversariante.FieldByName('profissao').AsString;
                   CDParticipante.post;
                end;
                ZQAniversariante.close;
              end;
            end;
          end;
        end;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;
      ACBrBoleto1.GerarRemessa( varremes );
      zqcarne.Edit;
      zqcarne.FieldByName('n_seq').AsInteger                                           :=carne;
      zqcarne.post;
      zqcarne.Close;
      mensagem('Relatório de Boletos gerado com sucesso!!!'+chr(13)+'O relatório foi gravado em '+VARARQ);
    end;
  end
  else
  begin
      varhoje                                                      := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);



      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible                                               := True;
      Gauge1.Progress                                              :=0;
      Gauge1.MaxValue                                              := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  //    xnosso                                                     :=DM_Tabelas.ZQContaBancarianosso_numero.Value;

      ACBrBoleto1.DirArqRemessa                                    := VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
      ACBrBoleto1.LayoutRemessa                                    :=c240;
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.Banco.TipoCobranca                               :=cobSantander;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum                      :=12;

      ACBrBoletoFCFortes1.DirLogo                                  :=ExtractFilePath( Application.ExeName)+'Logos_bco';

      ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=40;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=40;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=30;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=30;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;

//      if zqcarne.active=false then
//         zqcarne.open;
//      zqcarne.last;
//      carne                                                        :=zqcarnen_seq.Value;
//      inc(carne);

      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge1.Progress                                            :=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
        begin
          if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0' Then
          begin
            if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            begin
               mensagem('O boleto '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente');
            end
            else
            begin

              //Dados do Cedente
              DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

              ACBrBoleto1.DataArquivo                              := Date;
              ACBrBoleto1.Banco.Numero                             := strtoint(trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString));
              ACBrBoleto1.Cedente.Agencia                          :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString)<>emptystr then
                 ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
              //ACBrBoleto1.Cedente.AgenciaDigito                  :=trim(DM_Tabelas.ZQContaBancarian_agencia_v.value);
              ACBrBoleto1.Cedente.Conta                            := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
              ACBrBoleto1.Cedente.ContaDigito                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
              ACBrBoleto1.Cedente.CodigoCedente                    := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.CodigoTransmissao                := trim(DM_tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString);
              case ComboBox3.ItemIndex of
                   0: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lPadrao);
                   1: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lCarne);
              end;

              ACBrBoleto1.Cedente.Convenio                         := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.TipoDocumento                    := escritural;
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
                 ACBrBoleto1.Cedente.Modalidade                    :='101'
              else
                 ACBrBoleto1.Cedente.Modalidade                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);  {Tipo de Modalidade Carteira
                                                                                                      101-Cobrança Rápida COM Registro  cliente emite os boletos
                                                                                                      104-Cobrança Eletrônica COM Registro bco emite os boletos}

              if CBBancoemite.checked then
                 ACBrBoleto1.Cedente.ResponEmissao := tbBancoEmite
              else
                ACBrBoleto1.Cedente.ResponEmissao := tbCliEmite;
              ACBrBoleto1.Cedente.CaracTitulo                      := tcSimplesRapComReg; //(tcSimples, tcVinculada, tcCaucionada, tcDescontada, tcVendor, tcDireta, tcSimplesRapComReg, tcCaucionadaRapComReg, tcDiretaEspecial)
              ACBrBoleto1.Cedente.TipoCarteira                     := tctRegistrada;//tctSimples
              ACBrBoleto1.Cedente.Nome                             := trim(DM_tabelas.ZQContaBancaria.FieldByName('nometitular').AsString);
              if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pJuridica
              else  if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))=11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pFisica
              else
              begin
                mensagem('Falta o documento (CNPJ/CPF) do cedente');
                exit;
              end;
              ACBrBoleto1.Cedente.CNPJCPF                          := tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString);

              try
                zqaux:= tzquery.Create(nil);
                zqaux.Connection := DM_Tabelas.zconeccao;
                zqaux.Close;
                zqaux.sql.clear;
                zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
                zqaux.Open;
                if zqaux.RecordCount>0 then
                begin
                    ACBrBoleto1.Cedente.Logradouro        := Removeacentos(uppercase(trim(zqaux.FieldByName('ende_cob').AsString)));
                  //  ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                    ACBrBoleto1.Cedente.Complemento       := '';
                    ACBrBoleto1.Cedente.Bairro            := Removeacentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                    ACBrBoleto1.Cedente.Cidade            := Removeacentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                    ACBrBoleto1.Cedente.UF                := Removeacentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                    ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
                end;
              finally
                zqaux.Close;
                zqaux.free;
              end;



              Titulo                                               := ACBrBoleto1.CriarTituloNaLista;
              with Titulo do
              begin
                //dados do sacado
                Sacado.NomeSacado                                  := Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
                if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pJuridica;;
                end
                else
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pFisica;
                end;
                Sacado.Logradouro                                  := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
                Sacado.CEP                                         := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
                Sacado.Cidade                                      := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
                Sacado.UF                                          := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
                Sacado.Bairro                                      := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);

                //dados da cobranca

                ValorAbatimento                                    := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
                LocalPagamento                                     := 'PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO';

                CodigoMoraJuros                                    := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
                CodigoMora                                         := '1'; //1-por dia 2-mensal 3-isento
                CodigoMulta                                        := cmPercentual;//(cmValorFixo, cmPercentual)
                DataMulta                                          := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

                if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                   PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
                else
                   PercentualMulta                                 := 0;

       //            MultaValorFixo                                := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
                MultaValorFixo                                     := False;

      //            DataProtesto                                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;
               DataProtesto                                        := 0;


                eita                                               :=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
                if eita=1 then
                   codcli                                          :=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
                if eita=2 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
                if eita=3 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
                if eita=4 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
                if eita>=5 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
                eita                                               :=codcli;
                codcli                                             :=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                vardoc                                             := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(25),1,25);

                xnosso                                             := strtoint(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value);

                varnossonumero                                     := inttostrZero(xnosso,12);

               //       vary                                       := strtoint(Modulo11(varnossonumero,9));

//                varnossnum                                       :='234567892345';
//                vary                                             := 0;
//                for varx                                         := 1 to 12 do
//                    vary                                         := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(varnossonumero,13-varx,1)));
//                vary                                             := (vary mod 11);
//                if vary = 10 Then
//                begin
//                  varnossnum                                     := inttostrZero(xnosso,7)+'1';
//                end
//                else if vary = 1 Then
//                begin
//                 varnossnum                                      := inttostrZero(xnosso,7)+'0';
//                end
//                else if vary = 0 Then
//                begin
//                  varnossnum                                     := inttostrZero(xnosso,7)+'0';
//                end
//                else
//                begin
//                  vary                                           := 11 - vary ;
//                  varnossnum                                     := inttostrZero(xnosso,7)+inttostr(vary);
//                end;
//
//                varnossonumero                                   := varnossnum;

               // 14/12/2012
               if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
                  varnossonumero                                   :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;
               //     varnossonumero                               :=inttostr(strtoint(copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,1,7)));


                codcli                                             :=0;
                eita                                               :=0;


                 ///////////////  28/11/2022
                ZQRemesRec.close;
                ZQRemesRec.SQL.Clear;
                ZQRemesRec.SQL.Add('Select  idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
                ZQRemesRec.open;
                ZQRemesRec.First;
        {        ZQRemesRec.Filtered                               :=false;
                ZQRemesRec.Filter                                  :='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                ZQRemesRec.Filtered                                :=true;}
                if ZQRemesRec.RecordCount=0 then
                begin
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      := varnossonumero;
                  ZQRemesRec.Post;
                end;

                numeroremessa                                      :=ZQRemesRec.FieldByName('remessa').AsInteger;
                ZQRemesRec.Filtered                                :=false;
                //02/08/2022
                FrmRelRecebimento.CDS_MarcaTit.edit;
                FrmRelRecebimento.CDS_MarcaTitnumboleto.Value      :=varnossonumero;
                FrmRelRecebimento.CDS_MarcaTit.Post;
                //
                VarDif                                             :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


            /////////////////////////////////////////////////////////////////////


                NossoNumero                                        := copy(varnossonumero,1,12);

                ValorDocumento                                     := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
                if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
                begin
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   inc(vpar);
                   TotalParcelas     := strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   dec(vpar);
                   parcela           :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
                end
                else
                begin
                   TotalParcelas     :=  10;
                   parcela           :=  1;
                end;


                DataDocumento                                      := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
                if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
                   ValorMoraJuros                                  := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
                else
                  ValorMoraJuros                                   := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
                DataMoraJuros                                      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
                Vencimento                                         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;



                NumeroDocumento                                    := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);

                SeuNumero                                          := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
                EspecieMod                                         := '$';
                Aceite                                             := atNao; //atSim
                DataProcessamento                                  := Now;
                Carteira                                           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);//copy(trim(DM_Tabelas.ZQContaBancariacarteira.Value),3,1);  // verificar se é só um digito mesmo

                if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
                begin
    //              RLBTitulo2.TipoOcorrencia                      := toRemessaConcederDesconto;
                  CodigoDesconto                                   := cdValorFixo; //(cdSemDesconto, cdValorFixo);
                  TipoDesconto                                     := tdValorFixoAteDataInformada;
                  DataDesconto                                     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
                  ValorDesconto                                    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
                end;
               // tony pediu para trocar a mensagem no dia 28/11/2012
    //            RLBTitulo2.Instrucoes.Text                       := trim('NºO RECEBER APÓS 30 DIAS DE VENCIMENTO.Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            RLBTitulo2.Instrucoes.Text                       := trim('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100),2) >0 then
    //               RLBTitulo2.Instrucoes.Text                    := 'Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'))+#13+
    //                                             ' COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
    //            else
                if zqcarne.active=false then
                   zqcarne.open;
                zqcarne.last;
                carne                                              :=zqcarne.FieldByName('n_seq').AsInteger;

                ql                                                 :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
                begin
                  ql                                               :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                  Mensagem.Text                                    := 'NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+' N.Carnê:'+strzero(carne,5);
                end;
                case CBIdentOcorr.ItemIndex of
                    1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                    2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                    3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                    4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                    5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                    6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                    7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                    8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                    9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                    10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                    11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                    12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                    13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                    14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                    15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                    16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                    17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
                end;
              end;

              if cbemail.Checked=true then
              begin
                 ZQAniversariante.close;
                 ZQAniversariante.SQL.clear;
                 ZQAniversariante.SQL.add('select * from participante where nome_parte='+quotedstr(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+' order by nome_parte');
                 ZQAniversariante.open;
                 if not empty(ZQAniversariante.FieldByName('email').AsString) then
                 begin
                   CDParticipante.Insert;
                   CDParticipanteidpaticipante.Value               :=ZQAniversariante.FieldByName('idpaticipante').AsLargeInt;
                   CDParticipantenome_parte.Value                  :=ZQAniversariante.FieldByName('nome_parte').AsString;
                   CDParticipanteemail.Value                       :=ZQAniversariante.FieldByName('email').AsString;
                   CDParticipanteinativo.Value                     :=ZQAniversariante.FieldByName('Inativo').AsString;
                   CDParticipantenascimento.Value                  :=ZQAniversariante.FieldByName('aniversario').AsDateTime;
                   CDParticipanteProfissao.Value                   :=ZQAniversariante.FieldByName('profissao').AsString;
                   CDParticipante.post;
                end;
                ZQAniversariante.close;
              end;
            end;
          end;
        end;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;

  end;


    //Emissão dos Boletos
  if CBBancoemite.Checked=false then
  begin
    ACBrBoleto1.Imprimir;
  end;

  FrmRelRecebimento.CDS_MarcaTit.EnableControls;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  DM_Tabelas.ZQContaBancaria.edit;
  DM_Tabelas.ZQContaBancaria.FieldByName('n_sequencial').AsLargeInt:=varremes;
  if rgnosso.Checked then
     DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt:=xnosso;
  DM_Tabelas.ZQContaBancaria.post;
  DM_Tabelas.ZQContaBancaria.ApplyUpdates;
  DM_Tabelas.ZQContaBancaria.Refresh;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

  Gauge1.Visible      := False;
//  if cbemail.Checked=false then
  //   RLBTitulo2.PreviewModal;

  // 14/12/2012
  if  rgnosso.Checked then
  begin
    Gauge1.Visible    := True;
    Gauge1.Progress   :=0;

    //17/09/2013
    ZQRemesRec.close;
    ZQRemesRec.SQL.Clear;
    ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where remessa='+quotedstr(inttostr(numeroremessa)));
    ZQRemesRec.open;



    ZQRemesRec.First;
    Gauge1.MaxValue   := ZQRemesRec.RecordCount;
//    DM_Tabelas.ZQRemes_Receb_atualiza.open;
  //  DM_Tabelas.ZQRemes_Receb_atualiza.first;
    ZQRemesRec.DisableControls;
    while not ZQRemesRec.Eof do
    begin
      Gauge1.Progress := ZQRemesRec.RecNo;

      DM_Tabelas.ZQRemes_Receb_atualiza.Close;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text));
      DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;



      ZQRemesRec.Next;
    end;
    ZQRemesRec.EnableControls;
//    DM_Tabelas.ZQRemes_Receb_atualiza.refresh;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRecebimento.refresh;
  end;
  Gauge1.Progress     :=0;
  Gauge1.Visible      := false;
  varremes            :=0;
  xnosso              :=0;

end;




procedure TFrmCobrancaBancaria.caixa_240;
Var
  F : TextFile;
  ArqIni : tIniFile;
  VARARQ, Vararqnome, Varmens1, Varmens2, Varmens3, Varmens4, varnossonumero,
  varnossnum, varmora, vardoc, varcpfcnpj, varendereco, varhoje, VarNomepasta,
  varhostname, vardir, varpastabanco,VarPath, ext,VarDif,ql : string;
  vpar,NN,sq,totlin,xnosso,tm,numeroremessa,codcli,eita,carne,tam,tam2,varregistro, varremes, varx, vary : integer;
  data:Tdatetime;
  totvr:double;
  Titulo : TACBrTitulo;
  I : TACBrBolLayOut ;
begin
  totlin:=0;
  toTvr:=0;
  carne:=0;
  codcli:=0;
  tam:=0;
  VarDif:='';
  ql:='';
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  VarNomepasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );

//  varpastabanco := FrmRelRecebimento.CDSEmpreendimentoapelido.Value+'_'+FrmRelRecebimento.CDSEmpreendimentonomebanco.Value+FrmRelRecebimento.CDSEmpreendimentoconta.Value;
  if cb2via.Checked=false then
  begin

    vardir := '\\'+varhostname+'\'+VarNomePasta;
    // aqui gera a remessa
    if copy(CBIdentOcorr.Text,1,2)<>'00' then
    begin
      SaveDialog1.Execute;
      tam2:=length(SaveDialog1.FileName);
      tam2:=tam2-11;
      VarPath := copy(SaveDialog1.FileName,tam2+1,11);
      if trim(VarPath)<>'remessa.rem' then
      begin
        mensagem('O nome do Arquivo tem que ser remessa.rem');
        exit;
      end;

      VarPath := copy(SaveDialog1.FileName,1,tam2);
  //  if not DirectoryExists(vardir) then
  //    vardir := VarPath+varpastabanco;
      vardir := VarPath;
  //  else
  //    vardir := vardir +'\'+varpastabanco;
      if empty(trim(VarPath)) then
      begin
        mensagem('O nome do Arquivo não pode ficar em branco, tem que ser remessa.rem');
        exit;
      end;


      DBGerar.Enabled:=false;

      //CreateDir(vardir);

  //  vardir := vardir +'\'+copy(datetostr(date),7,4);
  //  CreateDir(vardir);
  //  vardir := vardir +'\'+copy(datetostr(date),4,2);
  //  CreateDir(vardir);
      vardir := vardir +'\'+copy(datetostr(date),1,2);
      CreateDir(vardir);
      Ext := '.REM';
      if CBTipoArq.Checked Then
        Ext := '.TST';

      varregistro:=2;

      for  varremes:= 1 to 99 do begin
        vararq := inttostrZero(varremes,2);
        Vararqnome := 'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+Ext;
        VARARQ :=vardir+'\'+Vararqnome;
        If not(fileexists(VARARQ)) then
          Break;
      end;
      Varnomearq := VARARQ;
      if DM_Tabelas.ZQBol_men.State in [DsEdit, DsInsert] Then
        DM_Tabelas.ZQBol_men.Edit;


      //pegar o ultimo registro auto incremento criado

      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

      DM_tabelas.ZQsequencia_remessa.close;
      DM_tabelas.ZQsequencia_remessa.SQL.Clear;
      DM_tabelas.ZQsequencia_remessa.SQL.Add('select * from remessa_sequencia where idconta_bancaria='+quotedstr(FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Text));
      DM_tabelas.ZQsequencia_remessa.Open;

      if DM_tabelas.ZQsequencia_remessa.RecordCount>0 then
      begin
        DM_tabelas.ZQsequencia_remessa.edit;
        DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt:=DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt+1;
        DM_tabelas.ZQsequencia_remessa.Post;
        DM_tabelas.ZQsequencia_remessa.Refresh;
      end;
      varremes                                                     := DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt;//DM_Tabelas.ZQContaBancarian_sequencial.Value+1;
      DM_tabelas.ZQsequencia_remessa.close;

      xnosso                                                       := DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;

      Frm_NumRemessa.XNENumRemessa.Value                           := varremes;
      Frm_NumRemessa.Label2.Caption                                := VARARQ;
      Frm_NumRemessa.Top                                           := FrmCobrancaBancaria.Top + 30;
      Frm_NumRemessa.Left                                          := FrmCobrancaBancaria.Left + 300;
      AbrirModal(Self, Frm_NumRemessa);
      //numero da remessa
      varremes                                                     := strtoint(floattostr(Frm_NumRemessa.XNENumRemessa.value));


      DM_Tabelas.ZQBancRemes.Insert;
      DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger                          := FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value;
      DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger                          := varremes;
      DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString                          := Vararqnome;
      DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime                           := date;
      DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString                       := FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
      DM_Tabelas.ZQBancRemes.Post;

      varhoje                                                      := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);



      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible                                               := True;
      Gauge1.Progress                                              :=0;
      Gauge1.MaxValue                                              := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  //    xnosso                                                     :=DM_Tabelas.ZQContaBancarianosso_numero.Value;

      ACBrBoleto1.DirArqRemessa                                    := VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
      ACBrBoleto1.LayoutRemessa                                    :=c240;
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.Banco.TipoCobranca                               :=cobCaixaEconomica;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum                      :=15;
//      if trim(DM_Tabelas.ZQContaBancariacomplemento.Value)<>emptystr then
//      begin
//         ACBrBoleto1.Banco.LayoutVersaoArquivo                     :=trim(DM_Tabelas.ZQContaBancariacomplemento.Value).ToInteger;
//         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo  :=trim(DM_Tabelas.ZQContaBancariacomplemento.Value).ToInteger;
//      end
//      else
//      begin
//        ACBrBoleto1.Banco.LayoutVersaoArquivo                      :=107;
//        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo   :=107;
//      end;

      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=107;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=107;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=67;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=67;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;


//      if trim(DM_Tabelas.ZQContaBancarian_dif_empreed.Value)<>emptystr then
//      begin
//         ACBrBoleto1.Banco.LayoutVersaoLote                        :=trim(DM_Tabelas.ZQContaBancarian_dif_empreed.Value).ToInteger;
//         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote     :=trim(DM_Tabelas.ZQContaBancarian_dif_empreed.Value).ToInteger;
//      end
//      else
//      begin
//         ACBrBoleto1.Banco.LayoutVersaoLote                        :=67;
//         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote     :=67;
//      end;
      ACBrBoletoFCFortes1.ACBrBoleto.Banco.Numero                  :=104;
      ACBrBoletoFCFortes1.ACBrBoleto.Banco.TamanhoMaximoNossoNum   :=15;
      ACBrBoletoFCFortes1.DirLogo                                  :=ExtractFilePath( Application.ExeName)+'Logos_bco';

      ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

      if zqcarne.active=false then
         zqcarne.open;
      zqcarne.last;
      carne                                                        :=zqcarne.FieldByName('n_seq').AsInteger;
      inc(carne);

      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge1.Progress                                            :=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
        begin
          if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0' Then
          begin
            if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            begin
               mensagem('O boleto '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente');
            end
            else
            begin

              //Dados do Cedente
              DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

              ACBrBoleto1.DataArquivo                              := Date;
              ACBrBoleto1.Banco.Numero                             := strtoint(trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString));
               ACBrBoleto1.Cedente.Agencia                         :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
              if (CBBancoemite.Checked) then
                ACBrBoleto1.Cedente.ResponEmissao:=tbBancoEmite     // 2 Cliente Emite - 1 Banco Emite
              else
                ACBrBoleto1.Cedente.ResponEmissao:=tbCliEmite;
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString)<>emptystr then
                 ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
              //ACBrBoleto1.Cedente.AgenciaDigito                  :=trim(DM_Tabelas.ZQContaBancarian_agencia_v.value);
              ACBrBoleto1.Cedente.Conta                            := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
              ACBrBoleto1.Cedente.ContaDigito                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
              ACBrBoleto1.Cedente.CodigoCedente                    := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.CodigoTransmissao                := trim(DM_tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString);
              case ComboBox3.ItemIndex of
                   0: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lPadrao);
                   1: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lCarne);
              end;

              ACBrBoleto1.Cedente.Convenio                         := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.TipoDocumento                    := escritural;
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
                 ACBrBoleto1.Cedente.Modalidade                    :='101'
              else
                 ACBrBoleto1.Cedente.Modalidade                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);  {Tipo de Modalidade Carteira
                                                                                                      101-Cobrança Rápida COM Registro  cliente emite os boletos
                                                                                                      104-Cobrança Eletrônica COM Registro bco emite os boletos}

              if CBBancoemite.checked then
                 ACBrBoleto1.Cedente.ResponEmissao := tbBancoEmite
              else
                ACBrBoleto1.Cedente.ResponEmissao := tbCliEmite;
              ACBrBoleto1.Cedente.CaracTitulo                      := tcSimplesRapComReg; //(tcSimples, tcVinculada, tcCaucionada, tcDescontada, tcVendor, tcDireta, tcSimplesRapComReg, tcCaucionadaRapComReg, tcDiretaEspecial)
              ACBrBoleto1.Cedente.TipoCarteira                     := tctRegistrada;//tctSimples
              ACBrBoleto1.Cedente.Nome                             := RemoveAcentos(uppercase(trim(DM_tabelas.ZQContaBancaria.FieldByName('nometitular').AsString)));
              if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pJuridica
              else  if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))=11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pFisica
              else
              begin
                mensagem('Falta o documento (CNPJ/CPF) do cedente');
                exit;
              end;
              ACBrBoleto1.Cedente.CNPJCPF                          := tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString);

              try
                zqaux:= tzquery.Create(nil);
                zqaux.Connection := DM_Tabelas.zconeccao;
                zqaux.Close;
                zqaux.sql.clear;
                zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
                zqaux.Open;
                if zqaux.RecordCount>0 then
                begin
                    ACBrBoleto1.Cedente.Logradouro        := Removeacentos(uppercase(trim(zqaux.FieldByName('ende_cob').AsString)));
                  //  ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                    ACBrBoleto1.Cedente.Complemento       := '';
                    ACBrBoleto1.Cedente.Bairro            := Removeacentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                    ACBrBoleto1.Cedente.Cidade            := Removeacentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                    ACBrBoleto1.Cedente.UF                := Removeacentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                    ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
                end;
              finally
                zqaux.Close;
                zqaux.free;
              end;



              Titulo                                               := ACBrBoleto1.CriarTituloNaLista;
              with Titulo do
              begin
                //dados do sacado
                Sacado.NomeSacado                                  := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)));
                if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
                begin
                  Sacado.CNPJCPF                                   := trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pJuridica;;
                end
                else
                begin
                  Sacado.CNPJCPF                                   := trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pFisica;
                end;
                Sacado.Logradouro                                  := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value)));
                Sacado.CEP                                         := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
                Sacado.Cidade                                      := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value)));
                Sacado.UF                                          := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value)));
                Sacado.Bairro                                      := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value)));

                //dados da cobranca

                ValorAbatimento                                    := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
                LocalPagamento                                     := 'EM TODA A REDE BANCÁRIA E SEUS CORRESPONDENTES ATÉ O VALOR LIMITE';

                CodigoMoraJuros                                    := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
                CodigoMora                                         := '1'; //1-por dia 2-mensal 3-isento
                CodigoMulta                                        := cmPercentual;//(cmValorFixo, cmPercentual)
                DataMulta                                          := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

                if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                   PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
                else
                   PercentualMulta                                 := 0;

       //            MultaValorFixo                                := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
                MultaValorFixo                                     := False;

      //            DataProtesto                                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;
               DataProtesto                                        := 0;
               DataBaixa                                           := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+60;

                eita                                               :=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
                if eita=1 then
                   codcli                                          :=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
                if eita=2 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
                if eita=3 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
                if eita=4 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
                if eita>=5 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
                eita                                               :=codcli;
                codcli                                             :=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                vardoc                                             := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(25),1,25);

                inc(xnosso);

                varnossonumero                                     := inttostrZero(xnosso,15);

               //       vary                                       := strtoint(Modulo11(varnossonumero,9));

//                varnossnum                                       :='234567892345';
//                vary                                             := 0;
//                for varx                                         := 1 to 12 do
//                    vary                                         := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(varnossonumero,13-varx,1)));
//                vary                                             := (vary mod 11);
//                if vary = 10 Then
//                begin
//                  varnossnum                                     := inttostrZero(xnosso,7)+'1';
//                end
//                else if vary = 1 Then
//                begin
//                 varnossnum                                      := inttostrZero(xnosso,7)+'0';
//                end
//                else if vary = 0 Then
//                begin
//                  varnossnum                                     := inttostrZero(xnosso,7)+'0';
//                end
//                else
//                begin
//                  vary                                           := 11 - vary ;
//                  varnossnum                                     := inttostrZero(xnosso,7)+inttostr(vary);
//                end;
//
//                varnossonumero                                   := varnossnum;

               // 14/12/2012
               if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
                  varnossonumero                                   :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;
               //     varnossonumero                               :=inttostr(strtoint(copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,1,7)));


                codcli                                             :=0;
                eita                                               :=0;


                 ///////////////  28/11/2022
                ZQRemesRec.close;
                ZQRemesRec.SQL.Clear;
                ZQRemesRec.SQL.Add('Select  idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
                ZQRemesRec.open;
        {        ZQRemesRec.Filtered                               :=false;
                ZQRemesRec.Filter                                  :='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                ZQRemesRec.Filtered                                :=true;}
                if ZQRemesRec.RecordCount>0 then
                begin
                  ZQRemesRec.First;
                  ZQRemesRec.DisableControls;
                  while not ZQRemesRec.Eof do
                  begin
                    ZQRemesRec.delete;
                  end;
                  ZQRemesRec.EnableControls;
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      :=varnossonumero;
                  ZQRemesRec.Post;
                end
                else
                begin
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      :=varnossonumero;
                  ZQRemesRec.Post;
                end;
                ZQRemesRec.Filtered                                :=false;

                //02/08/2022
                FrmRelRecebimento.CDS_MarcaTit.edit;
                FrmRelRecebimento.CDS_MarcaTitnumboleto.Value      :=varnossonumero;
                FrmRelRecebimento.CDS_MarcaTit.Post;
                //
                VarDif                                             :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


                // 17/09/2013

                numeroremessa                                      :=ZQRemesRec.FieldByName('remessa').AsInteger;

            /////////////////////////////////////////////////////////////////////


                NossoNumero                                        := copy(varnossonumero,1,15);

                ValorDocumento                                     := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
                if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
                begin
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   inc(vpar);
                   TotalParcelas     := strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   dec(vpar);
                   parcela           :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
                end
                else
                begin
                   TotalParcelas     :=  10;
                   parcela           :=  1;
                end;


                DataDocumento                                      := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
                if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
                   ValorMoraJuros                                  := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
                else
                  ValorMoraJuros                                   := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
                DataMoraJuros                                      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
                Vencimento                                         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;



                NumeroDocumento                                    := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text);//+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);

                SeuNumero                                          := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
                EspecieMod                                         := '$';
                Aceite                                             := atNao; //atSim
                DataProcessamento                                  := Now;
                Carteira                                           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);//copy(trim(DM_Tabelas.ZQContaBancariacarteira.Value),3,1);  // verificar se é só um digito mesmo
                if CBCarteiraEnvio.Checked then
                   CarteiraEnvio     := tceBanco   // vai retornar '1'   tceCedente; // vai retornar '2'    coluna 93
                else
                  CarteiraEnvio     := tceCedente;       // coluna 93 cnab400
                if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
                begin
    //              RLBTitulo2.TipoOcorrencia                      := toRemessaConcederDesconto;
                  CodigoDesconto                                   := cdValorFixo; //(cdSemDesconto, cdValorFixo);
                  TipoDesconto                                     := tdValorFixoAteDataInformada;
                  DataDesconto                                     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
                  ValorDesconto                                    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
                end;
               // tony pediu para trocar a mensagem no dia 28/11/2012
    //            RLBTitulo2.Instrucoes.Text                       := trim('NºO RECEBER APÓS 30 DIAS DE VENCIMENTO.Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            RLBTitulo2.Instrucoes.Text                       := trim('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100),2) >0 then
    //               RLBTitulo2.Instrucoes.Text                    := 'Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'))+#13+
    //                                             ' COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
    //            else

                ql                                                 :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
                begin
                  ql                                               :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                  Mensagem.Text                                    := RemoveAcentos(uppercase('NAO RECEBER APOS 30 DIAS DO VENCIMENTO. PAGTO. REF.'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+' N.CARNE.'+strzero(carne,5)));
                  inc(carne);
                end;
                case CBIdentOcorr.ItemIndex of
                    1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                    2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                    3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                    4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                    5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                    6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                    7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                    8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                    9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                    10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                    11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                    12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                    13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                    14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                    15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                    16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                    17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
                end;
              end;

              if cbemail.Checked=true then
              begin
                 ZQAniversariante.close;
                 ZQAniversariante.SQL.clear;
                 ZQAniversariante.SQL.add('select * from participante where nome_parte='+quotedstr(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+' order by nome_parte');
                 ZQAniversariante.open;
                 if not empty(ZQAniversariante.FieldByName('email').AsString) then
                 begin
                   CDParticipante.Insert;
                   CDParticipanteidpaticipante.Value               :=ZQAniversariante.FieldByName('idpaticipante').AsLargeInt;
                   CDParticipantenome_parte.Value                  :=ZQAniversariante.FieldByName('nome_parte').AsString;
                   CDParticipanteemail.Value                       :=ZQAniversariante.FieldByName('email').AsString;
                   CDParticipanteinativo.Value                     :=ZQAniversariante.FieldByName('Inativo').AsString;
                   CDParticipantenascimento.Value                  :=ZQAniversariante.FieldByName('aniversario').AsDateTime;
                   CDParticipanteProfissao.Value                   :=ZQAniversariante.FieldByName('profissao').AsString;
                   CDParticipante.post;
                end;
                ZQAniversariante.close;
              end;
            end;
          end;
        end;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;
      ACBrBoleto1.GerarRemessa( varremes );
      zqcarne.Edit;
      zqcarne.FieldByName('n_seq').AsInteger                                           :=carne;
      zqcarne.post;
      zqcarne.Close;
      mensagem('Relatório de Boletos gerado com sucesso!!!'+chr(13)+'O relatório foi gravado em '+VARARQ);
    end;
  end
  else
  begin
      varhoje                                                      := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);



      FrmRelRecebimento.CDS_MarcaTit.DisableControls;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge1.Visible                                               := True;
      Gauge1.Progress                                              :=0;
      Gauge1.MaxValue                                              := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  //    xnosso                                                     :=DM_Tabelas.ZQContaBancarianosso_numero.Value;

      ACBrBoleto1.DirArqRemessa                                    := VARARQ; //ExtractFileDir(Application.ExeName)+'\REMESSAS\SICOOB';
      ACBrBoleto1.LayoutRemessa                                    :=c240;
      case CBIdentOcorr.ItemIndex of
          1  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaRegistrar);                       {Remessa}
          2  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaBaixar);                          {Pedido de Baixa}
          3  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestoFinsFalimentares);        {Pedido de Protesto Falimentar}
          4  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaConcederAbatimento);              {Concessão de Abatimento}
          5  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarAbatimento);              {Cancelamento de Abatimento concedido}
          6  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarVencimento);               {Alteração de vencimento}
          7  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarControleParticipante);     {Alteração do controle do participante}
          8  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAlterarNumeroControle);           {Alteração de seu número}
          9  : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaProtestar);                       {Pedido de protesto}
          10 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtestoBaixa);  {Sustar protesto e baixar}
          11 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarInstrucaoProtesto);       {Sustar protesto e manter na carteira}
          12 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransfCessaoCreditoIDProd10);     {Transferência Cessão crédito ID. Prod.10}
          13 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaTransferenciaCarteira);           {Transferência entre Carteiras}
          14 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaDevTransferenciaCarteira);        {Dev. Transferência entre Carteiras}
          15 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaOutrasOcorrencias);               {Alteração de Outros Dados}
          16 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaAcertarRateioCredito);            {Acerto nos dados do rateio de Crédito}
          17 : ACBrBoleto1.Banco.TipoOcorrenciaToCodRemessa(toRemessaCancelarRateioCredito);           {Cancelamento do rateio de crédito.}
      end;

      ACBrBoleto1.Banco.TipoCobranca                               :=cobCaixaEconomica;
      ACBrBoleto1.Banco.TamanhoMaximoNossoNum                      :=15;
//      if trim(DM_Tabelas.ZQContaBancariacomplemento.Value)<>emptystr then
//      begin
//         ACBrBoleto1.Banco.LayoutVersaoArquivo                     :=trim(DM_Tabelas.ZQContaBancariacomplemento.Value).ToInteger;
//         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo  :=trim(DM_Tabelas.ZQContaBancariacomplemento.Value).ToInteger;
//      end
//      else
//      begin
//        ACBrBoleto1.Banco.LayoutVersaoArquivo                      :=107;
//        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo   :=107;
//      end;
//      if trim(DM_Tabelas.ZQContaBancarian_dif_empreed.Value)<>emptystr then
//      begin
//         ACBrBoleto1.Banco.LayoutVersaoLote                        :=trim(DM_Tabelas.ZQContaBancarian_dif_empreed.Value).ToInteger;
//         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote     :=trim(DM_Tabelas.ZQContaBancarian_dif_empreed.Value).ToInteger;
//      end
//      else
//      begin
//         ACBrBoleto1.Banco.LayoutVersaoLote                        :=67;
//         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote     :=67;
//      end;

      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=107;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo:=107;
      end
      else
      begin
         ACBrBoleto1.Banco.LayoutVersaoArquivo:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoArquivo :=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoArquivo').AsInteger;
      end;
      if DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger=0 then
      begin
         ACBrBoleto1.Banco.LayoutVersaoLote:=67;
         ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=67;
      end
      else
      begin
        ACBrBoleto1.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
        ACBrBoletoFCFortes1.ACBrBoleto.Banco.LayoutVersaoLote:=DM_Tabelas.ZQContaBancaria.FieldByName('layoutversaoLote').AsInteger;
      end;
      ACBrBoletoFCFortes1.DirLogo                                  :=ExtractFilePath( Application.ExeName)+'Logos_bco';

      ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

//      if zqcarne.active=false then
//         zqcarne.open;
//      zqcarne.last;
//      carne                                                        :=zqcarnen_seq.Value;
//      inc(carne);

      ACBrBoleto1.ListadeBoletos.Clear;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge1.Progress                                            :=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
        begin
          if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0' Then
          begin
            if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
            begin
               mensagem('O boleto '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente');
            end
            else
            begin

              //Dados do Cedente
              DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

              ACBrBoleto1.DataArquivo                              := Date;
              ACBrBoleto1.Banco.Numero                             := strtoint(trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString));
            if (CBBancoemite.Checked) then
              ACBrBoleto1.Cedente.ResponEmissao:=tbBancoEmite     // 2 Cliente Emite - 1 Banco Emite
            else
              ACBrBoleto1.Cedente.ResponEmissao:=tbCliEmite;
              ACBrBoleto1.Cedente.Agencia                          :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString)<>emptystr then
                 ACBrBoleto1.Cedente.DigitoVerificadorAgenciaConta :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
              //ACBrBoleto1.Cedente.AgenciaDigito                  :=trim(DM_Tabelas.ZQContaBancarian_agencia_v.value);
              ACBrBoleto1.Cedente.Conta                            := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
              ACBrBoleto1.Cedente.ContaDigito                      := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
              ACBrBoleto1.Cedente.CodigoCedente                    := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.CodigoTransmissao                := trim(DM_tabelas.ZQContaBancaria.FieldByName('cod_transmissao').AsString);
              case ComboBox3.ItemIndex of
                   0: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lPadrao);
                   1: ACBrBoleto1.ACBrBoletoFC.LayOut              := TACBrBolLayOut(lCarne);
              end;

              ACBrBoleto1.Cedente.Convenio                         := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
              ACBrBoleto1.Cedente.TipoDocumento                    := escritural;
              if trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString)=Emptystr then
                 ACBrBoleto1.Cedente.Modalidade                    :='101'
              else
                 ACBrBoleto1.Cedente.Modalidade                    := trim(DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString);  {Tipo de Modalidade Carteira
                                                                                                      101-Cobrança Rápida COM Registro  cliente emite os boletos
                                                                                                      104-Cobrança Eletrônica COM Registro bco emite os boletos}

              if CBBancoemite.checked then
                 ACBrBoleto1.Cedente.ResponEmissao := tbBancoEmite
              else
                ACBrBoleto1.Cedente.ResponEmissao := tbCliEmite;
              ACBrBoleto1.Cedente.CaracTitulo                      := tcSimplesRapComReg; //(tcSimples, tcVinculada, tcCaucionada, tcDescontada, tcVendor, tcDireta, tcSimplesRapComReg, tcCaucionadaRapComReg, tcDiretaEspecial)
              ACBrBoleto1.Cedente.TipoCarteira                     := tctRegistrada;//tctSimples
              ACBrBoleto1.Cedente.Nome                             := RemoveAcentos(uppercase(trim(DM_tabelas.ZQContaBancaria.FieldByName('nometitular').AsString)));
              if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pJuridica
              else  if (Length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))=11 ) and (length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))>0 ) then
                 ACBrBoleto1.Cedente.TipoInscricao                 := pFisica
              else
              begin
                mensagem('Falta o documento (CNPJ/CPF) do cedente');
                exit;
              end;
              ACBrBoleto1.Cedente.CNPJCPF                          := tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString);

              try
                zqaux:= tzquery.Create(nil);
                zqaux.Connection := DM_Tabelas.zconeccao;
                zqaux.Close;
                zqaux.sql.clear;
                zqaux.SQL.Text := 'SELECT * FROM participante as p join cidade as c on p.cidade_cob=c.idcidade where p.idpaticipante='+quotedstr(FrmRelRecebimento.CDS_MarcaTitcliente.Text);
                zqaux.Open;
                if zqaux.RecordCount>0 then
                begin
                    ACBrBoleto1.Cedente.Logradouro        := RemoveAcentos(uppercase(trim(zqaux.FieldByName('ende_cob').AsString)));
                  //  ACBrBoleto1.Cedente.NumeroRes         := RemoveLetras(removeacentos(zqaux.FieldByName('ende_cob').AsString));
                    ACBrBoleto1.Cedente.Complemento       := '';
                    ACBrBoleto1.Cedente.Bairro            := RemoveAcentos(uppercase(trim(zqaux.FieldByName('bairro_cob').AsString)));
                    ACBrBoleto1.Cedente.Cidade            := RemoveAcentos(uppercase(trim(zqaux.FieldByName('nomecid').AsString)));
                    ACBrBoleto1.Cedente.UF                := RemoveAcentos(uppercase(trim(zqaux.FieldByName('estado').AsString)));
                    ACBrBoleto1.Cedente.CEP               := trim(zqaux.FieldByName('cep_cob').AsString);
                end;
              finally
                zqaux.Close;
                zqaux.free;
              end;



              Titulo                                               := ACBrBoleto1.CriarTituloNaLista;
              with Titulo do
              begin
                //dados do sacado
                Sacado.NomeSacado                                  := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)));
                if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pJuridica;;
                end
                else
                begin
                  Sacado.CNPJCPF                                   :=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
                  Sacado.Pessoa                                    := pFisica;
                end;
                Sacado.Logradouro                                  := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value)));
                Sacado.CEP                                         := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
                Sacado.Cidade                                      := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value)));
                Sacado.UF                                          := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value)));
                Sacado.Bairro                                      := RemoveAcentos(uppercase(Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value)));

                //dados da cobranca

                ValorAbatimento                                    := 0; //StrToCurrDef(edtValorAbatimento.Text,0);
                LocalPagamento                                     := 'EM TODA A REDE BANCÁRIA E SEUS CORRESPONDENTES ATÉ O VALOR LIMITE';

                CodigoMoraJuros                                    := cjValorDia; //(cjValorDia, cjTaxaMensal, cjIsento, cjValorMensal, cjTaxaDiaria);
                CodigoMora                                         := '1'; //1-por dia 2-mensal 3-isento
                CodigoMulta                                        := cmPercentual;//(cmValorFixo, cmPercentual)
                DataMulta                                          := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+1;

                if DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat > 0 then
                   PercentualMulta                                 := DM_Tabelas.ZQContaBancaria.FieldByName('jurosemboleto').AsFloat
                else
                   PercentualMulta                                 := 0;

       //            MultaValorFixo                                := 1; //1-Cobrar Multa Valor Fixo / 2-Percentual / 0-Não cobrar multa
                MultaValorFixo                                     := False;

      //            DataProtesto                                   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value+5;

               DataProtesto                                        := 0;
               DiasDeProtesto                                      := 0;

                eita                                               :=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
                if eita=1 then
                   codcli                                          :=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
                if eita=2 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
                if eita=3 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
                if eita=4 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
                if eita>=5 then
                   codcli                                          :=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
                eita                                               :=codcli;
                codcli                                             :=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                vardoc                                             := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(25),1,25);

                xnosso                                             := strtoint(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value);

                varnossonumero                                     := inttostrZero(xnosso,15);

               //       vary                                       := strtoint(Modulo11(varnossonumero,9));

//                varnossnum                                       :='234567892345';
//                vary                                             := 0;
//                for varx                                         := 1 to 12 do
//                    vary                                         := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(varnossonumero,13-varx,1)));
//                vary                                             := (vary mod 11);
//                if vary = 10 Then
//                begin
//                  varnossnum                                     := inttostrZero(xnosso,7)+'1';
//                end
//                else if vary = 1 Then
//                begin
//                 varnossnum                                      := inttostrZero(xnosso,7)+'0';
//                end
//                else if vary = 0 Then
//                begin
//                  varnossnum                                     := inttostrZero(xnosso,7)+'0';
//                end
//                else
//                begin
//                  vary                                           := 11 - vary ;
//                  varnossnum                                     := inttostrZero(xnosso,7)+inttostr(vary);
//                end;
//
//                varnossonumero                                   := varnossnum;

               // 14/12/2012
               if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
                  varnossonumero                                   :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;
               //     varnossonumero                               :=inttostr(strtoint(copy(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value,1,7)));


                codcli                                             :=0;
                eita                                               :=0;


                 ///////////////  28/11/2022
                ZQRemesRec.close;
                ZQRemesRec.SQL.Clear;
                ZQRemesRec.SQL.Add('Select  idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
                ZQRemesRec.open;
                ZQRemesRec.First;
        {        ZQRemesRec.Filtered                               :=false;
                ZQRemesRec.Filter                                  :='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
                ZQRemesRec.Filtered                                :=true;}
                if ZQRemesRec.RecordCount=0 then
                begin
                  ZQRemesRec.Filtered                              :=false;
                  ZQRemesRec.Insert;
                  ZQRemesRec.FieldByName('remessa').AsInteger                          := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
                  ZQRemesRec.FieldByName('idrec').AsInteger                            := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
                  ZQRemesRec.FieldByName('nossonumero').AsString                      := varnossonumero;
                  ZQRemesRec.Post;
                end;

                numeroremessa                                      :=ZQRemesRec.FieldByName('remessa').AsInteger;
                ZQRemesRec.Filtered                                :=false;
                //02/08/2022
                FrmRelRecebimento.CDS_MarcaTit.edit;
                FrmRelRecebimento.CDS_MarcaTitnumboleto.Value      :=varnossonumero;
                FrmRelRecebimento.CDS_MarcaTit.Post;
                //
                VarDif                                             :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


            /////////////////////////////////////////////////////////////////////


                NossoNumero                                        := copy(varnossonumero,1,15);

                ValorDocumento                                     := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
                if FrmRelRecebimento.CDS_MarcaTitordem.Value<>emptystr then
                begin
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   inc(vpar);
                   TotalParcelas     := strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,vpar,6)));
                   vpar              :=  pos('/',FrmRelRecebimento.CDS_MarcaTitordem.Value);
                   dec(vpar);
                   parcela           :=  strtoint(copy(FrmRelRecebimento.CDS_MarcaTitordem.Value,1,vpar));
                end
                else
                begin
                   TotalParcelas     :=  10;
                   parcela           :=  1;
                end;


                DataDocumento                                      := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
                if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
                   ValorMoraJuros                                  := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
                else
                  ValorMoraJuros                                   := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
                DataMoraJuros                                      := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
                Vencimento                                         := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;



                NumeroDocumento                                    := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text);//+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);

                SeuNumero                                          := FrmRelRecebimento.CDS_MarcaTitIdreceb.text;
                EspecieMod                                         := '$';
                Aceite                                             := atNao; //atSim
                DataProcessamento                                  := Now;
                Carteira                                           := trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString);//copy(trim(DM_Tabelas.ZQContaBancariacarteira.Value),3,1);  // verificar se é só um digito mesmo
                if CBCarteiraEnvio.Checked then
                   CarteiraEnvio     := tceBanco   // vai retornar '1'   tceCedente; // vai retornar '2'    coluna 93
                else
                  CarteiraEnvio     := tceCedente;       // coluna 93 cnab400
                if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
                begin
    //              RLBTitulo2.TipoOcorrencia                      := toRemessaConcederDesconto;
                  CodigoDesconto                                   := cdValorFixo; //(cdSemDesconto, cdValorFixo);
                  TipoDesconto                                     := tdValorFixoAteDataInformada;
                  DataDesconto                                     := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
                  ValorDesconto                                    := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
                end;
               // tony pediu para trocar a mensagem no dia 28/11/2012
    //            RLBTitulo2.Instrucoes.Text                       := trim('NºO RECEBER APÓS 30 DIAS DE VENCIMENTO.Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            RLBTitulo2.Instrucoes.Text                       := trim('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

    //            if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100),2) >0 then
    //               RLBTitulo2.Instrucoes.Text                    := 'Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'))+#13+
    //                                             ' COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
    //            else
                if zqcarne.active=false then
                   zqcarne.open;
                zqcarne.last;
                carne                                                        :=zqcarne.FieldByName('n_seq').AsInteger;
                ql                                                           :=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
                begin
                  ql                                               := FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
                  Mensagem.Text                                    := RemoveAcentos(uppercase('NAO RECEBER APOS 30 DIAS DO VENCIMENTO. PAGTO. REF.'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+' N.CARNE.'+strzero(carne,5)));
                end;
                case CBIdentOcorr.ItemIndex of
                    1  : OcorrenciaOriginal.Tipo := toRemessaRegistrar;                       {Remessa}
                    2  : OcorrenciaOriginal.Tipo := toRemessaBaixar;                          {Pedido de Baixa}
                    3  : OcorrenciaOriginal.Tipo := toRemessaProtestoFinsFalimentares;        {Pedido de Protesto Falimentar}
                    4  : OcorrenciaOriginal.Tipo := toRemessaConcederAbatimento;              {Concessão de Abatimento}
                    5  : OcorrenciaOriginal.Tipo := toRemessaCancelarAbatimento;              {Cancelamento de Abatimento concedido}
                    6  : OcorrenciaOriginal.Tipo := toRemessaAlterarVencimento;               {Alteração de vencimento}
                    7  : OcorrenciaOriginal.Tipo := toRemessaAlterarControleParticipante;     {Alteração do controle do participante}
                    8  : OcorrenciaOriginal.Tipo := toRemessaAlterarNumeroControle;           {Alteração de seu número}
                    9  : OcorrenciaOriginal.Tipo := toRemessaProtestar;                       {Pedido de protesto}
                    10 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtestoBaixa;  {Sustar protesto e baixar}
                    11 : OcorrenciaOriginal.Tipo := toRemessaCancelarInstrucaoProtesto;       {Sustar protesto e manter na carteira}
                    12 : OcorrenciaOriginal.Tipo := toRemessaTransfCessaoCreditoIDProd10;     {Transferência Cessão crédito ID. Prod.10}
                    13 : OcorrenciaOriginal.Tipo := toRemessaTransferenciaCarteira;           {Transferência entre Carteiras}
                    14 : OcorrenciaOriginal.Tipo := toRemessaDevTransferenciaCarteira;        {Dev. Transferência entre Carteiras}
                    15 : OcorrenciaOriginal.Tipo := toRemessaOutrasOcorrencias;               {Alteração de Outros Dados}
                    16 : OcorrenciaOriginal.Tipo := toRemessaAcertarRateioCredito;            {Acerto nos dados do rateio de Crédito}
                    17 : OcorrenciaOriginal.Tipo := toRemessaCancelarRateioCredito;           {Cancelamento do rateio de crédito.}
                end;
              end;

              if cbemail.Checked=true then
              begin
                 ZQAniversariante.close;
                 ZQAniversariante.SQL.clear;
                 ZQAniversariante.SQL.add('select * from participante where nome_parte='+quotedstr(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+' order by nome_parte');
                 ZQAniversariante.open;
                 if not empty(ZQAniversariante.FieldByName('email').AsString) then
                 begin
                   CDParticipante.Insert;
                   CDParticipanteidpaticipante.Value               :=ZQAniversariante.FieldByName('idpaticipante').AsLargeInt;
                   CDParticipantenome_parte.Value                  :=ZQAniversariante.FieldByName('nome_parte').AsString;
                   CDParticipanteemail.Value                       :=ZQAniversariante.FieldByName('email').AsString;
                   CDParticipanteinativo.Value                     :=ZQAniversariante.FieldByName('Inativo').AsString;
                   CDParticipantenascimento.Value                  :=ZQAniversariante.FieldByName('aniversario').AsDateTime;
                   CDParticipanteProfissao.Value                   :=ZQAniversariante.FieldByName('profissao').AsString;
                   CDParticipante.post;
                end;
                ZQAniversariante.close;
              end;
            end;
          end;
        end;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;

  end;


    //Emissão dos Boletos
  if CBBancoemite.Checked=false then
  begin
    ACBrBoleto1.Imprimir;
  end;

  FrmRelRecebimento.CDS_MarcaTit.EnableControls;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  DM_Tabelas.ZQContaBancaria.edit;
  DM_Tabelas.ZQContaBancaria.FieldByName('n_sequencial').AsLargeInt:=varremes;
  if rgnosso.Checked then
     DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt:=xnosso;
  DM_Tabelas.ZQContaBancaria.post;
  DM_Tabelas.ZQContaBancaria.ApplyUpdates;
  DM_Tabelas.ZQContaBancaria.Refresh;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

  Gauge1.Visible      := False;
//  if cbemail.Checked=false then
  //   RLBTitulo2.PreviewModal;

  // 14/12/2012
  if  rgnosso.Checked then
  begin
    Gauge1.Visible    := True;
    Gauge1.Progress   :=0;

    //17/09/2013
    ZQRemesRec.close;
    ZQRemesRec.SQL.Clear;
    ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where remessa='+quotedstr(inttostr(numeroremessa)));
    ZQRemesRec.open;



    ZQRemesRec.First;
    Gauge1.MaxValue   := ZQRemesRec.RecordCount;
//    DM_Tabelas.ZQRemes_Receb_atualiza.open;
  //  DM_Tabelas.ZQRemes_Receb_atualiza.first;
    ZQRemesRec.DisableControls;
    while not ZQRemesRec.Eof do
    begin
      Gauge1.Progress := ZQRemesRec.RecNo;

      DM_Tabelas.ZQRemes_Receb_atualiza.Close;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text));
      DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;



      ZQRemesRec.Next;
    end;
    ZQRemesRec.EnableControls;
//    DM_Tabelas.ZQRemes_Receb_atualiza.refresh;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRecebimento.refresh;
  end;
  Gauge1.Progress     :=0;
  Gauge1.Visible      := false;
  varremes            :=0;
  xnosso              :=0;

end;

procedure TFrmCobrancaBancaria.hsbc_400;
Var
  F : TextFile;
  ArqIni : tIniFile;
  VARARQ, Vararqnome, Varmens1, Varmens2, Varmens3, Varmens4, varnossonumero,
  varnossnum, varmora, vardoc, varcpfcnpj, varendereco, varhoje, VarNomepasta,
  varhostname, vardir, varpastabanco,VarPath, ext,VarDif,ql : string;
  xnosso,tm,numeroremessa,codcli,eita,carne,tam,tam2,varregistro, varremes, varx, vary : integer;
  data:Tdatetime;
begin
  carne:=0;
  codcli:=0;
  tam:=0;
  VarDif:='';
  ql:='';
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni := tIniFile.Create(varpath+'siai.Ini');
  VarNomepasta := ArqIni.ReadString('BANCO_DE_DADOS','pasta_servidor', VarNomePasta );
  varhostname := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', varhostname );

//  varpastabanco := FrmRelRecebimento.CDSEmpreendimentoapelido.Value+'_'+FrmRelRecebimento.CDSEmpreendimentonomebanco.Value+FrmRelRecebimento.CDSEmpreendimentoconta.Value;
  if cb2via.Checked=false then
  begin

  vardir := '\\'+varhostname+'\'+VarNomePasta;
  // aqui gera a remessa
  if copy(CBIdentOcorr.Text,1,2)<>'00' then
  begin
    SaveDialog1.Execute;
    tam2:=length(SaveDialog1.FileName);
    tam2:=tam2-11;
    VarPath := copy(SaveDialog1.FileName,tam2+1,11);
    if trim(VarPath)<>'remessa.rem' then
    begin
      mensagem('O nome do Arquivo tem que ser remessa.rem');
      exit;
    end;

    VarPath := copy(SaveDialog1.FileName,1,tam2);
//  if not DirectoryExists(vardir) then
//    vardir := VarPath+varpastabanco;
    vardir := VarPath;
//  else
//    vardir := vardir +'\'+varpastabanco;
    if empty(trim(VarPath)) then
    begin
      mensagem('O nome do Arquivo não pode ficar em branco, tem que ser remessa.rem');
      exit;
    end;


    DBGerar.Enabled:=false;

    //CreateDir(vardir);

//  vardir := vardir +'\'+copy(datetostr(date),7,4);
//  CreateDir(vardir);
//  vardir := vardir +'\'+copy(datetostr(date),4,2);
//  CreateDir(vardir);
    vardir := vardir +'\'+copy(datetostr(date),1,2);
    CreateDir(vardir);
    Ext := '.REM';
    if CBTipoArq.Checked Then
      Ext := '.TST';

    varregistro:=2;

    for  varremes:= 1 to 99 do begin
      vararq := inttostrZero(varremes,2);
      Vararqnome := 'CB'+copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+vararq+Ext;
      VARARQ :=vardir+'\'+Vararqnome;
      If not(fileexists(VARARQ)) then
        Break;
    end;
    Varnomearq := VARARQ;
    if DM_Tabelas.ZQBol_men.State in [DsEdit, DsInsert] Then
      DM_Tabelas.ZQBol_men.Edit;
    AssignFile(f,VARARQ);

    //pegar o ultimo registro auto incremento criado

    DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

    DM_tabelas.ZQsequencia_remessa.close;
    DM_tabelas.ZQsequencia_remessa.SQL.Clear;
    DM_tabelas.ZQsequencia_remessa.SQL.Add('select * from remessa_sequencia where idconta_bancaria='+quotedstr(FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Text));
    DM_tabelas.ZQsequencia_remessa.Open;
    if DM_tabelas.ZQsequencia_remessa.RecordCount>0 then
    begin
      DM_tabelas.ZQsequencia_remessa.edit;
      DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt:=DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt+1;
      DM_tabelas.ZQsequencia_remessa.Post;
      DM_tabelas.ZQsequencia_remessa.Refresh;
    end;
    varremes := DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt;//DM_Tabelas.ZQContaBancarian_sequencial.Value+1;
    DM_tabelas.ZQsequencia_remessa.close;

    xnosso   := DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;

    Frm_NumRemessa.XNENumRemessa.Value := varremes;
    Frm_NumRemessa.Label2.Caption := VARARQ;
    Frm_NumRemessa.Top := FrmCobrancaBancaria.Top + 30;
    Frm_NumRemessa.Left := FrmCobrancaBancaria.Left + 300;
    AbrirModal(Self, Frm_NumRemessa);
    //numero da remessa
    varremes := strtoint(floattostr(Frm_NumRemessa.XNENumRemessa.value));


    DM_Tabelas.ZQBancRemes.Insert;
    DM_Tabelas.ZQBancRemes.FieldByName('idbanco').AsInteger := FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value;
    DM_Tabelas.ZQBancRemes.FieldByName('remessa').AsInteger := varremes;
    DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := Vararqnome;
    DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := date;
    DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
    DM_Tabelas.ZQBancRemes.Post;


    varhoje := copy(datetostr(date),1,2)+copy(datetostr(date),4,2)+copy(datetostr(date),9,2);
    Rewrite(f); //abre o arquivo para escrita
    Write(f,'0'); // Identificação do Registro
    Write(f,'1'); // Identificação d arquivo remessa
    if CBTipoArq.Checked=false Then
       Write(f,'REMESSA')  // literal remessa
    else
       Write(f,'TESTE  ');  // literal remessa
    Write(f,'01'); // codigo de serviço
    Write(f,'COBRANCA       '); // literal do serviço
//    Write(f,space(7));
    Write(f,'0');
    Write(f,alltrim(strzero(strtoint(FrmRelRecebimento.CDSEmpreendimenton_agencia.value),4)));
    Write(f,'55');
//    Write(f,copy(FrmRelRecebimento.CDSEmpreendimenton_agencia_v.value,1,1));
    // 11 digitos aqui linha abaixo
//    Write(f,strzero(strtoint(FrmRelRecebimento.CDSEmpreendimenton_agencia.value+FrmRelRecebimento.CDSEmpreendimentoconta.value+FrmRelRecebimento.CDSEmpreendimentoconta_v.value),11));
    Write(f,FrmRelRecebimento.CDSEmpreendimenton_agencia.value+FrmRelRecebimento.CDSEmpreendimentoconta.value+FrmRelRecebimento.CDSEmpreendimentoconta_v.value);
//    Write(f,copy(FrmRelRecebimento.CDSEmpreendimentoconta_v.value,1,1));
    Write(f,space(2));
    Write(f,copy(FrmRelRecebimento.CDSEmpreendimentonometitular.Value+space(30),1,30)); //  razão social da empresa
    Write(f,'399'); // numero do brasil na camara da compensação
    Write(f,'HSBC           '); //  Nome do banco por extenso//
    Write(f,varhoje); // data de hoje
    Write(f,'01600'); // densidade de gravação
    Write(f,'BPI'); // Unidade de densidade de gravação
    Write(f,space(2));
    Write(f,'LANCV08'); // Unidade de densidade de gravação
    Write(f,space(277));
//    Write(f,inttostrZero(DM_Tabelas.ZQBancRemesremessa.Value,7));  //  nº sequencial de remessa
//    Write(f,espacos(22));  //  branco
  //  Write(f,FrmRelRecebimento.CDSEmpreendimentoconvenio.Value); // convenio 7 digitos
//    Write(f,espacos(258)); // branco
    Writeln(f,'000001');

    FrmRelRecebimento.CDS_MarcaTit.First;
    Gauge1.Visible := True;
    Gauge1.MaxValue := FrmRelRecebimento.CDS_MarcaTit.RecordCount;


    ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;

    if zqcarne.active=false then
       zqcarne.open;
    zqcarne.last;
    carne:=zqcarne.FieldByName('n_seq').AsInteger;
    inc(carne);

    FrmRelRecebimento.CDS_MarcaTit.DisableControls;
    Gauge2.Progress:=0;
//    Gauge2.Visible:=true;
  //  gauge2.MaxValue:=FrmRelRecebimento.CDS_MarcaTit.RecordCount;
    while not FrmRelRecebimento.CDS_MarcaTit.Eof do
    begin
    //  gauge2.Progress:=FrmRelRecebimento.CDS_MarcaTit.RecNo;
     if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
     begin
      if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '1' Then
      else
        if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
           mensagem('A parcela '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente')
      else
      Begin
        if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
          varmora := inttostrZero(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,13)
        else
          varmora := inttostrZero(strtofloat(charrem(alltrim(transform(ExRound((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2),'###,##0.00')))),13); // valor da multa
//          varmora := inttostrZero(strtofloat(charrem(alltrim(transform(truncar(DM_Tabelas.ZQContaBancariataxadiaria.Value,2),'###,##0.00')))),13);// precentual da multa

        varendereco := FrmRelRecebimento.CDS_MarcaTitende_cob.Value;

     // aqui hsbc

//      ZQRemesRec.Insert;
//      ZQRemesRecremessa.Value := DM_Tabelas.ZQBancRemesidbanco_remessa.Value;
//      ZQRemesRecidrec.Value := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
//      ZQRemesRecnossonumero.value :=varnossnum;
//      ZQRemesRec.Post;

       Gauge1.Progress := FrmRelRecebimento.CDS_MarcaTit.RecNo;


       eita:=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
       if eita=1 then
         codcli:=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
       if eita=2 then
           codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
       if eita=3 then
           codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
       if eita=4 then
           codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
       if eita>=5 then
           codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
       eita:=codcli;
       codcli:=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
       vardoc := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);

//        varnossonumero := DM_Tabelas.ZQContaBancariacarteira.Value+inttostrZero(codcli,11);
        inc(xnosso);
        varnossonumero := copy(FrmRelRecebimento.CDSEmpreendimentoconvenio.Value,1,5)+inttostrZero(xnosso,5);
         // qdo precisar trocar o numero do nosso numero //comenta as 2 linhs anteriores e liberar as 2 seguinte
//      vardoc := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);
//      varnossonumero := DM_Tabelas.ZQContaBancariacarteira.Value+inttostrZero( FrmRelRecebimento.CDS_MarcaTitIdreceb.Value,11);

{        varnossnum:='32765432765432';
        vary := 0;
        for varx := 1 to 14 do
          vary := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(varnossonumero,varx,1)));
        vary := (vary mod 11);
        vary := 11 - vary ;
        if vary = 10 Then
          varnossnum := inttostrZero(codcli,11)+'P'
        else if vary = 11
         Then
          varnossnum := inttostrZero(codcli,11)+'0'
        else
          varnossnum := inttostrZero(codcli,11)+inttostr(vary);
 }



        if CBBancoemite.Checked=false then
        begin

          varnossnum:='5432765432';
          vary := 0;
          for varx := 1 to 10 do
            vary := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(varnossonumero,varx,1)));
          vary := (vary mod 11);
          vary := 11 - vary ;
          if vary = 10 Then
            varnossnum := '0'
          else if vary = 11
           Then
            varnossnum := '0'
          else
            varnossnum := inttostr(vary);

         //varnossonumero  := varnossnum;
         varnossonumero := varnossonumero +varnossnum;




          // 14/12/2012
          if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
                varnossonumero :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;
        end
        else
        begin
          varnossonumero:='00000000000';
        end;
        codcli:=0;
        eita:=0;

        ZQRemesRec.close;
        ZQRemesRec.SQL.Clear;
        ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text));
        ZQRemesRec.open;
{        ZQRemesRec.Filtered:=false;
        ZQRemesRec.Filter:='idrec='+quotedstr(FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
        ZQRemesRec.Filtered:=true;}
        if ZQRemesRec.RecordCount>0 then
        begin
          ZQRemesRec.First;
          while not ZQRemesRec.Eof do
          begin
            ZQRemesRec.delete;
          end;
          ZQRemesRec.Filtered:=false;
          ZQRemesRec.Insert;
          ZQRemesRec.FieldByName('remessa').AsInteger := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
          ZQRemesRec.FieldByName('idrec').AsInteger := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
          ZQRemesRec.FieldByName('nossonumero').AsString :=varnossonumero;
          ZQRemesRec.Post;
        end
        else
        begin
          ZQRemesRec.Filtered:=false;
          ZQRemesRec.Insert;
          ZQRemesRec.FieldByName('remessa').AsInteger := DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').AsLargeInt;
          ZQRemesRec.FieldByName('idrec').AsInteger := FrmRelRecebimento.CDS_MarcaTitIdreceb.Value;
          ZQRemesRec.FieldByName('nossonumero').AsString :=varnossonumero;
          ZQRemesRec.Post;
        end;
        ZQRemesRec.Filtered:=false;
        VarDif :=FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value;


        // 17/09/2013

        numeroremessa:=ZQRemesRec.FieldByName('remessa').AsInteger;

        Write(f,'1'); // Identificação do Registro
        if length(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString))=11 Then
           Write(f,'01')   // tipo cpf
        else
           Write(f,'02');  // tipo cnpj
        varcpfcnpj := inttostrZero(strtofloat(tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString)),14);
        Write(f,varcpfcnpj);  // cpf/cnpj
        Write(f,'0');
        Write(f,strzero(strtoint(FrmRelRecebimento.CDSEmpreendimenton_agencia.value),4));  // codigo da agencia cedente, sem o digito
//        Write(f,copy(FrmRelRecebimento.CDSEmpreendimenton_agencia_v.value,1,1));

        Write(f,'55');
//        Write(f,strzero(strtoint(FrmRelRecebimento.CDSEmpreendimenton_agencia.value+FrmRelRecebimento.CDSEmpreendimentoconta.value+FrmRelRecebimento.CDSEmpreendimentoconta_v.value),11));
        Write(f,FrmRelRecebimento.CDSEmpreendimenton_agencia.value+FrmRelRecebimento.CDSEmpreendimentoconta.value+FrmRelRecebimento.CDSEmpreendimentoconta_v.value);
        Write(f,space(2));

//        Write(f,strzero(strtoint(FrmRelRecebimento.CDSEmpreendimentoconta.value),8));  // conta corrente
  //      Write(f,copy(FrmRelRecebimento.CDSEmpreendimentoconta_v.Value,1,1)); // digito da conta
    //    Write(f,FrmRelRecebimento.CDSEmpreendimentoconvenio.Value); // convenio 7 digitos

        tam:=length(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text);
        tam:=37+tam;
        tam:=49-tam;
         //carne

        if ql=FrmRelRecebimento.CDS_MarcaTitquadralote.Value then
        begin
          if empty(VarDif) then // coluna 38 até 62
          begin
            Write(f,copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,25));  // uso da empresa, nº de controle da empresa;
          end
          else
          begin
            Write(f,copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(7),1,24));  // uso da empresa, nº de controle da empresa;
            Write(f,VarDif);
          end;
        end
        else
        begin
          inc(carne);
          ql:=FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
          if empty(VarDif) then // coluna 38 até 62
          begin
            Write(f,copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(8),1,25));  // uso da empresa, nº de controle da empresa;
          end
          else
          begin
            Write(f,copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(tam)+strzero(carne,5)+espacos(7),1,24));  // uso da empresa, nº de controle da empresa;
            Write(f,VarDif);
          end;
        end;
        VarDif:='';
        if CBBancoemite.Checked=false then
           Write(f,varnossonumero) //Nº bancario para cobrana com e em registro    63 a 73
        else
           Write(f,'00000000000');//Nº bancario para cobrana com e em registro    63 a 73

        Write(f,'0000000000000000000000000000000000'); // da 74 a 107

//        Write(f,DM_Tabelas.ZQContaBancariavariacao.value);  // variação da carteira  108 a 108
        if not empty(FrmRelRecebimento.CDSEmpreendimentocarteira.Value) then
            Write(f,alltrim(COPY(inttostr(strtoint(FrmRelRecebimento.CDSEmpreendimentocarteira.Value)),1,1)))  // Código da carteira 108 a 108
        else
            Write(f,'1');  // Código da carteira 108 a 108
        Write(f,copy(CBIdentOcorr.Items.Strings[CBIdentOcorr.ItemIndex],1,2));  // codigo de ocorrencia 109 a 110
        Write(f,vardoc);  //  nº do documento // Seu Número/Número do Título Atribuído pelo Cedente 111 a 120
        Write(f,copy(FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Text,1,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Text,4,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Text,9,2));  // Vencimento 121 a 126
        Write(f,inttostrZero(ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2)*100,13));   //  Valor do titulo 127  a 139
        Write(f,'399');
        Write(f,'00000');
        Write(f,'09'); // 01-DP-duplicata mercantil 09-CE-Cobrança Escritural 148 a 149
        Write(f,'N'); // aceite 150 a 150
        Write(f,copy(FrmRelRecebimento.CDS_MarcaTitDt_Entrada.Text,1,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Entrada.Text,4,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Entrada.Text,9,2));  // emissao 151 a 156
        Write(f,'1500'); //  1ª e 2ª instrução 157 a 160
        Write(f,varmora); // Mora por dia de atraso   161 a 173
        if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100),2) >0then
        begin
          Write(f,copy(FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Text,1,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Text,4,2)+ copy(FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Text,9,2));  // Limite para desconto
          Write(f,inttostrZero(truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100),2)*100,13));  //  Valor do desconto
        end
        else
        begin
          Write(f,'000000');  // Limite para desconto
          Write(f,'0000000000000');  //  Valor do desconto
        end;
        Write(f,'0000000000000');  //  Valor IOF  193 a 205

        {if FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Value>0 then
        begin
           Write(f,'599'); //
           Write(f,'2'); // cobrar multa (1-valor  2-percentual  9-sem multa)
           data:=FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;

           Write(f,copy(datetostr(data),1,2)+ copy(datetostr(data),4,2)+ copy(datetostr(data),9,2));  // Data de Inicio da Cobrança da Multa
           Write(f,inttostrZero(truncar(FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Value*100,0),12));  // % de multa a ser considerado
           Write(f,espacos(372));
           Writeln(f,inttostrZero(varregistro,6)); //  nº seguencial 395 a 400
           inc(varregistro);
        end;
        }
        data:=FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
        Write(f,copy(datetostr(data),1,2)+ copy(datetostr(data),4,2)+ copy(datetostr(data),9,2));  // Inicio da multa 206 a 211
        Write(f,inttostrZero(strtofloat(charrem(alltrim(transform(truncar(FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Value,2),'#0.00')))),4));  // % de multa a ser considerado 212 a 215
        Write(f,space(3));

       // Write(f,'0000000000000');  //  Valor abatimento   ou multa  206 a 2018

        if length(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)=14 Then
          Write(f,'01')   // tipo cpf       219 a 220
        else
          Write(f,'02');  // tipo cnpj
        varcpfcnpj := inttostrZero(strtofloat(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)),14);
        Write(f,varcpfcnpj);  // cpf/cnpj   221 a 234
        Write(f,copy(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+espacos(40),1,40));  // nome do sacado  coluna 235  a 274
        Write(f,copy(varendereco+espacos(38),1,38));  // Endereco   275 a 312
        Write(f,space(2));  //313 a 314
        Write(f,copy(FrmRelRecebimento.CDS_MarcaTitbairro.Value+espacos(12),1,12));  // Endereco   315 a 326
        Write(f,copy(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value)+espacos(8),1,8));  // cep  //327 a 334
        Write(f,copy(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value+espacos(15),1,15));  //335 a 349
        Write(f,copy(FrmRelRecebimento.CDS_MarcaTituf.Value+espacos(2),1,2));  //350 a 351
        tam:=length(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
        tam:=abs(39-tam);
        Write(f,copy(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+espacos(tam),1,39));  //352 a 390
        Write(f,'S');  //391 a 391
        Write(f,space(2));  //392 a 393 prazo para protesto numero de dias
        Write(f,'9');  //394 a 394  tipo de moeda
        varmens1 := 'APOS VENCIMENTO, PAGAVEL NO BANCO BRADESCO';
        varmens2 := copy(' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+' '+DM_Tabelas.ZQBol_men.FieldByName('linha2').AsString,1,40);
        if truncar(((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100),2) >0 then
           varmens3 := copy(' Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2),'###,###,##0.00')),1,40)
        else
           varmens3 := copy(' '+DM_Tabelas.ZQBol_men.FieldByName('linha3').AsString,1,40);
        varmens4 := copy(' '+DM_Tabelas.ZQBol_men.FieldByName('linha4').AsString,1,40);

        // tony pediu para tirar o numero da parcela 25/11/2013
//        Write(f,COPY('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value+espacos(40),1,40));  // decomposição

        Write(f,COPY(varmens1+varmens2+varmens3+varmens4+espacos(40),1,40));  // decomposição
  //      Write(f,espacos(3)); //392 a 394
        //inc(varregistro);
        Writeln(f,inttostrZero(varregistro,6)); //  nº seguencial 395 a 400
        inc(varregistro);


      end;
      FrmRelRecebimento.CDS_MarcaTit.Next;
     end
     else
     begin
       FrmRelRecebimento.CDS_MarcaTit.Next;
     end;
    end;
    //gauge2.Progress:=0;
    Gauge2.Visible:=false;
    FrmRelRecebimento.CDS_MarcaTit.EnableControls;

    zqcarne.Edit;
    zqcarne.FieldByName('n_seq').AsInteger:=carne;
    zqcarne.post;
    zqcarne.Close;

//    inc(varregistro);
    Write(f,'9');  // identificação do registro
    Write(f,espacos(393));   // branco
    Writeln(f,inttostrZero(varregistro,6)); //  nº seguencial
    Closefile(f); //fecha o handle de arquivo

    mensagem('Relatório de Boletos gerado com sucesso!!!'+chr(13)+'O relatório foi gravado em '+VARARQ);
  end;
  end;

  //Emissão dos Boletos
  if CBBancoemite.Checked=false then
  begin
    FrmRelRecebimento.CDS_MarcaTit.First;
    Gauge1.Visible := True;
    Gauge1.Progress:=0;
    Gauge1.MaxValue := FrmRelRecebimento.CDS_MarcaTit.RecordCount;
    DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
    xnosso:=DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt;
    FrmRelRecebimento.CDS_MarcaTit.DisableControls;
    while not FrmRelRecebimento.CDS_MarcaTit.Eof do
    begin
      Gauge1.Progress:=FrmRelRecebimento.CDS_MarcaTit.RecNo;
      if FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value = FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value then
      begin
        if FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0' Then
        begin
          if empty(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value)) Then
          begin
             mensagem('O boleto '+FrmRelRecebimento.CDS_MarcaTitordem.Value+' do cliente '+FrmRelRecebimento.CDS_MarcaTitnome_parte.Value+chr(13)+'não será anexada a cobrança bancaria,'+chr(13)+'pois está com o cadastro incompleto da documentação do cliente');
          end
          else
          begin
            case ComboBox3.ItemIndex of
                 0: RLBTitulo2.BoletoLayout := blPadrao;
                 1: RLBTitulo2.BoletoLayout := blCarne;
            end;
            //Dados do Cedente
            DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
            RLBTitulo2.DataProcessamento := Date;
            RLBTitulo2.Cedente.ContaBancaria.Banco.Codigo := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString);
            RLBTitulo2.Cedente.ContaBancaria.CodigoAgencia :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString);
            RLBTitulo2.Cedente.ContaBancaria.DigitoAgencia :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia_v').AsString);
            RLBTitulo2.Cedente.ContaBancaria.NumeroConta := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString);
            RLBTitulo2.Cedente.ContaBancaria.DigitoConta := trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
            RLBTitulo2.Cedente.CodigoCedente :=trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString)+'-'+trim(DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString);
//            RLBTitulo2.Cedente.CodigoCedente :=trim(DM_tabelas.ZQContaBancarian_no_banco.Value);
//          RLBTitulo2.Cedente.DigitoCodigoCedente := '6';
            RLBTitulo2.Carteira := copy(trim(DM_Tabelas.ZQContaBancaria.FieldByName('carteira').AsString),2,2)+'/'+DM_Tabelas.ZQContaBancaria.FieldByName('variacao').AsString;
            RLBTitulo2.Cedente.ContaBancaria.Convenio := trim(DM_tabelas.ZQContaBancaria.FieldByName('convenio').AsString);
            RLBTitulo2.Cedente.ContaBancaria.NomeCliente := trim(DM_tabelas.ZQContaBancaria.FieldByName('nometitular').AsString);
            RLBTitulo2.Cedente.TipoInscricao := tiPessoaJuridica;
            RLBTitulo2.Cedente.NumeroCPFCGC := tiramascara(DM_tabelas.ZQContaBancaria.FieldByName('doc_titular').AsString);
            RLBTitulo2.Cedente.Nome := trim(FrmRelRecebimento.CDSEmpreendimentonometitular.Value);
            RLBTitulo2.Cedente.EnderecoCedente:=copy(trim(DM_Tabelas.ZQEmpresa.FieldByName('endereco').AsString)+space(1)+trim(DM_Tabelas.ZQEmpresa.FieldByName('bairro').AsString)+space(1)+trim(DM_Tabelas.ZQEmpresa.FieldByName('nomecidade').AsString)+'-'+trim(DM_Tabelas.ZQEmpresa.FieldByName('estado').AsString)+space(1)+trim(DM_Tabelas.ZQEmpresa.FieldByName('cep').AsString),1,80);
            //dados do sacado
            RLBTitulo2.Sacado.Nome := Trim(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value);
            if length(charrem(FrmRelRecebimento.CDS_MarcaTitdoc1.Value))=14 Then
            begin
              RLBTitulo2.Sacado.NumeroCPFCGC:=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
              RLBTitulo2.Sacado.TipoInscricao := tiPessoaJuridica;
            end
            else
            begin
              RLBTitulo2.Sacado.NumeroCPFCGC:=trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitdoc1.Value));
              RLBTitulo2.Sacado.TipoInscricao := tiPessoaFisica;
            end;
            RLBTitulo2.Sacado.Endereco.Rua := Trim(FrmRelRecebimento.CDS_MarcaTitende_cob.Value);
            RLBTitulo2.Sacado.Endereco.CEP := Trim(tiramascara(FrmRelRecebimento.CDS_MarcaTitcep_cob.Value));
            RLBTitulo2.Sacado.Endereco.Cidade := Trim(FrmRelRecebimento.CDS_MarcaTitnomecidade.Value);
            RLBTitulo2.Sacado.Endereco.Estado := Trim(FrmRelRecebimento.CDS_MarcaTitUF.Value);
            RLBTitulo2.Sacado.Endereco.Bairro := Trim(FrmRelRecebimento.CDS_MarcaTitbairro.value);

            //dados da cobranca

            eita:=length(FrmRelRecebimento.CDS_MarcaTitidparti.text);
            if eita=1 then
               codcli:=FrmRelRecebimento.CDS_MarcaTitidparti.Value;
            if eita=2 then
               codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1)));
            if eita=3 then
               codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,1))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,2,1))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,1)));
            if eita=4 then
               codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2)));
            if eita>=5 then
               codcli:=strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,1,2))+strtoint(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,3,2))+strtoint(trim(copy(FrmRelRecebimento.CDS_MarcaTitidparti.text,5,12)));
            eita:=codcli;
            codcli:=strtoint(inttostr(eita)+FrmRelRecebimento.CDS_MarcaTitIdreceb.text);
            vardoc := copy(inttostr(codcli)+FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);

            inc(xnosso);
            varnossonumero := FrmRelRecebimento.CDSEmpreendimentoconvenio.Value+inttostrZero(xnosso,10);
               // qdo precisar trocar o numero do nosso numero //comenta as 2 linhs anteriores e liberar as 2 seguinte
    //        vardoc := copy(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+espacos(10),1,10);
//            varnossonumero := DM_Tabelas.ZQContaBancariacarteira.Value+inttostrZero( FrmRelRecebimento.CDS_MarcaTitIdreceb.Value,11);

{            varnossnum:='32765432765432';
            vary := 0;
            for varx := 1 to 14 do
              vary := vary + (strtoint(copy(varnossnum,varx,1)) * strtoint(copy(varnossonumero,varx,1)));
            vary := (vary mod 11);
            vary := 11 - vary ;
            if vary = 10 Then
              varnossnum := inttostrZero(codcli,11)+'P'
            else if vary = 11 Then
              varnossnum := inttostrZero(codcli,11)+'0'
            else
            varnossnum := inttostrZero(codcli,11)+inttostr(vary);}

           // 14/12/2012
           if (not empty(FrmRelRecebimento.CDS_MarcaTitnumboleto.Value)) and (rgnosso.Checked=false) then
              varnossonumero :=FrmRelRecebimento.CDS_MarcaTitnumboleto.Value;


            codcli:=0;
            eita:=0;


            RLBTitulo2.NossoNumero := copy(varnossonumero,1,17);

            RLBTitulo2.ValorDocumento  := ExRound(FrmRelRecebimento.CDS_MarcaTitvalor.Value,2);
            RLBTitulo2.DataDocumento   := FrmRelRecebimento.CDS_MarcaTitDt_Entrada.value;
            if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then
               RLBTitulo2.ValorMoraJuros  := truncar(DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat*100,2)
            else
              RLBTitulo2.ValorMoraJuros  := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('taxadiaria').AsFloat)/100,2);
            RLBTitulo2.DataMoraJuros   := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value+1;
            RLBTitulo2.DataVencimento  := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.value;
            RLBTitulo2.NumeroDocumento := Trim(FrmRelRecebimento.CDS_MarcaTitIdreceb.Text+FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value);
            if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2)>0 then
            begin
//              RLBTitulo2.TipoOcorrencia  := toRemessaConcederDesconto;
              RLBTitulo2.DataDesconto    := FrmRelRecebimento.CDS_MarcaTitDt_Vencimento.Value;
              RLBTitulo2.ValorDesconto   := truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancaria.FieldByName('Perc_descontos').AsFloat)/100,2);
            end;
           // tony pediu para trocar a mensagem no dia 28/11/2012
//            RLBTitulo2.Instrucoes.Text := trim('NºO RECEBER APÓS 30 DIAS DE VENCIMENTO.Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

//            RLBTitulo2.Instrucoes.Text := trim('N. Parcela:'+FrmRelRecebimento.CDS_MarcaTitordem.Value+' Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value);

//            if truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2)>0 then
//               RLBTitulo2.Instrucoes.Text :='Até o vencimento desconto de R$ '+trim(transform(truncar((FrmRelRecebimento.CDS_MarcaTitvalor.Value * DM_Tabelas.ZQContaBancariaPerc_descontos.Value)/100,2),'###,###,##0.00'))+#13+
//                                            ' COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value
//            else
               RLBTitulo2.Instrucoes.Text := 'COBRAR MULTA DE '+FrmRelRecebimento.CDSEmpreendimentojurosemboleto.Text+'% SOBRE O VALOR DA PARCELA'+#13+' NAO RECEBER APOS 30 DIAS DO VENCIMENTO. '+#13+' Pagto. Ref.:'+FrmRelRecebimento.CDS_MarcaTitquadralote.Value;
            RLBTitulo2.InsertRecord;
            if cbemail.Checked=true then
            begin
               RLBTitulo2.PreviewModal;
               ZQAniversariante.close;
               ZQAniversariante.SQL.clear;
               ZQAniversariante.SQL.add('select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
               ZQAniversariante.SQL.add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
               ZQAniversariante.SQL.add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico ');
               ZQAniversariante.SQL.add(' from participante where nome_parte='+quotedstr(FrmRelRecebimento.CDS_MarcaTitnome_parte.Value)+' order by nome_parte');
               ZQAniversariante.open;
               if not empty(ZQAniversariante.FieldByName('email').AsString) then
               begin
                 CDParticipante.Insert;
                 CDParticipanteidpaticipante.Value:=ZQAniversariante.FieldByName('idpaticipante').AsLargeInt;
                 CDParticipantenome_parte.Value:=ZQAniversariante.FieldByName('nome_parte').AsString;
                 CDParticipanteemail.Value:=ZQAniversariante.FieldByName('email').AsString;
                 CDParticipanteinativo.Value:=ZQAniversariante.FieldByName('Inativo').AsString;
                 CDParticipantenascimento.Value:=ZQAniversariante.FieldByName('aniversario').AsDateTime;
                 CDParticipanteProfissao.Value:=ZQAniversariante.FieldByName('profissao').AsString;
                 CDParticipante.post;
              end;
              ZQAniversariante.close;
            end;

{          // arquivo de remessa


            RLBTitulo2.Cedente.ContaBancaria.Banco.Codigo := trim(configuracao_boletoCodigo_bco.value);
            RLBTitulo2.Cedente.ContaBancaria.CodigoAgencia :=trim(configuracao_boletoAgencia.value);
            RLBTitulo2.Cedente.ContaBancaria.DigitoAgencia :=trim(configuracao_boletoDIGITO_AGE.value);
            RLBTitulo2.Cedente.ContaBancaria.NumeroConta := trim(configuracao_boletoConta.value);
//          RLBTitulo2.Cedente.ContaBancaria.DigitoConta := trim(configuracao_boletoDIGITO_CONTA.value);
            RLBTitulo2.Cedente.CodigoCedente := '2053098';
    //      RLBTitulo2.Cedente.DigitoCodigoCedente := '6';
            RLBTitulo2.Carteira := trim(configuracao_boletoCarteira.Value);
//          RLBTitulo2.Cedente.ContaBancaria.Convenio := trim(configuracao_boletoConvenio.Value);
            RLBTitulo2.Cedente.ContaBancaria.NomeCliente := trim(CadempRAZ_EMP.Value);
            RLBTitulo2.Cedente.TipoInscricao := tiPessoaJuridica;
            RLBTitulo2.Cedente.NumeroCPFCGC := charrem(CadempCNPJ.Value);
            RLBTitulo2.Cedente.Nome := trim(CadempRAZ_EMP.Value);
            //dados do sacado
            RLBTitulo2.Sacado.Nome := Trim(CadcliRAZAO.Value);
            if cadclitip_pessoa.Value='JURÍDICA' then
            begin
              RLBTitulo2.Sacado.NumeroCPFCGC:=charrem(trim(CadcliCGC.value));
              RLBTitulo2.Sacado.TipoInscricao := tiPessoaJuridica;
            end
            else
            begin
              RLBTitulo2.Sacado.NumeroCPFCGC:=charrem(trim(CadcliCPF.value));
              RLBTitulo2.Sacado.TipoInscricao := tiPessoaFisica;
            end;
            RLBTitulo2.Sacado.Endereco.Rua := Trim(CadcliENDE.value);
            RLBTitulo2.Sacado.Endereco.CEP := charrem(Trim(CadcliCEP.Value));
            RLBTitulo2.Sacado.Endereco.Cidade := Trim(Cadclicidade.Value);
            RLBTitulo2.Sacado.Endereco.Estado := Trim(CadcliUF.Value);
            RLBTitulo2.Sacado.Endereco.Bairro :=Trim(CadcliBairro.Value);

            //dados da cobranca
            RLBTitulo2.TipoOcorrencia := toRemessaRegistrar;
            RLBTitulo2.EspecieDocumento := edDuplicataMercantil;
            RLBTitulo2.NossoNumero := Trim(nossonumero)+RLBTitulo2.DigitoNossoNumero;
            RLBTitulo2.ValorDocumento := CadmovDUP_VAL.Value;
            RLBTitulo2.DataDocumento := CadmovDUP_EMI.Value;
            RLBTitulo2.ValorMoraJuros := (CadmovDUP_VAL.Value*0.07)/100;
            RLBTitulo2.DataMoraJuros:=CadmovDUP_VEN.Value+1;
            RLBTitulo2.DataVencimento := CadmovDUP_VEN.Value;
            RLBTitulo2.NumeroDocumento := Trim(CadmovN_DOC.Value+CadmovSEQ.Value);
            RLBTitulo2.SeuNumero :=  Trim(CadmovN_DOC.Value+CadmovSEQ.Value);
            RLBTitulo2.Instrucoes.Text := trim(configuracao_boletoInstrucoes.Value);

            //Inserindo o boleto na remessa
            RLBRemessa1.Titulos.Add(RLBTitulo2);}
          end;
        end;
      end;
      FrmRelRecebimento.CDS_MarcaTit.next;
    end;
    FrmRelRecebimento.CDS_MarcaTit.EnableControls;
    if cbemail.Checked=false then
     RLBTitulo2.PreviewModal;
  end;

  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);
  DM_Tabelas.ZQContaBancaria.edit;
  DM_Tabelas.ZQContaBancaria.FieldByName('n_sequencial').AsLargeInt:=varremes;
  if rgnosso.Checked then
     DM_Tabelas.ZQContaBancaria.FieldByName('nosso_numero').AsLargeInt:=xnosso;
  DM_Tabelas.ZQContaBancaria.post;
  DM_Tabelas.ZQContaBancaria.ApplyUpdates;
  DM_Tabelas.ZQContaBancaria.Refresh;
  DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',FrmRelRecebimento.CDSEmpreendimentocodcontabancaria.Value, []);

  Gauge1.Visible := False;
//  if cbemail.Checked=false then
  //   RLBTitulo2.PreviewModal;

  // 14/12/2012
  if  rgnosso.Checked then
  begin
    Gauge1.Visible := True;
    Gauge1.Progress:=0;

    //17/09/2013
    ZQRemesRec.close;
    ZQRemesRec.SQL.Clear;
    ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where remessa='+quotedstr(inttostr(numeroremessa)));
    ZQRemesRec.open;



    ZQRemesRec.First;
    Gauge1.MaxValue := ZQRemesRec.RecordCount;
//    DM_Tabelas.ZQRemes_Receb_atualiza.open;
  //  DM_Tabelas.ZQRemes_Receb_atualiza.first;
    while not ZQRemesRec.Eof do
    begin
      Gauge1.Progress := ZQRemesRec.RecNo;

      DM_Tabelas.ZQRemes_Receb_atualiza.Close;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text));
      DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;



      ZQRemesRec.Next;
    end;
//    DM_Tabelas.ZQRemes_Receb_atualiza.refresh;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRecebimento.refresh;
  end;
  Gauge1.Progress:=0;
  Gauge1.Visible := false;
  varremes:=0;
  xnosso:=0;

end;


procedure TFrmCobrancaBancaria.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmCobrancaBancaria.FormShow(Sender: TObject);
begin
  DBGerar.Enabled := false;
//  XBRelat.Enabled := False;
  DM_Tabelas.ZQBancRemes.Open;
  ZQRemesRec.Open;
  DM_Tabelas.ZQBol_men.open;

  CBIdentOcorr.ItemIndex := 1;
  DBGEmpre.SetFocus;
  FrmRelRecebimento.CDSEmpreendimento.Next;
  FrmRelRecebimento.CDSEmpreendimento.First;
  if CBBancoemite.Checked Then
  begin
    Label7.Visible:=false;
    ComboBox3.Visible:=false;
    chkemissao.Visible:=false;
  end
  else
  begin
    Label7.Visible:=true;
    ComboBox3.Visible:=true;
    chkemissao.Visible:=true;
  end;
  CDParticipante.Close;
  CDParticipante.CreateDataSet;
  CDParticipante.open;
end;

procedure TFrmCobrancaBancaria.DBGerarClick(Sender: TObject);
begin
// DBGerar.Enabled:=false;
 FrmRelRecebimento.CDSEmpreendimento.First;
 FrmRelRecebimento.CDS_MarcaTit.First;
  if copy(FrmRelRecebimento.XDEVencimentoInicio.DateText,1,2)='00' Then Begin
    mensagem('A data de vencimento não foi informada...  Arquivo não será gerado....');
    close;
    exit;
  end;
  DBGEmpre.SetFocus;
  FrmRelRecebimento.CDSEmpreendimento.First;
  while not FrmRelRecebimento.CDSEmpreendimento.Eof do
  begin
    if (FrmRelRecebimento.CDSEmpreendimentomarca.Value = '0') Then
      if FrmRelRecebimento.CDS_MarcaTit.RecordCount>0 Then
      Begin
        FrmRelRecebimento.CDS_MarcaTit.Filter := 'cliente='+FrmRelRecebimento.CDSEmpreendimentoidparticipante.Text;
        if FrmRelRecebimento.CDSEmpreendimentolayoutBoleto.Value='BRADESCO_400' Then bradesco_400
        else if FrmRelRecebimento.CDSEmpreendimentolayoutBoleto.Value='ITAU_400' Then itau_400
        else if FrmRelRecebimento.CDSEmpreendimentolayoutBoleto.Value='BANCOOB_400' Then BANCOOB_400
        else if FrmRelRecebimento.CDSEmpreendimentolayoutBoleto.Value='BRASIL_400' Then BRASIL_400
        else if FrmRelRecebimento.CDSEmpreendimentolayoutBoleto.Value='SANTANDER_400' Then santander_400
        else if FrmRelRecebimento.CDSEmpreendimentolayoutBoleto.Value='SANTANDER_240' Then santander_240
        else if FrmRelRecebimento.CDSEmpreendimentolayoutBoleto.Value='HSBC_400' Then hsbc_400
        else if FrmRelRecebimento.CDSEmpreendimentolayoutBoleto.Value='SICOOB_240' Then SICOOB_240
        else if FrmRelRecebimento.CDSEmpreendimentolayoutBoleto.Value='CAIXA_240' Then caixa_240

        else mensagem('Este banco não esta configurado para gerar arquivo magnético...');
      end;
//    end;
    FrmRelRecebimento.CDSEmpreendimento.Next;
  end;
//  XBRelat.Enabled := true;
//  XBRelat.SetFocus;
//  XBRelat.SetFocus;
end;

procedure TFrmCobrancaBancaria.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmCobrancaBancaria.DBETitKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  DM_Tabelas.ZQBol_men.Edit;
end;

procedure TFrmCobrancaBancaria.DBETitMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  DM_Tabelas.ZQBol_men.Edit;
end;

procedure TFrmCobrancaBancaria.XDBEdit1Change(Sender: TObject);
begin
  Label3.Caption := inttostr(length(XDBEdit1.Text));
end;

procedure TFrmCobrancaBancaria.XDBEdit2Change(Sender: TObject);
begin
  Label4.Caption := inttostr(length(XDBEdit2.Text));
end;

procedure TFrmCobrancaBancaria.XDBEdit3Change(Sender: TObject);
begin
  Label5.Caption := inttostr(length(XDBEdit3.Text));
end;

procedure TFrmCobrancaBancaria.XDBEdit4Change(Sender: TObject);
begin
  Label6.Caption := inttostr(length(XDBEdit4.Text));
end;

procedure TFrmCobrancaBancaria.DXBGravarClick(Sender: TObject);
begin
  if DM_Tabelas.ZQContaBancaria.FieldByName('boleto').AsString='BRADESCO_400' Then bradesco_400
  else if DM_Tabelas.ZQContaBancaria.FieldByName('boleto').AsString='ITAU_400' Then itau_400
  else if DM_Tabelas.ZQContaBancaria.FieldByName('boleto').AsString='BANCOOB_400' Then BANCOOB_400
  else if DM_Tabelas.ZQContaBancaria.FieldByName('boleto').AsString='SICOOB_240' Then SICOOB_240
  else if DM_Tabelas.ZQContaBancaria.FieldByName('boleto').AsString='BRASIL_400' Then BRASIL_400
  else if DM_Tabelas.ZQContaBancaria.FieldByName('boleto').AsString='SANTANDER_400' Then santander_400
  else mensagem('Este banco não esta configurado para gerar arquivo magnético...');
//  XBRelat.Enabled := true;
//  XBRelat.SetFocus;
end;

procedure TFrmCobrancaBancaria.XBRelatClick(Sender: TObject);
begin
  if Frm_RelBolBanco=nil then
     Frm_RelBolBanco:= TFrm_RelBolBanco.Create(Application);

  Frm_RelBolBanco.marca0.Value :=0;
  Frm_RelBolBanco.marca1.Value :=0;
//  Frm_RelBolBanco.RLLabel2.Caption := 'Gerado arquivo em '+Varnomearq;
  Frm_RelBolBanco.RLReport1.PreviewModal;
  Frm_RelBolBanco:=nil;
end;

procedure TFrmCobrancaBancaria.DBGEmpreColEnter(Sender: TObject);
begin
{  if DBGEmpre.SelectedField = FrmRelRecebimento.CDSEmpreendimentomarca then
    DBGEmpre.Options := DBGEmpre.Options - [dgEditing]
  else
    DBGEmpre.Options := DBGEmpre.Options + [dgEditing];
 }
end;

procedure TFrmCobrancaBancaria.DBGEmpreDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBEEmpre.Text = FrmRelRecebimento.CDSEmpreendimentoidordem.Text Then Begin
    DBGEmpre.Canvas.Brush.Color :=$006CFFFF;
    DBGEmpre.Canvas.Font.Color := $00A80000;
  end;
  DBGEmpre.DefaultDrawDataCell(Rect, DBGEmpre.columns[datacol].field, State);
  if Column.Field = FrmRelRecebimento.CDSEmpreendimentomarca then begin
    DBGEmpre.Canvas.FillRect(Rect);
    if (FrmRelRecebimento.CDSEmpreendimentomarca.Value = '0') or (FrmRelRecebimento.CDSEmpreendimentomarca.Value = '') then
      ImageList1.Draw(DBGEmpre.Canvas, Rect.Left + 2,Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBGEmpre.Canvas, Rect.Left + 2,Rect.Top + 1, 1);
  end;
end;

procedure TFrmCobrancaBancaria.DBGEmpreKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if key = 13 Then
    mudamarcaEmpreend;
end;

procedure TFrmCobrancaBancaria.mudamarcaEmpreend;
Begin
  if DBGEmpre.SelectedField.FieldName = 'marca' Then Begin
    FrmRelRecebimento.CDSEmpreendimento.Edit;
    if (FrmRelRecebimento.CDSEmpreendimentomarca.Value = '0') or (FrmRelRecebimento.CDSEmpreendimentomarca.Value = '') Then
      FrmRelRecebimento.CDSEmpreendimentomarca.Value := '1'
    else
      FrmRelRecebimento.CDSEmpreendimentomarca.Value := '0';
    FrmRelRecebimento.CDSEmpreendimento.Post;
  end;
End;

procedure TFrmCobrancaBancaria.DBGEmpreMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  mudamarcaEmpreend;
end;

procedure TFrmCobrancaBancaria.DBGTitDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBEIdRec.Text = FrmRelRecebimento.CDS_MarcaTitIdreceb.Text Then Begin
    DBGTit.Canvas.Brush.Color :=$006CFFFF;
    DBGTit.Canvas.Font.Color := $00A80000;
  end;
  DBGTit.DefaultDrawDataCell(Rect, DBGTit.columns[datacol].field, State);
  if Column.Field = FrmRelRecebimento.CDS_MarcaTitmarca then begin
    DBGTit.Canvas.FillRect(Rect);
    if (FrmRelRecebimento.CDS_MarcaTitmarca.Value= '0') or (FrmRelRecebimento.CDS_MarcaTitmarca.Value = '') then
      ImageList1.Draw(DBGTit.Canvas, Rect.Left + 2,Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBGTit.Canvas, Rect.Left + 2,Rect.Top + 1, 1);
  end;
end;

procedure TFrmCobrancaBancaria.DBEEmpreChange(Sender: TObject);
begin
  if (DBGEmpre.Focused) and (not empty(FrmRelRecebimento.CDSEmpreendimentoidparticipante.Text)) Then
  Begin
   if chkemissao.Checked=false then
   begin
      FrmRelRecebimento.CDS_MarcaTit.Filtered := False;
      FrmRelRecebimento.CDS_MarcaTit.Filter := 'idloteam='+FrmRelRecebimento.CDSEmpreendimentoidloteamento.Text+' and cliente='+FrmRelRecebimento.CDSEmpreendimentoidparticipante.Text;
      FrmRelRecebimento.CDS_MarcaTit.Filtered := True;
   end
   else
   begin
      FrmRelRecebimento.CDS_MarcaTit.Filtered := False;
      FrmRelRecebimento.CDS_MarcaTit.Filter := 'numboleto<>'''' and idloteam='+FrmRelRecebimento.CDSEmpreendimentoidloteamento.Text+' and cliente='+FrmRelRecebimento.CDSEmpreendimentoidparticipante.Text;
      FrmRelRecebimento.CDS_MarcaTit.Filtered := True;
   end;
  end;
end;

procedure TFrmCobrancaBancaria.DBGrid1MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  DM_Tabelas.ZQBol_men.Edit;
  mudamarcaTit;
end;

procedure TFrmCobrancaBancaria.DBGrid1Exit(Sender: TObject);
begin
  button1.SetFocus;
end;

procedure TFrmCobrancaBancaria.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  DM_Tabelas.ZQBol_men.Edit;
  if key = 13 Then
    mudamarcaTit;
end;

procedure TFrmCobrancaBancaria.mudamarcaTit;
Begin
  if DBGTit.SelectedField.FieldName = 'marca' Then Begin
    FrmRelRecebimento.CDS_MarcaTit.Edit;
    if (FrmRelRecebimento.CDS_MarcaTitmarca.Value = '0') or (FrmRelRecebimento.CDS_MarcaTitmarca.Value = '') Then
      FrmRelRecebimento.CDS_MarcaTitmarca.Value := '1'
    else
      FrmRelRecebimento.CDS_MarcaTitmarca.Value := '0';
    FrmRelRecebimento.CDS_MarcaTit.Post;
  end;
End;

procedure TFrmCobrancaBancaria.DXBMarcTitClick(Sender: TObject);
begin
  FrmRelRecebimento.CDS_MarcaTit.First;
  Gauge2.MaxValue:=FrmRelRecebimento.CDS_MarcaTit.RecordCount;
  Gauge2.Progress:=0;
  Gauge2.Visible:=true;
  while not FrmRelRecebimento.CDS_MarcaTit.Eof do
  begin
    Gauge2.Progress:=FrmRelRecebimento.CDS_MarcaTit.RecNo;
    FrmRelRecebimento.CDS_MarcaTit.Edit;
    if DXBMarcTit.Caption = 'Desmarcar' Then
      FrmRelRecebimento.CDS_MarcaTitmarca.Value := '1'
    else
      FrmRelRecebimento.CDS_MarcaTitmarca.Value := '0';
    FrmRelRecebimento.CDS_MarcaTit.Post;
    FrmRelRecebimento.CDS_MarcaTit.Next;
  end;
  if DXBMarcTit.Caption = 'Desmarcar' Then
    DXBMarcTit.Caption := 'Marcar'
  Else
    DXBMarcTit.Caption := 'Desmarcar';
  Gauge2.Progress:=0;
  Gauge2.Visible:=false;
end;

procedure TFrmCobrancaBancaria.DXBMarcEmpClick(Sender: TObject);
begin
  FrmRelRecebimento.CDSEmpreendimento.First;

  application.ProcessMessages;
  Gauge1.maxvalue:=CDSEmpreendimento.recordcount;
  Gauge1.progress:=0;
  Gauge1.Visible:=true;
  FrmRelRecebimento.CDSEmpreendimento.DisableControls;
  CDSEmpreendimento.first;
  while not CDSEmpreendimento.Eof do
  begin
    Gauge1.progress:=CDSEmpreendimento.recno;
    application.ProcessMessages;
    CDSEmpreendimento.Edit;
    if DXBMarcEmp.Caption = 'Desmarcar' Then
      CDSEmpreendimentomarca.Value := '1'
    else
      CDSEmpreendimentomarca.Value := '0';
    CDSEmpreendimento.Post;
    CDSEmpreendimento.Next;
  end;


  FrmRelRecebimento.CDSEmpreendimento.Close;
  FrmRelRecebimento.CDSEmpreendimento.CreateDataSet;
  FrmRelRecebimento.CDSEmpreendimento.open;
  FrmRelRecebimento.CDSEmpreendimento.CloneCursor(CDSEmpreendimento, true, true);
  FrmRelRecebimento.CDSEmpreendimento.Data:=CDSEmpreendimento.Data;



  FrmRelRecebimento.CDSEmpreendimento.EnableControls;
  FrmRelRecebimento.CDSEmpreendimento.first;
  Gauge1.progress:=0;
  Gauge1.Visible:=false;
  application.ProcessMessages;  
  if DXBMarcEmp.Caption = 'Desmarcar' Then
    DXBMarcEmp.Caption := 'Marcar'
  Else
    DXBMarcEmp.Caption := 'Desmarcar';
end;

procedure TFrmCobrancaBancaria.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
//  FrmRelRecebimento.CDS_MarcaTit.close;
  CDParticipante.Close;
  ZQRemesRec.close;
  FrmCobrancaBancaria:=nil;
  Action:=CaFree;
end;

procedure TFrmCobrancaBancaria.dxButton1Click(Sender: TObject);
var
vazio:string;
begin
//  DM_Tabelas.zqempresa.Close;
//  DM_Tabelas.zqempresa.SQL.Clear;
//  DM_Tabelas.zqempresa.SQL.Add('SELECT * FROM empresa  order by razao');
//  DM_Tabelas.zqempresa.Open;
//  DM_Tabelas.zqempresa.First;
//  if DM_Tabelas.zqempresa.RecordCount>1 then
//  begin
//    GroupBox2.Visible:=true;
//    dbgrid1.SetFocus;
//    exit;
//  end;

   vazio:='';
  if  CBatualizar.Checked then
  begin
    Gauge1.Visible := True;
    Gauge1.Progress:=0;

    //17/09/2013
    ZQRemesRec.close;
    ZQRemesRec.SQL.Clear;
    ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,nossonumero from remessa_receb ');
    ZQRemesRec.open;

    ZQRemesRec.First;
    Gauge1.MaxValue := ZQRemesRec.RecordCount;
//    DM_Tabelas.ZQRemes_Receb_atualiza.open;
  //  DM_Tabelas.ZQRemes_Receb_atualiza.first;
    while not ZQRemesRec.Eof do
    begin
      Gauge1.Progress := ZQRemesRec.RecNo;

      DM_Tabelas.ZQRemes_Receb_atualiza.Close;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Clear;
      DM_Tabelas.ZQRemes_Receb_atualiza.SQL.Add('update  recebimento set numboleto=('+quotedstr(ZQRemesRec.FieldByName('nossonumero').AsString)+') where ( idrecebimento='+quotedstr(ZQRemesRec.FieldByName('idrec').Text)+') and ((numboleto='+quotedstr(vazio)+') or (numboleto is null))');
      DM_Tabelas.ZQRemes_Receb_atualiza.ExecSQL;
      ZQRemesRec.Next;
    end;
    DM_Tabelas.ZQRemes_Receb_atualiza.close;
    DM_Tabelas.ZQRecebimento.refresh;
  end;
  Gauge1.Progress:=0;
  Gauge1.Visible := false;


   FrmRelRecebimento.CDSEmpreendimento.First;
   FrmRelRecebimento.CDSEmpreendimento.DisableControls;
   while not FrmRelRecebimento.CDSEmpreendimento.Eof do
   begin
      FrmRelRecebimento.CDS_MarcaTit.Filtered:=false;
      FrmRelRecebimento.CDS_MarcaTit.Filter:='idloteam='+quotedstr(FrmRelRecebimento.CDSEmpreendimentoidloteamento.Text);
      FrmRelRecebimento.CDS_MarcaTit.Filtered:=true;
      FrmRelRecebimento.CDS_MarcaTit.First;
      Gauge2.MaxValue:=FrmRelRecebimento.CDS_MarcaTit.RecordCount;
      Gauge2.Progress:=0;
      Gauge2.Visible:=true;
      while not FrmRelRecebimento.CDS_MarcaTit.Eof do
      begin
        Gauge2.Progress:=FrmRelRecebimento.CDS_MarcaTit.RecNo;
        FrmRelRecebimento.CDS_MarcaTit.Edit;
        FrmRelRecebimento.CDS_MarcaTitdigito_dif.Value:=FrmRelRecebimento.CDSEmpreendimentodigito_dif.Value;
        FrmRelRecebimento.CDS_MarcaTit.post;
        FrmRelRecebimento.CDS_MarcaTit.next;
      end;
      Gauge2.Progress:=0;
      Gauge2.Visible:=false;
      FrmRelRecebimento.CDSEmpreendimento.Next;
   end;
   FrmRelRecebimento.CDSEmpreendimento.EnableControls;
   FrmRelRecebimento.CDSEmpreendimento.First;
   FrmRelRecebimento.CDS_MarcaTit.Filtered := False;
   FrmRelRecebimento.CDS_MarcaTit.Filter := 'idloteam='+FrmRelRecebimento.CDSEmpreendimentoidloteamento.Text+' and cliente='+FrmRelRecebimento.CDSEmpreendimentoidparticipante.Text;
   FrmRelRecebimento.CDS_MarcaTit.Filtered := True;
   DBGerar.Enabled:=true;
end;

procedure TFrmCobrancaBancaria.CBBancoemiteClick(Sender: TObject);
begin
  if CBBancoemite.Checked Then
  begin
    Label7.Visible:=false;
    ComboBox3.Visible:=false;
    chkemissao.Visible:=false;
    cbemail.Visible:=false;
    cb2via.Checked:=false;
    cb2via.Visible:=false;
    CBCarteiraEnvio.Checked:=true;
  end
  else
  begin
    Label7.Visible:=true;
    ComboBox3.Visible:=true;
    chkemissao.Visible:=true;
    cbemail.Visible:=true;
    cb2via.Visible:=true;
    CBCarteiraEnvio.Checked:=false;
  end;

end;

procedure TFrmCobrancaBancaria.chkemissaoClick(Sender: TObject);
begin
  if (not empty(FrmRelRecebimento.CDSEmpreendimentoidparticipante.Text)) Then
  Begin
   if chkemissao.Checked=false then
   begin
      FrmRelRecebimento.CDS_MarcaTit.Filtered := False;
      FrmRelRecebimento.CDS_MarcaTit.Filter := 'idloteam='+FrmRelRecebimento.CDSEmpreendimentoidloteamento.Text+' and cliente='+FrmRelRecebimento.CDSEmpreendimentoidparticipante.Text;
      FrmRelRecebimento.CDS_MarcaTit.Filtered := True;
   end
   else
   begin
      FrmRelRecebimento.CDS_MarcaTit.Filtered := False;
      FrmRelRecebimento.CDS_MarcaTit.Filter := 'numboleto<>'''' and idloteam='+FrmRelRecebimento.CDSEmpreendimentoidloteamento.Text+' and cliente='+FrmRelRecebimento.CDSEmpreendimentoidparticipante.Text;
      FrmRelRecebimento.CDS_MarcaTit.Filtered := True;
   end;
  end;
end;

procedure TFrmCobrancaBancaria.cbemailClick(Sender: TObject);
begin
  if cbemail.Checked=true then
  begin
    CDParticipante.Close;
    CDParticipante.CreateDataSet;
    CDParticipante.open;
  end
  else
  begin
    CDParticipante.Close;
  end;

end;


procedure TFrmCobrancaBancaria.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmCobrancaBancaria, 'zqcarne', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'zqcarne', 'zqcarneidcarne', 'idcarne', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'zqcarne', 'zqcarnen_seq', 'n_seq', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCobrancaBancaria, 'ZQAniversariante', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantenome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantedoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantedoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantebairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantebairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantefone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantefone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantetipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantenacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantenaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantelocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianterenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantefone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversariantecadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQAniversariante', 'ZQAniversarianteInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCobrancaBancaria, 'ZQRemesRec', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQRemesRec', 'ZQRemesRecidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQRemesRec', 'ZQRemesRecremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQRemesRec', 'ZQRemesRecidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQRemesRec', 'ZQRemesRecnossonumero', 'nossonumero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCobrancaBancaria, 'ZQNumero_remessa', False);
  RegisterRuntimeField(TFrmCobrancaBancaria, 'ZQNumero_remessa', 'ZQNumero_remessaidnumero_remessa', 'idnumero_remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
