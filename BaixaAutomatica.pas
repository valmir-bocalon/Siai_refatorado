unit BaixaAutomatica;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XLabel3D, XBanner, Grids, DBGrids,
  DBCtrls, ExtCtrls, IniFiles, ComCtrls, DB, DBClient, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Mask, XDate, ACBrBase, ACBrBoleto;

//  , , , , , , ,
//  , , WinSkinData, , , , jpeg, ,
//  , Graphics;            Menus

type
  TFrm_BaixaAutomatica = class(TForm)
    XLTitulo: TXLabel3D;
    OpenDialog1: TOpenDialog;
    DXBFechar: TdxButtonArround;
    DBAbrir: TdxButtonArround;
    DBGBaixaAuto: TDBGrid;
    DXBProcessar: TdxButtonArround;
    DXBImprimir: TdxButtonArround;
    DS_ArqRetorno: TDataSource;
    CDSArqRetorno: TClientDataSet;
    CDSArqRetornoidrece: TIntegerField;
    Label1: TLabel;
    dxButton1: TdxButtonArround;
    CDSArqRetornoident_ocorr: TStringField;
    CDSArqRetornodesc_ocorr: TStringField;
    CDSArqRetornoident_motivo: TStringField;
    CDSArqRetornovrtit: TFloatField;
    CDSArqRetornovrarq: TFloatField;
    CDSArqRetornoacao: TStringField;
    DS_Receb: TDataSource;
    ZQReceb: TZQuery;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    EData: TXDateEdit;
    CDSArqRetornovrjuros: TFloatField;
    CDSArqRetornovrdesconto: TFloatField;
    CDSArqRetornodtvenci: TDateField;
    Ebxautoplc: TDBEdit;
    Label5: TLabel;
    Label6: TLabel;
    DBEdit2: TDBEdit;
    CDSArqRetornoordem: TStringField;
    DS_ComisTemp: TDataSource;
    ZQComisTemp: TZQuery;
    XBanner7: TXBanner;
    CDSArqRetornodesc_motivo: TStringField;
    datacredito: TXDateEdit;
    CDSArqRetornodt_credito: TDateField;
    CDSArqRetornonossonumero: TStringField;
    CDSArqRetornoDigito: TStringField;
    CDSArqRetornoidconta_bancaria: TLargeintField;
    ProgressBar1: TProgressBar;
    Panel1: TPanel;
    ProgressBar2: TProgressBar;
    CDSArqRetornoNomecliente: TStringField;
    CDSArqRetornoquadralote: TStringField;
    ACBrBoleto1: TACBrBoleto;
    CDSArqRetornobanco: TStringField;
    CDSArqRetornoadversa: TIntegerField;
    lbl240: TLabel;
    procedure DXBFecharClick(Sender: TObject);
    procedure DBAbrirClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure dxButton1Click(Sender: TObject);
    procedure DBGBaixaAutoDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DXBImprimirClick(Sender: TObject);
    procedure DXBProcessarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_BaixaAutomatica: TFrm_BaixaAutomatica;
  processou:string;
  
implementation

uses Tabelas, funcoes, TabelasDeMensagens, RelRetorno,
     StrUtils, math,
      DateUtils, FileCtrl,
      ACBrUtil, uRuntimeFields;



{$R *.dfm}

procedure TFrm_BaixaAutomatica.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_BaixaAutomatica.DBAbrirClick(Sender: TObject);
var
  arq: TextFile;
  dg,digit,linha, vartext,VarDif,dtscr : String;
  posi:integer;
  Varx : integer;
  zbco,nn,hj,varnome, vardata : string;
  vidrece,ArqLido,ArqSize :Integer;

begin
//  if ZQReceb.active=false then
  //  ZQReceb.open;
  DBAbrir.Enabled:=false;
  processou:='N';
  VarDif:='';
  digit:='';
  hj:='';
  dg:='';
  lbl240.Caption:='';
  if OpenDialog1.Execute Then Begin
    AssignFile ( arq, OpenDialog1.FileName );
    Reset ( arq );
    ReadLn ( arq, linha );
    if linha=emptystr then
       ReadLn ( arq, linha );
    zbco:=copy(linha,1,3);
    if copy(linha,1,3)<>'756' then
    begin
      if copy(linha,1,3)='033' then   //santander 240
      begin
        for varx := 1 to length(OpenDialog1.FileName) do
          if copy(OpenDialog1.FileName,varx,1)='/' Then
             break;
        varnome := copy(OpenDialog1.FileName,varx-29,29);
        lbl240.Caption:='033';
        vardata := Copy(Linha,148,4)+'-'+Copy(Linha,146,2)+'-'+Copy(Linha,144,2);
      end
      else
      if copy(linha,1,3)='104' then   //caixa 240
      begin
        for varx := 1 to length(OpenDialog1.FileName) do
          if copy(OpenDialog1.FileName,varx,1)='/' Then
             break;
        varnome := copy(OpenDialog1.FileName,varx-20,20);
        lbl240.Caption:='104';
        vardata := Copy(Linha,148,4)+'-'+Copy(Linha,146,2)+'-'+Copy(Linha,144,2);
      end
      else
      begin
        EData.DateValue:= strtodate(copy(linha,95,2)+'/'+copy(linha,97,2)+'/20'+copy(linha,99,2));
        if copy(linha,3,7) <> 'RETORNO' Then
          exit;

    {    for varx := length(OpenDialog1.FileName) to 1 do
            if copy(OpenDialog1.FileName,varx,1)='/' Then
               break;
        varnome := copy(OpenDialog1.FileName,varx-11,12);}

        for varx := 1 to length(OpenDialog1.FileName) do
            if copy(OpenDialog1.FileName,varx,1)='/' Then
               break;
        varnome := copy(OpenDialog1.FileName,varx-12,12);
        vardata := copy(edata.DateText,7,4)+'-'+copy(edata.DateText,4,2)+'-'+copy(Edata.DateText,1,2);

      end;

    end
    else
    begin
      for varx := 1 to length(OpenDialog1.FileName) do
          if copy(OpenDialog1.FileName,varx,1)='/' Then
             break;
      varnome := copy(OpenDialog1.FileName,varx-34,34);
      vardata := Copy(Linha,148,4)+'-'+Copy(Linha,146,2)+'-'+Copy(Linha,144,2);


    end;
//    if DM_Tabelas.ZQRemes_Receb.Active=false then
//       DM_Tabelas.ZQRemes_Receb.open;

    if copy(linha,1,3)<>'756' then
    begin

      DM_TAbelas.ZQBancRemes2.close;
      DM_TAbelas.ZQBancRemes2.SQL.Clear;
      DM_TAbelas.ZQBancRemes2.SQL.Add('select idbanco_remessa,remessa,idbanco,gerado,nomearq,arquivo,digito_dif,remessa2 from banco_remessa where nomearq='+quotedstr(varnome)+' and gerado='+quotedstr(vardata));
      DM_TAbelas.ZQBancRemes2.Open;


      {DM_TAbelas.ZQBancRemes2.Filtered:=false;
      DM_TAbelas.ZQBancRemes2.Filter:='nomearq='+quotedstr(varnome)+' and gerado='+quotedstr(vardata);
      DM_TAbelas.ZQBancRemes2.Filtered:=true;}

      if DM_TAbelas.ZQBancRemes2.RecordCount>0 then
      begin
  //      if DM_TAbelas.ZQBancRemes.locate('nomearq;gerado',VarArrayOf([varnome,vardata]),[]) Then Begin
    //       mensagem('Este arquivo ja foi baixado no dia '+DM_Tabelas.ZQBancRemes2gerado.Text+', ele não pode mais ser baixado!!!!');
        mensagem('Existe baixa com o mesmo nome do arquivo no dia '+DM_Tabelas.ZQBancRemes2.FieldByName('gerado').Text+' prosseguindo...');
        DM_TAbelas.ZQRemesRec.Open;
        DM_TAbelas.ZQBancRemes2.First;
        ProgressBar1.position:=0;
        ProgressBar1.max:= DM_TAbelas.ZQBancRemes2.recordcount;
        ProgressBar1.visible:=true;
        DM_TAbelas.ZQBancRemes2.DisableControls;
        while not DM_TAbelas.ZQBancRemes2.Eof do
        begin
          application.ProcessMessages;
          ProgressBar1.position:= DM_TAbelas.ZQBancRemes2.recno;
          DM_TAbelas.ZQRemesRec.Close;
          DM_TAbelas.ZQRemesRec.SQL.clear;
          DM_TAbelas.ZQRemesRec.SQL.Add('delete from remessa_receb where remessa='+quotedstr(DM_TAbelas.ZQBancRemes2.FieldByName('idbanco_remessa').Text));
          DM_TAbelas.ZQRemesRec.ExecSQL;

  {        DM_TAbelas.ZQRemesRec.Filtered:=false;
          DM_TAbelas.ZQRemesRec.Filter:='remessa='+quotedstr(DM_TAbelas.ZQBancRemes2idbanco_remessa.Text);
          DM_TAbelas.ZQRemesRec.Filtered:=true;
          DM_TAbelas.ZQRemesRec.First;
          if DM_TAbelas.ZQRemesRec.RecordCount>0 then
          begin
            while not DM_TAbelas.ZQRemesRec.Eof do
            begin
              DM_TAbelas.ZQRemesRec.Delete;
            end;
          end;}
          DM_TAbelas.ZQBancRemes2.delete;
        end;
        DM_TAbelas.ZQBancRemes2.EnableControls;
        ProgressBar1.position:=0;
        ProgressBar1.visible:=false;

        DM_TAbelas.ZQRemesRec.Filtered:=false;
        DM_TAbelas.ZQRemesRec.Close;
        DM_TAbelas.ZQRemesRec.SQL.clear;
        DM_TAbelas.ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,retorno_motivo,acao,nossonumero,ocorrencia,desc_motivo,credito,dt_ocorrencia from remessa_receb ');
  //      DM_TAbelas.ZQRemesRec.Open;
  //
  //      DM_TAbelas.ZQRemesRec.Filtered:=false;
  //      DM_TAbelas.ZQRemesRec.Refresh;

      //  exit;
      end;
      DM_TAbelas.ZQBancRemes2.Filtered:=false;
      DM_TAbelas.ZQBancRemes2.close;
    end
    else
    begin

      DM_TAbelas.ZQBancRemes2.close;
      DM_TAbelas.ZQBancRemes2.SQL.Clear;
      DM_TAbelas.ZQBancRemes2.SQL.Add('select idbanco_remessa,remessa,idbanco,gerado,nomearq,arquivo,digito_dif,remessa2 from banco_remessa where nomearq='+quotedstr(copy(varnome,15,20))+' and gerado='+quotedstr(vardata));
      DM_TAbelas.ZQBancRemes2.Open;


      {DM_TAbelas.ZQBancRemes2.Filtered:=false;
      DM_TAbelas.ZQBancRemes2.Filter:='nomearq='+quotedstr(varnome)+' and gerado='+quotedstr(vardata);
      DM_TAbelas.ZQBancRemes2.Filtered:=true;}

      if DM_TAbelas.ZQBancRemes2.RecordCount>0 then
      begin
  //      if DM_TAbelas.ZQBancRemes.locate('nomearq;gerado',VarArrayOf([varnome,vardata]),[]) Then Begin
    //       mensagem('Este arquivo ja foi baixado no dia '+DM_Tabelas.ZQBancRemes2gerado.Text+', ele não pode mais ser baixado!!!!');
        mensagem('Existe baixa com o mesmo nome do arquivo no dia '+DM_Tabelas.ZQBancRemes2.FieldByName('gerado').Text+' prosseguindo...');
        DM_TAbelas.ZQRemesRec.Open;
        DM_TAbelas.ZQBancRemes2.First;
        ProgressBar1.position:=0;
        ProgressBar1.max:= DM_TAbelas.ZQBancRemes2.recordcount;
        ProgressBar1.visible:=true;
        DM_TAbelas.ZQBancRemes2.DisableControls;
        while not DM_TAbelas.ZQBancRemes2.Eof do
        begin
          application.ProcessMessages;
          ProgressBar1.position:= DM_TAbelas.ZQBancRemes2.recno;
          DM_TAbelas.ZQRemesRec.Close;
          DM_TAbelas.ZQRemesRec.SQL.clear;
          DM_TAbelas.ZQRemesRec.SQL.Add('delete from remessa_receb where remessa='+quotedstr(DM_TAbelas.ZQBancRemes2.FieldByName('idbanco_remessa').Text));
          DM_TAbelas.ZQRemesRec.ExecSQL;

  {        DM_TAbelas.ZQRemesRec.Filtered:=false;
          DM_TAbelas.ZQRemesRec.Filter:='remessa='+quotedstr(DM_TAbelas.ZQBancRemes2idbanco_remessa.Text);
          DM_TAbelas.ZQRemesRec.Filtered:=true;
          DM_TAbelas.ZQRemesRec.First;
          if DM_TAbelas.ZQRemesRec.RecordCount>0 then
          begin
            while not DM_TAbelas.ZQRemesRec.Eof do
            begin
              DM_TAbelas.ZQRemesRec.Delete;
            end;
          end;}
          DM_TAbelas.ZQBancRemes2.delete;
        end;
        DM_TAbelas.ZQBancRemes2.EnableControls;
        ProgressBar1.position:=0;
        ProgressBar1.visible:=false;

        DM_TAbelas.ZQRemesRec.Filtered:=false;
        DM_TAbelas.ZQRemesRec.Close;
        DM_TAbelas.ZQRemesRec.SQL.clear;
        DM_TAbelas.ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,retorno_motivo,acao,nossonumero,ocorrencia,desc_motivo,credito,dt_ocorrencia from remessa_receb ');
  //      DM_TAbelas.ZQRemesRec.Open;
  //
  //      DM_TAbelas.ZQRemesRec.Filtered:=false;
  //      DM_TAbelas.ZQRemesRec.Refresh;

      //  exit;
      end;
      DM_TAbelas.ZQBancRemes2.Filtered:=false;
      DM_TAbelas.ZQBancRemes2.close;

    end;
    CDSArqRetorno.Close;
    CDSArqRetorno.CreateDataSet;
    CDSArqRetorno.open;
    CDSArqRetorno.DisableControls;
    if copy(linha,1,3)='756' then   //bco sicoob
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.filter:='n_conta='+quotedstr(inttostr(strtoint(copy(linha,64,7))));
      DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
    end;

    if copy(linha,77,3)='001' then   //bco brasil
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.filter:='n_conta='+quotedstr(inttostr(strtoint(copy(linha,32,8))));
      DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
    end;
    if copy(linha,1,3)='104' then   //bco caixa federal
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.close;
      DM_Tabelas.ZQContaBancaria.SQL.Clear;
      DM_Tabelas.ZQContaBancaria.SQL.Add('select idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('nometitular,aberta,senha_interna,senha_conta,ativa,jurosemboleto,taxadiaria,valoroupercent,carteira,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,apelido,cod_transmissao,complemento,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote from conta_bancaria where n_banco='+quotedstr(copy(linha,1,3))+' and  convenio like '+quotedstr('%'+inttostr(strtoint(copy(linha,60,6)))+'%'));
      DM_Tabelas.ZQContaBancaria.Open;

      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
    end;
    if copy(linha,77,3)='237' then  // bco bradesco
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.filter:='n_no_banco='+quotedstr(copy(linha,40,7));
      DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;

    end;

    if copy(linha,77,3)='033' then     // bco santander   400
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.close;
      DM_Tabelas.ZQContaBancaria.SQL.Clear;
      DM_Tabelas.ZQContaBancaria.SQL.Add('select idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('nometitular,aberta,senha_interna,senha_conta,ativa,jurosemboleto,taxadiaria,valoroupercent,carteira,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,apelido,cod_transmissao,complemento,');
      // 03/06/2025 o final do where coloquei o left e troquei a coluna do filtro da conta, pelo motivo que no cadastro a conta tem 9 digitos e no retorno 8
      DM_Tabelas.ZQContaBancaria.SQL.Add('moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote from conta_bancaria where left(n_conta,8) like '+quotedstr('%'+inttostr(strtoint(copy(linha,31,8)))+'%'));//+quotedstr('%'+inttostr(strtoint(copy(linha,39,8)))+'%'));
      DM_Tabelas.ZQContaBancaria.Open;

//      DM_Tabelas.ZQContaBancaria.filter:='n_conta='+quotedstr(inttostr(strtoint(copy(linha,39,8))));
  //    DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;

    end;

    if copy(linha,1,3)='033' then     // bco santander  240
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.close;
      DM_Tabelas.ZQContaBancaria.SQL.Clear;
      DM_Tabelas.ZQContaBancaria.SQL.Add('select idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('nometitular,aberta,senha_interna,senha_conta,ativa,jurosemboleto,taxadiaria,valoroupercent,carteira,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,apelido,cod_transmissao,complemento,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote from conta_bancaria where n_conta like '+quotedstr('%'+inttostr(strtoint(copy(linha,38,9)))+'%'));
      DM_Tabelas.ZQContaBancaria.Open;

//      DM_Tabelas.ZQContaBancaria.filter:='n_conta='+quotedstr(inttostr(strtoint(copy(linha,39,8))));
  //    DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;

    end;

    if copy(linha,77,3)='399' then     // bco hsbc
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.close;
      DM_Tabelas.ZQContaBancaria.SQL.Clear;
      DM_Tabelas.ZQContaBancaria.SQL.Add('select idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('nometitular,aberta,senha_interna,senha_conta,ativa,jurosemboleto,taxadiaria,valoroupercent,carteira,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,apelido,cod_transmissao,complemento,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote from conta_bancaria where n_conta like '+quotedstr('%'+inttostr(strtoint(copy(linha,38,5)))+'%'));
      DM_Tabelas.ZQContaBancaria.Open;

//      DM_Tabelas.ZQContaBancaria.filter:='n_conta='+quotedstr(inttostr(strtoint(copy(linha,39,8))));
  //    DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;

    end;



    if copy(linha,77,3)='341' then     // bco itau
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.close;
      DM_Tabelas.ZQContaBancaria.SQL.Clear;
      DM_Tabelas.ZQContaBancaria.SQL.Add('select idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('nometitular,aberta,senha_interna,senha_conta,ativa,jurosemboleto,taxadiaria,valoroupercent,carteira,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,apelido,cod_transmissao,complemento,');
      DM_Tabelas.ZQContaBancaria.SQL.Add('moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote from conta_bancaria where n_conta like '+quotedstr('%'+inttostr(strtoint(copy(linha,33,5)))+'%'));
      DM_Tabelas.ZQContaBancaria.Open;

//      DM_Tabelas.ZQContaBancaria.filter:='n_conta='+quotedstr(inttostr(strtoint(copy(linha,39,8))));
  //    DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;

    end;

    ProgressBar1.position:=0;
    ProgressBar1.max:= DM_Tabelas.ZQContaBancaria.recordcount;
    ProgressBar1.visible:=true;



    if copy(linha,1,3)='756' then
    begin
      while not DM_Tabelas.ZQContaBancaria.eof do
//    if DM_Tabelas.ZQContaBancaria.Locate('n_no_banco',copy(linha,40,7),[]) Then
      Begin
        ProgressBar1.position:= DM_Tabelas.ZQContaBancaria.recno;
        application.ProcessMessages;

        EData.DateValue:= StringToDateTimeDef(Copy(Linha,144,2)+'/'+
                                          Copy(Linha,146,2)+'/'+
                                          Copy(Linha,148,4),0, 'DD/MM/YYYY' );
//        for varx := 1 to length(OpenDialog1.FileName) do
//           if copy(OpenDialog1.FileName,varx,1)='/' Then
//              break;
//        varnome := copy(OpenDialog1.FileName,varx-12,12);
//        vardata := copy(edata.DateText,7,4)+'-'+copy(edata.DateText,4,2)+'-'+copy(Edata.DateText,1,2);
        Label1.caption := 'Banco: '+trim(copy(linha,103,30))+ ' - '+inttostr(strtoint(copy(linha,64,7)));
        Label2.caption := 'Gerado em : '+EData.DateText;
        Label3.Caption := varnome;

        // ler o arquivo texto
        ArqLido := 0;
        ArqSize := FileSize(Arq)*128;

        ProgressBar2.Position:=0;
        ProgressBar2.Max := 104; //Round(fileSize(Arq) * 1.7);
        ProgressBar2.Visible:=true;
        ReadLn ( arq, linha );

        While not eof (arq) do
        Begin
          ReadLn ( arq, linha );
          ArqLido := ArqLido +Length(Linha)+2; // 2 = CR+LN
          ProgressBar2.Position:=Trunc(ArqLido/ArqSize *100);
          application.ProcessMessages;
          if Copy(Linha,14,1)= 'T' then
          begin
            vartext := alltrim(copy(linha,106,24));
            // coloquei essas 4 linhas seguinte, pq a laura recebe retorno de boletos nao gerados pelo sistema
            if alltrim(RemoveNumeros(alltrim(vartext)))<>'' then
            begin
              vartext :='';
            end;

            VarDif  := '';
            posi:=pos('|',vartext);
            if not empty(vartext) then
            begin
              if posi=0 then
              begin
  //              posi:=length(vartext);
//                posi:=posi-2;
        //        if alltrim(copy(vartext,1,posi))<>emptystr then
                if alltrim(vartext)<>emptystr then
                begin
                  CDSArqRetorno.append;
                  CDSArqRetornobanco.Value:=zbco;

//                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,posi)));
                  CDSArqRetornoidrece.Value := strtoint(alltrim(vartext));
                  ZQReceb.Close;
                  ZQReceb.SQL.Clear;
                  ZQReceb.SQL.Add('Select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
                  ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
                  ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
                  ZQReceb.Open;

                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    nn:= copy(linha,40,7);
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').AsString;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').AsFloat;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').AsDateTime;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').AsString;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').AsString;
                    CDSArqRetornonossonumero.Value :=nn;
                    CDSArqRetornoident_ocorr.Value := copy(linha,16,2);
                    CDSArqRetornoident_motivo.Value := '00';

                    ReadLn ( arq, linha );
                    if Copy(Linha,14,1)= 'U' then
                    begin
                      datacredito.Clear;
                      dtscr:=alltrim(Copy(Linha,138,2)+'/'+Copy(Linha,140,2)+'/'+Copy(Linha,142,4));
                      if not empty(dtscr) and (dtscr<>'00000000') then
                      begin
                        hj:='N';
                        datacredito.DateValue:= StringToDateTimeDef(Copy(Linha,138,2)+'/'+
                                                           Copy(Linha,140,2)+'/'+
                                                           Copy(Linha,142,4),0, 'DD/MM/YYYY' );

                        datacredito.DateText;
                      end
                      else
                      begin
                        hj:='S';
                        datacredito.datevalue:=date;
                      end;
                      //datacredito.datevalue:=datacredito.datevalue-1;
                      // verifica o dia da semana
                      if DayOfWeek(datacredito.datevalue) = 7 then
                      begin
                        datacredito.datevalue := datacredito.datevalue - 1;
                      end
                      else if DayOfWeek(datacredito.datevalue) = 1 then
                      begin
                        datacredito.datevalue := datacredito.datevalue - 2;
                      end;

                     // coloquei no dia 13/10/2010 ref. data feriado
                    // datacredito.datevalue := EData.DateValue;

  //                   if hj='N' then
                     CDSArqRetornodt_credito.value := datacredito.DateValue;



                      CDSArqRetornovrarq.Value := StrToFloatDef(Copy(Linha,78,15),0)/100;
                      CDSArqRetornovrjuros.Value := StrToFloatDef(Copy(Linha,18,15),0)/100;
                      CDSArqRetornovrdesconto.Value := StrToFloatDef(Copy(Linha,33,15),0)/100;
                       //  DM_TAbelas.ZQMensMot.Locate('codmotivo,retorno',VarArrayOf([CDSArqRetornoident_ocorr.Value,CDSArqRetornoident_motivo.Value]),[]);
                      DM_TAbelas.ZQMensMot.close;
                      DM_TAbelas.ZQMensMot.SQL.Clear;
                      DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                      DM_TAbelas.ZQMensMot.Open;
                      if DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString='' then
                         CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                      else
                         CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString;


                      CDSArqRetornoDigito.Value:=digit;
                      CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
                      CDSArqRetorno.Post;
                      hj:='';
                    end;
                  end;
                end;
            end;
           end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := copy(varnome,15,20);
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;


        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
        Reset ( arq );
        ReadLn ( arq, linha );
        DBAbrir.Enabled:=true;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
      end;
    end;



    if copy(linha,1,3)='104' then
    begin
      while not DM_Tabelas.ZQContaBancaria.eof do
//    if DM_Tabelas.ZQContaBancaria.Locate('n_no_banco',copy(linha,40,7),[]) Then
      Begin
        ProgressBar1.position:= DM_Tabelas.ZQContaBancaria.recno;
        application.ProcessMessages;

        EData.DateValue:= StringToDateTimeDef(Copy(Linha,144,2)+'/'+
                                          Copy(Linha,146,2)+'/'+
                                          Copy(Linha,148,4),0, 'DD/MM/YYYY' );
//        for varx := 1 to length(OpenDialog1.FileName) do
//           if copy(OpenDialog1.FileName,varx,1)='/' Then
//              break;
//        varnome := copy(OpenDialog1.FileName,varx-12,12);
//        vardata := copy(edata.DateText,7,4)+'-'+copy(edata.DateText,4,2)+'-'+copy(Edata.DateText,1,2);
        Label1.caption := 'Banco: '+trim(copy(linha,103,30))+ ' - '+DM_Tabelas.ZQContaBancaria.FieldByName('n_conta').AsString+'-'+DM_Tabelas.ZQContaBancaria.FieldByName('n_conta_v').AsString;
        Label2.caption := 'Gerado em : '+EData.DateText;
        Label3.Caption := varnome;

        // ler o arquivo texto
        ArqLido := 0;
        ArqSize := FileSize(Arq)*128;

        ProgressBar2.Position:=0;
        ProgressBar2.Max := 104; //Round(fileSize(Arq) * 1.7);
        ProgressBar2.Visible:=true;
        ReadLn ( arq, linha );

        While not eof (arq) do
        Begin
          ReadLn ( arq, linha );
          ArqLido := ArqLido +Length(Linha)+2; // 2 = CR+LN
          ProgressBar2.Position:=Trunc(ArqLido/ArqSize *100);
          application.ProcessMessages;
          if Copy(Linha,14,1)= 'T' then
          begin
            vartext := alltrim(copy(linha,59,11));
            // coloquei essas 4 linhas seguinte, pq a laura recebe retorno de boletos nao gerados pelo sistema
            if alltrim(RemoveNumeros(alltrim(vartext)))<>'' then
            begin
              vartext :='';
            end;

            VarDif  := '';
            posi:=pos('|',vartext);
            if posi>0 then
            begin
              dec(posi);
              vartext:= alltrim(copy(vartext,1,posi));
              posi:=0;
            end;
            posi:=pos(' ',vartext);
            if posi>0 then
            begin
              dec(posi);
              vartext:= alltrim(copy(vartext,1,posi));
              posi:=0;
            end;
            posi:=pos('/',vartext);
            if posi>0 then
            begin
              dec(posi);
              vartext:= alltrim(copy(vartext,1,posi));
              posi:=0;
            end;
            if not empty(vartext) then
            begin
              if posi=0 then
              begin
  //              posi:=length(vartext);
//                posi:=posi-2;
        //        if alltrim(copy(vartext,1,posi))<>emptystr then
                if alltrim(vartext)<>emptystr then
                begin
                  CDSArqRetorno.append;
                  CDSArqRetornobanco.Value:=zbco;

//                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,posi)));
                  CDSArqRetornoidrece.Value := strtoint(alltrim(vartext));
                  ZQReceb.Close;
                  ZQReceb.SQL.Clear;
                  ZQReceb.SQL.Add('Select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
                  ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
                  ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
                  ZQReceb.Open;

                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    nn:= copy(linha,42,15);
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').AsString;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').AsFloat;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').AsDateTime;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').AsString;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').AsString;
                    CDSArqRetornonossonumero.Value :=nn;
                    CDSArqRetornoident_ocorr.Value := copy(linha,16,2);
                    CDSArqRetornoident_motivo.Value := '00';

                    ReadLn ( arq, linha );
                    if Copy(Linha,14,1)= 'U' then
                    begin
                      datacredito.Clear;
                      dtscr:=alltrim(Copy(Linha,146,2)+'/'+Copy(Linha,148,2)+'/'+Copy(Linha,150,4));
                      if not empty(dtscr) and (charrem(dtscr)<>'00000000') then
                      begin
                        hj:='N';
                        datacredito.DateValue:= StringToDateTimeDef(Copy(Linha,146,2)+'/'+
                                                           Copy(Linha,148,2)+'/'+
                                                           Copy(Linha,150,4),0, 'DD/MM/YYYY' );

                        datacredito.DateText;
                      end
                      else
                      begin
                        hj:='S';
                        datacredito.datevalue:=date;
                      end;
                      //datacredito.datevalue:=datacredito.datevalue-1;
                      // verifica o dia da semana
                      if DayOfWeek(datacredito.datevalue) = 7 then
                      begin
                        datacredito.datevalue := datacredito.datevalue - 1;
                      end
                      else if DayOfWeek(datacredito.datevalue) = 1 then
                      begin
                        datacredito.datevalue := datacredito.datevalue - 2;
                      end;

                     // coloquei no dia 13/10/2010 ref. data feriado
                    // datacredito.datevalue := EData.DateValue;

  //                   if hj='N' then
                     CDSArqRetornodt_credito.value := datacredito.DateValue;



                      CDSArqRetornovrarq.Value := StrToFloatDef(Copy(Linha,78,15),0)/100;
                      CDSArqRetornovrjuros.Value := StrToFloatDef(Copy(Linha,18,15),0)/100;
                      CDSArqRetornovrdesconto.Value := StrToFloatDef(Copy(Linha,33,15),0)/100;
                       //  DM_TAbelas.ZQMensMot.Locate('codmotivo,retorno',VarArrayOf([CDSArqRetornoident_ocorr.Value,CDSArqRetornoident_motivo.Value]),[]);
                      DM_TAbelas.ZQMensMot.close;
                      DM_TAbelas.ZQMensMot.SQL.Clear;
                      DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                      DM_TAbelas.ZQMensMot.Open;
                      if DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString='' then
                         CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                      else
                         CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString;


                      CDSArqRetornoDigito.Value:=digit;
                      CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
                      CDSArqRetorno.Post;
                      hj:='';
                    end;
                  end;
                end;
              end;
           end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime     := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString    := varnome;
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;


        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
        Reset ( arq );
        ReadLn ( arq, linha );
        DBAbrir.Enabled:=true;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
      end;
    end;



    if copy(linha,77,3)='001' then
    begin
      while not DM_Tabelas.ZQContaBancaria.eof do
//    if DM_Tabelas.ZQContaBancaria.Locate('n_no_banco',copy(linha,40,7),[]) Then
      Begin
        ProgressBar1.position:= DM_Tabelas.ZQContaBancaria.recno;
        application.ProcessMessages;

        EData.DateValue:= strtodate(copy(linha,95,2)+'/'+copy(linha,97,2)+'/20'+copy(linha,99,2));
        Label1.caption := 'Banco: '+copy(linha,80,15)+ ' - '+inttostr(strtoint(copy(linha,32,8)));
        Label2.caption := 'Gerado em : '+EData.DateText;
        Label3.Caption := OpenDialog1.FileName;

        // ler o arquivo texto
        ArqLido := 0;
        ArqSize := FileSize(Arq)*128;

        ProgressBar2.Position:=0;
        ProgressBar2.Max := 104; //Round(fileSize(Arq) * 1.7);
        ProgressBar2.Visible:=true;
        While not eof (arq) do
        Begin
          ReadLn ( arq, linha );
          ArqLido := ArqLido +Length(Linha)+2; // 2 = CR+LN
          ProgressBar2.Position:=Trunc(ArqLido/ArqSize *100);
          application.ProcessMessages;
          datacredito.Clear;
//          dg:=alltrim(copy(linha,62,1));
          if copy(linha,1,1) = '7' Then
          Begin
              dtscr:=alltrim(copy(linha,176,6));
              if not empty(dtscr) and (dtscr<>'000000') then
              begin
                hj:='N';
                datacredito.DateValue:= strtodate(copy(linha,176,2)+'/'+copy(linha,178,2)+'/20'+copy(linha,180,2));
                datacredito.DateText;
              end
              else
              begin
                hj:='S';
                datacredito.datevalue:=date;
              end;
              datacredito.datevalue:=datacredito.datevalue-1;
              // verifica o dia da semana
              if DayOfWeek(datacredito.datevalue) = 7 then
              begin
                datacredito.datevalue := datacredito.datevalue - 1;
              end
              else if DayOfWeek(datacredito.datevalue) = 1 then
              begin
                datacredito.datevalue := datacredito.datevalue - 2;
              end;

              // coloquei no dia 13/10/2010 ref. data feriado
              datacredito.datevalue := EData.DateValue;


//            vartext := alltrim(copy(linha,38,25));
              vartext := alltrim(copy(linha,39,25));
              // coloquei essas 4 linhas seguinte, pq a laura recebe retorno de boletos nao gerados pelo sistema
              if (Trim(RemoveNumeros(vartext)) <> '') or (Pos(' ', Trim(vartext)) > 0) then
              begin
                vartext :='';
              end;

              VarDif  := '';
              posi:=pos('|',vartext);
              if not empty(vartext) then
              begin
                if posi=0 then
                begin
                  CDSArqRetorno.append;
                  CDSArqRetornobanco.Value:=zbco;

                  if SomenteNumeros(vartext) and TryStrToInt(Copy(vartext, 1, 11), vidrece) then
                  begin
                    CDSArqRetornoidrece.Value := vidrece;
                  end
                  else
                  begin
                    vartext := '';
                    CDSArqRetornoidrece.Clear; // se quiser deixar vazio quando inválido
                  end;

                 // CDSArqRetornoidrece.Value := strtoint(trim(copy(vartext,1,11)));
                  ZQReceb.Close;
                  ZQReceb.SQL.Clear;
                  ZQReceb.SQL.Add('Select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
                  ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
                  ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
                  ZQReceb.Open;

                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    nn:= copy(linha,64,17);
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').AsString;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').AsFloat;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').AsDateTime;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').AsString;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').AsString;
                    if hj='N' then
                       CDSArqRetornodt_credito.value := datacredito.DateValue;
                    CDSArqRetornonossonumero.Value :=nn;
                    CDSArqRetornoident_ocorr.Value := copy(linha,109,2);
                    CDSArqRetornoident_motivo.Value := '00';
                    CDSArqRetornovrarq.Value := strtofloat(copy(linha,254,13))/100;
                    CDSArqRetornovrjuros.Value := strtofloat(copy(linha,202,13))/100 + strtofloat(copy(linha,267,13))/100;
                    CDSArqRetornovrdesconto.Value := strtofloat(copy(linha,241,13))/100;
                     //  DM_TAbelas.ZQMensMot.Locate('codmotivo,retorno',VarArrayOf([CDSArqRetornoident_ocorr.Value,CDSArqRetornoident_motivo.Value]),[]);
                    DM_TAbelas.ZQMensMot.close;
                    DM_TAbelas.ZQMensMot.SQL.Clear;
                    DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                    DM_TAbelas.ZQMensMot.Open;
                    if DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString='' then
                       CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                    else
                       CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString;
                  end;

                  CDSArqRetornoDigito.Value:=digit;
                  CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
                  CDSArqRetorno.Post;
                  hj:='';
                end;
              end
              else
              begin
                mensagem('Código do Recebimento em Branco ou inválido. Coluna 39.');
              end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := varnome;
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;

        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
        Reset ( arq );
        ReadLn ( arq, linha );
        DBAbrir.Enabled:=true;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
      end;
    end;
    if copy(linha,77,3)='237' then
    begin
      while not DM_Tabelas.ZQContaBancaria.eof do
//    if DM_Tabelas.ZQContaBancaria.Locate('n_no_banco',copy(linha,40,7),[]) Then
      Begin
        ProgressBar1.position:= DM_Tabelas.ZQContaBancaria.recno;
        application.ProcessMessages;

        EData.DateValue:= strtodate(copy(linha,95,2)+'/'+copy(linha,97,2)+'/20'+copy(linha,99,2));
        Label1.caption := 'Banco: '+copy(linha,80,15)+ ' - '+copy(linha,40,07);
        Label2.caption := 'Gerado em : '+EData.DateText;
        Label3.Caption := OpenDialog1.FileName;

        // ler o arquivo texto
        ArqLido := 0;
        ArqSize := FileSize(Arq) *128;

        ProgressBar2.Position:=0;
        ProgressBar2.Max := 104; //Round(fileSize(Arq) * 1.7);
        ProgressBar2.Visible:=true;
        While not eof (arq) do
        Begin
          ReadLn ( arq, linha );
          ArqLido := ArqLido +Length(Linha)+2; // 2 = CR+LN
          ProgressBar2.Position:=Trunc(ArqLido/ArqSize *100);
          application.ProcessMessages;
          datacredito.Clear;
          dg:=alltrim(copy(linha,62,1));
          if copy(linha,1,1) = '1' Then
          Begin
            if digit=dg then
            begin
              dtscr:=alltrim(copy(linha,296,6));
              if not empty(dtscr) and (dtscr<>'000000') then
              begin
                hj:='N';
                datacredito.DateValue:= strtodate(copy(linha,296,2)+'/'+copy(linha,298,2)+'/20'+copy(linha,300,2));
                datacredito.DateText;
              end
              else
              begin
                hj:='S';
                datacredito.datevalue:=date;
              end;
              datacredito.datevalue:=datacredito.datevalue-1;
              // verifica o dia da semana
              if DayOfWeek(datacredito.datevalue) = 7 then
              begin
                datacredito.datevalue := datacredito.datevalue - 1;
              end
              else if DayOfWeek(datacredito.datevalue) = 1 then
              begin
                datacredito.datevalue := datacredito.datevalue - 2;
              end;

              // coloquei no dia 13/10/2010 ref. data feriado
              datacredito.datevalue := EData.DateValue;


//            vartext := alltrim(copy(linha,38,25));
              vartext := alltrim(copy(linha,38,24));
              // coloquei essas 4 linhas seguinte, pq a laura recebe retorno de boletos nao gerados pelo sistema
              if alltrim(RemoveNumeros(alltrim(vartext)))<>'' then
              begin
                vartext :='';
              end;

              // coloquei essas 4 linhas seguinte, pq a laura recebe retorno de boletos nao gerados pelo sistema
              if alltrim(RemoveNumeros(alltrim(vartext)))<>'' then
              begin
                vartext :='';
              end;

              VarDif  := alltrim(copy(linha,62,1));
              posi:=pos('|',vartext);
              if not empty(vartext) then
              begin
                if posi=0 then
                begin
                  CDSArqRetorno.append;
                  CDSArqRetornobanco.Value:=zbco;

                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,11)));
                  ZQReceb.Close;
                  ZQReceb.SQL.Clear;
                  ZQReceb.SQL.Add('select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
                  ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
                  ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
                  ZQReceb.Open;

                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').AsString;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').AsFloat;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').AsDateTime;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').AsString;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').AsString;
                    if hj='N' then
                       CDSArqRetornodt_credito.value := datacredito.DateValue;
                    CDSArqRetornonossonumero.Value :=copy(linha,71,12);
                    CDSArqRetornoident_ocorr.Value := copy(linha,109,2);
                    CDSArqRetornoident_motivo.Value := copy(linha,319,2);
                    CDSArqRetornovrarq.Value := strtofloat(copy(linha,254,13))/100;
                    CDSArqRetornovrjuros.Value := strtofloat(copy(linha,202,13))/100 + strtofloat(copy(linha,267,13))/100;
                    CDSArqRetornovrdesconto.Value := strtofloat(copy(linha,241,13))/100;
                     //  DM_TAbelas.ZQMensMot.Locate('codmotivo,retorno',VarArrayOf([CDSArqRetornoident_ocorr.Value,CDSArqRetornoident_motivo.Value]),[]);
                    DM_TAbelas.ZQMensMot.close;
                    DM_TAbelas.ZQMensMot.SQL.Clear;
                    DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                    DM_TAbelas.ZQMensMot.Open;
                    if DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString='' then
                       CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                    else
                       CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString;
                  end;
                  CDSArqRetornoDigito.Value:=digit;
                  CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
                  CDSArqRetorno.Post;
                  hj:='';
                end;
              end
              else
              begin
             //   mensagem('Código do Recebimento em Branco. Coluna 38.');
              end;
            end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;

        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := varnome;
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;

        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
        Reset ( arq );
        ReadLn ( arq, linha );
        DBAbrir.Enabled:=true;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;

      end;
    end;
    if copy(linha,77,3)='033' then // santander 400
    begin
      while not DM_Tabelas.ZQContaBancaria.eof do
//    if DM_Tabelas.ZQContaBancaria.Locate('n_no_banco',copy(linha,40,7),[]) Then
      Begin
        ProgressBar1.position:= DM_Tabelas.ZQContaBancaria.recno;
        application.ProcessMessages;

        EData.DateValue:= strtodate(copy(linha,95,2)+'/'+copy(linha,97,2)+'/20'+copy(linha,99,2));
        Label1.caption := 'Banco: '+copy(linha,80,15)+ ' - '+inttostr(strtoint(copy(linha,31,8)));
        Label2.caption := 'Gerado em : '+EData.DateText;
        Label3.Caption := OpenDialog1.FileName;

        // ler o arquivo texto
        ArqLido := 0;
        ArqSize := FileSize(Arq) *128;

        ProgressBar2.Position:=0;
        ProgressBar2.Max := 104; //Round(fileSize(Arq) * 1.7);
        ProgressBar2.Visible:=true;
        While not eof (arq) do
        Begin
          ReadLn ( arq, linha );
          ArqLido := ArqLido +Length(Linha)+2; // 2 = CR+LN
          ProgressBar2.Position:=Trunc(ArqLido/ArqSize *100);
          application.ProcessMessages;
          datacredito.Clear;
//          dg:=alltrim(copy(linha,62,1));
          if copy(linha,1,1) = '1' Then
          Begin
              dtscr:=alltrim(copy(linha,296,6));
              if not empty(dtscr) and (dtscr<>'000000') then
              begin
                hj:='N';
                datacredito.DateValue:= strtodate(copy(linha,296,2)+'/'+copy(linha,298,2)+'/20'+copy(linha,300,2));
                datacredito.DateText;
              end
              else
              begin
                hj:='S';
                datacredito.datevalue:=date;
              end;
              datacredito.datevalue:=datacredito.datevalue-1;
              // verifica o dia da semana
              if DayOfWeek(datacredito.datevalue) = 7 then
              begin
                datacredito.datevalue := datacredito.datevalue - 1;
              end
              else if DayOfWeek(datacredito.datevalue) = 1 then
              begin
                datacredito.datevalue := datacredito.datevalue - 2;
              end;

              // coloquei no dia 13/10/2010 ref. data feriado
              datacredito.datevalue := EData.DateValue;


              vartext := alltrim(copy(linha,38,25));
              // coloquei essas 4 linhas seguinte, pq a laura recebe retorno de boletos nao gerados pelo sistema
              if alltrim(RemoveNumeros(alltrim(vartext)))<>'' then
              begin
                vartext :='';
              end;

              VarDif  := '';
              posi:=pos('|',vartext);
              if not empty(vartext) then
              begin
                if posi=0 then
                begin
                  CDSArqRetorno.append;
                  CDSArqRetornobanco.Value:=zbco;

                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,11)));
                  ZQReceb.Close;
                  ZQReceb.SQL.Clear;
                  ZQReceb.SQL.Add('select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
                  ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
                  ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
                  ZQReceb.Open;

                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    nn:= copy(linha,63,8);
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').AsString;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').AsFloat;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').AsDateTime;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').AsString;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').AsString;
                    if hj='N' then
                       CDSArqRetornodt_credito.value := datacredito.DateValue;
                    CDSArqRetornonossonumero.Value :=nn;
                    CDSArqRetornoident_ocorr.Value := copy(linha,109,2);
                    CDSArqRetornoident_motivo.Value := '00';
                    CDSArqRetornovrarq.Value := strtofloat(copy(linha,254,13))/100;
                    CDSArqRetornovrjuros.Value := strtofloat(copy(linha,202,13))/100 + strtofloat(copy(linha,267,13))/100;
                    CDSArqRetornovrdesconto.Value := strtofloat(copy(linha,241,13))/100;
                     //  DM_TAbelas.ZQMensMot.Locate('codmotivo,retorno',VarArrayOf([CDSArqRetornoident_ocorr.Value,CDSArqRetornoident_motivo.Value]),[]);
                    DM_TAbelas.ZQMensMot.close;
                    DM_TAbelas.ZQMensMot.SQL.Clear;
                    DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                    DM_TAbelas.ZQMensMot.Open;
                    if DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString='' then
                       CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                    else
                       CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString;
                  end;

                  CDSArqRetornoDigito.Value:=digit;
                  CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
                  CDSArqRetorno.Post;
                  hj:='';
                end;
              end
              else
              begin
                mensagem('Código do Recebimento em Branco. Coluna 39.');
              end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;

        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := varnome;
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;

        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
        Reset ( arq );
        ReadLn ( arq, linha );
        DBAbrir.Enabled:=true;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
      end;
    end;


    if copy(linha,1,3)='033' then  // santander 240
    begin
      while not DM_Tabelas.ZQContaBancaria.eof do
//    if DM_Tabelas.ZQContaBancaria.Locate('n_no_banco',copy(linha,40,7),[]) Then
      Begin
        ProgressBar1.position:= DM_Tabelas.ZQContaBancaria.recno;
        application.ProcessMessages;

        EData.DateValue:= StringToDateTimeDef(Copy(Linha,144,2)+'/'+
                                          Copy(Linha,146,2)+'/'+
                                          Copy(Linha,148,4),0, 'DD/MM/YYYY' );
//        for varx := 1 to length(OpenDialog1.FileName) do
//           if copy(OpenDialog1.FileName,varx,1)='/' Then
//              break;
//        varnome := copy(OpenDialog1.FileName,varx-12,12);
//        vardata := copy(edata.DateText,7,4)+'-'+copy(edata.DateText,4,2)+'-'+copy(Edata.DateText,1,2);
        Label1.caption := 'Banco: '+trim(copy(linha,103,30))+ ' - '+inttostr(strtoint(copy(linha,38,9)));
        Label2.caption := 'Gerado em : '+EData.DateText;
        Label3.Caption := varnome;

        // ler o arquivo texto
        ArqLido := 0;
        ArqSize := FileSize(Arq)*128;

        ProgressBar2.Position:=0;
        ProgressBar2.Max := 104; //Round(fileSize(Arq) * 1.7);
        ProgressBar2.Visible:=true;
        ReadLn ( arq, linha );

        While not eof (arq) do
        Begin
          ReadLn ( arq, linha );
          ArqLido := ArqLido +Length(Linha)+2; // 2 = CR+LN
          ProgressBar2.Position:=Trunc(ArqLido/ArqSize *100);
          application.ProcessMessages;
          if Copy(Linha,14,1)= 'T' then
          begin
            vartext := alltrim(copy(linha,101,25));
            // coloquei essas 4 linhas seguinte, pq a laura recebe retorno de boletos nao gerados pelo sistema
            if alltrim(RemoveNumeros(alltrim(vartext)))<>'' then
            begin
              vartext :='';
            end;

            VarDif  := '';
            posi:=pos('|',vartext);
            if not empty(vartext) then
            begin
              if posi=0 then
              begin
  //              posi:=length(vartext);
//                posi:=posi-2;
        //        if alltrim(copy(vartext,1,posi))<>emptystr then
                if alltrim(vartext)<>emptystr then
                begin
                  CDSArqRetorno.append;
                  CDSArqRetornobanco.Value:=zbco;

//                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,posi)));
                  CDSArqRetornoidrece.Value := strtoint(alltrim(vartext));
                  ZQReceb.Close;
                  ZQReceb.SQL.Clear;
                  ZQReceb.SQL.Add('Select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
                  ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
                  ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
                  ZQReceb.Open;

                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    nn:= copy(linha,41,12);
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').AsString;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').AsFloat;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').AsDateTime;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').AsString;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').AsString;
                    CDSArqRetornonossonumero.Value :=nn;
                    CDSArqRetornoident_ocorr.Value := copy(linha,16,2);
                    CDSArqRetornoident_motivo.Value := '00';

                    ReadLn ( arq, linha );
                    if Copy(Linha,14,1)= 'U' then
                    begin
                      datacredito.Clear;
                      dtscr:=alltrim(Copy(Linha,138,2)+'/'+Copy(Linha,140,2)+'/'+Copy(Linha,142,4));   // data da ocorrencia
                     // dtscr:=alltrim(Copy(Linha,146,2)+'/'+Copy(Linha,148,2)+'/'+Copy(Linha,150,4)); // data da efetivação do credito
                      if not empty(dtscr) and (dtscr<>'00000000') then
                      begin
                        hj:='N';
                        datacredito.DateValue:= StringToDateTimeDef(Copy(Linha,138,2)+'/'+
                                                           Copy(Linha,140,2)+'/'+
                                                           Copy(Linha,142,4),0, 'DD/MM/YYYY' );
//                        datacredito.DateValue:= StringToDateTimeDef(Copy(Linha,146,2)+'/'+
//                                                           Copy(Linha,148,2)+'/'+
//                                                           Copy(Linha,150,4),0, 'DD/MM/YYYY' );

                        datacredito.DateText;
                      end
                      else
                      begin
                        hj:='S';
                        datacredito.datevalue:=date;
                      end;
                      //datacredito.datevalue:=datacredito.datevalue-1;
                      // verifica o dia da semana
                      if DayOfWeek(datacredito.datevalue) = 7 then
                      begin
                        datacredito.datevalue := datacredito.datevalue - 1;
                      end
                      else if DayOfWeek(datacredito.datevalue) = 1 then
                      begin
                        datacredito.datevalue := datacredito.datevalue - 2;
                      end;

                     // coloquei no dia 13/10/2010 ref. data feriado
                    // datacredito.datevalue := EData.DateValue;

  //                   if hj='N' then
                     CDSArqRetornodt_credito.value := datacredito.DateValue;



                      CDSArqRetornovrarq.Value := StrToFloatDef(Copy(Linha,78,15),0)/100;
                      CDSArqRetornovrjuros.Value := StrToFloatDef(Copy(Linha,18,15),0)/100;
                      CDSArqRetornovrdesconto.Value := StrToFloatDef(Copy(Linha,33,15),0)/100;
                       //  DM_TAbelas.ZQMensMot.Locate('codmotivo,retorno',VarArrayOf([CDSArqRetornoident_ocorr.Value,CDSArqRetornoident_motivo.Value]),[]);
                      DM_TAbelas.ZQMensMot.close;
                      DM_TAbelas.ZQMensMot.SQL.Clear;
                      DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                      DM_TAbelas.ZQMensMot.Open;
                      if DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString='' then
                         CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                      else
                         CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString;


                      CDSArqRetornoDigito.Value:=digit;
                      CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
                      CDSArqRetorno.Post;
                      hj:='';
                    end;
                  end;
                end;
            end;
           end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := copy(varnome,10,20);
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;


        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
        Reset ( arq );
        ReadLn ( arq, linha );
        DBAbrir.Enabled:=true;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
      end;
    end;

    if copy(linha,77,3)='399' then
    begin
      while not DM_Tabelas.ZQContaBancaria.eof do
//    if DM_Tabelas.ZQContaBancaria.Locate('n_no_banco',copy(linha,40,7),[]) Then
      Begin
        ProgressBar1.position:= DM_Tabelas.ZQContaBancaria.recno;
        application.ProcessMessages;

        EData.DateValue:= strtodate(copy(linha,95,2)+'/'+copy(linha,97,2)+'/20'+copy(linha,99,2));
        Label1.caption := 'Banco: '+copy(linha,80,15)+ ' - '+inttostr(strtoint(copy(linha,39,5)));
        Label2.caption := 'Gerado em : '+EData.DateText;
        Label3.Caption := OpenDialog1.FileName;

        // ler o arquivo texto
        ArqLido := 0;
        ArqSize := FileSize(Arq) *128;

        ProgressBar2.Position:=0;
        ProgressBar2.Max := 104; //Round(fileSize(Arq) * 1.7);
        ProgressBar2.Visible:=true;
        While not eof (arq) do
        Begin
          ReadLn ( arq, linha );
          ArqLido := ArqLido +Length(Linha)+2; // 2 = CR+LN
          ProgressBar2.Position:=Trunc(ArqLido/ArqSize *100);
          application.ProcessMessages;
          datacredito.Clear;
//          dg:=alltrim(copy(linha,62,1));
          if copy(linha,1,1) = '1' Then
          Begin
              dtscr:=alltrim(copy(linha,111,6));
              if not empty(dtscr) and (dtscr<>'000000') then
              begin
                hj:='N';
                datacredito.DateValue:= strtodate(copy(linha,111,2)+'/'+copy(linha,113,2)+'/20'+copy(linha,115,2));
                datacredito.DateText;
              end
              else
              begin
                hj:='S';
                datacredito.datevalue:=date;
              end;
              datacredito.datevalue:=datacredito.datevalue-1;
              // verifica o dia da semana
              if DayOfWeek(datacredito.datevalue) = 7 then
              begin
                datacredito.datevalue := datacredito.datevalue - 1;
              end
              else if DayOfWeek(datacredito.datevalue) = 1 then
              begin
                datacredito.datevalue := datacredito.datevalue - 2;
              end;

              // coloquei no dia 13/10/2010 ref. data feriado
              datacredito.datevalue := EData.DateValue;


              vartext := alltrim(copy(linha,38,25));
              // coloquei essas 4 linhas seguinte, pq a laura recebe retorno de boletos nao gerados pelo sistema
              if alltrim(RemoveNumeros(alltrim(vartext)))<>'' then
              begin
                vartext :='';
              end;

              VarDif  := '';
              posi:=pos('|',vartext);
              if not empty(vartext) then
              begin
                if posi=0 then
                begin
                  CDSArqRetorno.append;
                  CDSArqRetornobanco.Value:=zbco;

                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,11)));
                  ZQReceb.Close;
                  ZQReceb.SQL.Clear;
                  ZQReceb.SQL.Add('select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
                  ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
                  ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
                  ZQReceb.Open;

                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    nn:= copy(linha,63,11);
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').AsString;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').AsFloat;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').AsDateTime;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').AsString;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').AsString;
                    if hj='N' then
                       CDSArqRetornodt_credito.value := datacredito.DateValue;
                    CDSArqRetornonossonumero.Value :=nn;
                    CDSArqRetornoident_ocorr.Value := copy(linha,109,2);
                    CDSArqRetornoident_motivo.Value := '00';
                    CDSArqRetornovrarq.Value := strtofloat(copy(linha,254,13))/100;
                    CDSArqRetornovrjuros.Value := strtofloat(copy(linha,267,13))/100;
                    CDSArqRetornovrdesconto.Value := strtofloat(copy(linha,241,13))/100;
                     //  DM_TAbelas.ZQMensMot.Locate('codmotivo,retorno',VarArrayOf([CDSArqRetornoident_ocorr.Value,CDSArqRetornoident_motivo.Value]),[]);
                    DM_TAbelas.ZQMensMot.close;
                    DM_TAbelas.ZQMensMot.SQL.Clear;
                    DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                    DM_TAbelas.ZQMensMot.Open;
                    if DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString='' then
                       CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                    else
                       CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString;
                  end;

                  CDSArqRetornoDigito.Value:=digit;
                  CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
                  CDSArqRetorno.Post;
                  hj:='';
                end;
              end
              else
              begin
                mensagem('Código do Recebimento em Branco. Coluna 39.');
              end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;

        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := varnome;
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;

        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
        Reset ( arq );
        ReadLn ( arq, linha );
        DBAbrir.Enabled:=true;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
      end;
    end;


    if copy(linha,77,3)='341' then
    begin
      while not DM_Tabelas.ZQContaBancaria.eof do
//    if DM_Tabelas.ZQContaBancaria.Locate('n_no_banco',copy(linha,40,7),[]) Then
      Begin
        ProgressBar1.position:= DM_Tabelas.ZQContaBancaria.recno;
        application.ProcessMessages;

        EData.DateValue:= strtodate(copy(linha,95,2)+'/'+copy(linha,97,2)+'/20'+copy(linha,99,2));
        Label1.caption := 'Banco: '+copy(linha,80,15);
        Label2.caption := 'Gerado em : '+EData.DateText;
        Label3.Caption := OpenDialog1.FileName;

        // ler o arquivo texto
        ArqLido := 0;
        ArqSize := FileSize(Arq) *128;

        ProgressBar2.Position:=0;
        ProgressBar2.Max := 104; //Round(fileSize(Arq) * 1.7);
        ProgressBar2.Visible:=true;
        While not eof (arq) do
        Begin
          ReadLn ( arq, linha );
          ArqLido := ArqLido +Length(Linha)+2; // 2 = CR+LN
          ProgressBar2.Position:=Trunc(ArqLido/ArqSize *100);
          application.ProcessMessages;
          datacredito.Clear;
//          dg:=alltrim(copy(linha,62,1));
          if copy(linha,1,1) = '1' Then
          Begin
              // data de credito
              dtscr:=alltrim(copy(linha,296,6));
              if not empty(dtscr) and (dtscr<>'000000') then
              begin
                hj:='N';
                datacredito.DateValue:= strtodate(copy(linha,296,2)+'/'+copy(linha,298,2)+'/20'+copy(linha,300,2));
                datacredito.DateText;
              end
              else
              begin
                hj:='S';
                datacredito.datevalue:=date;
              end;
              datacredito.datevalue:=datacredito.datevalue-1;
              // verifica o dia da semana
              if DayOfWeek(datacredito.datevalue) = 7 then
              begin
                datacredito.datevalue := datacredito.datevalue - 1;
              end
              else if DayOfWeek(datacredito.datevalue) = 1 then
              begin
                datacredito.datevalue := datacredito.datevalue - 2;
              end;

              // coloquei no dia 13/10/2010 ref. data feriado
              datacredito.datevalue := EData.DateValue;


              vartext := alltrim(copy(linha,38,25));
              // coloquei essas 4 linhas seguinte, pq a laura recebe retorno de boletos nao gerados pelo sistema
              if alltrim(RemoveNumeros(alltrim(vartext)))<>'' then
              begin
                vartext :='';
              end;

              VarDif  := '';
              posi:=pos('|',vartext);
              if not empty(vartext) then
              begin
                if posi=0 then
                begin
                  CDSArqRetorno.append;
                  CDSArqRetornobanco.Value:=zbco;

                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,11)));
                  ZQReceb.Close;
                  ZQReceb.SQL.Clear;
                  ZQReceb.SQL.Add('select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
                  ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
                  ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
                  ZQReceb.Open;

                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    nn:= copy(linha,127,8);
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').AsString;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').AsFloat;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').AsDateTime;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').AsString;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').AsString;
                    if hj='N' then
                       CDSArqRetornodt_credito.value := datacredito.DateValue;
                    CDSArqRetornonossonumero.Value :=nn;
                    CDSArqRetornoident_ocorr.Value := copy(linha,109,2);
                    CDSArqRetornoident_motivo.Value := copy(linha,378,2);
                    CDSArqRetornovrarq.Value := (strtofloat(copy(linha,254,13))/100)+(strtofloatdef(copy(linha,176,13),0)/100);  // essa linha 176 e da tarifa do bco que ja vem descontado
                    CDSArqRetornovrjuros.Value := strtofloat(copy(linha,202,13))/100 + strtofloat(copy(linha,267,13))/100;
                    CDSArqRetornovrdesconto.Value := strtofloat(copy(linha,241,13))/100;
                    if CDSArqRetornovrtit.Value<>CDSArqRetornovrarq.Value then
                    begin
                      if (CDSArqRetornovrjuros.Value=0) and (CDSArqRetornovrdesconto.Value=0) then
                         CDSArqRetornovrarq.Value:= CDSArqRetornovrtit.Value;
                    end;

                     //  DM_TAbelas.ZQMensMot.Locate('codmotivo,retorno',VarArrayOf([CDSArqRetornoident_ocorr.Value,CDSArqRetornoident_motivo.Value]),[]);
                    DM_TAbelas.ZQMensMot.close;
                    DM_TAbelas.ZQMensMot.SQL.Clear;
                    DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                    DM_TAbelas.ZQMensMot.Open;
                    if DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString='' then
                       CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                    else
                       CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').AsString;
                  end;

                  CDSArqRetornoDigito.Value:=digit;
                  CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
                  CDSArqRetorno.Post;
                  hj:='';
                end;
              end
              else
              begin
                mensagem('Código do Recebimento em Branco. Coluna 39.');
              end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;

        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select idmens_motivo,retorno,codmotivo,mensagem from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').AsDateTime := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString := varnome;
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').AsString := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;

        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').AsString;
        Reset ( arq );
        ReadLn ( arq, linha );
        DBAbrir.Enabled:=true;
        CDSArqRetorno.EnableControls;
        CDSArqRetorno.First;
      end;
    end;

    Label4.Caption := inttostr(CDSArqRetorno.RecordCount);
    CloseFile ( arq );

   // DBGBaixaAuto.Refresh;
//    DBGBaixaAuto.Repaint;
    
    application.ProcessMessages;
    ProgressBar1.position:=0;
    ProgressBar1.visible:=false;


    DM_Tabelas.ZQContaBancaria.Filtered:=false;
    DM_Tabelas.ZQContaBancaria.close;
    DM_Tabelas.ZQContaBancaria.SQL.Clear;
    DM_Tabelas.ZQContaBancaria.SQL.Add('select idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,nometitular,aberta,senha_interna,senha_conta,ativa,jurosemboleto,taxadiaria,valoroupercent,carteira,boleto,idparticipante,n_dif_empreed,convenio,');
    DM_Tabelas.ZQContaBancaria.SQL.Add('doc_titular,n_sequencial,nosso_numero,apelido,cod_transmissao,complemento,moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote from conta_bancaria');
    DM_Tabelas.ZQContaBancaria.Open;
    DM_Tabelas.ZQBancRemes.Refresh;
    if DM_Tabelas.ZQRemes_Receb.Active=false then
       DM_Tabelas.ZQRemes_Receb.open;

    DXBProcessar.Enabled:=true;
    CDSArqRetorno.EnableControls;
    CDSArqRetorno.First;

  end;
end;

procedure TFrm_BaixaAutomatica.FormShow(Sender: TObject);
begin
  DBAbrir.enabled:=true;
  DM_Tabelas.ZQMensMot.Open;
  //ZQReceb.open;
  DM_Tabelas.ZQContaBancaria.open;
  DM_Tabelas.ZQMensMot.open;
  DM_TAbelas.ZQBancRemes2.open;
//  DM_Tabelas.ZQRecebimento.Open;
  DM_tabelas.ZQTipodoc.Open;
  DM_Tabelas.ZQCobaRe.Open;
//  DM_Tabelas.ZQReBxHi.Open;
//  DM_tabelas.ZQReceb_Baixa.Open;
  DM_tabelas.ZQCaixa.Open;
  DM_Tabelas.ZQPlanodeContas.open;
  DM_Tabelas.ZQMens_Motivo.Open;
  DM_Tabelas.ZQMens_Retorno.Open;
  DM_Tabelas.ZQBancRemes.open;
  if DM_Tabelas.ZQRemes_Receb.Active=false then
     DM_Tabelas.ZQRemes_Receb.open;
  DXBImprimir.Enabled := False;
  DXBProcessar.Enabled := False;
  DBAbrir.SetFocus;
  CDSArqRetorno.Close;
  if DM_Tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('bxautoplc').AsInteger,[]) Then
    Ebxautoplc.Text := DM_Tabelas.ZQPlanoDeContas.FieldByName('codigo').Text+' - '+DM_Tabelas.ZQPlanoDeContas.FieldByName('descricao').AsString;
//  CDSArqRetorno.CreateDataSet;
end;

procedure TFrm_BaixaAutomatica.dxButton1Click(Sender: TObject);
begin
  if not Verif_senha('Recebimento','Mensagens de Retorno ','') then exit;
  if Frm_TabelasDeMensagens=nil then
     Frm_TabelasDeMensagens:=TFrm_TabelasDeMensagens.Create(Self);
  Frm_TabelasDeMensagens.Left := Frm_BaixaAutomatica.Left;
  Frm_TabelasDeMensagens.Top := Frm_BaixaAutomatica.Top+60;
  Frm_TabelasDeMensagens.showmodal;
  Frm_TabelasDeMensagens:=nil;
end;

procedure TFrm_BaixaAutomatica.DBGBaixaAutoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if Column.Field = CDSArqRetornoacao then begin
    DBGBaixaAuto.Canvas.FillRect(Rect);
    if CDSArqRetornoacao.Value = 'B' Then
      DM_TAbelas.ImageList1.draw(DBGBaixaAuto.Canvas, Rect.Left + 5, Rect.Top + 1, 0)
    else if CDSArqRetornoacao.Value = 'R' Then
      DM_TAbelas.ImageList1.draw(DBGBaixaAuto.Canvas, Rect.Left + 5, Rect.Top + 1, 1)
    else if (CDSArqRetornoacao.Value = 'I') or (empty(CDSArqRetornoacao.Value)) Then
      DM_TAbelas.ImageList1.draw(DBGBaixaAuto.Canvas, Rect.Left + 5, Rect.Top + 1, 3);
  end;
end;

procedure TFrm_BaixaAutomatica.DXBImprimirClick(Sender: TObject);
begin
  if Frm_RelRetorno=nil then
     Frm_RelRetorno:= TFrm_RelRetorno.Create(Application);
  Frm_RelRetorno.RLLabel1.Caption := 'Retorno do Banco '+Label1.Caption;
  Frm_RelRetorno.RLLabel2.Caption := OpenDialog1.FileName+'  referente '+EData.DateText;
  Frm_RelRetorno.RLReport1.PreviewModal;
  Frm_RelRetorno:=nil;
end;

procedure TFrm_BaixaAutomatica.DXBProcessarClick(Sender: TObject);
Var
  xref,Varx : integer;
  bco,ctbl,perc,varnome, vardata : string;
  saldo:double;
begin
  DXBProcessar.Enabled:=false;
  processou:='S';
  if not Verif_senha('Recebimento','Processa baixa automática','') then exit;
  DM_TAbelas.ZQBancRemes2.Open;
{  for varx := length(OpenDialog1.FileName) to 1 do
    if copy(OpenDialog1.FileName,varx,1)='/' Then
      break;
  varnome := copy(OpenDialog1.FileName,varx-11,12);}
  bco:=alltrim(CDSArqRetornobanco.Value);
   if alltrim(CDSArqRetornobanco.Value)='756' then
   begin
      for varx := 1 to length(OpenDialog1.FileName) do
          if copy(OpenDialog1.FileName,varx,1)='/' Then
             break;
      varnome := copy(OpenDialog1.FileName,varx-34,34);
   end
   else
   if trim(CDSArqRetornobanco.Value)='104' then   //caixa 240
   begin
     for varx := 1 to length(OpenDialog1.FileName) do
       if copy(OpenDialog1.FileName,varx,1)='/' Then
          break;
     varnome := copy(OpenDialog1.FileName,varx-20,20);
//     varnome := copy(varnome,15,20);
   end
   else
   begin
     for varx := 1 to length(OpenDialog1.FileName) do
     if copy(OpenDialog1.FileName,varx,1)='/' Then
        break;
     varnome := copy(OpenDialog1.FileName,varx-12,12);
     if lbl240.Caption='033' then //cnab 240
        varnome := copy(OpenDialog1.FileName,varx-20,20);
   end;
//    if DM_Tabelas.ZQRemes_Receb.Active=false then
//       DM_Tabelas.ZQRemes_Receb.open;

  vardata := copy(edata.DateText,7,4)+'-'+copy(edata.DateText,4,2)+'-'+copy(Edata.DateText,1,2);

  {DM_TAbelas.ZQBancRemes2.Filtered:=false;
  DM_TAbelas.ZQBancRemes2.Filter:='nomearq='+quotedstr(varnome)+' and gerado='+quotedstr(vardata);
  DM_TAbelas.ZQBancRemes2.Filtered:=true;
  if DM_TAbelas.ZQBancRemes2.RecordCount>0 then
  begin
    if DM_TAbelas.ZQBancRemes.locate('nomearq;gerado',VarArrayOf([varnome,vardata]),[]) Then Begin
       mensagem('Este arquivo ja foi baixado no dia '+DM_Tabelas.ZQBancRemes2gerado.Text+', ele não pode mais ser baixado!!!!');
      // mensagem('Existe baixa com o mesmo nome do arquivo no dia '+DM_Tabelas.ZQBancRemes2gerado.Text+' prosseguindo...');
       DM_TAbelas.ZQBancRemes2.close;
       exit;
    end
    else
    begin
      DM_TAbelas.ZQBancRemes.Open;
      DM_TAbelas.ZQBancRemes.Insert;
      DM_Tabelas.ZQBancRemesgerado.Value := Edata.DateValue;
      DM_Tabelas.ZQBancRemesnomearq.Value := varnome;
      DM_Tabelas.ZQBancRemes.Post;
    end;
    DM_TAbelas.ZQBancRemes2.Filtered:=false; }



  DM_TAbelas.ZQBancRemes2.close;
  DM_Tabelas.ZQRemesRec.Close;
//  DM_Tabelas.ZQRemesRec.SQL.Clear;
//  DM_Tabelas.ZQRemesRec.SQL.Add('select idremessa_receb,remessa,idrec,retorno_motivo,acao,nossonumero,ocorrencia,desc_motivo,credito,dt_ocorrencia from remessa_receb');
//  DM_Tabelas.ZQRemesRec.Open;

  CDSArqRetorno.First;
  ProgressBar1.Position:=0;
  ProgressBar1.Max:=CDSArqRetorno.RecordCount;
  ProgressBar1.Visible:=true;
  CDSArqRetorno.DisableControls;



  while not CDSArqRetorno.Eof do
  begin
    saldo:=0;
    application.ProcessMessages;
    ProgressBar1.Position:=CDSArqRetorno.RecNo;
    if CDSArqRetornoacao.Value = 'R' Then
    Begin
{      DM_Tabelas.ZQRemesRec.SQL.Clear;
      DM_Tabelas.ZQRemesRec.SQL.Add('Select * from remessa_receb where idrec='+CDSArqRetornoidrece.Text+' and acao is null');
      DM_Tabelas.ZQRemesRec.Open;
      if DM_Tabelas.ZQRemesRec.RecordCount>0 Then Begin
        DM_Tabelas.ZQRemesRec.Edit;
        DM_TAbelas.ZQRemesRecacao.Value := 'R';
        DM_TAbelas.ZQRemesRecretorno_motivo.Value := CDSArqRetornoident_ocorr.Value+CDSArqRetornoident_motivo.Value;
        DM_Tabelas.ZQRemesRec.Post;}

        DM_Tabelas.ZQRemesRec.Close;
        DM_Tabelas.ZQRemesRec.SQL.Clear;
        DM_Tabelas.ZQRemesRec.SQL.Add('update remessa_receb set acao=''R'', retorno_motivo=('+quotedstr(CDSArqRetornoident_ocorr.Value+CDSArqRetornoident_motivo.Value)+') where idrec='+CDSArqRetornoidrece.Text+' and acao is null');
        DM_Tabelas.ZQRemesRec.ExecSQL;

        DM_Tabelas.ZQRemesRec.Close;
        DM_Tabelas.ZQRemesRec.SQL.Clear;
        DM_Tabelas.ZQRemesRec.SQL.Add('select idremessa_receb,remessa,idrec,retorno_motivo,acao,nossonumero,ocorrencia,desc_motivo,credito,dt_ocorrencia from remessa_receb');
       // DM_Tabelas.ZQRemesRec.Open;


      //end;
      ZQReceb.Close;
      ZQReceb.SQL.Clear;
      ZQReceb.SQL.Add('update recebimento set observ='+quotedstr(datetostr(date))+quotedstr(' ')+quotedstr(CDSArqRetornodesc_ocorr.Value)+quotedstr(' ')+quotedstr(copy(CDSArqRetornodesc_motivo.Value,1,100))+' where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text)));
      ZQReceb.ExecSQL;

//      ZQReceb.Edit;
//      ZQRecebobserv.Value := datetostr(date)+chr(13)+CDSArqRetornodesc_ocorr.Value+chr(13)+CDSArqRetornodesc_motivo.Value;
//      ZQReceb.Post;
    end
    else if (CDSArqRetornoacao.Value = 'B') and (CDSArqRetornovrarq.Value>0) Then Begin
      ZQReceb.Close;
      ZQReceb.SQL.Clear;
      ZQReceb.SQL.Add('select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
      ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
      ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
      ZQReceb.Open;

      if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
      begin
        DM_Tabelas.ZQReBxHi.Close;
        DM_Tabelas.ZQReBxHi.SQL.Clear;
        DM_Tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,idrecib,refer,valor,descon,juros,percent_usado,data,sq,valor_parcela from RecBxHist where idrecib='+quotedstr(CDSArqRetornoidrece.text));
        DM_Tabelas.ZQReBxHi.open;
        xref:=DM_Tabelas.ZQReBxHi.FieldByName('refer').AsInteger;

//        if ZQRecebsaldo.Value>0 then
        begin
          saldo:=ZQReceb.FieldByName('saldo').AsFloat;

          DM_tabelas.ZQCobaRe.Append;
          DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
          DM_Tabelas.ZQCobaRe.FieldByName('entrada').AsDateTime := EData.DateValue;
          DM_Tabelas.ZQCobaRe.Post;

          perc:='100';


          DM_Tabelas.ZQReBxHi.Close;
          DM_Tabelas.ZQReBxHi.SQL.Clear;
          DM_Tabelas.ZQReBxHi.SQL.Add('delete from RecBxHist where idrecib='+quotedstr(CDSArqRetornoidrece.text));
          DM_Tabelas.ZQReBxHi.ExecSQL;


          if ZQReceb.FieldByName('tip').AsString<>'J' then
          begin
            DM_Tabelas.ZQReBxHi.Close;
            DM_Tabelas.ZQReBxHi.SQL.Clear;
//            DM_Tabelas.ZQReBxHi.SQL.Add('INSERT IGNORE INTO RecBxHist (juros,valor,data,percent_usado,refer,idrecib,valor_parcela) VALUES ('+quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrjuros.text,'.')))+','+
            DM_Tabelas.ZQReBxHi.SQL.Add('INSERT INTO RecBxHist (descon,juros,valor,data,percent_usado,refer,idrecib,valor_parcela) VALUES ('+quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrdesconto.text,'.')))+','+quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrjuros.text,'.')))+','+
                                    quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrarq.text,'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',CDSArqRetornodt_credito.Value))+','+quotedstr(perc)+','+quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').text)+','+
                                    quotedstr(CDSArqRetornoidrece.text)+','+quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrtit.text,'.')))+')');
          end
          else
          begin
            DM_Tabelas.ZQReBxHi.Close;
            DM_Tabelas.ZQReBxHi.SQL.Clear;
//            DM_Tabelas.ZQReBxHi.SQL.Add('INSERT IGNORE INTO RecBxHist (juros,valor,data,percent_usado,refer,idrecib) VALUES ('+quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrarq.text,'.')))+','+
            DM_Tabelas.ZQReBxHi.SQL.Add('INSERT INTO RecBxHist (juros,valor,data,percent_usado,refer,idrecib) VALUES ('+quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrarq.text,'.')))+','+
                                    quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrarq.text,'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',CDSArqRetornodt_credito.Value))+','+quotedstr(perc)+','+quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').text)+','+
                                    quotedstr(CDSArqRetornoidrece.text)+')');
          end;
          DM_Tabelas.ZQReBxHi.ExecSQL;
          DM_Tabelas.ZQReBxHi.Close;


{         DM_Tabelas.ZQReBxHi.Insert;
          DM_Tabelas.ZQReBxHijuros.Value := CDSArqRetornovrjuros.Value;
          DM_Tabelas.ZQReBxHivalor.Value := CDSArqRetornovrarq.Value;
          DM_Tabelas.ZQReBxHidata.Value := CDSArqRetornodt_credito.value; //EData.DateValue;
          DM_Tabelas.ZQReBxHipercent_usado.Value := 100;
          DM_tabelas.ZQReBxHirefer.Value := DM_tabelas.ZQCobaRecotagem.Value;
          DM_tabelas.ZQReBxHiidrecib.Value := CDSArqRetornoidrece.Value;
          DM_tabelas.ZQReBxHivalor_parcela.Value := CDSArqRetornovrtit.Value;
          DM_Tabelas.ZQReBxHi.Post;


          DM_tabelas.ZQReceb_Baixa.Insert;
          DM_tabelas.ZQReceb_BaixaVr_rec.Value := CDSArqRetornovrarq.Value;
          DM_tabelas.ZQReceb_BaixaDt_rec.Value := EData.DateValue;
          DM_TAbelas.ZQReceb_Baixadataref.Value := CDSArqRetornodt_credito.value; //EData.DateValue;
          DM_TAbelas.ZQReceb_Baixavencimento.Value := CDSArqRetornodtvenci.Value;
          DM_tabelas.ZQReceb_BaixaTipdoc.Value := 'BB';
//        DM_tabelas.ZQReceb_BaixaTipdoc.Value := DM_Tabelas.ZQconfiguracoesbxautodoc.Value;
          DM_tabelas.ZQReceb_BaixaDocum.Value  := Label3.Caption;
          DM_tabelas.ZQReceb_BaixaJuros_Vr.Value := CDSArqRetornovrjuros.Value;
          DM_tabelas.ZQReceb_BaixaDesc_Vr.Value  := CDSArqRetornovrdesconto.Value;
          DM_tabelas.ZQReceb_Baixarefbaixa.Value := DM_tabelas.ZQCobaRecotagem.Value;
          DM_tabelas.ZQReceb_Baixacodcontabil.Value := DM_Tabelas.ZQconfiguracoesbxautoplc.Value;
          DM_tabelas.ZQReceb_BaixaObsebx.Value := 'Baixa Automática';
          DM_tabelas.ZQReceb_Baixavalor_parcela.Value := CDSArqRetornovrtit.Value;
          DM_tabelas.ZQReceb_Baixa.Post;}

          if DM_Tabelas.ZQConfiguracoes.FieldByName('bxautoplc').AsInteger=0 then
             ctbl:='0'
          else
             ctbl:=DM_Tabelas.ZQConfiguracoes.FieldByName('bxautoplc').Text;


          DM_tabelas.ZQReceb_Baixa.Close;
          DM_tabelas.ZQReceb_Baixa.SQL.Clear;
          DM_tabelas.ZQReceb_Baixa.SQL.Add('delete from  Receb_Baixa where refbaixa='+quotedstr(inttostr(xref)));
    //      mensagem(DM_tabelas.ZQReceb_Baixa.SQL.Text);
          DM_tabelas.ZQReceb_Baixa.ExecSQL;
          xref:=0;


          if ZQReceb.FieldByName('tip').AsString<>'J' then
          begin
            DM_tabelas.ZQReceb_Baixa.Close;
            DM_tabelas.ZQReceb_Baixa.SQL.Clear;
//            DM_tabelas.ZQReceb_Baixa.SQL.Add('INSERT IGNORE INTO  Receb_Baixa (Vr_rec,Dt_rec,dataref,vencimento,Tipdoc,Docum,Juros_Vr,Desc_Vr,refbaixa,codcontabil,Obsebx,valor_parcela) VALUES ('+
            DM_tabelas.ZQReceb_Baixa.SQL.Add('INSERT INTO  Receb_Baixa (Vr_rec,Dt_rec,dataref,vencimento,Tipdoc,Docum,Juros_Vr,Desc_Vr,refbaixa,codcontabil,Obsebx,valor_parcela) VALUES ('+
                                            quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrarq.Text,'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',EData.DateValue))+','+quotedstr(FormatDateTime('YYYY-MM-DD',CDSArqRetornodt_credito.value))+','+
                                            quotedstr(FormatDateTime('YYYY-MM-DD',CDSArqRetornodtvenci.Value))+','+quotedstr('BB')+','+quotedstr(alltrim(copy(Label3.Caption,1,19)))+','+quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrjuros.Text,'.')))+','+
                                            quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrdesconto.Text,'.')))+','+quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(ctbl)+','+quotedstr('Baixa Automática')+','+
                                            quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrtit.Text,'.')))+')');

          end
          else
          begin

            DM_tabelas.ZQReceb_Baixa.Close;
            DM_tabelas.ZQReceb_Baixa.SQL.Clear;
//            DM_tabelas.ZQReceb_Baixa.SQL.Add('INSERT IGNORE INTO  Receb_Baixa (Vr_rec,Dt_rec,dataref,vencimento,Tipdoc,Docum,Juros_Vr,Desc_Vr,refbaixa,codcontabil,Obsebx) VALUES ('+
            DM_tabelas.ZQReceb_Baixa.SQL.Add('INSERT INTO  Receb_Baixa (Vr_rec,Dt_rec,dataref,vencimento,Tipdoc,Docum,Juros_Vr,Desc_Vr,refbaixa,codcontabil,Obsebx) VALUES ('+
                                            quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrarq.Text,'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',EData.DateValue))+','+quotedstr(FormatDateTime('YYYY-MM-DD',CDSArqRetornodt_credito.value))+','+
                                            quotedstr(FormatDateTime('YYYY-MM-DD',CDSArqRetornodtvenci.Value))+','+quotedstr('BB')+','+quotedstr(alltrim(copy(Label3.Caption,1,19)))+','+quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrarq.Text,'.')))+','+
                                            quotedstr(TrocaVirgPPto(TrimChar(CDSArqRetornovrdesconto.Text,'.')))+','+quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(ctbl)+','+quotedstr('Baixa Automática')+')');


          end;
    //      mensagem(DM_tabelas.ZQReceb_Baixa.SQL.Text);
          DM_tabelas.ZQReceb_Baixa.ExecSQL;
          DM_tabelas.ZQReceb_Baixa.Close;



          DM_tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.ZQConfiguracoes.FieldByName('bxautodoc').AsString,[]);
          //20/08/2010
          ZQReceb.Close;
          ZQReceb.SQL.Clear;
          ZQReceb.SQL.Add('select numboleto,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,empresa,custodaparcela,origem,recpag,numordem,contabil,adversa,');
          ZQReceb.SQL.Add('idloteamento,venda_idvenda,quadralote,Substituicao,sq,nomeadversa,tip,observ');
          ZQReceb.SQL.Add(' from recebimento where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text))+' order by idrecebimento');
          ZQReceb.Open;

          if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
          begin
            ZQReceb.Edit;
            if empty(ZQReceb.FieldByName('numboleto').AsString) then
               ZQReceb.FieldByName('numboleto').AsString:=CDSArqRetornonossonumero.Value;
            ZQReceb.FieldByName('Observ').AsString := 'Baixa Automática';

            // 21/05/2014
            if (CDSArqRetornovrdesconto.Value=0) and (abs(CDSArqRetornovrarq.Value)<ZQReceb.FieldByName('saldo').AsFloat) then
            begin
              ZQReceb.FieldByName('saldo').AsFloat := abs(ExRound(ZQReceb.FieldByName('saldo').AsFloat,2)-ExRound(CDSArqRetornovrarq.Value,2));
              ZQReceb.FieldByName('Valor').AsFloat := abs(ExRound(ZQReceb.FieldByName('saldo').AsFloat,2)-ExRound(CDSArqRetornovrarq.Value,2));
              ZQReceb.FieldByName('tip').AsString:='P';
            end
            else
            begin
              ZQReceb.FieldByName('saldo').AsFloat := 0;
              ZQReceb.FieldByName('tip').AsString:='';
            end;  

            // vinculo para exclusao de baixa coloquei 21/05/2014
            ZQReceb.FieldByName('RefBaixa').AsInteger := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;

            ZQReceb.Post;
          end;


          DM_tabelas.ZQCaixa.Close;
          DM_tabelas.ZQCaixa.SQL.Clear;
          DM_tabelas.ZQCaixa.SQL.Add('delete from caixa where vinculo = ''R'''+quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text));
          DM_tabelas.ZQCaixa.ExecSQL;

          DM_tabelas.ZQCaixa.Close;
          DM_tabelas.ZQCaixa.SQL.Clear;
          DM_tabelas.ZQCaixa.SQL.Add('select codigo,plano_contas_codigo,data_lan,vr_lan,obs,documento,empresa,vinculo,saldo,credeb,contraquem,contracodigo,Substituicao,sq from caixa');
          DM_tabelas.ZQCaixa.Open;


          DM_tabelas.ZQCaixa.Insert;
          DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').AsInteger := DM_Tabelas.ZQConfiguracoes.FieldByName('bxautoplc').AsInteger;
          DM_tabelas.ZQCaixa.FieldByName('data_lan').AsDateTime := date;
          DM_tabelas.ZQCaixa.FieldByName('vr_lan').AsFloat := CDSArqRetornovrarq.Value;
          DM_tabelas.ZQCaixa.FieldByName('documento').AsString := Label3.Caption;
          DM_tabelas.ZQCaixa.FieldByName('empresa').AsInteger := 1;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('bxautoplc').AsInteger,[]);
          DM_tabelas.ZQCaixa.FieldByName('credeb').AsString := 'C';
          DM_tabelas.ZQCaixa.FieldByName('vinculo').AsString := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_tabelas.ZQCaixa.FieldByName('contracodigo').AsInteger :=ZQReceb.FieldByName('cliente').AsInteger;
          DM_tabelas.ZQCaixa.FieldByName('obs').AsString := 'Baixa Automática';
          IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
              DM_tabelas.ZQCaixa.FieldByName('obs').AsString :=DM_tabelas.ZQCaixa.FieldByName('obs').AsString + ' Vencimento em: '+datetostr(CDSArqRetornodtvenci.Value);
          end;
          DM_tabelas.ZQCaixa.FieldByName('obs').AsString :=DM_tabelas.ZQCaixa.FieldByName('obs').AsString;
          DM_tabelas.ZQCaixa.Post;
        end;
      end;
    end
    else
    Begin
      ZQReceb.Close;
      ZQReceb.SQL.Clear;
      ZQReceb.SQL.Add('update recebimento set observ='+quotedstr(datetostr(date))+quotedstr(' ')+quotedstr(CDSArqRetornodesc_ocorr.Value)+quotedstr(' ')+quotedstr(copy(CDSArqRetornodesc_motivo.Value,1,100))+' where idrecebimento='+quotedstr(trim(CDSArqRetornoidrece.Text)));
      ZQReceb.ExecSQL;

//      ZQReceb.Edit;
//      ZQRecebobserv.Value := datetostr(date)+chr(13)+CDSArqRetornodesc_ocorr.Value+chr(13)+CDSArqRetornodesc_motivo.Value;
//      ZQReceb.Post;
    end;
    // aqui grava retorno
//    if saldo>0 then
    begin

      application.ProcessMessages;
      DM_TAbelas.ZQBancRemes2.close;
      DM_Tabelas.ZQBancRemes2.SQL.Clear;
      if (bco<>'756') then
         DM_Tabelas.ZQBancRemes2.SQL.Add('Select idbanco_remessa,remessa,idbanco,gerado,nomearq,arquivo,digito_dif,remessa2 from banco_remessa where idbanco='+quotedstr(CDSArqRetornoidconta_bancaria.Text)+' and digito_dif='+quotedstr(trim(CDSArqRetornoDigito.Value))+' and gerado='+quotedstr(vardata)+' and nomearq='+quotedstr(varnome))
      else
         DM_Tabelas.ZQBancRemes2.SQL.Add('Select idbanco_remessa,remessa,idbanco,gerado,nomearq,arquivo,digito_dif,remessa2 from banco_remessa where idbanco='+quotedstr(CDSArqRetornoidconta_bancaria.Text)+' and digito_dif='+quotedstr(trim(CDSArqRetornoDigito.Value))+' and gerado='+quotedstr(vardata)+' and nomearq='+quotedstr(copy(varnome,15,20)));
      DM_TAbelas.ZQBancRemes2.open;
      DM_TAbelas.ZQBancRemes2.RecordCount;
      application.ProcessMessages;

      DM_Tabelas.ZQRemesRec.Close;
      DM_Tabelas.ZQRemesRec.SQL.Clear;
//      DM_Tabelas.ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,retorno_motivo,acao,nossonumero,ocorrencia,desc_motivo,credito,dt_ocorrencia from remessa_receb');
//      DM_Tabelas.ZQRemesRec.Open;
      if CDSArqRetornodt_credito.value>0 then
      begin
        DM_Tabelas.ZQRemesRec.SQL.Add('INSERT INTO  remessa_receb (remessa,idrec,nossonumero,ocorrencia,desc_motivo,credito,dt_ocorrencia) VALUES ('+
                                          quotedstr(DM_Tabelas.ZQBancRemes2.FieldByName('idbanco_remessa').AsString)+','+quotedstr(CDSArqRetornoidrece.Value.ToString)+','+quotedstr(CDSArqRetornonossonumero.Value)+','+quotedstr(CDSArqRetornodesc_ocorr.Value)+','+
                                          quotedstr(copy(CDSArqRetornodesc_motivo.Value,1,100))+','+quotedstr(FormatDateTime('yyyy/mm/dd',CDSArqRetornodt_credito.value))+','+quotedstr(FormatDateTime('yyyy/mm/dd',EData.DateValue))+')');
      end
      else
      begin
        DM_Tabelas.ZQRemesRec.SQL.Add('INSERT INTO  remessa_receb (remessa,idrec,nossonumero,ocorrencia,desc_motivo,dt_ocorrencia) VALUES ('+
                                          quotedstr(DM_Tabelas.ZQBancRemes2.FieldByName('idbanco_remessa').AsString)+','+quotedstr(CDSArqRetornoidrece.Value.ToString)+','+quotedstr(CDSArqRetornonossonumero.Value)+','+quotedstr(CDSArqRetornodesc_ocorr.Value)+','+
                                          quotedstr(copy(CDSArqRetornodesc_motivo.Value,1,100))+','+quotedstr(FormatDateTime('yyyy/mm/dd',EData.DateValue))+')');
      end;
      DM_Tabelas.ZQRemesRec.ExecSQL;

//      DM_tabelas.ZQRemesRec.Append;
//      DM_Tabelas.ZQRemesRecremessa.Value := DM_Tabelas.ZQBancRemes2idbanco_remessa.Value;
//      DM_Tabelas.ZQRemesRecidrec.Value := CDSArqRetornoidrece.Value;
//      DM_Tabelas.ZQRemesRecnossonumero.Value := CDSArqRetornonossonumero.Value;
//      DM_Tabelas.ZQRemesRecocorrencia.Value  := CDSArqRetornodesc_ocorr.Value;
//      DM_Tabelas.ZQRemesRecdesc_motivo.Value := CDSArqRetornodesc_motivo.Value;
//      if CDSArqRetornodt_credito.value>0 then
//         DM_Tabelas.ZQRemesReccredito.Value := CDSArqRetornodt_credito.Value;
//      DM_Tabelas.ZQRemesRecdt_ocorrencia.value := EData.DateValue;
//      DM_tabelas.ZQRemesRec.Post;
    end;
//    ZQReceb.refresh;
    CDSArqRetorno.Next;
  end;
  CDSArqRetorno.First;
  CDSArqRetorno.EnableControls;
  ProgressBar1.Position:=0;
  ProgressBar1.Visible:=false;
  application.ProcessMessages;

{  DM_Tabelas.ZQReBxHi.Close;
  DM_Tabelas.ZQReBxHi.SQL.Clear;
  DM_Tabelas.ZQReBxHi.SQL.Add('select * from RecBxHist');
  DM_Tabelas.ZQReBxHi.open;

  DM_tabelas.ZQReceb_Baixa.Close;
  DM_tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_tabelas.ZQReceb_Baixa.SQL.Add('Select * from Receb_Baixa');
  DM_tabelas.ZQReceb_Baixa.Open;
 }


  DM_Tabelas.ZQBancRemes.Filtered:=false;
  DM_Tabelas.ZQRemesRec.Close;
  DM_Tabelas.ZQRemesRec.SQL.Clear;
  DM_Tabelas.ZQRemesRec.SQL.Add('select idremessa_receb,remessa,idrec,retorno_motivo,acao,nossonumero,ocorrencia,desc_motivo,credito,dt_ocorrencia from remessa_receb');
  DM_Tabelas.ZQRemesRec.Close;

  DM_TAbelas.ZQBancRemes2.close;
  DM_Tabelas.ZQBancRemes2.SQL.Clear;
  DM_Tabelas.ZQBancRemes2.SQL.Add('Select idbanco_remessa,remessa,idbanco,gerado,nomearq,arquivo,digito_dif,remessa2 from banco_remessa');
//  DM_TAbelas.ZQBancRemes2.open;
  DM_TAbelas.ZQBancRemes2.close;
  ZQReceb.close;


  Panel1.Visible:=true;
  application.ProcessMessages;
  DM_Tabelas.ZQRecebimento.Refresh;
  DM_TAbelas.ZQBancRemes.close;
  DM_Tabelas.ZQRemesRec.Close;
  DM_TAbelas.ZQBancRemes2.close;
  CDSArqRetorno.close;
  DBGBaixaAuto.Refresh;
  DBGBaixaAuto.Repaint;
  Panel1.Visible:=false;
  application.ProcessMessages;
end;


procedure TFrm_BaixaAutomatica.FormClose(Sender: TObject;
  var Action: TCloseAction);
  var
  varx :integer;
  varnome,vardata:string;
begin
  if processou<>'S' then
  begin
    for varx := length(OpenDialog1.FileName) to 1 do
      if copy(OpenDialog1.FileName,varx,1)='/' Then
         break;
    varnome := copy(OpenDialog1.FileName,varx-11,12);
    vardata := copy(edata.DateText,7,4)+'-'+copy(edata.DateText,4,2)+'-'+copy(Edata.DateText,1,2);

    DM_TAbelas.ZQBancRemes.Open;
    DM_TAbelas.ZQBancRemes.filtered:=false;
    DM_TAbelas.ZQBancRemes.Filter:='nomearq='+quotedstr(varnome);
    DM_TAbelas.ZQBancRemes.filtered:=true;
    if DM_TAbelas.ZQBancRemes.RecordCount>0 then
    begin
      DM_TAbelas.ZQBancRemes.first;
      DM_TAbelas.ZQBancRemes.DisableControls;
      while not DM_TAbelas.ZQBancRemes.eof do
      begin
        DM_Tabelas.ZQBancRemes.delete;
      end;
      DM_TAbelas.ZQBancRemes.EnableControls;
    end;
    DM_TAbelas.ZQBancRemes.filtered:=false;
    DM_TAbelas.ZQBancRemes.Refresh;
  end;
//  DM_Tabelas.ZQContaBancaria.close;
  DM_Tabelas.ZQMensMot.close;
  DM_TAbelas.ZQBancRemes2.close;
  DM_Tabelas.ZQBancRemes.close;  
//  DM_Tabelas.ZQRecebimento.close;
//  DM_tabelas.ZQTipodoc.Close;
//  DM_Tabelas.ZQCobaRe.close;
  DM_Tabelas.ZQReBxHi.close;
  DM_tabelas.ZQReceb_Baixa.close;
  DM_tabelas.ZQCaixa.close;
//  DM_Tabelas.ZQPlanodeContas.close;
  DM_Tabelas.ZQMens_Motivo.close;
  DM_Tabelas.ZQMens_Retorno.close;
  ZQReceb.close;
  Frm_BaixaAutomatica:=nil;
  Action:=CaFree;
end;


procedure TFrm_BaixaAutomatica.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_BaixaAutomatica, 'ZQReceb', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebnomecliente', 'nomecliente', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQReceborigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebtip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebobserv', 'observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_BaixaAutomatica, 'ZQComisTemp', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempidcomissao', 'idcomissao', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempcontrat_idcontrato', 'contrat_idcontrato', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempvrvenda', 'vrvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTemppercent', 'percent', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempvendedor', 'vendedor', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempvrcomis', 'vrcomis', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTemppago', 'pago', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempliberado', 'liberado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempordem', 'ordem', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempvenci', 'venci', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempcodcliente', 'codcliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
