unit BaixaAutomatica;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,   dxButton, StdCtrls, XLabel3D, XBanner, Grids, DBGrids,
  DBCtrls, ExtCtrls, IniFiles, ComCtrls, DB, DBClient, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Mask, XDate, dxCore2 ;

//  , , , , , , ,
//  , , WinSkinData, , , , jpeg, ,
//  , Graphics;            Menus

type
  TFrm_BaixaAutomatica = class(TForm)
    XLTitulo: TXLabel3D;
    OpenDialog1: TOpenDialog;
    DXBFechar: TdxButton;
    DBAbrir: TdxButton;
    DBGBaixaAuto: TDBGrid;
    DXBProcessar: TdxButton;
    DXBImprimir: TdxButton;
    Label1: TLabel;
    dxButton1: TdxButton;
    DS_Receb: TDataSource;
    ZQReceb: TZQuery;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    EData: TXDateEdit;
    Ebxautoplc: TDBEdit;
    Label5: TLabel;
    Label6: TLabel;
    DBEdit2: TDBEdit;

    DS_ComisTemp: TDataSource;
    ZQComisTemp: TZQuery;













    XBanner7: TXBanner;




























    datacredito: TXDateEdit;



    ProgressBar1: TProgressBar;
    ProgressBar2: TProgressBar;
    Panel1: TPanel;
    CDSArqRetorno: TClientDataSet;
    CDSArqRetornoidrece: TIntegerField;
    CDSArqRetornoident_ocorr: TStringField;
    CDSArqRetornodesc_ocorr: TStringField;
    CDSArqRetornoident_motivo: TStringField;
    CDSArqRetornovrtit: TFloatField;
    CDSArqRetornovrarq: TFloatField;
    CDSArqRetornoacao: TStringField;
    CDSArqRetornovrjuros: TFloatField;
    CDSArqRetornovrdesconto: TFloatField;
    CDSArqRetornodtvenci: TDateField;
    CDSArqRetornoordem: TStringField;
    CDSArqRetornodesc_motivo: TStringField;
    CDSArqRetornodt_credito: TDateField;
    CDSArqRetornonossonumero: TStringField;
    CDSArqRetornoDigito: TStringField;
    CDSArqRetornoidconta_bancaria: TLargeintField;
    CDSArqRetornoNomecliente: TStringField;
    CDSArqRetornoquadralote: TStringField;
    DS_ArqRetorno: TDataSource;
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

uses Tabelas, funcoes, TabelasDeMensagens, RelRetorno, uRuntimeFields;

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
  nn,hj,varnome, vardata : string;
  ArqLido,ArqSize :Integer;
  
begin
  ZQReceb.close;
  ZQReceb.open;
  processou:='N';
  VarDif:='';
  digit:='';
  hj:='';
  dg:='';
  if OpenDialog1.Execute Then Begin
    AssignFile ( arq, OpenDialog1.FileName );
    Reset ( arq );
    ReadLn ( arq, linha );
    EData.DateValue:= strtodate(copy(linha,95,2)+'/'+copy(linha,97,2)+'/20'+copy(linha,99,2));
    if copy(linha,3,7) <> 'RETORNO' Then
      exit;
    for varx := length(OpenDialog1.FileName) to 1 do
        if copy(OpenDialog1.FileName,varx,1)='/' Then
           break;
    varnome := copy(OpenDialog1.FileName,varx-11,12);
    vardata := copy(edata.DateText,7,4)+'-'+copy(edata.DateText,4,2)+'-'+copy(Edata.DateText,1,2);

    if DM_Tabelas.ZQRemes_Receb.Active=false then
       DM_Tabelas.ZQRemes_Receb.open;

    DM_TAbelas.ZQBancRemes2.close;
    DM_TAbelas.ZQBancRemes2.SQL.Clear;
    DM_TAbelas.ZQBancRemes2.SQL.Add('select * from banco_remessa where nomearq='+quotedstr(varnome)+' and gerado='+quotedstr(vardata));
    DM_TAbelas.ZQBancRemes2.Open;


    {DM_TAbelas.ZQBancRemes2.Filtered:=false;
    DM_TAbelas.ZQBancRemes2.Filter:='nomearq='+quotedstr(varnome)+' and gerado='+quotedstr(vardata);
    DM_TAbelas.ZQBancRemes2.Filtered:=true;}
    
    if DM_TAbelas.ZQBancRemes2.RecordCount>0 then
    begin
  //  if DM_TAbelas.ZQBancRemes.locate('nomearq;gerado',VarArrayOf([varnome,vardata]),[]) Then Begin
      //showmessage('Este arquivo ja foi baixado no dia '+DM_Tabelas.ZQBancRemes2gerado.Text+', ele não pode mais ser baixado!!!!');
      showmessage('Existe baixa com o mesmo nome do arquivo no dia '+DM_Tabelas.ZQBancRemes2.FieldByName('gerado').Text+' prosseguindo...');
      DM_TAbelas.ZQRemesRec.Open;
      DM_TAbelas.ZQBancRemes2.First;
      ProgressBar1.position:=0;
      ProgressBar1.max:= DM_TAbelas.ZQBancRemes2.recordcount;
      ProgressBar1.visible:=true;

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

      ProgressBar1.position:=0;
      ProgressBar1.visible:=false;


      DM_TAbelas.ZQRemesRec.Close;
      DM_TAbelas.ZQRemesRec.SQL.clear;
      DM_TAbelas.ZQRemesRec.SQL.Add('Select * from remessa_receb ');
      DM_TAbelas.ZQRemesRec.Open;

      DM_TAbelas.ZQRemesRec.Filtered:=false;
      DM_TAbelas.ZQRemesRec.Refresh;

  //    exit;
    end;
    DM_TAbelas.ZQBancRemes2.Filtered:=false;
    DM_TAbelas.ZQBancRemes2.close;

    CDSArqRetorno.Close;
    CDSArqRetorno.CreateDataSet;
    CDSArqRetorno.open;
    if copy(linha,77,3)='001' then   //bco brasil
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.filter:='n_conta='+quotedstr(inttostr(strtoint(copy(linha,32,8))));
      DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').value;
    end;
    if copy(linha,77,3)='237' then  // bco bradesco
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.filter:='n_no_banco='+quotedstr(copy(linha,40,7));
      DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').value;

    end;

    if copy(linha,77,3)='033' then     // bco santander
    begin
      DM_Tabelas.ZQContaBancaria.filtered:=false;
      DM_Tabelas.ZQContaBancaria.close;
      DM_Tabelas.ZQContaBancaria.SQL.Clear;
      DM_Tabelas.ZQContaBancaria.SQL.Add('select * from conta_bancaria where n_conta like '+quotedstr('%'+inttostr(strtoint(copy(linha,39,8)))+'%'));
      DM_Tabelas.ZQContaBancaria.Open;

//      DM_Tabelas.ZQContaBancaria.filter:='n_conta='+quotedstr(inttostr(strtoint(copy(linha,39,8))));
  //    DM_Tabelas.ZQContaBancaria.filtered:=true;
      DM_Tabelas.ZQContaBancaria.first;
      DM_Tabelas.ZQContaBancaria.recordcount;
      digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').value;

    end;

    ProgressBar1.position:=0;
    ProgressBar1.max:= DM_Tabelas.ZQContaBancaria.recordcount;
    ProgressBar1.visible:=true;
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
              VarDif  := '';
              posi:=pos('|',vartext);
              if not empty(vartext) then
              begin
                if posi=0 then
                begin
                  CDSArqRetorno.Insert;
                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,11)));
                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    nn:= copy(linha,64,17);
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').Value;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').Value;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').Value;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').Value;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').Value;
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
                    DM_TAbelas.ZQMensMot.SQL.add('select * from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                    DM_TAbelas.ZQMensMot.Open;
                    if DM_Tabelas.ZQMensMot.FieldByName('mensagem').Value='' then
                       CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                    else
                       CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').Value;
                  end;

                  CDSArqRetornoDigito.Value:=digit;
                  CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Value;
                  CDSArqRetorno.Post;
                  hj:='';
                end;
              end
              else
              begin
                showmessage('Código do Recebimento em Branco. Coluna 39.');
              end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;

        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select * from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').Value := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').Value := varnome;
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').Value := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;

        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').value;
        Reset ( arq );
        ReadLn ( arq, linha );
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
              VarDif  := alltrim(copy(linha,62,1));
              posi:=pos('|',vartext);
              if not empty(vartext) then
              begin
                if posi=0 then
                begin
                  CDSArqRetorno.Insert;
                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,11)));
                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').Value;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').Value;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').Value;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').Value;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').Value;
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
                    DM_TAbelas.ZQMensMot.SQL.add('select * from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                    DM_TAbelas.ZQMensMot.Open;
                    if DM_Tabelas.ZQMensMot.FieldByName('mensagem').Value='' then
                       CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                    else
                       CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').Value;
                  end;
                  CDSArqRetornoDigito.Value:=digit;
                  CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Value;
                  CDSArqRetorno.Post;
                  hj:='';
                end;
              end
              else
              begin
                showmessage('Código do Recebimento em Branco. Coluna 38.');
              end;
            end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;

        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select * from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').Value := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').Value := varnome;
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').Value := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;

        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').value;
        Reset ( arq );
        ReadLn ( arq, linha );
      end;
    end;
    if copy(linha,77,3)='033' then
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
              VarDif  := '';
              posi:=pos('|',vartext);
              if not empty(vartext) then
              begin
                if posi=0 then
                begin
                  CDSArqRetorno.Insert;
                  CDSArqRetornoidrece.Value := strtoint(alltrim(copy(vartext,1,11)));
                  if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
                  Begin
                    nn:= copy(linha,63,8);
                    CDSArqRetornoordem.Value := ZQReceb.FieldByName('ordem').Value;
                    CDSArqRetornovrtit.Value := ZQReceb.FieldByName('Valor').Value;
                    CDSArqRetornodtvenci.Value := ZQReceb.FieldByName('Dt_Vencimento').Value;
                    CDSArqRetornoNomecliente.Value := ZQReceb.FieldByName('nomeadversa').Value;
                    CDSArqRetornoquadralote.Value:=ZQReceb.FieldByName('quadralote').Value;
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
                    DM_TAbelas.ZQMensMot.SQL.add('select * from Mens_Motivo where codmotivo='+quotedstr(CDSArqRetornoident_ocorr.Value)+' and retorno='+quotedstr(CDSArqRetornoident_motivo.Value));
                    DM_TAbelas.ZQMensMot.Open;
                    if DM_Tabelas.ZQMensMot.FieldByName('mensagem').Value='' then
                       CDSArqRetornodesc_motivo.Value := 'Motivo Inexistente'
                    else
                       CDSArqRetornodesc_motivo.Value := DM_Tabelas.ZQMensMot.FieldByName('mensagem').Value;
                  end;

                  CDSArqRetornoDigito.Value:=digit;
                  CDSArqRetornoidconta_bancaria.Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Value;
                  CDSArqRetorno.Post;
                  hj:='';
                end;
              end
              else
              begin
                showmessage('Código do Recebimento em Branco. Coluna 39.');
              end;
          end;
        end;
        ProgressBar2.Position:=0;
        ProgressBar2.Visible:=false;

        application.ProcessMessages;
        DM_TAbelas.ZQMensMot.close;
        DM_TAbelas.ZQMensMot.SQL.Clear;
        DM_TAbelas.ZQMensMot.SQL.add('select * from Mens_Motivo');
        DM_TAbelas.ZQMensMot.Open;
        DXBImprimir.Enabled := True;
        DXBProcessar.Enabled := True;

        application.ProcessMessages;
        // gravar o nome do arquivo e a data
        DM_TAbelas.ZQBancRemes.Open;
        DM_TAbelas.ZQBancRemes.Insert;
        DM_Tabelas.ZQBancRemes.FieldByName('gerado').Value := Edata.DateValue;
        DM_Tabelas.ZQBancRemes.FieldByName('nomearq').Value := varnome;
        DM_Tabelas.ZQBancRemes.FieldByName('digito_dif').Value := digit;
        DM_Tabelas.ZQBancRemes.Post;

        DM_Tabelas.ZQContaBancaria.next;

        CloseFile ( arq );

        AssignFile ( arq, OpenDialog1.FileName );
        digit:=DM_Tabelas.ZQContaBancaria.FieldByName('n_dif_empreed').value;
        Reset ( arq );
        ReadLn ( arq, linha );
      end;
    end;
    Label4.Caption := inttostr(CDSArqRetorno.RecordCount);
    CloseFile ( arq );

    DBGBaixaAuto.Refresh;
    DBGBaixaAuto.Repaint;
    
    application.ProcessMessages;
    ProgressBar1.position:=0;
    ProgressBar1.visible:=false;


    DM_Tabelas.ZQContaBancaria.Filtered:=false;
    DM_Tabelas.ZQContaBancaria.close;
    DM_Tabelas.ZQContaBancaria.SQL.Clear;
    DM_Tabelas.ZQContaBancaria.SQL.Add('select * from conta_bancaria');
    DM_Tabelas.ZQContaBancaria.Open;
    DM_Tabelas.ZQBancRemes.Refresh;
    DM_Tabelas.ZQRemes_Receb.Refresh;
    DXBProcessar.Enabled:=true;
  end;
end;

procedure TFrm_BaixaAutomatica.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZQMensMot.Open;
  ZQReceb.open;
  DM_Tabelas.ZQContaBancaria.open;
  DM_Tabelas.ZQMensMot.open;
  DM_TAbelas.ZQBancRemes2.open;
//  DM_Tabelas.ZQRecebimento.Open;
  DM_tabelas.ZQTipodoc.Open;
  DM_Tabelas.ZQCobaRe.Open;
  DM_Tabelas.ZQReBxHi.Open;
  DM_tabelas.ZQReceb_Baixa.Open;
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
  if DM_Tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('bxautoplc').Value,[]) Then
    Ebxautoplc.Text := DM_Tabelas.ZQPlanoDeContas.FieldByName('codigo').Text+' - '+DM_Tabelas.ZQPlanoDeContas.FieldByName('descricao').Value;
//  CDSArqRetorno.CreateDataSet;
end;

procedure TFrm_BaixaAutomatica.dxButton1Click(Sender: TObject);
begin
  if not Verif_senha('Recebimento','Mensagens de Retorno ','') then exit;
  Frm_TabelasDeMensagens.Left := Frm_BaixaAutomatica.Left;
  Frm_TabelasDeMensagens.Top := Frm_BaixaAutomatica.Top+60;
  Frm_TabelasDeMensagens.showmodal;
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
  Frm_RelRetorno.RLLabel1.Caption := 'Retorno do Banco '+Label1.Caption;
  Frm_RelRetorno.RLLabel2.Caption := OpenDialog1.FileName+'  referente '+EData.DateText;
  if Frm_RelRetorno=nil then
    Frm_RelRetorno:= TFrm_RelRetorno.Create(Application);
  Frm_RelRetorno.RLReport1.Preview;

end;

procedure TFrm_BaixaAutomatica.DXBProcessarClick(Sender: TObject);
Var
  xref,Varx : integer;
  varnome, vardata : string;
  saldo:double;
begin
  DXBProcessar.Enabled:=false;
  processou:='S';
  if not Verif_senha('Recebimento','Processa baixa automática','') then exit;
  DM_TAbelas.ZQBancRemes2.Open;
  for varx := length(OpenDialog1.FileName) to 1 do
    if copy(OpenDialog1.FileName,varx,1)='/' Then
      break;
  varnome := copy(OpenDialog1.FileName,varx-11,12);
  vardata := copy(edata.DateText,7,4)+'-'+copy(edata.DateText,4,2)+'-'+copy(Edata.DateText,1,2);

{  DM_TAbelas.ZQBancRemes2.Filtered:=false;
  DM_TAbelas.ZQBancRemes2.Filter:='nomearq='+quotedstr(varnome)+' and gerado='+quotedstr(vardata);
  DM_TAbelas.ZQBancRemes2.Filtered:=true;
  if DM_TAbelas.ZQBancRemes2.RecordCount>0 then
  begin
//  if DM_TAbelas.ZQBancRemes.locate('nomearq;gerado',VarArrayOf([varnome,vardata]),[]) Then Begin
    //showmessage('Este arquivo ja foi baixado no dia '+DM_Tabelas.ZQBancRemes2gerado.Text+', ele não pode mais ser baixado!!!!');
    showmessage('Existe baixa com o mesmo nome do arquivo no dia '+DM_Tabelas.ZQBancRemes2gerado.Text+' prosseguindo...');
//    exit;
  end
  else
  begin
    DM_TAbelas.ZQBancRemes.Open;
    DM_TAbelas.ZQBancRemes.Insert;
    DM_Tabelas.ZQBancRemesgerado.Value := Edata.DateValue;
    DM_Tabelas.ZQBancRemesnomearq.Value := varnome;
    DM_Tabelas.ZQBancRemes.Post;
  end;
  DM_TAbelas.ZQBancRemes2.Filtered:=false;}
  DM_TAbelas.ZQBancRemes2.close;

  DM_Tabelas.ZQRemesRec.SQL.Clear;
  DM_Tabelas.ZQRemesRec.SQL.Add('select * from remessa_receb');
  DM_Tabelas.ZQRemesRec.Open;

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

    if CDSArqRetornoacao.Value = 'R' Then Begin
      DM_Tabelas.ZQRemesRec.SQL.Clear;
      DM_Tabelas.ZQRemesRec.SQL.Add('Select * from remessa_receb where idrec='+CDSArqRetornoidrece.Text+' and acao is null');
      DM_Tabelas.ZQRemesRec.Open;
      if DM_Tabelas.ZQRemesRec.RecordCount>0 Then Begin
        DM_Tabelas.ZQRemesRec.Edit;
        DM_TAbelas.ZQRemesRec.FieldByName('acao').Value := 'R';
        DM_TAbelas.ZQRemesRec.FieldByName('retorno_motivo').Value := CDSArqRetornoident_ocorr.Value+CDSArqRetornoident_motivo.Value;
        DM_Tabelas.ZQRemesRec.Post;
      end;
      ZQReceb.Edit;
      ZQReceb.FieldByName('Observ').Value := datetostr(date)+chr(13)+CDSArqRetornodesc_ocorr.Value+chr(13)+CDSArqRetornodesc_motivo.Value;
      ZQReceb.Post;
    end
    else if (CDSArqRetornoacao.Value = 'B') and (CDSArqRetornovrarq.Value>0) Then Begin
      if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
      begin
        DM_Tabelas.ZQReBxHi.Close;
        DM_Tabelas.ZQReBxHi.SQL.Clear;
        DM_Tabelas.ZQReBxHi.SQL.Add('select * from RecBxHist where idrecib='+quotedstr(CDSArqRetornoidrece.text));
        DM_Tabelas.ZQReBxHi.open;
        xref:=DM_Tabelas.ZQReBxHi.FieldByName('refer').Value;

      //  if ZQRecebsaldo.Value>0 then
        begin
          saldo:=ZQReceb.FieldByName('saldo').Value;
          DM_tabelas.ZQCobaRe.Insert;
          DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
          DM_Tabelas.ZQCobaRe.FieldByName('entrada').Value := EData.DateValue;
          DM_Tabelas.ZQCobaRe.Post;

          DM_Tabelas.ZQReBxHi.Close;
          DM_Tabelas.ZQReBxHi.SQL.Clear;
          DM_Tabelas.ZQReBxHi.SQL.Add('delete from RecBxHist where idrecib='+quotedstr(CDSArqRetornoidrece.text));
          DM_Tabelas.ZQReBxHi.ExecSQL;

          DM_Tabelas.ZQReBxHi.Close;
          DM_Tabelas.ZQReBxHi.SQL.Clear;
          DM_Tabelas.ZQReBxHi.SQL.Add('select * from RecBxHist ');
          DM_Tabelas.ZQReBxHi.open;

          DM_Tabelas.ZQReBxHi.Insert;
          DM_Tabelas.ZQReBxHi.FieldByName('juros').Value := CDSArqRetornovrjuros.Value;
          DM_Tabelas.ZQReBxHi.FieldByName('valor').Value := CDSArqRetornovrarq.Value;
          DM_Tabelas.ZQReBxHi.FieldByName('data').Value := CDSArqRetornodt_credito.value; //EData.DateValue;
          DM_Tabelas.ZQReBxHi.FieldByName('percent_usado').Value := 100;
          DM_tabelas.ZQReBxHi.FieldByName('refer').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_tabelas.ZQReBxHi.FieldByName('idrecib').Value := CDSArqRetornoidrece.Value;
          DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').Value := CDSArqRetornovrtit.Value;
          DM_Tabelas.ZQReBxHi.Post;



          DM_tabelas.ZQReceb_Baixa.Close;
          DM_tabelas.ZQReceb_Baixa.SQL.Clear;
          DM_tabelas.ZQReceb_Baixa.SQL.Add('delete from Receb_Baixa where refbaixa='+quotedstr(inttostr(xref)));
  //        showmessage(DM_tabelas.ZQReceb_Baixa.SQL.Text);
          DM_tabelas.ZQReceb_Baixa.ExecSQL;
          xref:=0;

          DM_tabelas.ZQReceb_Baixa.Close;
          DM_tabelas.ZQReceb_Baixa.SQL.Clear;
          DM_tabelas.ZQReceb_Baixa.SQL.Add('select * from  Receb_Baixa ');
    //      showmessage(DM_tabelas.ZQReceb_Baixa.SQL.Text);
          DM_tabelas.ZQReceb_Baixa.Open;

          DM_tabelas.ZQReceb_Baixa.Insert;
          DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').Value := CDSArqRetornovrarq.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').Value := EData.DateValue;
          DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := CDSArqRetornodt_credito.value; //EData.DateValue;
          DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSArqRetornodtvenci.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := 'BB';
//          DM_tabelas.ZQReceb_BaixaTipdoc.Value := DM_Tabelas.ZQconfiguracoesbxautodoc.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := Label3.Caption;
          DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSArqRetornovrjuros.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Desc_Vr').Value  := CDSArqRetornovrdesconto.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := DM_Tabelas.ZQConfiguracoes.FieldByName('bxautoplc').Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := 'Baixa Automática';
          DM_tabelas.ZQReceb_Baixa.FieldByName('valor_parcela').Value := CDSArqRetornovrtit.Value;
          DM_tabelas.ZQReceb_Baixa.Post;
          DM_tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.ZQConfiguracoes.FieldByName('bxautodoc').Value,[]);
          //20/08/2010
          if ZQReceb.Locate('idrecebimento',CDSArqRetornoidrece.Value,[]) Then
          begin
            ZQReceb.Edit;
            if empty(ZQReceb.FieldByName('numboleto').value) then
               ZQReceb.FieldByName('numboleto').value:=CDSArqRetornonossonumero.Value;
            ZQReceb.FieldByName('Observ').Value := 'Baixa Automática';
                // 21/05/2014
            if (CDSArqRetornovrdesconto.Value=0) and (abs(CDSArqRetornovrarq.Value)<ZQReceb.FieldByName('saldo').Value) then
            begin
               ZQReceb.FieldByName('saldo').Value := abs(ExRound(ZQReceb.FieldByName('saldo').Value,2)-ExRound(CDSArqRetornovrarq.Value,2));
                ZQReceb.FieldByName('Valor').Value := abs(ExRound(ZQReceb.FieldByName('saldo').Value,2)-ExRound(CDSArqRetornovrarq.Value,2));
                ZQReceb.FieldByName('tip').Value:='P';
            end
            else
            begin
              ZQReceb.FieldByName('saldo').Value := 0;
              ZQReceb.FieldByName('tip').Value:='';
            end;
            // vinculo para exclusao de baixa coloquei 21/05/2014
            ZQReceb.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;

            ZQReceb.Post;
          end;
          DM_tabelas.ZQCaixa.Insert;
          DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').Value := DM_Tabelas.ZQConfiguracoes.FieldByName('bxautoplc').Value;
          DM_tabelas.ZQCaixa.FieldByName('data_lan').Value := date;
          DM_tabelas.ZQCaixa.FieldByName('vr_lan').Value := CDSArqRetornovrarq.Value;
          DM_tabelas.ZQCaixa.FieldByName('documento').Value := Label3.Caption;
          DM_tabelas.ZQCaixa.FieldByName('empresa').Value := 1;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('bxautoplc').Value,[]);
          DM_tabelas.ZQCaixa.FieldByName('credeb').Value := 'C';
          DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_tabelas.ZQCaixa.FieldByName('contracodigo').Value :=ZQReceb.FieldByName('cliente').Value;
          DM_tabelas.ZQCaixa.FieldByName('obs').Value := 'Baixa Automática';
          IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
              DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value + ' Vencimento em: '+datetostr(CDSArqRetornodtvenci.Value);
          end;
          DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value;
          DM_tabelas.ZQCaixa.Post;
        end;
      end;
    end
    else  Begin
      ZQReceb.Edit;
      ZQReceb.FieldByName('Observ').Value := datetostr(date)+chr(13)+CDSArqRetornodesc_ocorr.Value+chr(13)+CDSArqRetornodesc_motivo.Value;
      ZQReceb.Post;
    end;
    // aqui grava retorno

    DM_TAbelas.ZQBancRemes2.close;
    DM_Tabelas.ZQBancRemes2.SQL.Clear;
    DM_Tabelas.ZQBancRemes2.SQL.Add('Select * from banco_remessa where idbanco='+quotedstr(CDSArqRetornoidconta_bancaria.Text)+' and digito_dif='+quotedstr(CDSArqRetornoDigito.Value)+' and gerado='+quotedstr(vardata)+' and nomearq='+quotedstr(varnome));
    DM_TAbelas.ZQBancRemes2.open;

    DM_Tabelas.ZQRemesRec.SQL.Clear;
    DM_Tabelas.ZQRemesRec.SQL.Add('Select * from remessa_receb');
    DM_Tabelas.ZQRemesRec.Open;
    DM_tabelas.ZQRemesRec.Insert;
    DM_Tabelas.ZQRemesRec.FieldByName('remessa').Value := DM_Tabelas.ZQBancRemes2.FieldByName('idbanco_remessa').Value;
    DM_Tabelas.ZQRemesRec.FieldByName('idrec').Value := CDSArqRetornoidrece.Value;
    DM_Tabelas.ZQRemesRec.FieldByName('nossonumero').Value := CDSArqRetornonossonumero.Value;
    DM_Tabelas.ZQRemesRec.FieldByName('ocorrencia').Value  := CDSArqRetornodesc_ocorr.Value;
    DM_Tabelas.ZQRemesRec.FieldByName('desc_motivo').Value := CDSArqRetornodesc_motivo.Value;
    if CDSArqRetornodt_credito.value>0 then
       DM_Tabelas.ZQRemesRec.FieldByName('credito').Value := CDSArqRetornodt_credito.value;
    DM_Tabelas.ZQRemesRec.FieldByName('dt_ocorrencia').value := EData.DateValue;
    DM_tabelas.ZQRemesRec.Post;
    ZQReceb.refresh;
    CDSArqRetorno.Next;
  end;
  DM_Tabelas.ZQBancRemes.Filtered:=false;
  DM_TAbelas.ZQBancRemes2.close;
  DM_Tabelas.ZQBancRemes2.SQL.Clear;
  DM_Tabelas.ZQBancRemes2.SQL.Add('Select * from banco_remessa');
  DM_TAbelas.ZQBancRemes2.open;
  DM_TAbelas.ZQBancRemes2.close;
  ZQReceb.close;
  ProgressBar1.position:=0;
  ProgressBar1.visible:=false;
  Panel1.Visible:=true;
  application.ProcessMessages;
  DM_Tabelas.ZQRecebimento.Refresh;

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
      while not DM_TAbelas.ZQBancRemes.eof do
      begin
        DM_Tabelas.ZQBancRemes.delete;
      end;
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
end;


procedure TFrm_BaixaAutomatica.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_BaixaAutomatica, 'ZQReceb', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebnomecliente', 'nomecliente', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQReceborigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebsomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebnomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQReceb', 'ZQRecebtip', 'tip', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempordem', 'ordem', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempvenci', 'venci', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_BaixaAutomatica, 'ZQComisTemp', 'ZQComisTempcodcliente', 'codcliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
