unit PesqEndereco;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XLabel3D, XBanner, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Grids, DBGrids, Mask, DBCtrls;

type
  TFrmPesqEndereco = class(TForm)
    XLTit: TXLabel3D;
    DBGrid1: TDBGrid;
    Label1: TLabel;
    edlogradouro: TEdit;
    XBanner6: TXBanner;
    Label11: TLabel;
    LReg: TLabel;
    Label4: TLabel;
    cbuf: TComboBox;
    Label5: TLabel;
    Label6: TLabel;
    cbtipo: TComboBox;
    cblocalidade: TComboBox;
    log_tipo_logr: TZQuery;
    Datalog_tipo_logr: TDataSource;

    log_faixa_uf: TZQuery;
    Datalog_faixa_uf: TDataSource;


    DataSource1: TDataSource;
    log_localidade: TZQuery;


    Datalocalizacao: TDataSource;
    localizacao: TZQuery;
    Label2: TLabel;

    DS_zqbairro: TDataSource;
    ZQbairro: TZQuery;











    procedure FormShow(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure edlogradouroKeyPress(Sender: TObject; var Key: Char);
    procedure DBGrid1KeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure cbtipoDropDown(Sender: TObject);
    procedure cbufDropDown(Sender: TObject);
    procedure cblocalidadeDropDown(Sender: TObject);
    procedure cbufKeyPress(Sender: TObject; var Key: Char);
    procedure cblocalidadeKeyPress(Sender: TObject; var Key: Char);
    procedure cbtipoKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmPesqEndereco: TFrmPesqEndereco;
  limite,totreg:integer;
implementation

uses Tabelas, Funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrmPesqEndereco.FormShow(Sender: TObject);
begin
  ZQbairro.open;
  limite:=1;
  cbuf.Clear;
  cblocalidade.Clear;
  cbtipo.Clear;
  edlogradouro.Clear;
  localizacao.SQL.Clear;
  localizacao.SQL.Clear;
  localizacao.SQL.Add('select lobairro.bai_no,loca.loc_no,loca.ufe_sg,loga.cep, loga.loc_nu_sequencial,loga.log_tipo_logradouro,loga.log_no,log_no_sem_acento,loga.bai_nu_sequencial_ini from log_logradouro as loga ');
  localizacao.SQL.Add('join log_localidade as loca on loga.loc_nu_sequencial=loca.loc_nu_sequencial ');
  localizacao.SQL.Add(' join log_bairro as lobairro on loga.bai_nu_sequencial_ini=lobairro.bai_nu_sequencial ');
  localizacao.SQL.Add(' LIMIT '+alltrim(inttostr(limite))+',15000 ');
  localizacao.Open;
  totreg:=totreg+localizacao.RecordCount;
  LReg.Caption := inttostr(totreg);
  if (localizacao.Active) and (localizacao.RecordCount>0) Then
    DBGrid1.SetFocus
  else
    close;
  cbuf.SetFocus;
end;

procedure TFrmPesqEndereco.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 then
    close;

end;

procedure TFrmPesqEndereco.edlogradouroKeyPress(Sender: TObject;
  var Key: Char);
  var
  ps:integer;
begin
  if Key = #13 then begin
     Label2.Visible:=true;
     ps:=pos('|',cblocalidade.Text);
     inc(ps);
     application.ProcessMessages;
     localizacao.SQL.Clear;

     localizacao.SQL.Add('select loga.ufe_sg,loga.cep,loga.loc_nu_sequencial,loga.log_tipo_logradouro,loga.log_no,log_no_sem_acento,loga.bai_nu_sequencial_ini from log_logradouro as loga');
     localizacao.SQL.Add(' WHERE loga.loc_nu_sequencial='+trim(copy(cblocalidade.Text,ps,20))+' and (loga.log_tipo_logradouro='+quotedstr(cbtipo.text)+') and (loga.ufe_sg='+quotedstr(cbuf.text)+')');
     localizacao.SQL.Add(' and ((loga.log_no like '+quotedstr('%'+edlogradouro.text+'%')+') or (loga.log_no_sem_acento like '+quotedstr('%'+edlogradouro.text+'%')+'))');
     localizacao.Open;
     totreg:=localizacao.RecordCount;
     LReg.Caption := inttostr(totreg);
     Label2.Visible:=false;
     application.ProcessMessages;
     if (localizacao.Active) and (localizacao.RecordCount>0) Then
        DBGrid1.SetFocus
      else
        close;
  end;
end;

procedure TFrmPesqEndereco.DBGrid1KeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) and (localizacao.RecordCount>0) Then
  Begin
     DM_tabelas.localizacao.close;
     DM_tabelas.localizacao.SQL.Clear;
     DM_tabelas.localizacao.SQL.Add('select loga.ufe_sg,loga.cep,loga.loc_nu_sequencial,loga.log_tipo_logradouro,loga.log_no,log_no_sem_acento,loga.bai_nu_sequencial_ini from log_logradouro as loga');
     DM_tabelas.localizacao.SQL.Add(' where (loga.cep='+quotedstr(localizacao.FieldByName('cep').AsString)+')');
     DM_tabelas.localizacao.Open;
     cbuf.Clear;
     cblocalidade.Clear;
     cbtipo.Clear;
     edlogradouro.Clear;
     Close;
  end;
end;

procedure TFrmPesqEndereco.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  localizacao.Close;
  ZQbairro.close;
  log_tipo_logr.close;
  log_faixa_uf.close;
  log_localidade.close;
end;

procedure TFrmPesqEndereco.cbtipoDropDown(Sender: TObject);
var
nom:string;
begin
  cbtipo.Clear;
  log_tipo_logr.Open;
  log_tipo_logr.First;
  nom:=log_tipo_logr.FieldByName('tipologradouro').AsString;
  cbtipo.Items.Add(log_tipo_logr.FieldByName('tipologradouro').AsString);
  while not log_tipo_logr.Eof do
  begin
    if nom<>log_tipo_logr.FieldByName('tipologradouro').AsString then
    begin
      nom:=log_tipo_logr.FieldByName('tipologradouro').AsString;
      cbtipo.Items.Add(log_tipo_logr.FieldByName('tipologradouro').AsString);
      log_tipo_logr.Next;
    end
    else
    begin
      log_tipo_logr.Next;
    end;
  end;
  log_tipo_logr.Close;
end;

procedure TFrmPesqEndereco.cbufDropDown(Sender: TObject);
var
nom:string;
begin
  cbuf.Clear;
  log_faixa_uf.Open;
  log_faixa_uf.First;
  nom:=log_faixa_uf.FieldByName('ufe_sg').AsString;
  cbuf.Items.Add(log_faixa_uf.FieldByName('ufe_sg').AsString);
  while not log_faixa_uf.Eof do
  begin
    if nom<>log_faixa_uf.FieldByName('ufe_sg').AsString then
    begin
      nom:=log_faixa_uf.FieldByName('ufe_sg').AsString;
      cbuf.Items.Add(log_faixa_uf.FieldByName('ufe_sg').AsString);
      log_faixa_uf.Next;
    end
    else
    begin
      log_faixa_uf.Next;
    end;
  end;
  log_faixa_uf.Close;
end;

procedure TFrmPesqEndereco.cblocalidadeDropDown(Sender: TObject);
var
nom:string;
begin
  cblocalidade.Clear;
  log_localidade.close;
  log_localidade.SQL.clear;
  log_localidade.SQL.Add('Select loc_no,ufe_sg,loc_nu_sequencial from log_localidade where ufe_sg='+quotedstr(cbuf.text)+' order by loc_no');
  log_localidade.open;
  log_localidade.First;
  nom:=log_localidade.FieldByName('loc_no').AsString;
  cblocalidade.Items.Add(log_localidade.FieldByName('loc_no').AsString+'|'+log_localidade.FieldByName('loc_nu_sequencial').AsString);
  while not log_localidade.Eof do
  begin
    if nom<>log_localidade.FieldByName('loc_no').AsString then
    begin
      nom:=log_localidade.FieldByName('loc_no').AsString;
      cblocalidade.Items.Add(log_localidade.FieldByName('loc_no').AsString+'|'+log_localidade.FieldByName('loc_nu_sequencial').AsString);
      log_localidade.Next;
    end
    else
    begin
      log_localidade.Next;
    end;
  end;
  log_localidade.Close;
end;

procedure TFrmPesqEndereco.cbufKeyPress(Sender: TObject; var Key: Char);
begin
  if key=#13 then
     Perform(Wm_NextDlgCtl,0,0);

end;

procedure TFrmPesqEndereco.cblocalidadeKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key=#13 then
     Perform(Wm_NextDlgCtl,0,0);

end;

procedure TFrmPesqEndereco.cbtipoKeyPress(Sender: TObject; var Key: Char);
begin
  if key=#13 then
     Perform(Wm_NextDlgCtl,0,0);

end;


procedure TFrmPesqEndereco.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmPesqEndereco, 'log_tipo_logr', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'log_tipo_logr', 'log_tipo_logrtipologradouro', 'tipologradouro', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmPesqEndereco, 'log_faixa_uf', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'log_faixa_uf', 'log_faixa_ufufe_sg', 'ufe_sg', TWideStringField, fkData, 2, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'log_faixa_uf', 'log_faixa_ufufe_no', 'ufe_no', TWideStringField, fkData, 72, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmPesqEndereco, 'log_localidade', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'log_localidade', 'log_localidadeloc_no', 'loc_no', TWideStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'log_localidade', 'log_localidadeufe_sg', 'ufe_sg', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'log_localidade', 'log_localidadeloc_nu_sequencial', 'loc_nu_sequencial', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmPesqEndereco, 'localizacao', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'localizacao', 'localizacaoufe_sg', 'ufe_sg', TWideStringField, fkData, 2, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'localizacao', 'localizacaocep', 'cep', TWideStringField, fkData, 16, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'localizacao', 'localizacaoloc_nu_sequencial', 'loc_nu_sequencial', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'localizacao', 'localizacaolog_tipo_logradouro', 'log_tipo_logradouro', TWideStringField, fkData, 72, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'localizacao', 'localizacaolog_no', 'log_no', TWideStringField, fkData, 70, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'localizacao', 'localizacaolog_no_sem_acento', 'log_no_sem_acento', TWideStringField, fkData, 70, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'localizacao', 'localizacaobai_nu_sequencial_ini', 'bai_nu_sequencial_ini', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'localizacao', 'localizacaoloc_no', 'loc_no', TWideStringField, fkLookup, 60, 0, False, '', '', '', '', 0, 'loc_nu_sequencial', 'log_localidade', 'loc_nu_sequencial', 'loc_no', True);
  RegisterRuntimeField(TFrmPesqEndereco, 'localizacao', 'localizacaobai_no', 'bai_no', TWideStringField, fkLookup, 60, 0, False, '', '', '', '', 0, 'bai_nu_sequencial_ini', 'ZQbairro', 'bai_nu_sequencial', 'bai_no', True);
  RegisterRuntimeDataSet(TFrmPesqEndereco, 'ZQbairro', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'ZQbairro', 'ZQbairrobai_no', 'bai_no', TWideStringField, fkData, 72, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqEndereco, 'ZQbairro', 'ZQbairrobai_nu_sequencial', 'bai_nu_sequencial', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);

end.
