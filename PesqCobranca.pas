unit PesqCobranca;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XLabel3D, XBanner, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Grids, DBGrids, Mask, DBCtrls, Vcl.ExtCtrls;

type
  TFrmPesqCobranca = class(TForm)
    XLTit: TXLabel3D;
    DBGrid1: TDBGrid;
    Label1: TLabel;
    DS_TempCliReceb: TDataSource;
    ZQTempCliReceb: TZQuery;
    Edit1: TEdit;
    XBanner6: TXBanner;
    DBECod: TDBEdit;
    Label11: TLabel;
    LReg: TLabel;
    ZQCheque: TZQuery;
    DS_zqcheque: TDataSource;

















































    Label2: TLabel;
    Edit2: TEdit;
    Label3: TLabel;
    DBGrid2: TDBGrid;
    Label4: TLabel;
    ednome: TComboBox;
    veio: TLabel;
































    RGFiltro: TRadioGroup;
    procedure FormShow(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
    procedure DBGrid1KeyPress(Sender: TObject; var Key: Char);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Edit2KeyPress(Sender: TObject; var Key: Char);
    procedure Edit2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBGrid2KeyPress(Sender: TObject; var Key: Char);
    procedure DS_TempCliRecebDataChange(Sender: TObject; Field: TField);
    procedure ednomeDropDown(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmPesqCobranca: TFrmPesqCobranca;
  limite,totreg:integer;
implementation

uses Tabelas, Cad_Recebimento, Funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrmPesqCobranca.FormShow(Sender: TObject);
begin
  ednome.Clear;
  edit1.Clear;
  edit2.Clear;
  limite:=1;
  ZQTempCliReceb.SQL.Clear;
  ZQTempCliReceb.close;
  ZQTempCliReceb.SQL.Add('select p1.idpaticipante,p1.nome_parte,p1.codpaginc,p1.codrecinc,p1.codpagBx,p1.codrecBx,rec.idrecebimento,rec.documento,rec.cliente,');
  ZQTempCliReceb.SQL.Add('rec.usuario,rec.Dt_Entrada,rec.Dt_Vencimento,rec.Valor,rec.ordem,rec.TipDoc,rec.saldo,rec.marcar,rec.RefBaixa,rec.refvinda,rec.contabil,');
  ZQTempCliReceb.SQL.Add('rec.empresa,rec.custodaparcela,rec.origem,rec.adversa,rec.recpag,rec.numordem,rec.nomeadversa,rec.idloteamento,rec.venda_idvenda,rec.quadralote,rec.numboleto,');
  ZQTempCliReceb.SQL.Add(' lt.apelido from recebimento as rec join participante as p1 on rec.cliente=p1.idpaticipante');
  ZQTempCliReceb.SQL.Add(' join participante as p2 on rec.adversa=p2.idpaticipante join loteamento as lt on lt.idloteamento=rec.idloteamento');
  ZQTempCliReceb.SQL.Add(' group by rec.quadralote ');
  ZQTempCliReceb.Open;
  totreg:=ZQTempCliReceb.RecordCount;
  LReg.Caption := inttostr(totreg);
  if (ZQTempCliReceb.Active) and (ZQTempCliReceb.RecordCount>0) Then
    DBGrid1.SetFocus
  else
    close;
  ednome.SetFocus;
end;

procedure TFrmPesqCobranca.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 then
    close;

end;

procedure TFrmPesqCobranca.Edit1KeyPress(Sender: TObject;
  var Key: Char);
  var
  posi:integer;
begin
  if Key = #13 then
  begin
    veio.Caption:='S';
    ZQTempCliReceb.close;
    ZQTempCliReceb.SQL.Clear;
    ZQTempCliReceb.SQL.Add('select p1.idpaticipante,p1.nome_parte,p1.codpaginc,p1.codrecinc,p1.codpagBx,p1.codrecBx,rec.idrecebimento,rec.documento,rec.cliente,');
    ZQTempCliReceb.SQL.Add('rec.usuario,rec.Dt_Entrada,rec.Dt_Vencimento,rec.Valor,rec.ordem,rec.TipDoc,rec.saldo,rec.marcar,rec.RefBaixa,rec.refvinda,rec.contabil,');
    ZQTempCliReceb.SQL.Add('rec.empresa,rec.custodaparcela,rec.origem,rec.adversa,rec.recpag,rec.numordem,rec.nomeadversa,rec.idloteamento,rec.venda_idvenda,rec.quadralote,rec.numboleto,');
    ZQTempCliReceb.SQL.Add(' lt.apelido from recebimento as rec join participante as p1 on rec.cliente=p1.idpaticipante');
    ZQTempCliReceb.SQL.Add(' join participante as p2 on rec.adversa=p2.idpaticipante join loteamento as lt on lt.idloteamento=rec.idloteamento');
    if not empty(ednome.Text) then
    begin
      posi:=pos('|',ednome.text);
      dec(posi);
      ZQTempCliReceb.SQL.Add(' where (rec.idloteamento='+quotedstr(copy(ednome.Text,1,posi))+') and (p1.nome_parte like '+quotedstr('%'+Edit1.Text+'%'));
    end
    else
    begin
      ZQTempCliReceb.SQL.Add(' where (p1.nome_parte like '+quotedstr('%'+Edit1.Text+'%'));
    end;
    ZQTempCliReceb.SQL.Add(' or p2.nome_parte like '+quotedstr('%'+Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.documento like '+quotedstr('%'+Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.idrecebimento like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.refbaixa like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.refvinda like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.quadralote like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.numboleto like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.ordem like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.quadralote like '+quotedstr('%'+Edit1.Text+'%'));
    if (not empty(Edit1.Text)) and (StrIsNumero(Edit1.Text)) Then
       ZQTempCliReceb.SQL.Add(' or rec.valor='+Edit1.Text);
    ZQTempCliReceb.SQL.Add(' or rec.numboleto like '+quotedstr('%'+Edit1.Text+'%')+')');
    ZQTempCliReceb.SQL.Add(' group by rec.quadralote');
    ZQTempCliReceb.Open;
    totreg:=ZQTempCliReceb.RecordCount;
    LReg.Caption := inttostr(totreg);
    if (ZQTempCliReceb.Active) and (ZQTempCliReceb.RecordCount>0) Then
        DBGrid1.SetFocus
    else
        close;
  end;
end;

procedure TFrmPesqCobranca.DBGrid1KeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) and (ZQTempCliReceb.RecordCount>0) Then Begin
    DM_tabelas.ZQRecebimento.Locate('idrecebimento',ZQTempCliReceb.FieldByName('idrecebimento').AsLargeInt,[]);
    Close;
  end;
end;

procedure TFrmPesqCobranca.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = ZQTempCliReceb.FieldByName('idrecebimento').Text Then Begin
    DBGrid1.Canvas.Brush.Color := $006CFFFF;
    DBGrid1.Canvas.Font.Color := $00A80000;
//    DBGReceb.Canvas.Font.Style := [FsBold];
  end;
  DBGrid1.DefaultDrawDataCell(Rect, DBGrid1.columns[datacol].field, State);

end;

procedure TFrmPesqCobranca.Edit2KeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
  begin
    ZQCheque.close;
    ZQCheque.SQL.Clear;
    ZQCheque.SQL.Add('select * from cheque_rec as ch join recebimento as rc on ch.idrecebimento=rc.idrecebimento');
    ZQCheque.SQL.Add(' where ch.numero like '+quotedstr('%'+edit2.Text+'%')+' order by numero');
    ZQCheque.Open;
    //  LReg.Caption := inttostr(ZQCheque.RecordCount);
    if (ZQCheque.Active) and (ZQCheque.RecordCount>0) Then
    begin
      ZQTempCliReceb.Locate('idrecebimento',ZQCheque.FieldByName('idrecebimento').AsInteger,[]);
      DM_tabelas.ZQRecebimento.Locate('idrecebimento',ZQTempCliReceb.FieldByName('idrecebimento').AsLargeInt,[]);
      DBGrid2.Visible:=true;
      DBGrid2.SetFocus;
    end
    else
    begin
      ZQCheque.close;
      DBGrid2.Visible:=false;
      showmessage('Número de Cheque não Encontrado.');
    end;
  end;
end;

procedure TFrmPesqCobranca.Edit2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_F2 then
  begin
    zqcheque.Next;
    ZQTempCliReceb.Locate('idrecebimento',ZQCheque.FieldByName('idrecebimento').AsInteger,[]);
    DM_tabelas.ZQRecebimento.Locate('idrecebimento',ZQTempCliReceb.FieldByName('idrecebimento').AsLargeInt,[]);
  end;
end;

procedure TFrmPesqCobranca.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQCheque.close;
end;

procedure TFrmPesqCobranca.DBGrid2KeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) and (ZQTempCliReceb.RecordCount>0) Then
  Begin
//    ZQTempCliReceb.Locate('idrecebimento',zqchequeidrecebimento.Value,[]);
    DM_tabelas.ZQRecebimento.Locate('idrecebimento',ZQCheque.FieldByName('idrecebimento').AsInteger,[]);
//    dbgrid1.setfocus;
    dbgrid2.Visible:=false;
    Close;
  end;

end;

procedure TFrmPesqCobranca.DS_TempCliRecebDataChange(Sender: TObject;
  Field: TField);
begin
  if ZQTempCliReceb.Eof then
  begin
    if veio.Caption='S' then
    begin
      Showmessage('Dados da Pesquisa Não Encontrados.');
    end
    else
    begin
      limite:=limite+500;
      ZQTempCliReceb.SQL.Clear;
      ZQTempCliReceb.SQL.Add('select p1.idpaticipante,p1.nome_parte,p1.codpaginc,p1.codrecinc,p1.codpagBx,p1.codrecBx,rec.idrecebimento,rec.documento,rec.cliente,');
      ZQTempCliReceb.SQL.Add('rec.usuario,rec.Dt_Entrada,rec.Dt_Vencimento,rec.Valor,rec.ordem,rec.TipDoc,rec.saldo,rec.marcar,rec.RefBaixa,rec.refvinda,rec.contabil,');
      ZQTempCliReceb.SQL.Add('empresa,custodaparcela,origem,adversa,recpag,numordem,nomeadversa,idloteamento,venda_idvenda,rec.quadralote,rec.numboleto,');
      ZQTempCliReceb.SQL.Add(' lt.apelido from recebimento as rec join participante as p1 on rec.cliente=p1.idpaticipante');
      ZQTempCliReceb.SQL.Add(' join participante as p2 on rec.adversa=p2.idpaticipante join loteamento as lt on lt.idloteamento=rec.idloteamento');
      ZQTempCliReceb.SQL.Add(' group by rec.quadralote');
      ZQTempCliReceb.Open;
    end;
    veio.Caption:='N';  
    totreg:=totreg+ZQTempCliReceb.RecordCount;
    LReg.Caption := inttostr(totreg);
  end;
end;

procedure TFrmPesqCobranca.ednomeDropDown(Sender: TObject);
begin
  edNome.Clear;
  DM_Tabelas.ZQLoteamento.First;
  while not DM_Tabelas.ZQLoteamento.Eof do
  begin
    edNome.Items.Add(DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').AsString+'|'+DM_Tabelas.ZQLoteamento.FieldByName('Apelido').AsString);
    DM_Tabelas.ZQLoteamento.Next;
  end;
  DM_Tabelas.ZQLoteamento.First;
end;


procedure TFrmPesqCobranca.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmPesqCobranca, 'ZQTempCliReceb', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliReceborigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQTempCliReceb', 'ZQTempCliRecebapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmPesqCobranca, 'ZQCheque', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequebanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequedono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequevalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequedeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequealias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequenumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequesq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeidrecebimento_1', 'idrecebimento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequedocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequecliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequesaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequemarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequerefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequecontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequecustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequerecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequenumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequevenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequequadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequenumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequesq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequesomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequenomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqCobranca, 'ZQCheque', 'ZQChequeData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
