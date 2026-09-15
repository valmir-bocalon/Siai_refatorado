unit PesqRecebimento2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XLabel3D, XBanner, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Grids, DBGrids, Mask, DBCtrls;

type
  TFrmPesqRecebimento2 = class(TForm)
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
  FrmPesqRecebimento2: TFrmPesqRecebimento2;
  FrmPesqRecebimento2Target: TForm;
  limite,totreg:integer;
implementation

uses Tabelas, Cad_Recebimento, Funcoes,  ReajusteDeParcelas, uRuntimeFields;

{$R *.dfm}

function PesquisaReajusteTarget: TForm;
begin
  Result := FrmPesqRecebimento2Target;
  if (Result = nil) and (Frm_ReajusteDeParcelas <> nil) then
    Result := Frm_ReajusteDeParcelas;
end;

function PesquisaReajusteDataSet: TDataSet;
var
  LTarget: TForm;
begin
  Result := nil;
  LTarget := PesquisaReajusteTarget;
  if LTarget <> nil then
    Result := LTarget.FindComponent('ZQRecebimento') as TDataSet;
end;

function PesquisaReajusteQuery: TZQuery;
var
  LTarget: TForm;
begin
  Result := nil;
  LTarget := PesquisaReajusteTarget;
  if LTarget <> nil then
    Result := LTarget.FindComponent('ZQRecebimento') as TZQuery;
end;

function PesquisaReajusteCheckBox: TCheckBox;
var
  LTarget: TForm;
begin
  Result := nil;
  LTarget := PesquisaReajusteTarget;
  if LTarget <> nil then
    Result := LTarget.FindComponent('CBDesconsidera') as TCheckBox;
end;

procedure TFrmPesqRecebimento2.FormShow(Sender: TObject);
begin
  ednome.Clear;
  edit1.Clear;
  edit2.Clear;
  limite:=1;
  ZQTempCliReceb.SQL.Clear;
  ZQTempCliReceb.SQL.Add('select p1.idpaticipante,p1.nome_parte,p1.codpaginc,p1.codrecinc,p1.codpagBx,p1.codrecBx,rec.idrecebimento,rec.documento,rec.cliente,');
  ZQTempCliReceb.SQL.Add('rec.usuario,rec.Dt_Entrada,rec.Dt_Vencimento,rec.Valor,rec.ordem,rec.TipDoc,rec.saldo,rec.marcar,rec.RefBaixa,rec.refvinda,rec.contabil,');
  ZQTempCliReceb.SQL.Add('empresa,custodaparcela,origem,adversa,recpag,numordem,nomeadversa,idloteamento,venda_idvenda,rec.quadralote,rec.numboleto');
  ZQTempCliReceb.SQL.Add(' from recebimento as rec left join participante as p1 on rec.cliente=p1.idpaticipante');
  ZQTempCliReceb.SQL.Add(' left join participante as p2 on rec.adversa=p2.idpaticipante where (rec.Parcelas_fixas<>''S'') and (rec.TipDoc=''BO'') and (rec.saldo>''0'')');
  ZQTempCliReceb.SQL.Add(' group by rec.idrecebimento LIMIT '+alltrim(inttostr(limite))+',500');
  ZQTempCliReceb.Open;

  totreg:=ZQTempCliReceb.RecordCount;
  LReg.Caption := inttostr(totreg);
  if (ZQTempCliReceb.Active) and (ZQTempCliReceb.RecordCount>0) Then
    DBGrid1.SetFocus
  else
    close;
  ednome.SetFocus;
end;

procedure TFrmPesqRecebimento2.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 then
    close;

end;

procedure TFrmPesqRecebimento2.Edit1KeyPress(Sender: TObject;
  var Key: Char);
  var
  posi:integer;
  ValorPesquisa: Double;
  CodigoValor: Integer;
  TextoValor: string;
begin
  if Key = #13 then
  begin
    veio.Caption:='S';
    ZQTempCliReceb.close;
    ZQTempCliReceb.SQL.Clear;
    ZQTempCliReceb.SQL.Add('select p1.idpaticipante,p1.nome_parte,p1.codpaginc,p1.codrecinc,p1.codpagBx,p1.codrecBx,rec.idrecebimento,rec.documento,rec.cliente,');
    ZQTempCliReceb.SQL.Add('rec.usuario,rec.Dt_Entrada,rec.Dt_Vencimento,rec.Valor,rec.ordem,rec.TipDoc,rec.saldo,rec.marcar,rec.RefBaixa,rec.refvinda,rec.contabil,');
    ZQTempCliReceb.SQL.Add('rec.empresa,rec.custodaparcela,rec.origem,rec.adversa,rec.recpag,rec.numordem,rec.nomeadversa,rec.idloteamento,rec.venda_idvenda,rec.quadralote,rec.numboleto');
    ZQTempCliReceb.SQL.Add(' from recebimento as rec left join participante as p1 on rec.cliente=p1.idpaticipante');
    ZQTempCliReceb.SQL.Add(' left join participante as p2 on rec.adversa=p2.idpaticipante');
    if not empty(ednome.Text) then
    begin
      posi:=pos('|',ednome.text);
      dec(posi);
      ZQTempCliReceb.SQL.Add(' where (rec.Parcelas_fixas<>''S'') and (rec.TipDoc=''BO'') and (rec.saldo>''0'') and (idloteamento='+quotedstr(copy(ednome.Text,1,posi))+') and (p1.nome_parte like '+quotedstr('%'+Edit1.Text+'%'));
    end
    else
    begin
      ZQTempCliReceb.SQL.Add(' where (rec.Parcelas_fixas<>''S'') and (rec.TipDoc=''BO'') and (rec.saldo>''0'') and (p1.nome_parte like '+quotedstr('%'+Edit1.Text+'%'));
    end;
    ZQTempCliReceb.SQL.Add(' or p2.nome_parte like '+quotedstr('%'+Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.documento like '+quotedstr('%'+Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.idrecebimento like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.refbaixa like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.refvinda like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.quadralote like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.numboleto like '+quotedstr(Edit1.Text+'%'));
    ZQTempCliReceb.SQL.Add(' or rec.ordem like '+quotedstr(Edit1.Text+'%'));
    TextoValor := StringReplace(Trim(Edit1.Text), ',', '.', [rfReplaceAll]);
    Val(TextoValor, ValorPesquisa, CodigoValor);
    if (TextoValor <> '') and (CodigoValor = 0) then
    begin
      ZQTempCliReceb.SQL.Add(' or rec.valor=:valor_pesquisa');
      ZQTempCliReceb.ParamByName('valor_pesquisa').AsFloat := ValorPesquisa;
    end;
    ZQTempCliReceb.SQL.Add(' or rec.numboleto like '+quotedstr('%'+Edit1.Text+'%')+')');
    ZQTempCliReceb.SQL.Add(' group by rec.idrecebimento');
    ZQTempCliReceb.Open;
//    showmessage(ZQTempCliReceb.SQL.Text);
    totreg:=ZQTempCliReceb.RecordCount;
    LReg.Caption := inttostr(totreg);
    if (ZQTempCliReceb.Active) and (ZQTempCliReceb.RecordCount>0) Then
        DBGrid1.SetFocus
    else
        close;
  end;
end;

procedure TFrmPesqRecebimento2.DBGrid1KeyPress(Sender: TObject;
  var Key: Char);
var
  LQuery: TZQuery;
  LCheckBox: TCheckBox;
begin
  if (Key = #13) and (ZQTempCliReceb.RecordCount>0) Then Begin
    LQuery := PesquisaReajusteQuery;
    LCheckBox := PesquisaReajusteCheckBox;
    if LQuery = nil then
      Exit;

    LQuery.Close;
    LQuery.SQL.Clear;
{  ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
  ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
  ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}

{  if CBDesconsidera.Checked=false then
     ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
  else
     ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');}


    if (LCheckBox <> nil) and not LCheckBox.Checked then
        LQuery.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') order by DT_Vencimento')
    else
        LQuery.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') order by DT_Vencimento');

    LQuery.Open;

    LQuery.Locate('idrecebimento',ZQTempCliReceb.FieldByName('idrecebimento').AsLargeInt,[]);
{    Frm_ReajusteDeParcelas.ZQRecebimento.Filtered:=false;
    Frm_ReajusteDeParcelas.ZQRecebimento.Filter:='idrecebimento='+ZQTempCliRecebidrecebimento.text;
    Frm_ReajusteDeParcelas.ZQRecebimento.Filtered:=true;}
    LQuery.RecordCount;
    Close;
  end;
end;

procedure TFrmPesqRecebimento2.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if not ZQTempCliReceb.Active or ZQTempCliReceb.IsEmpty then
  begin
    DBGrid1.DefaultDrawColumnCell(Rect, DataCol, Column, State);
    Exit;
  end;
  if DBECod.Text = ZQTempCliReceb.FieldByName('idrecebimento').Text Then Begin
    DBGrid1.Canvas.Brush.Color := $006CFFFF;
    DBGrid1.Canvas.Font.Color := $00A80000;
//    DBGReceb.Canvas.Font.Style := [FsBold];
  end;
  DBGrid1.DefaultDrawDataCell(Rect, DBGrid1.columns[datacol].field, State);

end;

procedure TFrmPesqRecebimento2.Edit2KeyPress(Sender: TObject;
  var Key: Char);
var
  LDataSet: TDataSet;
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
      LDataSet := PesquisaReajusteDataSet;
      ZQTempCliReceb.Locate('idrecebimento',ZQCheque.FieldByName('idrecebimento').AsInteger,[]);
      if LDataSet <> nil then
        LDataSet.Locate('idrecebimento',ZQTempCliReceb.FieldByName('idrecebimento').AsLargeInt,[]);
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

procedure TFrmPesqRecebimento2.Edit2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  LDataSet: TDataSet;
begin
  if (Key = VK_F2) and ZQCheque.Active and not ZQCheque.IsEmpty then
  begin
    LDataSet := PesquisaReajusteDataSet;
    zqcheque.Next;
    ZQTempCliReceb.Locate('idrecebimento',ZQCheque.FieldByName('idrecebimento').AsInteger,[]);
    if LDataSet <> nil then
      LDataSet.Locate('idrecebimento',ZQTempCliReceb.FieldByName('idrecebimento').AsLargeInt,[]);
  end;
end;

procedure TFrmPesqRecebimento2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQCheque.close;
end;

procedure TFrmPesqRecebimento2.DBGrid2KeyPress(Sender: TObject;
  var Key: Char);
var
  LDataSet: TDataSet;
begin
  if (Key = #13) and ZQCheque.Active and not ZQCheque.IsEmpty and
     (ZQTempCliReceb.RecordCount>0) Then
  Begin
//    ZQTempCliReceb.Locate('idrecebimento',zqchequeidrecebimento.Value,[]);
    LDataSet := PesquisaReajusteDataSet;
    if LDataSet <> nil then
      LDataSet.Locate('idrecebimento',ZQCheque.FieldByName('idrecebimento').AsInteger,[]);
//    dbgrid1.setfocus;
    dbgrid2.Visible:=false;
    Close;
  end;

end;

procedure TFrmPesqRecebimento2.DS_TempCliRecebDataChange(Sender: TObject;
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
      ZQTempCliReceb.SQL.Add('empresa,custodaparcela,origem,adversa,recpag,numordem,nomeadversa,idloteamento,venda_idvenda,rec.quadralote,rec.numboleto');
      ZQTempCliReceb.SQL.Add(' from recebimento as rec left join participante as p1 on rec.cliente=p1.idpaticipante');
      ZQTempCliReceb.SQL.Add(' left join participante as p2 on rec.adversa=p2.idpaticipante ');
      ZQTempCliReceb.SQL.Add(' where (rec.Parcelas_fixas<>''S'') and (rec.TipDoc=''BO'') and (rec.saldo>''0'')');
      ZQTempCliReceb.SQL.Add(' group by rec.idrecebimento LIMIT '+alltrim(inttostr(limite))+',500 ');
      ZQTempCliReceb.Open;
    end;
    veio.Caption:='N';  
    totreg:=totreg+ZQTempCliReceb.RecordCount;
    LReg.Caption := inttostr(totreg);
  end;
end;

procedure TFrmPesqRecebimento2.ednomeDropDown(Sender: TObject);
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


procedure TFrmPesqRecebimento2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmPesqRecebimento2, 'ZQTempCliReceb', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliReceborigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebnomeadversa', 'nomeadversa', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQTempCliReceb', 'ZQTempCliRecebNome_loteamento', 'Nome_loteamento', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeDataSet(TFrmPesqRecebimento2, 'ZQCheque', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequebanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequedono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequevalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequedeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequealias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequenumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequesq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeidrecebimento_1', 'idrecebimento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequedocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequecliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequesaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequemarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequerefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequecontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequecustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequerecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequenumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequevenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequequadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequenumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequesq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequesomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequenomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento2, 'ZQCheque', 'ZQChequeData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
