unit PesqRecebimento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XLabel3D, XBanner, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Grids, DBGrids, Mask, DBCtrls;

type
  TFrmPesqRecebimento = class(TForm)
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
  FrmPesqRecebimento: TFrmPesqRecebimento;
  limite,totreg:integer;
implementation

uses Tabelas, Cad_Recebimento, Funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrmPesqRecebimento.FormShow(Sender: TObject);
begin
  ednome.Clear;
  edit1.Clear;
  edit2.Clear;
  limite:=1;
//  ZQTempCliReceb.SQL.Clear;
//  ZQTempCliReceb.SQL.Add('select p1.idpaticipante,p1.nome_parte,p1.codpaginc,p1.codrecinc,p1.codpagBx,p1.codrecBx,rec.idrecebimento,rec.documento,rec.cliente,');
//  ZQTempCliReceb.SQL.Add('rec.usuario,rec.Dt_Entrada,rec.Dt_Vencimento,rec.Valor,rec.ordem,rec.TipDoc,rec.saldo,rec.marcar,rec.RefBaixa,rec.refvinda,rec.contabil,');
//  ZQTempCliReceb.SQL.Add('rec.empresa,rec.custodaparcela,rec.origem,rec.adversa,rec.recpag,rec.numordem,rec.nomeadversa,rec.idloteamento,rec.venda_idvenda,rec.quadralote,rec.numboleto');
//  ZQTempCliReceb.SQL.Add(' from recebimento as rec left join participante as p1 on rec.cliente=p1.idpaticipante');
//  ZQTempCliReceb.SQL.Add(' left join participante as p2 on rec.adversa=p2.idpaticipante ');
//  ZQTempCliReceb.SQL.Add(' group by rec.idrecebimento LIMIT '+alltrim(inttostr(limite))+',50 ');
//  ZQTempCliReceb.Open;
//  totreg:=ZQTempCliReceb.RecordCount;
//  LReg.Caption := inttostr(totreg);
//  if (ZQTempCliReceb.Active) and (ZQTempCliReceb.RecordCount>0) Then
//    DBGrid1.SetFocus
//  else
//    close;
  ednome.SetFocus;
end;

procedure TFrmPesqRecebimento.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 then
    close;

end;

procedure TFrmPesqRecebimento.Edit1KeyPress(Sender: TObject;
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
      ZQTempCliReceb.SQL.Add(' where (idloteamento='+quotedstr(copy(ednome.Text,1,posi))+') and (p1.nome_parte like '+quotedstr('%'+Edit1.Text+'%'));
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
    totreg:=ZQTempCliReceb.RecordCount;
    LReg.Caption := inttostr(totreg);
    if (ZQTempCliReceb.Active) and (ZQTempCliReceb.RecordCount>0) Then
        DBGrid1.SetFocus
    else
        close;
  end;
end;

procedure TFrmPesqRecebimento.DBGrid1KeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) and (ZQTempCliReceb.RecordCount>0) Then
  Begin

    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where idrecebimento='+quotedstr(ZQTempCliReceb.FieldByName('idrecebimento').text)+' order by DT_Vencimento ');
//    DM_tabelas.ZQRecebimento.SQL.Add('Select * from Recebimento order by DT_Vencimento ');
    DM_tabelas.ZQRecebimento.Open;

    DM_tabelas.ZQRecebimento.Locate('idrecebimento',ZQTempCliReceb.FieldByName('idrecebimento').AsLargeInt,[]);
    Close;
  end;
end;

procedure TFrmPesqRecebimento.DBGrid1DrawColumnCell(Sender: TObject;
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

procedure TFrmPesqRecebimento.Edit2KeyPress(Sender: TObject;
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
      showmessage('N�mero de Cheque n�o Encontrado.');
    end;
  end;
end;

procedure TFrmPesqRecebimento.Edit2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_F2 then
  begin
    zqcheque.Next;
    ZQTempCliReceb.Locate('idrecebimento',ZQCheque.FieldByName('idrecebimento').AsInteger,[]);
    DM_tabelas.ZQRecebimento.Locate('idrecebimento',ZQTempCliReceb.FieldByName('idrecebimento').AsLargeInt,[]);        
  end;
end;

procedure TFrmPesqRecebimento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQCheque.close;
end;

procedure TFrmPesqRecebimento.DBGrid2KeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) and (ZQTempCliReceb.RecordCount>0) Then
  Begin
//    ZQTempCliReceb.Locate('idrecebimento',zqchequeidrecebimento.Value,[]);
      DM_Tabelas.ZQRecebimento.Close;
      DM_Tabelas.ZQRecebimento.SQL.Clear;
      DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
      DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
      DM_Tabelas.ZQRecebimento.SQL.Add(' from Recebimento where idrecebimento='+quotedstr(ZQCheque.FieldByName('idrecebimento').Text)+' order by DT_Vencimento');
      DM_Tabelas.ZQRecebimento.open;


    DM_tabelas.ZQRecebimento.Locate('idrecebimento',ZQCheque.FieldByName('idrecebimento').AsInteger,[]);
//    dbgrid1.setfocus;
    dbgrid2.Visible:=false;
    Close;
  end;

end;

procedure TFrmPesqRecebimento.DS_TempCliRecebDataChange(Sender: TObject;
  Field: TField);
begin
  if ZQTempCliReceb.Eof then
  begin
    if veio.Caption='S' then
    begin
      Showmessage('Dados da Pesquisa N�o Encontrados.');
    end
    else
    begin
      limite:=limite+50;
      ZQTempCliReceb.SQL.Clear;
      ZQTempCliReceb.SQL.Add('select p1.idpaticipante,p1.nome_parte,p1.codpaginc,p1.codrecinc,p1.codpagBx,p1.codrecBx,rec.idrecebimento,rec.documento,rec.cliente,');
      ZQTempCliReceb.SQL.Add('rec.usuario,rec.Dt_Entrada,rec.Dt_Vencimento,rec.Valor,rec.ordem,rec.TipDoc,rec.saldo,rec.marcar,rec.RefBaixa,rec.refvinda,rec.contabil,');
      ZQTempCliReceb.SQL.Add('empresa,custodaparcela,origem,adversa,recpag,numordem,nomeadversa,idloteamento,venda_idvenda,rec.quadralote,rec.numboleto');
      ZQTempCliReceb.SQL.Add(' from recebimento as rec left join participante as p1 on rec.cliente=p1.idpaticipante');
      ZQTempCliReceb.SQL.Add(' left join participante as p2 on rec.adversa=p2.idpaticipante ');
      ZQTempCliReceb.SQL.Add(' group by rec.idrecebimento LIMIT '+alltrim(inttostr(limite))+',500 ');
      ZQTempCliReceb.Open;
    end;
    veio.Caption:='N';  
    totreg:=totreg+ZQTempCliReceb.RecordCount;
    LReg.Caption := inttostr(totreg);
  end;
end;

procedure TFrmPesqRecebimento.ednomeDropDown(Sender: TObject);
begin
  edNome.Clear;
  DM_Tabelas.ZQLoteamento.First;
  DM_Tabelas.ZQLoteamento.DisableControls;
  while not DM_Tabelas.ZQLoteamento.Eof do
  begin
    edNome.Items.Add(DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').AsString+'|'+DM_Tabelas.ZQLoteamento.FieldByName('Apelido').AsString);
    DM_Tabelas.ZQLoteamento.Next;
  end;
  DM_Tabelas.ZQLoteamento.First;
  DM_Tabelas.ZQLoteamento.EnableControls;
end;


procedure TFrmPesqRecebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmPesqRecebimento, 'ZQTempCliReceb', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliReceborigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebnomeadversa', 'nomeadversa', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQTempCliReceb', 'ZQTempCliRecebNome_loteamento', 'Nome_loteamento', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeDataSet(TFrmPesqRecebimento, 'ZQCheque', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequebanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequedono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequevalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequedeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequealias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequenumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequesq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeidrecebimento_1', 'idrecebimento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequedocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequecliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequesaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequemarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequerefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequecontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequecustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequerecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequenumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequevenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequequadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequenumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequesq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequesomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequenomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento, 'ZQCheque', 'ZQChequeData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
end.
