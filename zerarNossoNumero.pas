unit zerarNossoNumero;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, wwdbdatetimepicker, StdCtrls, Mask, DBCtrls,
  XBanner, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Grids,
  DBGrids, ImgList, XNum, DBClient, XDBDate, XDBNum, ComCtrls,
  JvExComCtrls, JvDateTimePicker, XEdit, XDBEdit, ExtCtrls, JvExControls,
  JvSimIndicator, TFlatGaugeUnit, XDate, Provider, System.ImageList,
  FnpNumericEdit;
type
  TFrmzerarNossoNumero = class(TForm)
    Label16: TLabel;
    DXBBaixaGravar: TdxButtonArround;
    DS_RecBai: TDataSource;
    DBGBaixando: TDBGrid;
    DXBFechar: TdxButtonArround;
    ImageList1: TImageList;
    Label9: TLabel;
    LREg: TLabel;
    Label31: TLabel;
    Label33: TLabel;
    XBanner1: TXBanner;
    DXBPesquisar: TdxButtonArround;
    btndes: TButton;
    Label35: TLabel;
    Elote: TEdit;
    barravertical: TFlatGauge;
    ZQRecBai: TZQuery;









































    ZQremessa_receb: TZQuery;
    procedure Atualiza_tela;
    procedure DXBFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DBGBaixandoDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGBaixandoColEnter(Sender: TObject);
    procedure DBGBaixandoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DXBBaixaGravarClick(Sender: TObject);
    procedure filtrabaixa;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ZQRecBaiCalcFields(DataSet: TDataSet);
    procedure XNEJurosEnter(Sender: TObject);
    procedure DXBPesquisarClick(Sender: TObject);
    procedure btndesClick(Sender: TObject);
    procedure Filtere(Sender: TObject);
    procedure DBGBaixandoExit(Sender: TObject);
    procedure Label9MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmzerarNossoNumero: TFrmzerarNossoNumero;
  ql,VArDoc, varrecpag : String;
  contador:integer;


implementation

uses tabelas, Cad_Recebimento, funcoes, ImpRecibo, PesqRecebimento_bx, uRuntimeFields;

{$R *.dfm}


// procedure de mover o formulario sem bordas
procedure MovimentaObject(Sender:TObject;Button:TMouseButton;Shift:TShiftState;X,Y:Integer;Formulario:TForm); 
 var ObjectPos,MousePosMov:TPoint;Pt:TPoint;fHandle:HWND; 
 begin 
   GetCursorPos(Pt); 
   ObjectPos.X:=Formulario.Left; 
   ObjectPos.Y:=Formulario.Top; 
   if(Sender is TForm)then 
     fHandle:=TWinControl(Sender).Handle 
   else 
     fHandle:=TWinControl(Sender).Parent.Handle; 
   while DragDetect(fHandle,ObjectPos) do 
   begin 
     GetCursorPos(MousePosMov); 
     Formulario.Left:=MousePosMov.X-X-3; 
     Formulario.Top:=MousePosMov.Y-Y-3; 
     Application.ProcessMessages;
   end; 
 end; 


procedure TFrmzerarNossoNumero.Atualiza_tela;
Var
  Varreg, varmarc, varvenc : integer;
Begin
  if ZQRecBai.active=false then
     ZQRecBai.open;

  Varreg := ZQRecBai.RecNo;
  ZQRecBai.First;


  BarraVertical.MaxValue:=ZQRecBai.RecordCount;
  BarraVertical.Progress:=0;
//  BarraVertical.Visible:=true;

  ZQRecBai.first;

  ZQRecBai.EnableControls;
  DS_RecBai.DataSet:=ZQRecBai;
  BarraVertical.Progress:=0;
  BarraVertical.Visible:=false;
  if ZQRecBai.RecordCount>0 then
     DXBBaixaGravar.Enabled := True;

  ZQRecBai.RecNo := Varreg;
End;

procedure TFrmzerarNossoNumero.DXBFecharClick(Sender: TObject);
begin
  elote.clear;
  Close;
end;

procedure TFrmzerarNossoNumero.FormShow(Sender: TObject);
begin
  Elote.SetFocus;
end;

procedure TFrmzerarNossoNumero.DBGBaixandoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);

begin
  if ZQRecBai.active=true then
  begin

    DBGBaixando.DefaultDrawDataCell(Rect, DBGBaixando.columns[datacol].field, State);
    if Column.Field = ZQRecBai.FieldByName('marcar') then
    begin
      DBGBaixando.Canvas.FillRect(Rect);
      if ZQRecBai.FieldByName('marcar').AsString = '0' then
        ImageList1.Draw(DBGBaixando.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
      else
        ImageList1.Draw(DBGBaixando.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
    end;
  end;
end;

procedure TFrmzerarNossoNumero.DBGBaixandoColEnter(Sender: TObject);
begin
  if DBGBaixando.SelectedField = ZQRecBai.FieldByName('marcar') then
    DBGBaixando.Options := DBGBaixando.Options - [dgEditing]
  else
    DBGBaixando.Options := DBGBaixando.Options + [dgEditing];
end;

procedure TFrmzerarNossoNumero.DBGBaixandoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DBGBaixando.SelectedField.fieldname = 'marcar' Then Begin
    ZQRecBai.Edit;
    if  ZQRecBai.FieldByName('marcar').AsString = '0' Then
      ZQRecBai.FieldByName('marcar').AsString := '1'
    else
      ZQRecBai.FieldByName('marcar').AsString := '0';
    ZQRecBai.Post;
//    Atualiza_tela;
  end;
end;

procedure TFrmzerarNossoNumero.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then
    close;
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmzerarNossoNumero.DXBBaixaGravarClick(Sender: TObject);
Var
  Varpago, percentbaixa, baixatot, varjuros, varx,
  vvrparc,varsomarepasse, varsomarepasseautomatico : Double;
  dd,posi,posi2,contador,varreg : Integer;
  Resp,dtvc,ctbl,dia,varordem, varnumordem1, varnumordem2 : string;

begin

  DXBBaixaGravar.Enabled := False;
  DXBFechar.Enabled := False;

  ZQRecBai.Filtered:=false;
  ZQRecBai.Filter := 'marcar=0';
  ZQRecBai.Filtered := True;

  ql:=ZQRecBai.FieldByName('quadralote').AsString;

  LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';

  Resp:='N';
  if simnao('Zerar Nosso Nº e Remessas Geradas até Agora ?','SIM') then
  begin
    ZQRecBai.close;
    ZQRecBai.SQL.Clear;
    ZQRecBai.SQL.Add('update recebimento set numboleto=''''  where saldo>''0'' and quadralote='+quotedstr(ql)+ ' and Marcar=''0''');
    ZQRecBai.ExecSQL;
    ZQRecBai.close;
    resp:='S';
  end;
  ZQRecBai.SQL.Clear;
  ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
  ZQRecBai.SQL.Add(' from recebimento where marcar=''0'' and  saldo>''0'' and quadralote='+quotedstr(ql));
//    ZQRecBai.SQL.Add('Select * from recebimento where   saldo>''0'' and quadralote='+quotedstr(ql));
  ZQRecBai.open;
  ZQRecBai.First;
  if Resp='S' then
  begin
    barravertical.MaxValue:=ZQRecBai.RecordCount;
    barravertical.Progress:=0;
    barravertical.Visible:=true;
    ZQrecBai.DisableControls;
    while not ZQrecBai.Eof do
    begin
      barravertical.Progress:=ZQRecBai.RecNo;
      ZQremessa_receb.Close;
      ZQremessa_receb.SQL.Clear;
      ZQremessa_receb.SQL.Add('delete from remessa_receb where idrec='+quotedstr(ZQRecBai.FieldByName('idrecebimento').Text));
      ZQremessa_receb.ExecSQL;

      ZQrecBai.Next;
    end;
    barravertical.Progress:=0;
    barravertical.Visible:=false;
    ZQremessa_receb.Close;
    ZQrecBai.EnableControls;
    ZQRecBai.First;
  end;
  Resp:='N';
  ql:='';
  dia:='';
  //ZQRecBai.Refresh;
  DBGBaixando.Refresh;
  ql:=elote.Text;
  elote.clear;

  DXBFechar.Enabled := true;
  zqrecbai.Filtered:=false;
  ZQRecBai.close;
  mensagem('Operação Terminada!');

end;

procedure TFrmzerarNossoNumero.filtrabaixa;
Begin
  if ( not empty(alltrim(Elote.Text))) then
  begin
    ZQRecBai.close;
    ZQRecBai.SQL.Clear;

    ZQRecBai.Close;
    ZQRecBai.SQL.Clear;
    ZQRecBai.SQL.Add('update recebimento set marcar=''0''');
    ZQRecBai.SQL.Add(' where (quadralote='+quotedstr(Elote.text)+') and (recpag=''R'') and ((saldo>0) and  (not saldo is null)))');
    ZQRecBai.ExecSQL;

    ZQRecBai.Close;
    ZQRecBai.SQL.Clear;
    ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
    ZQRecBai.SQL.Add(' from recebimento where (quadralote='+quotedstr(Elote.text)+') and (recpag=''R'') and ((saldo>0) and  (not saldo is null)))');
    ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
    ZQRecBai.Open;


    if ZQRecBai.active=true then
    begin
      LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
      ZQRecBai.Locate('idRecebimento',DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt,[]);

    end;
  end;

end;


procedure TFrmzerarNossoNumero.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if ZQRecBai.active=true then
     ZQRecBai.Close;

  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where quadralote='+quotedstr(ql)+ ' order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.open;
  DM_tabelas.ZQRecebimento.last;
  FrmzerarNossoNumero:=nil;
  Action:=CaFree;

end;

procedure TFrmzerarNossoNumero.ZQRecBaiCalcFields(DataSet: TDataSet);
begin
 if ZQRecBai.FieldByName('saldo').AsFloat=0 Then
    ZQRecBai.FieldByName('saldocorrig').AsFloat := 0
  else
    ZQRecBai.FieldByName('saldocorrig').AsFloat := ZQRecBai.FieldByName('saldo').AsFloat;
end;

procedure TFrmzerarNossoNumero.XNEJurosEnter(Sender: TObject);
begin
  if ZQRecBai.FieldByName('recpag').AsString = 'R' Then
    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrmzerarNossoNumero.DXBPesquisarClick(Sender: TObject);
begin
  FrmPesqRecebimento_bx.Top := self.Top+100;
  FrmPesqRecebimento_bx.Left := self.Left;
  FrmPesqRecebimento_bx.showmodal;

  DXBBaixaGravar.Enabled := false;


  filtrabaixa;
  DXBFechar.Enabled := True;

  Atualiza_tela;


  DBGBaixando.SetFocus;
end;

procedure TFrmzerarNossoNumero.btndesClick(Sender: TObject);
begin
  if empty(Elote.text) then
     Elote.text:=ZQRecBai.FieldByName('quadralote').AsString;

  ZQRecBai.Close;
  ZQRecBai.SQL.Clear;
  ZQRecBai.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end  ');
  ZQRecBai.SQL.Add(' where (quadralote='+quotedstr(Elote.text)+') and (recpag=''R'') and ((saldo>0) and  (not saldo is null)) ');
  ZQRecBai.ExecSQL;

  ZQRecBai.Close;
  ZQRecBai.SQL.Clear;
//  ZQRecBai.SQL.Add('select * from recebimento where (quadralote='+quotedstr(Elote.text)+') and ( tipdoc in ('+vartipo+')) and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (dt_vencimento between :dt5 and :dt6)');
  ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
  ZQRecBai.SQL.Add(' from recebimento where (quadralote='+quotedstr(Elote.text)+') and (recpag=''R'') and ((saldo>0) and  (not saldo is null)) ');

  ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
  ZQRecBai.Open;
  ZQRecBai.first;

  Atualiza_tela;
end;

procedure TFrmzerarNossoNumero.Filtere(Sender: TObject);
begin
 if not empty(Elote.Text) then
 begin
  ZQRecBai.SQL.Clear;
  ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
  ZQRecBai.SQL.Add(' from recebimento where  saldo>''0'' and quadralote='+quotedstr(Elote.Text));
  ZQRecBai.open;
  ZQRecBai.First;
  ZQRecBai.Filtered:=false;
  ql:=Elote.Text;
  ZQRecBai.RecordCount;
  Atualiza_tela;
 end;

end;

procedure TFrmzerarNossoNumero.DBGBaixandoExit(Sender: TObject);
begin
  Atualiza_tela;
end;

procedure TFrmzerarNossoNumero.Label9MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  //mover formulario sem bordas
  MovimentaObject(sender,button,shift,x,y,self);
end;


procedure TFrmzerarNossoNumero.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmzerarNossoNumero, 'ZQRecBai', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaisaldocorrig', 'saldocorrig', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaidocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaicliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaisaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaimarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBairefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaicontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaicustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBairecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBainumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBainomeadversa', 'nomeadversa', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaivenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBainumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBainomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaisomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaijuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaidescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaiobserv_estorno', 'observ_estorno', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmzerarNossoNumero, 'ZQRecBai', 'ZQRecBaitip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
