unit UFrmRelMovBanco;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,   dxButton, Mask, XDate, ExtCtrls, XBanner, Grids, DBGrids,
  DB, ZAbstractRODataset, ZDataset, ZAbstractDataset, JvExControls,
  JvWaitingProgress, JvSpecialProgress, TFlatGaugeUnit,
  TFlatProgressBarUnit, ComCtrls, DBClient;

type
  TFrmRelMovBanco = class(TForm)
    XBanner1: TXBanner;
    RGTipo: TRadioGroup;
    RGOrdem: TRadioGroup;
    DXBRelatorios: TdxButton;
    DXBFechar: TdxButton;
    CBMostrasaldo: TCheckBox;
    CBMostraObs: TCheckBox;
    Label3: TLabel;
    CBZebrado: TCheckBox;
    CBSomalimite: TCheckBox;
    DS_TempPlano: TDataSource;
    ZQtempPlano: TZQuery;






    CBTodasContas: TCheckBox;
    JSP1: TJvSpecialProgress;
    ZQTempMovBan: TZQuery;
    DS_tempmovban: TDataSource;


    CBMostLanc: TCheckBox;
    DS_Templanc: TDataSource;
    ZQTempLanc: TZQuery;




    GBData: TGroupBox;
    XBanner2: TXBanner;
    XDEFinal: TDateTimePicker;
    XDEInicial: TDateTimePicker;
    Label2: TLabel;
    Label1: TLabel;
    GBPlano: TGroupBox;
    Eplano: TEdit;
    DBGrid1: TDBGrid;
    DS_Plano: TDataSource;
    CDS_Plano: TClientDataSet;
    CDS_Planocodplanno: TIntegerField;
    CDS_Planodesclpano: TStringField;
    XBanner3: TXBanner;






































    procedure DXBRelatoriosClick(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure XDEInicialExit(Sender: TObject);
    procedure CBMostraObsClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CBTodasContasClick(Sender: TObject);
    procedure RGTipoExit(Sender: TObject);
    procedure XDEInicialEnter(Sender: TObject);
    procedure XDEFinalEnter(Sender: TObject);
    procedure XDEFinalExit(Sender: TObject);
    procedure ZQTempMovBanCalcFields(DataSet: TDataSet);
    procedure EplanoExit(Sender: TObject);
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelMovBanco: TFrmRelMovBanco;

implementation

uses Tabelas, funcoes, UFrmRelMovBanco01, UFrmRelMovBanco02, Jornal,
  uFrmRelMovBanco03, UachaPlanoDeConta, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelMovBanco.DXBRelatoriosClick(Sender: TObject);
var
  vardata,varmasc,varcrede, varplano : string;
  varsoma : double;
  varregi : integer;
begin
  if XDEInicial.Date <= 0 then begin
    mensagem('a data não pode ficar em branco....');
    XDEInicial.SetFocus;
    exit
  end;
  varplano:='';
  CDS_Plano.First;
  while not CDS_Plano.Eof do begin
    varplano := varplano + quotedstr(CDS_Planocodplanno.Text)+',';
    CDS_Plano.Next;
  end;
  varplano := copy(varplano,1,length(varplano)-1);
  if CBMostLanc.Checked then
    CBZebrado.Checked := False;
  vardata := 'dt_conciliado';
  if RGOrdem.ItemIndex = 0 then
    vardata := 'dt_lanc';
  if RGTipo.ItemIndex=0 then Begin
    ZQTempMovBan.SQL.Clear;
    ZQTempMovBan.SQL.Add('select * from Mov_Banco as M join plano_contas as p ON p.codigo=m.plano_contas_codigo');
    ZQTempMovBan.SQL.Add(' where ('+vardata+' between :dt1 and :dt2)');
    if not empty(varplano) Then
      ZQTempMovBan.SQL.Add(' and Plano_Contas_codigo in ('+varplano+')');
    if not CBTodasContas.Checked then
      ZQTempMovBan.SQL.Add(' and (Conta_Bancaria_cod_banco='+DM_tabelas.ZQConta_Bancariacod_banco.Text+') ');
    ZQTempMovBan.ParamByName('dt1').AsDate:=XDEInicial.Date;
    ZQTempMovBan.ParamByName('dt2').AsDate:=XDEFinal.Date;
    ZQTempMovBan.SQL.Add(' order by '+vardata+',credeb');
    ZQTempMovBan.Open;
    if Frmrelmovbanco03=nil then
      Frmrelmovbanco03:=tFrmrelmovbanco03.Create(Application);
//    Frmrelmovbanco01.QRmovbancario01.OnPreview:=FrmPrincipal.Prev;
    Frmrelmovbanco03.RLReport1.Preview;
  end
  else
  Begin
    FrmRelMovBanco02.QRSubDetail1.Enabled := False;
    ZQTempLanc.SQL.Clear;
    ZQTempLanc.SQL.Add('select * from mov_banco');
    ZQTempLanc.SQL.Add(' where ('+vardata+' between :dt1 and :dt2) ');
    if not empty(varplano) Then
      ZQTempLanc.SQL.Add(' and Plano_Contas_codigo in ('+varplano+')');
    ZQTempLanc.ParamByName('dt1').AsDate:=XDEInicial.Date;
    ZQTempLanc.ParamByName('dt2').AsDate:=XDEFinal.Date;
    ZQTempLanc.Open;
    ZQtempPlano.close;
    ZQtempPlano.SQL.Clear;
    ZQtempPlano.SQL.Add('UPDATE plano_contas SET valor = 0');
    ZQtempPlano.ExecSQL;
    ZQtempPlano.SQL.Clear;
    ZQtempPlano.SQL.Add('SELECT distinct c.codigo, c.mascara, c.descricao, c.valor, c.cred_debe, c.classificacao FROM plano_contas p , mov_banco m , plano_contas c');
    ZQtempPlano.SQL.Add(' where p.codigo=plano_contas_codigo and c.mascara=left(p.mascara,Length(c.mascara))');
    if not empty(varplano) Then
      ZQtempPlano.SQL.Add(' and Plano_Contas_codigo in ('+varplano+')');
    if not CBTodasContas.Checked then
      ZQtempPlano.SQL.Add(' and m.conta_bancaria_cod_banco='+DM_tabelas.ZQConta_Bancariacod_banco.Text);
    ZQtempPlano.SQL.Add(' and ('+vardata+' between :dt1 and :dt2) group by c.mascara');
    ZQtempPlano.ParamByName('dt1').AsDate:=XDEInicial.Date;
    ZQtempPlano.ParamByName('dt2').AsDate:=XDEFinal.Date;
    ZQtempPlano.open;
    ZQtempPlano.Filter := 'classificacao ='+QuotedStr('S');
    ZQtempPlano.Filtered := True;
    ZQtempPlano.First;
    JSP1.Visible := True;
    JSP1.Maximum:=ZQtempPlano.RecordCount;
    while not ZQtempPlano.Eof do begin
      JSP1.Position :=ZQtempPlano.RecNo;
      if ZQtempPlano.FieldByName('classificacao').Value = 'S' then Begin
        DM_tabelas.ZQMovBancaria.SQL.Clear;
        DM_tabelas.ZQMovBancaria.SQL.Add('select * from Mov_Banco where ('+vardata+' between :dt1 and :dt2)');
        if not empty(varplano) Then
          DM_tabelas.ZQMovBancaria.SQL.Add(' and Plano_Contas_codigo in ('+varplano+')');
        if not CBTodasContas.Checked then
          DM_tabelas.ZQMovBancaria.SQL.Add(' and (Conta_Bancaria_cod_banco='+DM_tabelas.ZQConta_Bancariacod_banco.Text+') ');
        DM_tabelas.ZQMovBancaria.SQL.Add(' and Plano_Contas_codigo='+ZQtempPlano.FieldByName('codigo').Text);
        DM_tabelas.ZQMovBancaria.ParamByName('dt1').AsDate:=XDEInicial.Date;
        DM_tabelas.ZQMovBancaria.ParamByName('dt2').AsDate:=XDEFinal.Date;
        DM_tabelas.ZQMovBancaria.Open;
        varsoma := 0;
        while not DM_tabelas.ZQMovBancaria.Eof do begin
          varsoma := varsoma+DM_tabelas.ZQMovBancaria.FieldByName('vr_lanc').Value;;
          DM_tabelas.ZQMovBancaria.Next;
        end;
        ZQtempPlano.Edit;
        ZQtempPlano.FieldByName('valor').Value := varsoma;
        ZQtempPlano.Post;
      end;
      ZQtempPlano.next;
    end;
    ZQtempPlano.Filter := '';
    ZQtempPlano.Filtered := False;
    ZQtempPlano.First;
    JSP1.Maximum:=ZQtempPlano.RecordCount;
    while not ZQtempPlano.Eof do begin
      JSP1.Position :=ZQtempPlano.RecNo;
      if ZQtempPlano.FieldByName('classificacao').Value = 'S' then Begin
        varregi :=ZQtempPlano.RecNo;
        varsoma:= ZQtempPlano.FieldByName('valor').Value;
        varmasc:= ZQtempPlano.FieldByName('mascara').Value;
        varcrede:=ZQtempPlano.FieldByName('cred_debe').Value;
        if Length(copy(varmasc,1,10))=10 then begin
          if (ZQtempplano.Locate('mascara',copy(varmasc,1,10),[])) and (ZQtempPlano.RecNo<>varregi) then Begin
            ZQtempPlano.Edit;
            if varcrede='C' then
              ZQtempPlano.FieldByName('valor').Value := ZQtempPlano.FieldByName('valor').Value+varsoma
            else
              ZQtempPlano.FieldByName('valor').Value := ZQtempPlano.FieldByName('valor').Value-varsoma;
            ZQtempPlano.Post;
//            varcrede:=ZQtempplanocred_debe.Value;
          end;
        end;
        if Length(copy(varmasc,1,6))=6 then begin
          if (ZQtempplano.Locate('mascara',copy(varmasc,1,6),[])) and (ZQtempPlano.RecNo<>varregi) then begin
            ZQtempPlano.Edit;
            if varcrede='C' then
              ZQtempPlano.FieldByName('valor').Value := ZQtempPlano.FieldByName('valor').Value+varsoma
            else
              ZQtempPlano.FieldByName('valor').Value := ZQtempPlano.FieldByName('valor').Value-varsoma;
            ZQtempPlano.Post;
//            varcrede:=ZQtempplanocred_debe.Value;
          end;
        end;
         if Length(copy(varmasc,1,3))=3 then begin
          if (ZQtempplano.Locate('mascara',copy(varmasc,1,3),[])) and (ZQtempPlano.RecNo<>varregi) then begin
            ZQtempPlano.Edit;
            if varcrede='C' then
              ZQtempPlano.FieldByName('valor').Value := ZQtempPlano.FieldByName('valor').Value+varsoma
            else
              ZQtempPlano.FieldByName('valor').Value := ZQtempPlano.FieldByName('valor').Value-varsoma;
            ZQtempPlano.Post;
//            varcrede:=ZQtempplanocred_debe.Value;
          end;
        end;
        if Length(copy(varmasc,1,1))=1 then begin
          if (ZQtempplano.Locate('mascara',copy(varmasc,1,1),[])) and (ZQtempPlano.RecNo<>varregi) then begin
            ZQtempPlano.Edit;
            if varcrede='C' then
              ZQtempPlano.FieldByName('valor').Value := ZQtempPlano.FieldByName('valor').Value+varsoma
            else
              ZQtempPlano.FieldByName('valor').Value := ZQtempPlano.FieldByName('valor').Value-varsoma;
            ZQtempPlano.Post;
          end;
        end;
        ZQtempPlano.RecNo:=varregi;
      end;
      ZQtempPlano.next;
    end;
    if Frmrelmovbanco02=nil then
      Frmrelmovbanco02:=tFrmrelmovbanco02.Create(Application);
    Frmrelmovbanco02.QRmovbancario02.OnPreview:=FrmPrincipal.Prev;
    Frmrelmovbanco02.QRmovbancario02.Preview;
  end;
  JSP1.Visible := False;
  DM_tabelas.ZQMovBancaria.SQL.Clear;
  DM_tabelas.ZQMovBancaria.SQL.Add('select * from Mov_Banco where Conta_Bancaria_cod_banco='+DM_tabelas.ZQConta_Bancariacod_banco.Text);
  DM_tabelas.ZQMovBancaria.SQL.Add(' order by '+vardata+',credeb');
  DM_tabelas.ZQMovBancaria.Open;
end;

procedure TFrmRelMovBanco.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmRelMovBanco.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmRelMovBanco.XDEInicialExit(Sender: TObject);
begin
  if XDEFinal.Date <= 0 then
    XDEFinal.Date := XDEInicial.Date;
  XDEInicial.Color:= clWindow;
end;

procedure TFrmRelMovBanco.CBMostraObsClick(Sender: TObject);
begin
{  if CBMostraObs.Checked then begin
    CBZebrado.Enabled := False;
    CBZebrado.Checked := False;
  end
  else
    CBZebrado.Enabled := True;}
end;

procedure TFrmRelMovBanco.FormShow(Sender: TObject);
begin
  XDEInicial.Date := DATE;
  XDEFinal.Date := date;
  JSP1.Visible := False;
  CBTodasContas.Checked := False;
  RGTipo.SetFocus;
  Label3.Font.Color :=DM_tabelas.ZQConfiguracaocorTopo.Value;
  XBanner1.ColorOf:=DM_tabelas.ZQconfiguracaocorup.Value;
  XBanner1.ColorFor:=DM_Tabelas.ZQconfiguracaocordw.Value;
  XBanner2.ColorOf:=DM_tabelas.ZQconfiguracaocorup.Value;
  XBanner2.ColorFor:=DM_Tabelas.ZQconfiguracaocordw.Value;
  XBanner3.ColorOf:=DM_tabelas.ZQconfiguracaocorup.Value;
  XBanner3.ColorFor:=DM_Tabelas.ZQconfiguracaocordw.Value;
  RGOrdem.Color :=DM_Tabelas.ZQconfiguracaocormEIO.Value;
  RGTipo.Color := DM_Tabelas.ZQconfiguracaocorMEIO.Value;
  CBMostrasaldo.Color := DM_Tabelas.ZQconfiguracaocorMEIO.Value;
  CBSomalimite.Color := DM_Tabelas.ZQconfiguracaocorMEIO.Value;
  CBMostraObs.Color := DM_Tabelas.ZQconfiguracaocorMEIO.Value;
  CBZebrado.Color := DM_Tabelas.ZQconfiguracaocorMEIO.Value;
  CBTodasContas.Color := DM_Tabelas.ZQconfiguracaocorMEIO.Value;
  CBMostLanc.Color := DM_Tabelas.ZQconfiguracaocorMEIO.Value;
  GBData.Color:=DM_tabelas.ZQconfiguracaocorup.Value;
  GBPlano.Color:=DM_tabelas.ZQconfiguracaocorup.Value;
  CDS_Plano.Close;
  CDS_Plano.CreateDataSet;
end;

procedure TFrmRelMovBanco.CBTodasContasClick(Sender: TObject);
begin
if CBTodasContas.Checked then
  if not senha('BANCO - Movimentação','Agrupar todas as contas no relatório de extrato por Plano de Contas','') then
    CBTodasContas.Checked := False;

end;

procedure TFrmRelMovBanco.RGTipoExit(Sender: TObject);
begin
  CBTodasContas.Enabled := False;
  CBMostLanc.Enabled := False;
  if RGTipo.ItemIndex = 1 then Begin
    CBTodasContas.Enabled := True;
    CBMostLanc.Enabled := True;
    GBPlano.Enabled := True;
  end
  else Begin
    CBMostLanc.Checked := False;
    GBPlano.Enabled := False;
    CDS_Plano.Close;
    CDS_Plano.CreatedataSet;
  end;
end;

procedure TFrmRelMovBanco.XDEInicialEnter(Sender: TObject);
begin
XDEInicial.Color:= $00FFF0E1
end;

procedure TFrmRelMovBanco.XDEFinalEnter(Sender: TObject);
begin
XDEFinal.Color:= $00FFF0E1
end;

procedure TFrmRelMovBanco.XDEFinalExit(Sender: TObject);
begin
XDEFinal.Color:= clWindow;
end;

procedure TFrmRelMovBanco.ZQTempMovBanCalcFields(DataSet: TDataSet);
begin
  DM_tabelas.ZQLimiteEspecial.SQL.Clear;
  DM_Tabelas.ZQLimiteEspecial.SQL.Add('select * from limite_especial where data_valida<=:datL');
  DM_Tabelas.ZQLimiteEspecial.ParamByName('datL').AsDate := ZQTempMovBan.FieldByName('dt_lanc').Value;
  DM_tabelas.ZQLimiteEspecial.Open;
  if DM_tabelas.ZQLimiteEspecial.RecordCount>0 Then Begin
    DM_tabelas.ZQLimiteEspecial.Last;
    ZQTempMovBan.FieldByName('Saldo_lanc_plus').Value := ZQTempMovBan.FieldByName('Saldo_lanc').Value + DM_Tabelas.ZQLimiteEspecialvalor.Value;
  end;
  DM_tabelas.ZQLimiteEspecial.SQL.Clear;
  DM_Tabelas.ZQLimiteEspecial.SQL.Add('select * from limite_especial where data_valida<=:datL');
  DM_Tabelas.ZQLimiteEspecial.ParamByName('datL').AsDate := ZQTempMovBan.FieldByName('dt_conciliado').Value;
  DM_tabelas.ZQLimiteEspecial.Open;
  if DM_tabelas.ZQLimiteEspecial.RecordCount>0 Then Begin
    DM_tabelas.ZQLimiteEspecial.Last;
    ZQTempMovBan.FieldByName('Saldo_conci_plus').Value := ZQTempMovBan.FieldByName('Saldo_conci').Value + DM_Tabelas.ZQLimiteEspecialvalor.Value;
  end;
end;

procedure TFrmRelMovBanco.EplanoExit(Sender: TObject);
begin
  If (not ((DXBRelatorios.Focused) or (DXBFechar.Focused))) and (not empty(Eplano.Text)) Then Begin
    if Acha_CodContabil(Eplano.Text,'','','') then Begin
      Eplano.Text := '';
      CDS_Plano.Insert;
      CDS_Planocodplanno.Value := FrmAchaPlanodeConta.ZQTempPlanoContascodigo.Value;
      CDS_Plano.Post;
      Eplano.SetFocus;
    end;
  End;
end;

procedure TFrmRelMovBanco.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDS_Plano.RecordCount>0) Then Begin
    CDS_Plano.Delete;
  end;
end;


procedure TFrmRelMovBanco.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmRelMovBanco, 'ZQtempPlano', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQtempPlano', 'ZQtempPlanocodigo', 'codigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQtempPlano', 'ZQtempPlanomascara', 'mascara', TStringField, fkData, 19, 0, False, '', '', '', '0\.00\.000\.0000\.00000;0;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQtempPlano', 'ZQtempPlanodescricao', 'descricao', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQtempPlano', 'ZQtempPlanovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQtempPlano', 'ZQtempPlanocred_debe', 'cred_debe', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQtempPlano', 'ZQtempPlanoclassificacao', 'classificacao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelMovBanco, 'ZQTempMovBan', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanSaldo_lanc_plus', 'Saldo_lanc_plus', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanSaldo_conci_plus', 'Saldo_conci_plus', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanlancamento', 'lancamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanConta_Bancaria_cod_banco', 'Conta_Bancaria_cod_banco', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanHistorico_codigo', 'Historico_codigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanPlano_Contas_codigo', 'Plano_Contas_codigo', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBandocumento', 'documento', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBandt_lanc', 'dt_lanc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanhora_lanc', 'hora_lanc', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBandt_conciliado', 'dt_conciliado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanvr_lanc', 'vr_lanc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanCreDeb', 'CreDeb', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanSaldo_lanc', 'Saldo_lanc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanSaldo_conci', 'Saldo_conci', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanobs', 'obs', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanvinculo', 'vinculo', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBancontracodigo', 'contracodigo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBancodigo', 'codigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanmascara', 'mascara', TStringField, fkData, 16, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBandescricao', 'descricao', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanclassificacao', 'classificacao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBancred_debe', 'cred_debe', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBancod_reduzido', 'cod_reduzido', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBandoccomum', 'doccomum', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanusoativo', 'usoativo', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanCDPlano', 'CDPlano', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempMovBan', 'ZQTempMovBanDespFixVar', 'DespFixVar', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelMovBanco, 'ZQTempLanc', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLanclancamento', 'lancamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancConta_Bancaria_cod_banco', 'Conta_Bancaria_cod_banco', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancvinculo', 'vinculo', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancparticipante', 'participante', TStringField, fkLookup, 40, 0, False, '', '', '', '', 0, 'contracodigo', '', 'idcliente', 'razao', True);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancHistorico_codigo', 'Historico_codigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancPlano_Contas_codigo', 'Plano_Contas_codigo', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancdocumento', 'documento', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancdt_lanc', 'dt_lanc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLanchora_lanc', 'hora_lanc', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancdt_conciliado', 'dt_conciliado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancvr_lanc', 'vr_lanc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancCreDeb', 'CreDeb', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancSaldo_lanc', 'Saldo_lanc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancSaldo_conci', 'Saldo_conci', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLancobs', 'obs', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelMovBanco, 'ZQTempLanc', 'ZQTempLanccontracodigo', 'contracodigo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
