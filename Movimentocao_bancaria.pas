unit Movimentocao_bancaria;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, DBCtrls, Mask,  
  dxButton, ExtCtrls, Buttons, XDBNum, TFlatPanelUnit, DB, JvExControls,
  JvButton, JvNavigationPane, TFlatGaugeUnit, ImgList, wwdbdatetimepicker,
  ComCtrls;
type
  TFrm_MovimentoBancario = class(TForm)
    XBanner1: TXBanner;
    Label5: TLabel;
    DBGMovBancaria: TDBGrid;
    Label1: TLabel;
    DBELancamento: TDBEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    DBEHora: TDBEdit;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    DBEDocumento: TDBEdit;
    Label10: TLabel;
    DBMObs: TDBMemo;
    PPainelBotoes: TPanel;
    XBanner2: TXBanner;
    DXBExcluir: TdxButton;
    DXBEditar: TdxButton;
    DXBIncluir: TdxButton;
    DXBCancelar: TdxButton;
    DXBGravar: TdxButton;
    DXBPesquisar: TdxButton;
    DXBRelatorios: TdxButton;
    DXBFechar: TdxButton;
    DXBUltimo: TdxButton;
    DXBProximo: TdxButton;
    DXBAnterior: TdxButton;
    DXBPrimeiro: TdxButton;
    L_DescricaoBanco: TLabel;
    XNEValor: TXDBNumEdit;
    Label11: TLabel;
    CXBDefiniContaBancaria: TdxButton;
    FPAchaBanco: TFlatPanel;
    DBGBanco: TDBGrid;
    EConfirmaSenhaBanco: TEdit;
    XBanner3: TXBanner;
    Label2: TLabel;
    FGBarraSaldo: TFlatGauge;
    IMenosMais: TImage;
    ImageList1: TImageList;
    DBEDataLanc: TwwDBDateTimePicker;
    DBEData_con: TwwDBDateTimePicker;
    LCorrentista: TLabel;
    EContabil: TEdit;
    Label12: TLabel;
    Enomes: TEdit;
    DBEVinculo: TDBEdit;
    Label13: TLabel;
    Dtl: TDateTimePicker;
    dtc: TDateTimePicker;
    procedure atualizar_tela;
    procedure Ativar_Campos;
    Procedure Desativar_campos;
    procedure FormShow(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure DXBPrimeiroClick(Sender: TObject);
    procedure DXBAnteriorClick(Sender: TObject);
    procedure DXBProximoClick(Sender: TObject);
    procedure DXBUltimoClick(Sender: TObject);
    procedure DXBIncluirClick(Sender: TObject);
    procedure DXBEditarClick(Sender: TObject);
    procedure DXBGravarClick(Sender: TObject);
    procedure DXBCancelarClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DBGBancoCellClick(Column: TColumn);
    procedure CXBDefiniContaBancariaClick(Sender: TObject);
    procedure DBGBancoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure EConfirmaSenhaBancoExit(Sender: TObject);
    procedure DBGBancoKeyPress(Sender: TObject; var Key: Char);
    procedure DBGMovBancariaTitleClick(Column: TColumn);
    procedure DBGMovBancariaDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure EConfirmaSenhaBancoEnter(Sender: TObject);
    procedure XBAchaPlanoDeContasClick(Sender: TObject);
    procedure DBGMovBancariaCellClick(Column: TColumn);
    procedure DBGMovBancariaKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DXBRelatoriosClick(Sender: TObject);
    procedure DBEHoraEnter(Sender: TObject);
    procedure DBEDocumentoExit(Sender: TObject);
    procedure DXBExcluirClick(Sender: TObject);
    procedure DBEDataLancKeyPress(Sender: TObject; var Key: Char);
    procedure DBELancamentoExit(Sender: TObject);
    procedure DBELancamentoEnter(Sender: TObject);
    procedure DBEDataLancExit(Sender: TObject);
    procedure DBEHoraExit(Sender: TObject);
    procedure DBEData_conExit(Sender: TObject);
    procedure XNEValorExit(Sender: TObject);
    procedure DBMObsExit(Sender: TObject);
    procedure DBMObsEnter(Sender: TObject);
    procedure DBEDocumentoEnter(Sender: TObject);
    procedure XNEValorEnter(Sender: TObject);
    procedure DBEData_conEnter(Sender: TObject);
    procedure DBEDataLancEnter(Sender: TObject);
    procedure EnomesExit(Sender: TObject);
    procedure RBClienteClick(Sender: TObject);
    procedure RBForneceClick(Sender: TObject);
    procedure EContabilExit(Sender: TObject);
    procedure EContabilEnter(Sender: TObject);
    procedure DXBPesquisarClick(Sender: TObject);
    procedure DBGMovBancariaEnter(Sender: TObject);
private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_MovimentoBancario: TFrm_MovimentoBancario;
  VarContaAtiva : integer;
implementation

uses Tabelas, funcoes, UachaPlanoDeConta, UFrmRelMovBanco,
  PesqMovBanco, uRuntimeFields;

{$R *.dfm}

procedure TFrm_MovimentoBancario.atualizar_tela;
Begin
  EContabil.Text := '';
  label11.Caption := '';
  if DM_Tabelas.ZQPlanodeContas.Locate('codigo',DM_Tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').Value,[]) then Begin
    EContabil.Text := DM_Tabelas.ZQPlanoDeContas.FieldByName('descricao').Value;
    label11.Caption := 'Conta: '+masccontabil(DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Value);
  end;
  IMenosMais.Visible := True;
  if DM_tabelas.ZQMovBancaria.FieldByName('credeb').Value = 'C' Then
    ImageList1.Draw(IMenosMais.Canvas, 1, 1, 1)
  else
    ImageList1.Draw(IMenosMais.Canvas, 1, 1, 0);
  Enomes.Text := '';
  DM_Tabelas.ZQCliente.Locate('idcliente',DM_Tabelas.ZQMovBancaria.FieldByName('contracodigo').Value,[]);
  Enomes.Text := DM_Tabelas.ZQClienterazao.Value;
  IMenosmais.Repaint;
End;

procedure TFrm_MovimentoBancario.Ativar_Campos;
Begin
 IMenosMais.Visible := True;
  CXBDefiniContaBancaria.Enabled := false;
  DBGMovBancaria.Enabled := False;
  EContabil.ReadOnly := false;
  DBEDataLanc.ReadOnly := false;
  DBEHora.ReadOnly := false;
  DBEData_con.ReadOnly := false;
  XNEValor.ReadOnly := false;
  DBEDocumento.ReadOnly := false;
  DBMObs.ReadOnly := false;
  DXBPrimeiro.Enabled := False;
  DXBAnterior.Enabled := False;
  DXBProximo.Enabled := False;
  DXBUltimo.Enabled := False;
  DXBIncluir.Enabled := False;
  DXBEditar.Enabled := False;
  DXBExcluir.Enabled := False;
  DXBPesquisar.Enabled := False;
  DXBRelatorios.Enabled := False;
  DXBGravar.Enabled := True;
  DXBCancelar.Enabled := True;
  DBGMovBancaria.enabled := false;
  Enomes.ReadOnly := False;

end;

Procedure TFrm_MovimentoBancario.Desativar_campos;
Begin
  IMenosMais.Visible := false;
  CXBDefiniContaBancaria.Enabled := True;
  EContabil.ReadOnly := True;
  Enomes.ReadOnly := True;
  DBGMovBancaria.Enabled := True;
  DBEDataLanc.ReadOnly := True;
  DBEHora.ReadOnly := True;
  DBEData_con.ReadOnly := True;
  XNEValor.ReadOnly := True;
  DBEDocumento.ReadOnly := True;
  DBMObs.ReadOnly := True;
  DXBPrimeiro.Enabled := True;
  DXBAnterior.Enabled := True;
  DXBProximo.Enabled := True;
  DXBUltimo.Enabled := True;
  DXBIncluir.Enabled := True;
  DXBEditar.Enabled := True;
  DXBExcluir.Enabled := True;
  DXBPesquisar.Enabled := True;
  DXBRelatorios.Enabled := True;
  DXBGravar.Enabled := False;
  DXBCancelar.Enabled := False;
  DBGMovBancaria.enabled := True;
End;

procedure TFrm_MovimentoBancario.FormShow(Sender: TObject);
begin
  L_DescricaoBanco.Caption := '';
  LCorrentista.Caption := '';
  label11.Caption := '';
  FGBarraSaldo.Visible := false;
  Desativar_campos;
  Frm_MovimentoBancario.Top := 38;
  DBELancamento.Enabled := False;
  DXBGravar.Enabled := False;
  DXBCancelar.Enabled := False;
  DM_tabelas.ZQConta_Bancaria.Open;
  DM_tabelas.ZQLimiteEspecial.Open;
  DM_tabelas.ZQMovBancaria.Open;
  DM_tabelas.ZQPlanoDeContas.Open;
  DM_Tabelas.ZQconfiguracao.Open;
  DM_Tabelas.ZQCliente.Open;
  XBanner1.ColorOf := DM_Tabelas.ZQconfiguracaocorup.Value;
  XBanner1.ColorFor := DM_TAbelas.ZQconfiguracaocordw.Value;
  XBanner2.ColorOf := DM_Tabelas.ZQconfiguracaocorup.Value;
  XBanner2.ColorFor := DM_TAbelas.ZQconfiguracaocordw.Value;
  XBanner3.ColorOf := DM_Tabelas.ZQconfiguracaocorup.Value;
  XBanner3.ColorFor := DM_TAbelas.ZQconfiguracaocordw.Value;
  Label5.Font.Color := DM_Tabelas.ZQconfiguracaocortopo.Value;
  Label2.Font.Color := DM_Tabelas.ZQconfiguracaocorup.Value;
  CXBDefiniContaBancaria.setFocus;
  DXBIncluir.Enabled := False;
  DXBEditar.Enabled := False;
  DXBExcluir.Enabled := False;
  DXBPesquisar.Enabled := False;
  DXBRelatorios.Enabled := False;
  DXBPrimeiro.Enabled := False;
  DXBAnterior.Enabled := False;
  DXBProximo.Enabled := False;
  DXBUltimo.Enabled := False;
  Label1.Visible := False;
  Label3.Visible := False;
  Label4.Visible := False;
  Label7.Visible := False;
  Label8.Visible := False;
  Label9.Visible := False;
  Label10.Visible := False;
  Label6.Visible := False;
  DBELancamento.Visible := False;
  EContabil.Visible := False;
  DBEDataLanc.Visible := False;
  DBEHora.Visible := False;
  DBEData_con.Visible := False;
  XNEValor.Visible := False;
  DBEVinculo.Visible := False;
  Label13.Visible := False;
  DBEDocumento.Visible := False;
  DBMObs.Visible := False;
  FPAchaBanco.Visible := false;
  DBGMovBancaria.visible := false;
  Label12.Visible := False;
  Enomes.Visible := false;
  DBGMovBancaria.Columns[5].Font.Style:=[FsBold];

end;

procedure TFrm_MovimentoBancario.DXBFecharClick(Sender: TObject);
begin
  DM_tabelas.ZQConta_Bancaria.Close;
  DM_tabelas.ZQLimiteEspecial.Close;
  DM_tabelas.ZQMovBancaria.Close;
  DM_tabelas.ZQPlanoDeContas.Close;
  Close;

end;

procedure TFrm_MovimentoBancario.DXBPrimeiroClick(Sender: TObject);
begin
  DM_tabelas.ZQMovBancaria.First;
  DXBPrimeiro.Enabled := false;
  DXBAnterior.Enabled := false;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  DXBGravar.Enabled := False;
  DXBCancelar.Enabled := false;
  atualizar_tela;
end;

procedure TFrm_MovimentoBancario.DXBAnteriorClick(Sender: TObject);
begin
  DM_tabelas.ZQMovBancaria.Prior;
  if DM_tabelas.ZQMovBancaria.RecNo = 1 then begin
    DXBPrimeiro.Enabled := false;
    DXBAnterior.Enabled := false;
  end;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  atualizar_tela;
end;

procedure TFrm_MovimentoBancario.DXBProximoClick(Sender: TObject);
begin
  DM_tabelas.ZQMovBancaria.Next;
  if DM_tabelas.ZQMovBancaria.RecNo = DM_tabelas.ZQMovBancaria.RecordCount then begin
    DXBProximo.Enabled := false;
    DXBUltimo.Enabled := false;
  end;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
  atualizar_tela;
end;

procedure TFrm_MovimentoBancario.DXBUltimoClick(Sender: TObject);
begin
  DM_tabelas.ZQMovBancaria.Last;
  DXBProximo.Enabled := false;
  DXBUltimo.Enabled := false;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
  atualizar_tela;
end;

procedure TFrm_MovimentoBancario.DXBIncluirClick(Sender: TObject);
begin
  if not senha('BANCO - Movimenta��o','Incluir Conta Banc�ria','') then exit;
  DM_tabelas.ZQMovBancaria.Insert;
  Ativar_Campos;
  CXBDefiniContaBancaria.Enabled := false;
  DM_tabelas.ZQMovBancaria.FieldByName('dt_lanc').Value := date;
  DM_tabelas.ZQMovBancaria.FieldByName('dt_conciliado').Value := date;
  DM_tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Value := DM_tabelas.ZQConta_Bancariacod_banco.Value;
  Label11.Caption := '';
  EContabil.setfocus;
end;

procedure TFrm_MovimentoBancario.DXBEditarClick(Sender: TObject);
begin
  if not senha('BANCO - Movimenta��o','Editar Conta Banc�ria','Conta: '+DM_tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Text+'  Lan�. em: '+DM_tabelas.ZQMovBancaria.FieldByName('dt_lanc').Text+'  Valor: '+DM_tabelas.ZQMovBancaria.FieldByName('vr_lanc').Text) then exit;
  DM_tabelas.ZQMovBancaria.Edit;
  Ativar_Campos;
  DM_tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Value := DM_tabelas.ZQConta_Bancariacod_banco.Value;
  Dtl.Date := DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Value;
  dtc.Date := DM_tabelas.ZQMovBancaria.FieldByName('dt_conciliado').Value;
  EContabil.setfocus;
end;

procedure TFrm_MovimentoBancario.DXBGravarClick(Sender: TObject);
Var
  Varrec : integer;
  varaltera : boolean;
begin
  if not senha('BANCO - Movimenta��o','Confirmar Inclus�o ou Edi��o (Bot�o GRAVAR)','Conta: '+DM_tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Text+'  Lan�. em: '+DM_tabelas.ZQMovBancaria.FieldByName('dt_lanc').Text+'  Valor: '+DM_tabelas.ZQMovBancaria.FieldByName('vr_lanc').Text) then exit;
  if DM_Tabelas.ZQMovBancaria.State in [DsEdit] Then
    varaltera := True
  else
    varaltera := False;
  DM_tabelas.ZQMovBancaria.Post;
  Desativar_campos;
  DBGMovBancaria.setfocus;
  varrec := DM_Tabelas.ZQMovBancaria.FieldByName('lancamento').Value;
  if varaltera Then Begin
    if Dtl.date < DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Value then
      CalcSaldo(datetostr(Dtl.date),'L',True)
    else
     CalcSaldo(DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Text,'L',True);
    if dtc.Date < DM_tabelas.ZQMovBancaria.FieldByName('dt_conciliado').Value then
      CalcSaldo(datetostr(dtc.date),'C',True)
    else
      CalcSaldo(DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Text,'C',True);
  end
  else
    CalcSaldo('','',True);
  DM_Tabelas.ZQMovBancaria.Locate('lancamento',varrec,[]);
  atualizar_tela;
end;

procedure TFrm_MovimentoBancario.DXBCancelarClick(Sender: TObject);
begin
  if not senha('BANCO - Movimenta��o','Cancelamento de Inclus�o ou Edi��o (Bot�o CANCELAR)','Conta: '+DM_tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Text+'  Lan�. em: '+DM_tabelas.ZQMovBancaria.FieldByName('dt_lanc').Text+'  Valor: '+DM_tabelas.ZQMovBancaria.FieldByName('vr_lanc').Text) then exit;
  DM_tabelas.ZQMovBancaria.Cancel;
  Desativar_campos;
  DBGMovBancaria.setfocus;
end;

procedure TFrm_MovimentoBancario.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_MovimentoBancario.DBGBancoCellClick(Column: TColumn);
begin
 EConfirmaSenhaBanco.SetFocus;
end;

procedure TFrm_MovimentoBancario.CXBDefiniContaBancariaClick(Sender: TObject);
begin
  FPAchaBanco.Visible := True;
  DBGBanco.Enabled := True;
  DBGBanco.setfocus;
end;

procedure TFrm_MovimentoBancario.DBGBancoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  EConfirmaSenhaBanco.SetFocus;
end;

procedure TFrm_MovimentoBancario.EConfirmaSenhaBancoExit(Sender: TObject);
begin
  DBGBanco.Enabled := false;
  if EConfirmaSenhaBanco.Text<>DM_tabelas.ZQConta_bancaria.FieldByName('senha_interna').Text then Begin
    showmessage('Senha incorreta....');
    CXBDefiniContaBancaria.SetFocus;
    FPAchaBanco.Visible := False;
    DM_tabelas.ZQConta_Bancaria.Locate('cod_banco',VarContaAtiva,[]);
    exit;
  end;
  DM_tabelas.ZQMovBancaria.SQL.Clear;
  DM_tabelas.ZQMovBancaria.SQL.Add('select * from Mov_Banco where Conta_Bancaria_cod_banco='+DM_tabelas.ZQConta_Bancariacod_banco.Text);
  DM_tabelas.ZQMovBancaria.SQL.Add(' order by dt_lanc,credeb');
  DM_tabelas.ZQMovBancaria.Open;
  DBGMovBancaria.Columns[0].Color := clWindow;
  DBGMovBancaria.Columns[1].Color := clWindow;
  DBGMovBancaria.Columns[2].Color := $00EDF0F1;
  DBGMovBancaria.Columns[3].Color := clWindow;
  DBGMovBancaria.Columns[4].Color := clWindow;
  DM_tabelas.ZQMovBancaria.last;
  DBGMovBancaria.Visible := True;
  FPAchaBanco.Visible := False;
  L_DescricaoBanco.Caption := 'C�d.: '+DM_tabelas.ZQConta_Bancariacod_banco.Text+'  Ag�ncia: '+DM_tabelas.ZQConta_BancariaAgencia.Value+'  Conta: '+DM_tabelas.ZQConta_Bancariaconta.Value+'  '+DM_tabelas.ZQConta_Bancarianome_banco.Value;
  LCorrentista.Caption := 'Correntista: '+DM_TAbelas.ZQConta_Bancarianome_titular.Value;
  VarContaAtiva := DM_tabelas.ZQConta_Bancariacod_banco.Value;
  DXBIncluir.Enabled := True;
  DXBEditar.Enabled := True;
  DXBExcluir.Enabled := True;
  DXBPesquisar.Enabled := True;
  DXBRelatorios.Enabled := True;
  DXBProximo.Enabled := True;
  DXBUltimo.Enabled := True;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
  Label1.Visible := True;
  Label3.Visible := True;
  Label4.Visible := True;
  Label7.Visible := True;
  Label8.Visible := True;
  Label9.Visible := True;
  Label10.Visible := True;
  Label6.Visible := True;
  DBELancamento.Visible := True;
  EContabil.Visible := True;
  DBEDataLanc.Visible := True;
  DBEHora.Visible := True;
  DBEData_con.Visible := True;
  XNEValor.Visible := True;
  DBEVinculo.Visible := True;
  Label13.Visible := True;
  DBEDocumento.Visible := True;
  DBMObs.Visible := True;
  Label12.Visible := True;
  Enomes.Visible := True;
  atualizar_tela;
  
  DBGMovBancaria.SetFocus;
end;

procedure TFrm_MovimentoBancario.DBGBancoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    EConfirmaSenhaBanco.SetFocus;
  if key = #27 Then Begin
    DM_tabelas.ZQConta_Bancaria.Locate('cod_banco',VarContaAtiva,[]);
    CXBDefiniContaBancaria.SetFocus;
    FPAchaBanco.Visible := False;
  end;
end;

procedure TFrm_MovimentoBancario.DBGMovBancariaTitleClick(Column: TColumn);
Var
  campo:string;
  varlancamento : integer;
begin
  campo:=column.fieldname;
  if (campo='lancamento') or (campo='dt_conciliado') or (campo='dt_lanc')  or (campo='documento') or (campo='vr_lanc') Then Begin
    DBGMovBancaria.Columns[0].Color := clWindow;
    DBGMovBancaria.Columns[1].Color := clWindow;
    DBGMovBancaria.Columns[2].Color := clWindow;
    DBGMovBancaria.Columns[3].Color := clWindow;
    DBGMovBancaria.Columns[4].Color := clWindow;
    DBGMovBancaria.Columns[5].Font.Style:=[];
    DBGMovBancaria.Columns[6].Font.Style:=[];
    application.processmessages;
    varlancamento := DM_tabelas.ZQMovBancaria.FieldByName('lancamento').Value;
    DM_tabelas.ZQMovBancaria.sql.clear;
    DM_tabelas.ZQMovBancaria.SQL.Add('select * from Mov_Banco where Conta_Bancaria_cod_banco='+DM_tabelas.ZQConta_Bancariacod_banco.Text);
    DM_tabelas.ZQMovBancaria.SQL.Add(' order by '+campo+', credeb');
    DM_tabelas.ZQMovBancaria.Open;
    DM_tabelas.ZQMovBancaria.Locate('lancamento',varlancamento,[]);
    column.color:=$00EDF0F1;
    if  campo='dt_lanc' Then
      DBGMovBancaria.Columns[5].Font.Style:=[FsBold];
    if  campo='dt_conciliado' Then
      DBGMovBancaria.Columns[6].Font.Style:=[FsBold];

  end;
end;

procedure TFrm_MovimentoBancario.DBGMovBancariaDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin

  if DBELancamento.Text = DM_tabelas.ZQMovBancaria.FieldByName('lancamento').Text Then
    DBGMovBancaria.Canvas.Brush.Color := DM_tabelas.ZQconfiguracaocorlinha.Value;


  If DM_tabelas.ZQMovBancaria.FieldByName('credeb').Value = 'C' then
    DBGMovBancaria.Canvas.Font.Color:= clBlue
  else
    DBGMovBancaria.Canvas.Font.Color:= clRed;
  if DBELancamento.Text = DM_Tabelas.ZQMovBancaria.FieldByName('lancamento').Text Then Begin
    DBGMovBancaria.Canvas.Brush.Color := DM_Tabelas.ZQConfiguracaocorlinha.Value;
    DBGMovBancaria.Canvas.Font.Color := DM_Tabelas.ZQConfiguracaocorFonte.Value;
  end;
  DBGMovBancaria.DefaultDrawDataCell(Rect, DBGMovBancaria.columns[datacol].field, State);
end;


procedure TFrm_MovimentoBancario.EConfirmaSenhaBancoEnter(Sender: TObject);
begin
EConfirmaSenhaBanco.Text := '';
IF EMPTY(DM_tabelas.ZQConta_bancaria.FieldByName('senha_interna').Text) Then
  DXBFechar.setfocus;
end;

procedure TFrm_MovimentoBancario.XBAchaPlanoDeContasClick(Sender: TObject);
begin
  FrmAchaPlanodeConta.showmodal;
  DM_tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').Value := DM_tabelas.ZQPlanoDeContas.FieldByName('codigo').Value;
  EContabil.setfocus;
end;

procedure TFrm_MovimentoBancario.DBGMovBancariaCellClick(Column: TColumn);
begin
  atualizar_tela;
end;

procedure TFrm_MovimentoBancario.DBGMovBancariaKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  atualizar_tela;
end;

procedure TFrm_MovimentoBancario.DXBRelatoriosClick(Sender: TObject);
begin
  if not senha('BANCO - Movimenta��o','Menu de Relat�rio','') then exit;
  FrmRelMovBanco.showmodal;
end;

procedure TFrm_MovimentoBancario.DBEHoraEnter(Sender: TObject);
begin
  IF DM_tabelas.ZQMovBancaria.State in [dsEdit, dsInsert] then
    DM_tabelas.ZQMovBancaria.FieldByName('hora_lanc').Value := Time;
  DBEHora.Color:= $00FFF0E1
end;

procedure TFrm_MovimentoBancario.DBEDocumentoExit(Sender: TObject);
begin
  if empty(DM_tabelas.ZQMovBancaria.FieldByName('documento').Value) then
    DBEDocumento.SetFocus;
  DBEDocumento.Color:= clWindow;
end;

procedure TFrm_MovimentoBancario.DXBExcluirClick(Sender: TObject);
Var
  Varrec : integer;
begin
  if not senha('BANCO - Movimenta��o','Exclus�o de um registro da movimenta��o','Conta: '+DM_tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Text+'  Lan�. em: '+DM_tabelas.ZQMovBancaria.FieldByName('dt_lanc').Text+'  Valor: '+DM_tabelas.ZQMovBancaria.FieldByName('vr_lanc').Text) then exit;
  if DM_Tabelas.ZQMovBancaria.RecordCount >0 Then Begin
    varrec := DM_Tabelas.ZQMovBancaria.RecNo;
    DM_tabelas.ZQMovBancaria.Delete;
    CalcSaldo('','',True);
    DM_Tabelas.ZQMovBancaria.RecNo :=varrec;
  end;
end;

procedure TFrm_MovimentoBancario.DBEDataLancKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then begin
    try

    StrToDate(DBEDataLanc.Text);
    except
    on EConvertError do
    ShowMessage ('Data Inv�lida!');
    end;
  end;
end;

procedure TFrm_MovimentoBancario.DBELancamentoExit(Sender: TObject);
begin
DBELancamento.Color:= clWindow;
end;

procedure TFrm_MovimentoBancario.DBELancamentoEnter(Sender: TObject);
begin
DBELancamento.Color:=$00FFF0E1;
end;

procedure TFrm_MovimentoBancario.DBEDataLancExit(Sender: TObject);
begin
DBEDataLanc.Color:= clWindow;
end;

procedure TFrm_MovimentoBancario.DBEHoraExit(Sender: TObject);
begin
DBEHora.Color:= clWindow;
end;

procedure TFrm_MovimentoBancario.DBEData_conExit(Sender: TObject);
begin
DBEData_con.Color:= clWindow;
end;

procedure TFrm_MovimentoBancario.XNEValorExit(Sender: TObject);
begin
XNEValor.Color:= clWindow;
end;

procedure TFrm_MovimentoBancario.DBMObsExit(Sender: TObject);
begin
DBMObs.Color:= clWindow;
end;

procedure TFrm_MovimentoBancario.DBMObsEnter(Sender: TObject);
begin
DBMObs.Color:= $00FFF0E1;
end;

procedure TFrm_MovimentoBancario.DBEDocumentoEnter(Sender: TObject);
begin
DBEDocumento.Color:= $00FFF0E1;
end;

procedure TFrm_MovimentoBancario.XNEValorEnter(Sender: TObject);
begin
XNEValor.Color:= $00FFF0E1;

end;

procedure TFrm_MovimentoBancario.DBEData_conEnter(Sender: TObject);
begin
DBEData_con.Color:= $00FFF0E1;
end;

procedure TFrm_MovimentoBancario.DBEDataLancEnter(Sender: TObject);
begin
DBEDataLanc.Color:= $00FFF0E1;
end;

procedure TFrm_MovimentoBancario.EnomesExit(Sender: TObject);
begin
  If (not ((DXBCancelar.Focused) or (DXBFechar.Focused))) and (DM_Tabelas.ZQMovBancaria.State in [DsInsert, DsEdit]) Then Begin
    if not AchaCLiente(450, 140,Enomes.Text) Then Begin
      Enomes.SetFocus;
      Exit;
    End;
    DM_Tabelas.ZQMovBancaria.FieldByName('contracodigo').Value := DM_Tabelas.ZQClienteidcliente.Value;
    Enomes.Text := DM_Tabelas.ZQClienterazao.Value;
  End;
end;

procedure TFrm_MovimentoBancario.RBClienteClick(Sender: TObject);
begin
  if DM_TAbelas.ZQMovBancaria.State in [DsInsert, DsEdit] Then
//    DM_Tabelas.ZQMovBancariacontraquem.Value := 'C';
end;

procedure TFrm_MovimentoBancario.RBForneceClick(Sender: TObject);
begin
  if DM_TAbelas.ZQMovBancaria.State in [DsInsert, DsEdit] Then
//    DM_Tabelas.ZQMovBancariacontraquem.Value := 'F';
end;

procedure TFrm_MovimentoBancario.EContabilExit(Sender: TObject);
begin
  if not DXBCancelar.Focused Then Begin
    if not Acha_CodContabil(EContabil.Text,'S',DM_Tabelas.ZQConta_Bancariaplanodeconta.Text,'') then Begin
      EContabil.SetFocus;
      exit;
    end;
    EContabil.Text := FrmAchaPlanodeConta.ZQTempPlanoContasdescricao.Value;
    DM_Tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').Value := FrmAchaPlanodeConta.ZQTempPlanoContascodigo.Value;
    DM_Tabelas.ZQMovBancaria.FieldByName('credeb').Value := FrmAchaPlanodeConta.ZQTempPlanoContascred_debe.Value;
    if DM_tabelas.ZQMovBancaria.FieldByName('credeb').Value = 'C' Then
      ImageList1.Draw(IMenosMais.Canvas, 1, 1, 1)
    else
      ImageList1.Draw(IMenosMais.Canvas, 1, 1, 0);
    IMenosmais.Repaint;
  end;
end;

procedure TFrm_MovimentoBancario.EContabilEnter(Sender: TObject);
begin
  EContabil.color :=$00FFF0E1;
end;

procedure TFrm_MovimentoBancario.DXBPesquisarClick(Sender: TObject);
begin
  FrmPesqMovBanco.showmodal;
  DBGMovBancaria.SetFocus;
end;

procedure TFrm_MovimentoBancario.DBGMovBancariaEnter(Sender: TObject);
begin
  atualizar_tela;
end;


procedure TFrm_MovimentoBancario.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
