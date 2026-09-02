unit UPlanoDeContas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, Grids, DBGrids, StdCtrls, Mask, DBCtrls,   dxButton,
  XBanner, ExtCtrls, XDBEdit, qrprntr, Data.DB, dxCore2;


//  ,  , DB, , , , jpeg, XBanner,  XEdit;


type
  TFrmCad_PlanodeContas = class(TForm)
    PPainelBotoes: TPanel;
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
    Label5: TLabel;
    Label1: TLabel;
    DBECodInterno: TDBEdit;
    Label2: TLabel;
    Label3: TLabel;
    DBEDescricao: TDBEdit;
    Label4: TLabel;
    Label6: TLabel;
    DBGPlanoConta: TDBGrid;
    DBEcodContabil: TDBEdit;
    DBCBClassificacao: TDBComboBox;
    DBCBCredDeb: TDBComboBox;
    Label8: TLabel;
    DBCBAtivo: TDBComboBox;
    DBCBTIPDOC: TDBComboBox;
    DBCBDesp: TDBComboBox;
    Label10: TLabel;
    XBanner8: TXBanner;
    XBanner1: TXBanner;
    procedure Ativar_Campos;
    Procedure Desativar_campos;
    procedure DXBPrimeiroClick(Sender: TObject);
    procedure DXBAnteriorClick(Sender: TObject);
    procedure DXBProximoClick(Sender: TObject);
    procedure DXBUltimoClick(Sender: TObject);
    procedure DXBGravarClick(Sender: TObject);
    procedure DXBCancelarClick(Sender: TObject);
    procedure DXBIncluirClick(Sender: TObject);
    procedure DXBEditarClick(Sender: TObject);
    procedure DXBExcluirClick(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DBGPlanoContaTitleClick(Column: TColumn);
    procedure DXBRelatoriosClick(Sender: TObject);
    procedure DBGPlanoContaDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmCad_PlanodeContas: TFrmCad_PlanodeContas;

implementation

uses Tabelas, funcoes, Principal, RelPlanoDeContas, uRuntimeFields;

{$R *.dfm}

procedure TFrmCad_PlanodeContas.Ativar_Campos;
Begin
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
  DBEDescricao.ReadOnly := False;
  DBEcodContabil.ReadOnly := False;
  DBCBClassificacao.Enabled := True;
  DBCBCredDeb.Enabled := True;
  DBCBTIPDOC.Enabled := True;
  DBCBAtivo.Enabled := True;
  DBCBDesp.Enabled := True;
  DBGPlanoConta.Enabled := False;
end;

Procedure TFrmCad_PlanodeContas.Desativar_campos;
Begin
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
  DBEDescricao.ReadOnly := True;
  DBEcodContabil.ReadOnly := True;
  DBCBClassificacao.Enabled := False;
  DBCBCredDeb.Enabled := False;
  DBCBTIPDOC.Enabled := False;
  DBCBAtivo.Enabled := False;
  DBCBDesp.Enabled := False;
  DBGPlanoConta.Enabled := True;
End;

procedure TFrmCad_PlanodeContas.DXBPrimeiroClick(Sender: TObject);
begin
  DM_tabelas.ZQPlanoDeContas.First;
  DXBPrimeiro.Enabled := false;
  DXBAnterior.Enabled := false;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  DXBGravar.Enabled := False;
  DXBCancelar.Enabled := false;
end;

procedure TFrmCad_PlanodeContas.DXBAnteriorClick(Sender: TObject);
begin
  DM_tabelas.ZQPlanoDeContas.Prior;
  if DM_tabelas.ZQPlanoDeContas.RecNo = 1 then begin
    DXBPrimeiro.Enabled := false;
    DXBAnterior.Enabled := false;
  end;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
end;

procedure TFrmCad_PlanodeContas.DXBProximoClick(Sender: TObject);
begin
  DM_tabelas.ZQPlanoDeContas.Next;
  if DM_tabelas.ZQPlanoDeContas.RecNo = DM_tabelas.ZQPlanoDeContas.RecordCount then begin
    DXBProximo.Enabled := false;
    DXBUltimo.Enabled := false;
  end;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
end;

procedure TFrmCad_PlanodeContas.DXBUltimoClick(Sender: TObject);
begin
  DM_tabelas.ZQPlanoDeContas.Last;
  DXBProximo.Enabled := false;
  DXBUltimo.Enabled := false;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
end;

procedure TFrmCad_PlanodeContas.DXBGravarClick(Sender: TObject);
begin
  if not Verif_senha('Plano de Contas','Confirmação de Inclusão ou Edição  (Botão GRAVAR)','Cód. interno: '+DM_tabelas.ZQPlanoDeContas.FieldByName('codigo').Text+'  '+DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Text) then exit;
  DM_tabelas.ZQPlanoDeContas.Post;
  DM_tabelas.ZQPlanoDeContas.Refresh;
  Desativar_campos;
  DBGPlanoConta.setfocus;
end;

procedure TFrmCad_PlanodeContas.DXBCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Plano de Contas','Cancelamento de Inclusão ou Edição (Botão CANCELAR) ','Cód. interno: '+DM_tabelas.ZQPlanoDeContas.FieldByName('codigo').Text+'  '+DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Text) then exit;
  DM_tabelas.ZQPlanoDeContas.Cancel;
  Desativar_campos;
  DBGPlanoConta.setfocus;
end;

procedure TFrmCad_PlanodeContas.DXBIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Plano de Contas','Incluir Conta Contabil','') then exit;
  DM_tabelas.ZQPlanoDeContas.Insert;
  Ativar_Campos;
  DBEcodContabil.SetFocus;
end;

procedure TFrmCad_PlanodeContas.DXBEditarClick(Sender: TObject);
begin
  if not Verif_senha('Plano de Contas','Editar Conta Contábil','Cód. interno: '+DM_tabelas.ZQPlanoDeContas.FieldByName('codigo').Text+'  '+DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Text) then exit;
  DM_tabelas.ZQPlanoDeContas.Edit;
  Ativar_campos;
  DBEcodContabil.SetFocus;
end;

procedure TFrmCad_PlanodeContas.DXBExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Plano de Contas','Excluir Conta Contábil','Cód. interno: '+DM_tabelas.ZQPlanoDeContas.FieldByName('codigo').Text+'  '+DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Text) then exit;
  if DM_tabelas.ZQPlanoDeContas.RecordCount>0 then begin
    DM_tabelas.ZQPlanoDeContas.Delete;
  end;
end;

procedure TFrmCad_PlanodeContas.DXBFecharClick(Sender: TObject);
begin
  DM_tabelas.ZQPlanoDeContas.Close;
  Close;
end;

procedure TFrmCad_PlanodeContas.FormShow(Sender: TObject);
begin
  DBGPlanoConta.ReadOnly := True;
  FrmCad_PlanodeContas.Top := 26;
  DM_tabelas.ZQPlanoDeContas.Open;
  DBCBTipDoc.Clear;
  DM_tabelas.ZQTipodoc.Open;
  DM_tabelas.ZQTipodoc.First;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DBGPlanoConta.Columns[1].Color:=$00EDF0F1;
  Desativar_campos;
  DBGPlanoConta.setfocus;
end;

procedure TFrmCad_PlanodeContas.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_tabelas.ZQPlanoDeContas.Close;
  FrmCad_PlanodeContas:=nil;
  action:=CaFree;
end;

procedure TFrmCad_PlanodeContas.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmCad_PlanodeContas.DBGPlanoContaTitleClick(Column: TColumn);
var
  campo:string;
  varcodigo : integer;
begin
  campo:=column.fieldname;
  if (campo='codigo') or (campo='mascara') or (campo='descricao') Then Begin
    DBGPlanoConta.Columns[0].Color := clWindow;
    DBGPlanoConta.Columns[1].Color := clWindow;
    DBGPlanoConta.Columns[2].Color := clWindow;
    application.processmessages;
    varcodigo := DM_tabelas.ZQPlanoDeContas.FieldByName('codigo').AsLargeInt;
    DM_tabelas.ZQPlanoDeContas.sql.clear;
    DM_tabelas.ZQPlanoDeContas.SQL.Add('select * from plano_contas');
    DM_tabelas.ZQPlanoDeContas.SQL.Add(' order by '+campo);
    DM_tabelas.ZQPlanoDeContas.Open;
    DM_tabelas.ZQPlanoDeContas.Locate('codigo',varcodigo,[]);
    column.color:=$00EDF0F1;
  end;
end;

procedure TFrmCad_PlanodeContas.DXBRelatoriosClick(Sender: TObject);
begin
  FrmCad_PlanodeContas.FormStyle:=fsNormal;
  if FrmRelPlanoDeContas=nil then
     FrmRelPlanoDeContas:=TFrmRelPlanoDeContas.Create(Application);
  FrmRelPlanoDeContas.RLReport1.Previewmodal;
  FrmRelPlanoDeContas:=nil;
  FrmCad_PlanodeContas.FormStyle:=fsStayOnTop;
end;

procedure TFrmCad_PlanodeContas.DBGPlanoContaDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECodInterno.Text = DM_Tabelas.ZQPlanoDeContas.FieldByName('codigo').Text Then Begin
    DBGPlanoConta.Canvas.Brush.Color := $006CFFFF;
    DBGPlanoConta.Canvas.Font.Color := $00A80000;
    DBGPlanoConta.Canvas.Font.Style := [FsBold];
  end;
  DBGPlanoConta.DefaultDrawDataCell(Rect, DBGPlanoConta.columns[datacol].field, State);

end;


procedure TFrmCad_PlanodeContas.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
