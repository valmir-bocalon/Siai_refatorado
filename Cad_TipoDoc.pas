unit Cad_TipoDoc;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, ExtCtrls, StdCtrls, XBanner,
  Mask, DBCtrls, Data.DB;

type
  TFrmCad_TipoDoc = class(TForm)
    Label5: TLabel;
    LReg: TLabel;
    Label11: TLabel;
    PPainelBotoes: TPanel;
    DXBExcluir: TdxButtonArround;
    DXBEditar: TdxButtonArround;
    DXBIncluir: TdxButtonArround;
    DXBCancelar: TdxButtonArround;
    DXBGravar: TdxButtonArround;
    DXBPesquisar: TdxButtonArround;
    DXBRelatorios: TdxButtonArround;
    DXBFechar: TdxButtonArround;
    DXBUltimo: TdxButtonArround;
    DXBProximo: TdxButtonArround;
    DXBAnterior: TdxButtonArround;
    DXBPrimeiro: TdxButtonArround;
    Label1: TLabel;
    DBECodInt: TDBEdit;
    Label2: TLabel;
    DBEAbrev: TDBEdit;
    Label3: TLabel;
    DBEDescr: TDBEdit;
    DBGTipoDoc: TDBGrid;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel5: TBevel;
    Bevel6: TBevel;
    Bevel7: TBevel;
    Bevel9: TBevel;
    Label14: TLabel;
    DBCB_R_C: TDBComboBox;
    DBCB_R_R: TDBComboBox;
    DBCBCompCh: TDBComboBox;
    DBCBAvista: TDBComboBox;
    DBCBLancBanco: TDBComboBox;
    DBCBLancInclu: TDBComboBox;
    DBComboBox1: TDBComboBox;
    DBComboBox2: TDBComboBox;
    DBComboBox3: TDBComboBox;
    Bevel3: TBevel;
    Label6: TLabel;
    Bevel4: TBevel;
    Label13: TLabel;
    Bevel8: TBevel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label16: TLabel;
    Label12: TLabel;
    XBanner10: TXBanner;
    XBanner1: TXBanner;
    XBanner2: TXBanner;
    procedure FormShow(Sender: TObject);
    procedure Atualiza_tela;
    procedure botoes_setas;
    procedure DXBPrimeiroClick(Sender: TObject);
    procedure DXBAnteriorClick(Sender: TObject);
    procedure DXBProximoClick(Sender: TObject);
    procedure DXBUltimoClick(Sender: TObject);
    procedure DXBGravarClick(Sender: TObject);
    procedure DXBCancelarClick(Sender: TObject);
    procedure DXBIncluirClick(Sender: TObject);
    procedure DXBEditarClick(Sender: TObject);
    procedure DXBExcluirClick(Sender: TObject);
    procedure DXBPesquisarClick(Sender: TObject);
    procedure DXBRelatoriosClick(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure DBGTipoDocEnter(Sender: TObject);
    procedure DBGTipoDocKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGTipoDocMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Desativa_campos;
    procedure Ativa_campos;
    procedure DBGTipoDocDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmCad_TipoDoc: TFrmCad_TipoDoc;

implementation

uses Funcoes, Tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrmCad_TipoDoc.FormShow(Sender: TObject);
begin
  DM_tabelas.ZQTipodoc.open;
  Desativa_campos;
  DBGTipoDoc.SetFocus;
end;

procedure TFrmCad_TipoDoc.Atualiza_tela;
Begin
  LReg.Caption := inttostr(DM_tabelas.ZQTipodoc.RecordCount);
End;

procedure TFrmCad_TipoDoc.botoes_setas;
Begin
  DXBPrimeiro.Enabled := True;
  DXBAnterior.Enabled := True;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  if DM_tabelas.ZQTipodoc.RecNo = 1 then begin
    DXBPrimeiro.Enabled := false;
    DXBAnterior.Enabled := false;
  end;
  if DM_tabelas.ZQTipodoc.RecNo = DM_tabelas.ZQTipodoc.RecordCount then begin
    DXBProximo.Enabled := false;
    DXBUltimo.Enabled := false;
  end;
  Atualiza_Tela;
End;


procedure TFrmCad_TipoDoc.DXBPrimeiroClick(Sender: TObject);
begin
  DM_tabelas.ZQTipodoc.First;
  botoes_setas;
end;

procedure TFrmCad_TipoDoc.DXBAnteriorClick(Sender: TObject);
begin
  DM_tabelas.ZQTipodoc.Prior;
  botoes_setas;
end;

procedure TFrmCad_TipoDoc.DXBProximoClick(Sender: TObject);
begin
  DM_tabelas.ZQTipodoc.Next;
  botoes_setas;
end;

procedure TFrmCad_TipoDoc.DXBUltimoClick(Sender: TObject);
begin
  DM_tabelas.ZQTipodoc.Last;
  botoes_setas;
end;

procedure TFrmCad_TipoDoc.DXBGravarClick(Sender: TObject);
begin
  if not Verif_senha('Tipo de Documento','Gravar Inclusão ou Edição (Botão GRAVAR)','Abrev: '+DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString+'  Descr: '+DM_tabelas.ZQTipoDoc.FieldByName('descricao').Text+' (V->R)='+DM_tabelas.ZQTipoDoc.FieldByName('vend_receb').AsString+' (V->C)='+DM_tabelas.ZQTipoDoc.FieldByName('vend_caixa').AsString+' (R->R)='+DM_tabelas.ZQTipoDoc.FieldByName('receb_receb').AsString+' (R->C)='+DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').AsString) then exit;
  DM_tabelas.ZQTipodoc.Post;
  Desativa_campos;
  DBGTipoDoc.SetFocus;
end;

procedure TFrmCad_TipoDoc.DXBCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Tipo de Documento','Cancelar Inclusão ou Edição (Botão CANCELAR)','Abrev: '+DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString+'  Descr: '+DM_tabelas.ZQTipoDoc.FieldByName('descricao').Text+' (V->R)='+DM_tabelas.ZQTipoDoc.FieldByName('vend_receb').AsString+' (V->C)='+DM_tabelas.ZQTipoDoc.FieldByName('vend_caixa').AsString+' (R->R)='+DM_tabelas.ZQTipoDoc.FieldByName('receb_receb').AsString+' (R->C)='+DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').AsString) then exit;
  DM_tabelas.ZQTipodoc.Cancel;
  Desativa_campos;
  DBGTipoDoc.SetFocus;
end;

procedure TFrmCad_TipoDoc.DXBIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Tipo de Documento','Inclusão de Registro(Botão INCLUIR)','') then exit;
  DM_tabelas.ZQTipodoc.Insert;
  Ativa_campos;
  DBEAbrev.SetFocus;
end;

procedure TFrmCad_TipoDoc.DXBEditarClick(Sender: TObject);
begin
  if not Verif_senha('Tipo de Documento','Edição de Registro(Botão EDITAR)','Abrev: '+DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString+'  Descr: '+DM_tabelas.ZQTipoDoc.FieldByName('descricao').Text+' (V->R)='+DM_tabelas.ZQTipoDoc.FieldByName('vend_receb').AsString+' (V->C)='+DM_tabelas.ZQTipoDoc.FieldByName('vend_caixa').AsString+' (R->R)='+DM_tabelas.ZQTipoDoc.FieldByName('receb_receb').AsString+' (R->C)='+DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').AsString) then exit;
  DM_tabelas.ZQTipodoc.Edit;
  Ativa_campos;
  DBEAbrev.SetFocus;
end;

procedure TFrmCad_TipoDoc.DXBExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Tipo de Documento','Exclusão de Registro (Botão EXCLUSÃO)','Abrev: '+DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString+'  Descr: '+DM_tabelas.ZQTipoDoc.FieldByName('descricao').Text+' (V->R)='+DM_tabelas.ZQTipoDoc.FieldByName('vend_receb').AsString+' (V->C)='+DM_tabelas.ZQTipoDoc.FieldByName('vend_caixa').AsString+' (R->R)='+DM_tabelas.ZQTipoDoc.FieldByName('receb_receb').AsString+' (R->C)='+DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').AsString) then exit;
  if simnao('Confirma a exclusão do tipo de documento ?','SIM') then
  begin
    IF DM_tabelas.ZQTipodoc.RecordCount>1 then
      DM_tabelas.ZQTipodoc.Delete;
    botoes_setas;
  end;
end;

procedure TFrmCad_TipoDoc.DXBPesquisarClick(Sender: TObject);
begin
// pesquisa
end;

procedure TFrmCad_TipoDoc.DXBRelatoriosClick(Sender: TObject);
begin
// relatorio
end;

procedure TFrmCad_TipoDoc.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmCad_TipoDoc.DBGTipoDocEnter(Sender: TObject);
begin
  botoes_setas;
end;

procedure TFrmCad_TipoDoc.DBGTipoDocKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes_setas;
end;

procedure TFrmCad_TipoDoc.DBGTipoDocMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  botoes_setas;
end;

procedure TFrmCad_TipoDoc.Desativa_campos;
Begin
  DBGTipoDoc.Enabled := True;
  DBEAbrev.ReadOnly := True;
  DBEDescr.ReadOnly := True;
  DBCB_R_R.ReadOnly := True;
  DBCB_R_C.ReadOnly := True;
  DBCBLancInclu.ReadOnly := True;
  DBCBLancBanco.ReadOnly := True;
  DBCBCompCh.ReadOnly := True;
  DBCBAvista.ReadOnly := True;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  DXBIncluir.Enabled := true;
  DXBEditar.Enabled := true;
  DXBExcluir.Enabled := true;
  DXBPesquisar.Enabled := true;
  DXBRelatorios.Enabled := true;
  DXBGravar.Enabled := False;
  DXBCancelar.Enabled := False;
end;

procedure TFrmCad_TipoDoc.Ativa_campos;
Begin
  DBGTipoDoc.Enabled := False;
  DBEAbrev.ReadOnly := False;
  DBEDescr.ReadOnly := False;
  DBCB_R_R.ReadOnly := False;
  DBCB_R_C.ReadOnly := False;
  DBCBLancInclu.ReadOnly := False;
  DBCBLancBanco.ReadOnly := False;
  DBCBCompCh.ReadOnly := False;
  DBCBAvista.ReadOnly := False;
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
end;

procedure TFrmCad_TipoDoc.DBGTipoDocDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECodInt.Text = DM_tabelas.ZQTipoDoc.FieldByName('idtipodocumento').Text Then Begin
    DBGTipoDoc.Canvas.Brush.Color :=$006CFFFF;
    DBGTipoDoc.Canvas.Font.Color := $00A80000;
    DBGTipoDoc.Canvas.Font.Style := [FsBold];
  end;
  DBGTipoDoc.DefaultDrawDataCell(Rect,DBGTipoDoc.columns[datacol].field, State);
end;

procedure TFrmCad_TipoDoc.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmCad_TipoDoc.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_tabelas.ZQTipodoc.close;
  FrmCad_TipoDoc:=nil;
  Action:=CaFree;
end;


procedure TFrmCad_TipoDoc.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
