unit Cad_Etiqueta;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids,   dxButton, ExtCtrls, StdCtrls, XBanner,
  XDBNum, Mask, XDBEdit, DBCtrls, DB;

type
  TFrmCad_Etiqueta = class(TForm)
    XBanner1: TXBanner;
    Label5: TLabel;
    LReg: TLabel;
    Label8: TLabel;
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
    DBGEtiqueta: TDBGrid;
    DBECod: TXDBEdit;
    DBEDesc: TXDBEdit;
    Label11: TLabel;
    Label12: TLabel;
    Panel1: TPanel;
    DBEdit1: TDBEdit;
    CheckBox1: TCheckBox;
    XDBNumEdit10: TXDBNumEdit;
    XDBNumEdit11: TXDBNumEdit;
    XDBNumEdit12: TXDBNumEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    DBEdit2: TDBEdit;
    CheckBox2: TCheckBox;
    XDBNumEdit7: TXDBNumEdit;
    XDBNumEdit8: TXDBNumEdit;
    XDBNumEdit9: TXDBNumEdit;
    DBEdit3: TDBEdit;
    CheckBox3: TCheckBox;
    XDBNumEdit13: TXDBNumEdit;
    XDBNumEdit14: TXDBNumEdit;
    XDBNumEdit15: TXDBNumEdit;
    CheckBox4: TCheckBox;
    DBEdit4: TDBEdit;
    XDBNumEdit16: TXDBNumEdit;
    XDBNumEdit17: TXDBNumEdit;
    XDBNumEdit18: TXDBNumEdit;
    GBMedFolha: TGroupBox;
    Label9: TLabel;
    DBEME: TXDBNumEdit;
    GBMedEti: TGroupBox;
    Label3: TLabel;
    DBEEsp: TXDBNumEdit;
    Label6: TLabel;
    DBELar: TXDBNumEdit;
    DBEAlt: TXDBNumEdit;
    Label4: TLabel;
    Label10: TLabel;
    Label15: TLabel;
    XDBNumEdit1: TXDBNumEdit;
    XDBNumEdit2: TXDBNumEdit;
    XBanner3: TXBanner;
    XBanner4: TXBanner;
    Label7: TLabel;
    DBECol: TXDBNumEdit;
    procedure FormShow(Sender: TObject);
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
    procedure Ativar_Campos;
    procedure Desativar_Campos;
    procedure DBGEtiquetaEnter(Sender: TObject);
    procedure DBGEtiquetaMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGEtiquetaKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure DBGEtiquetaDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Atualiza_Tela;
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmCad_Etiqueta: TFrmCad_Etiqueta;

implementation

uses tabelas, funcoes, uRuntimeFields;
{$R *.dfm}

procedure TFrmCad_Etiqueta.FormShow(Sender: TObject);
begin
  DM_TAbelas.ZQConfiguracao.Open;
  DM_Tabelas.ZQEtiqueta.Open;
  XBanner1.ColorOf := DM_Tabelas.ZQConfiguracaocorup.Value;
  XBanner1.ColorFor := DM_TAbelas.ZQConfiguracaocordw.Value;
  XBanner2.ColorOf := DM_Tabelas.ZQConfiguracaocorup.Value;
  XBanner2.ColorFor := DM_TAbelas.ZQConfiguracaocordw.Value;
  XBanner3.ColorOf := DM_Tabelas.ZQConfiguracaocorup.Value;
  XBanner3.ColorFor := DM_TAbelas.ZQConfiguracaocordw.Value;
  XBanner4.ColorOf := DM_Tabelas.ZQConfiguracaocorup.Value;
  XBanner4.ColorFor := DM_TAbelas.ZQConfiguracaocordw.Value;
  Label5.Font.Color := DM_Tabelas.ZQConfiguracaocortopo.Value;
  Panel1.Color := DM_Tabelas.ZQConfiguracaocormeio.Value;
  CheckBox1.Color := DM_Tabelas.ZQConfiguracaocormeio.Value;
  CheckBox2.Color := DM_Tabelas.ZQConfiguracaocormeio.Value;
  CheckBox3.Color := DM_Tabelas.ZQConfiguracaocormeio.Value;
  CheckBox4.Color := DM_Tabelas.ZQConfiguracaocormeio.Value;
  GBMedFolha.Color := DM_Tabelas.ZQConfiguracaocorup.Value;
  GBMedEti.Color := DM_Tabelas.ZQConfiguracaocorup.Value;
  Desativar_Campos;
  DBGEtiqueta.SetFocus;
end;

procedure TFrmCad_Etiqueta.botoes_setas;
Begin
  DXBPrimeiro.Enabled := True;
  DXBAnterior.Enabled := True;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  if DM_Tabelas.ZQEtiqueta.RecNo = 1 then begin
    DXBPrimeiro.Enabled := false;
    DXBAnterior.Enabled := false;
  end;
  if DM_Tabelas.ZQEtiqueta.RecNo = DM_Tabelas.ZQEtiqueta.RecordCount then begin
    DXBProximo.Enabled := false;
    DXBUltimo.Enabled := false;
  end;
  LReg.Caption := inttostr(DM_Tabelas.ZQEtiqueta.RecordCount);
  Atualiza_Tela;
End;

procedure TFrmCad_Etiqueta.DXBPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQEtiqueta.First;
  botoes_setas;
end;

procedure TFrmCad_Etiqueta.DXBAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQEtiqueta.Prior;
  botoes_setas;
end;

procedure TFrmCad_Etiqueta.DXBProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQEtiqueta.Next;
  botoes_setas;
end;

procedure TFrmCad_Etiqueta.DXBUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQEtiqueta.Last;
  botoes_setas;
end;

procedure TFrmCad_Etiqueta.DXBGravarClick(Sender: TObject);
begin
  if not senha('Etiqueta','Grava��o da Inclus�o ou Edi��o (Bot�o GRAVAR) ','C�d.:'+DM_Tabelas.ZQEtiquetaidetiqueta.Text+'  Descr.: '+DM_Tabelas.ZQEtiquetadescricao.Text) then exit;
  DM_TAbelas.ZQEtiqueta.Post;
  Desativar_Campos;
  DBGEtiqueta.SetFocus;
end;

procedure TFrmCad_Etiqueta.DXBCancelarClick(Sender: TObject);
begin
  if not senha('Etiqueta','Cancelamento da Inclus�o ou Edi��o (Bot�o CANCELAR) ','Descr.: '+DM_Tabelas.ZQEtiquetadescricao.Text) then exit;
  DM_TAbelas.ZQEtiqueta.Cancel;
  Desativar_Campos;
  DBGEtiqueta.SetFocus;
end;

procedure TFrmCad_Etiqueta.DXBIncluirClick(Sender: TObject);
begin
  if not senha('Etiqueta','Inclus�o de um novo registro (Bot�o INCLUIR) ','') then exit;
  DM_TAbelas.ZQEtiqueta.Insert;
  Ativar_Campos;
  CheckBox1.Checked := True;
  CheckBox2.Checked := True;
  CheckBox3.Checked := True;
  CheckBox4.Checked := True;
  DM_Tabelas.ZQEtiquetalinha1A.Value := 'S';
  DM_Tabelas.ZQEtiquetalinha2A.Value := 'S';
  DM_Tabelas.ZQEtiquetalinha3A.Value := 'S';
  DM_Tabelas.ZQEtiquetalinha4A.Value := 'S';
  DBEDesc.SetFocus;
end;

procedure TFrmCad_Etiqueta.DXBEditarClick(Sender: TObject);
begin
  if not senha('Etiqueta','Edit��o do registro da etiqueta (Bot�o EDITAR) ','C�d.:'+DM_Tabelas.ZQEtiquetaidetiqueta.Text+'  Descr.: '+DM_Tabelas.ZQEtiquetadescricao.Text) then exit;
  DM_TAbelas.ZQEtiqueta.Edit;
  Ativar_Campos;
  DBEDesc.SetFocus;
end;

procedure TFrmCad_Etiqueta.DXBExcluirClick(Sender: TObject);
begin
  if (not senha('Etiqueta','Exclus�o do registro da etiqueta (Bot�o EXCLUIR)','C�d.:'+DM_Tabelas.ZQEtiquetaidetiqueta.Text+'  Descr.: '+DM_Tabelas.ZQEtiquetadescricao.Text)) and (DM_TAbelas.ZQEtiqueta.RecordCount>0) then exit;
  DM_TAbelas.ZQEtiqueta.Delete;
end;

procedure TFrmCad_Etiqueta.DXBPesquisarClick(Sender: TObject);
begin
// pesquisa
end;

procedure TFrmCad_Etiqueta.DXBRelatoriosClick(Sender: TObject);
begin
  // relatorio
end;

procedure TFrmCad_Etiqueta.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmCad_Etiqueta.Ativar_Campos;
Begin
  DBECod.ReadOnly := False;
  DBEDesc.ReadOnly := False;
//  DBEMT.ReadOnly := False;
  DBEME.ReadOnly := False;
  Panel1.Enabled := True;
  GBMedFolha.Enabled := True;
  GBMedEti.Enabled := True;
  DXBPrimeiro.Enabled := False;
  DXBAnterior.Enabled := False;
  DXBProximo.Enabled := False;
  DXBUltimo.Enabled := False;
  DXBGravar.Enabled := True;
  DXBCancelar.Enabled := True;
  DXBIncluir.Enabled := False;
  DXBEditar.Enabled := False;
  DXBExcluir.Enabled := False;
  DBGEtiqueta.Enabled := False;
End;

procedure TFrmCad_Etiqueta.Desativar_Campos;
Begin
  DBECod.ReadOnly := True;
  DBEDesc.ReadOnly := True;
//  DBEMT.ReadOnly := True;
  DBEME.ReadOnly := True;
  Panel1.Enabled := False;
  GBMedFolha.Enabled := False;
  GBMedEti.Enabled := False;
  DXBPrimeiro.Enabled := True;
  DXBAnterior.Enabled := True;
  DXBProximo.Enabled := True;
  DXBUltimo.Enabled := True;
  DXBGravar.Enabled := False;
  DXBCancelar.Enabled := False;
  DXBIncluir.Enabled := True;
  DXBEditar.Enabled := True;
  DXBExcluir.Enabled := True;
  DBGEtiqueta.Enabled := True;
End;

procedure TFrmCad_Etiqueta.DBGEtiquetaEnter(Sender: TObject);
begin
  botoes_setas;
end;

procedure TFrmCad_Etiqueta.DBGEtiquetaMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  botoes_setas;
end;

procedure TFrmCad_Etiqueta.DBGEtiquetaKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes_setas;
end;

procedure TFrmCad_Etiqueta.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmCad_Etiqueta.CheckBox1Click(Sender: TObject);
begin
  if DM_TAbelas.ZQEtiqueta.State in [DsInsert, DsEdit] Then Begin
    if CheckBox1.Checked Then Begin
      DM_TAbelas.ZQEtiquetalinha1A.Value := 'S';
      DBEdit1.Enabled := True;
      XDBNumEdit10.Enabled := True;
      XDBNumEdit11.Enabled := True;
      XDBNumEdit12.Enabled := True;
    end
    else Begin
      DM_TAbelas.ZQEtiquetalinha1A.Value := ' ';
      DBEdit1.Enabled := False;
      XDBNumEdit10.Enabled := False;
      XDBNumEdit11.Enabled := False;
      XDBNumEdit12.Enabled := False;
    end;
  end;
end;

procedure TFrmCad_Etiqueta.CheckBox2Click(Sender: TObject);
begin
  if DM_TAbelas.ZQEtiqueta.State in [DsInsert, DsEdit] Then Begin
    if CheckBox2.Checked Then Begin
      DM_TAbelas.ZQEtiquetalinha2A.Value := 'S';
      DBEdit2.Enabled := True;
      XDBNumEdit7.Enabled := True;
      XDBNumEdit8.Enabled := True;
      XDBNumEdit9.Enabled := True;
    end
    else begin
      DM_TAbelas.ZQEtiquetalinha2A.Value := ' ';
      DBEdit2.Enabled := False;
      XDBNumEdit7.Enabled := False;
      XDBNumEdit8.Enabled := False;
      XDBNumEdit9.Enabled := False;
    end;
  end;
end;

procedure TFrmCad_Etiqueta.CheckBox3Click(Sender: TObject);
begin
  if DM_TAbelas.ZQEtiqueta.State in [DsInsert, DsEdit] Then Begin
    if CheckBox3.Checked Then Begin
      DM_TAbelas.ZQEtiquetalinha3A.Value := 'S';
      DBEdit3.Enabled := True;
      XDBNumEdit13.Enabled := True;
      XDBNumEdit14.Enabled := True;
      XDBNumEdit15.Enabled := True;
    end
    else begin
      DM_TAbelas.ZQEtiquetalinha3A.Value := ' ';
      DBEdit3.Enabled := False;
      XDBNumEdit13.Enabled := False;
      XDBNumEdit14.Enabled := False;
      XDBNumEdit15.Enabled := False;
    end;
  end;
end;

procedure TFrmCad_Etiqueta.CheckBox4Click(Sender: TObject);
begin
  if DM_TAbelas.ZQEtiqueta.State in [DsInsert, DsEdit] Then Begin
    if CheckBox4.Checked Then Begin
      DM_TAbelas.ZQEtiquetalinha4A.Value := 'S';
      DBEdit4.Enabled := True;
      XDBNumEdit16.Enabled := True;
      XDBNumEdit17.Enabled := True;
      XDBNumEdit18.Enabled := True;
    end
    else begin
      DM_TAbelas.ZQEtiquetalinha4A.Value := ' ';
      DBEdit4.Enabled := False;
      XDBNumEdit16.Enabled := False;
      XDBNumEdit17.Enabled := False;
      XDBNumEdit18.Enabled := False;
    end;
  end;
end;

procedure TFrmCad_Etiqueta.DBGEtiquetaDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQEtiquetaidetiqueta.Text Then Begin
    DBGEtiqueta.Canvas.Brush.Color := DM_Tabelas.ZQConfiguracaocorlinha.Value;
    DBGEtiqueta.Canvas.Font.Color := DM_Tabelas.ZQConfiguracaocorFonte.Value;
  end;
  DBGEtiqueta.DefaultDrawDataCell(Rect, DBGEtiqueta.columns[datacol].field, State);
end;

procedure TFrmCad_Etiqueta.Atualiza_Tela;
Begin
  CheckBox1.Checked := False;
  if DM_TAbelas.ZQEtiquetalinha1A.Value = 'S' Then
    CheckBox1.Checked := True;
  CheckBox2.Checked := False;
  if DM_TAbelas.ZQEtiquetalinha2A.Value = 'S' Then
    CheckBox2.Checked := True;
  CheckBox3.Checked := False;
  if DM_TAbelas.ZQEtiquetalinha3A.Value = 'S' Then
    CheckBox3.Checked := True;
  CheckBox4.Checked := False;
  if DM_TAbelas.ZQEtiquetalinha4A.Value = 'S' Then
    CheckBox4.Checked := True;

End;


procedure TFrmCad_Etiqueta.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;
end.
