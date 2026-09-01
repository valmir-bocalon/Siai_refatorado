unit Price;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, wwdbdatetimepicker, StdCtrls, DBCtrls, Mask, Grids, DBGrids,
    dxButton, ExtCtrls, XBanner, DB, dxCore2;

type
  TFrm_Price = class(TForm)
    XBanner14: TXBanner;
    Label19: TLabel;
    Label18: TLabel;
    Label20: TLabel;
    Label23: TLabel;
    Panel4: TPanel;
    XBanner13: TXBanner;
    BtExcluir: TdxButton;
    BtEditar: TdxButton;
    BtIncluir: TdxButton;
    BtCancelar: TdxButton;
    BtGravar: TdxButton;
    BtPesquisar: TdxButton;
    BtRelatorio: TdxButton;
    BtUltimo: TdxButton;
    BTProximo: TdxButton;
    BTAnterior: TdxButton;
    BTPrimeiro: TdxButton;
    DBGCid: TDBGrid;
    DBECodCid: TDBEdit;
    DBECidNome: TDBEdit;
    BtFechar: TdxButton;
    Label1: TLabel;
    Registros: TLabel;
    LReg: TLabel;
    Label2: TLabel;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    procedure BTPrimeiroClick(Sender: TObject);
    procedure BTAnteriorClick(Sender: TObject);
    procedure ativacidade;
    procedure desativacidade;
    procedure botoes_cidade;
    procedure BTProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtCancelarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BtExcluirClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure BtFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DBGCidKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGCidMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGCidEnter(Sender: TObject);
    procedure DBGCidDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGCidTitleClick(Column: TColumn);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BtRelatorioClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Price: TFrm_Price;

implementation

uses tabelas, Funcoes, RelPrice, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Price.BTPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQPrice.First;
  botoes_cidade;
end;

procedure TFrm_Price.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQPrice.Prior;
  botoes_cidade;
end;

procedure TFrm_Price.ativacidade;
Begin
  DBECidNome.ReadOnly := False;
  DBEdit1.ReadOnly    := False;
  DBEdit2.ReadOnly    := False;
  DBEdit3.ReadOnly    := False;
  BTPrimeiro.Enabled := False;
  BTAnterior.Enabled := False;
  BTProximo.Enabled := False;
  BtUltimo.Enabled := False;
  BtGravar.Enabled := True;
  BtCancelar.Enabled := True;
  BtIncluir.Enabled := False;
  BtEditar.Enabled := False;
  BtExcluir.Enabled := False;
  BtPesquisar.Enabled := False;
  BtRelatorio.Enabled := False;
  DBGCid.Enabled := False;
end;

procedure TFrm_Price.desativacidade;
Begin
  DBECidNome.ReadOnly := true;
  DBEdit1.ReadOnly    := true;
  DBEdit2.ReadOnly    := true;
  DBEdit3.ReadOnly    := true;
  BTPrimeiro.Enabled := true;
  BTAnterior.Enabled := true;
  BTProximo.Enabled := true;
  BtUltimo.Enabled := true;
  BtGravar.Enabled := False;
  BtCancelar.Enabled := False;
  BtIncluir.Enabled := true;
  BtEditar.Enabled := true;
  BtExcluir.Enabled := true;
  BtPesquisar.Enabled := true;
  BtRelatorio.Enabled := true;
  DBGCid.Enabled := True;
end;

procedure TFrm_Price.botoes_cidade;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQPrice.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQPrice.RecNo=DM_Tabelas.ZQPrice.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  LReg.Caption := inttostr(DM_Tabelas.ZQPrice.RecordCount);
End;

procedure TFrm_Price.BTProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQPrice.Next;;
  botoes_cidade;
end;

procedure TFrm_Price.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQPrice.Last;
  botoes_cidade;
end;

procedure TFrm_Price.BtGravarClick(Sender: TObject);
begin
  if not Verif_senha('Price','Gravar inclus�o ou edi��o','Taxa M�s: '+DM_Tabelas.ZQprice.FieldByName('Taxa').Text) Then Exit;
  DM_Tabelas.ZQPrice.Post;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_Price.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Price','Cancelar inclus�o ou edi��o','Taxa M�s: '+DM_Tabelas.ZQprice.FieldByName('Taxa').Text) Then Exit;
  DM_Tabelas.ZQPrice.Cancel;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_Price.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Price','Inclus�o','') Then Exit;
  DM_Tabelas.ZQPrice.Insert;
  ativacidade;
  DBECidNome.SetFocus;

end;

procedure TFrm_Price.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('Price','Edi��o','C�digo: '+DM_Tabelas.ZQprice.FieldByName('idPrice').Text+' - '+DM_Tabelas.ZQprice.FieldByName('Taxa').Text) Then Exit;
  DM_Tabelas.ZQPrice.Edit;
  ativacidade;
  DBECidNome.SetFocus;

end;

procedure TFrm_Price.BtExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Price','Exclus�o','C�digo: '+DM_Tabelas.ZQprice.FieldByName('idPrice').Text+' - '+DM_Tabelas.ZQprice.FieldByName('Taxa').Text) Then Exit;
  if DM_Tabelas.ZQPrice.RecordCount>0 Then
    DM_Tabelas.ZQPrice.Delete;

end;

procedure TFrm_Price.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Price.BtFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Price.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZQPrice.Open;
  desativacidade;
  DBGCid.SetFocus;
end;

procedure TFrm_Price.DBGCidKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes_cidade;
end;

procedure TFrm_Price.DBGCidMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  botoes_cidade;
end;

procedure TFrm_Price.DBGCidEnter(Sender: TObject);
begin
  botoes_cidade;
end;

procedure TFrm_Price.DBGCidDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECodCid.Text = DM_Tabelas.ZQprice.FieldByName('idPrice').Text Then Begin
    DBGCid.Canvas.Brush.Color :=$006CFFFF;
    DBGCid.Canvas.Font.Color := $00A80000;
    DBGCid.Canvas.Font.Style := [FsBold];
  end;
  DBGCid.DefaultDrawDataCell(Rect, DBGCid.columns[datacol].field, State);

end;

procedure TFrm_Price.DBGCidTitleClick(Column: TColumn);
var
  campo : string;
  varcodigo : integer;
begin
  campo:=column.fieldname;
  DBGCid.Columns[0].Color := clWindow;
  DBGCid.Columns[1].Color := clWindow;
  DBGCid.Columns[2].Color := clWindow;
  DBGCid.Columns[3].Color := clWindow;
  DBGCid.Columns[4].Color := clWindow;
  application.processmessages;
  varcodigo := DM_Tabelas.ZQprice.FieldByName('idPrice').AsLargeInt;
  DM_Tabelas.ZQPrice.sql.clear;
  DM_Tabelas.ZQPrice.SQL.Add('select * from price');
  DM_Tabelas.ZQPrice.SQL.Add(' order by '+campo);
  DM_Tabelas.ZQPrice.Open;
  DM_Tabelas.ZQPrice.Locate('idprice',varcodigo,[]);
  column.color:=$00EDF0F1;
  botoes_cidade;
end;

procedure TFrm_Price.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DM_Tabelas.ZQPrice.close;
  Frm_Price:=nil;
  Action:=CaFree;
end;

procedure TFrm_Price.BtRelatorioClick(Sender: TObject);
begin
  Frm_Price.FormStyle:=fsNormal;
  if Frm_Relprice=nil then
    Frm_Relprice:=TFrm_Relprice.Create(Application);
  Frm_Relprice.RLReport1.PreviewModal;
  Frm_Relprice:=nil;
  Frm_Price.FormStyle:=fsStayOntop;
end;


procedure TFrm_Price.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
