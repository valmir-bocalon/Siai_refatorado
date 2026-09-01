unit TabelasDeMensagens;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, Grids, DBGrids, StdCtrls,   dxButton, Data.DB, dxCore2;

type
  TFrm_TabelasDeMensagens = class(TForm)
    DXBFechar: TdxButton;
    Label5: TLabel;
    DBGMens: TDBGrid;
    DBGMotivo: TDBGrid;
    Label1: TLabel;
    Label2: TLabel;
    XBanner7: TXBanner;
    procedure DXBFecharClick(Sender: TObject);
    procedure DBGMensKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure DBGMensMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGMensDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGMensKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGMotivoKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_TabelasDeMensagens: TFrm_TabelasDeMensagens;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_TabelasDeMensagens.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_TabelasDeMensagens.DBGMensKeyPress(Sender: TObject;
  var Key: Char);
Var
  varcampo : string;
begin
  varcampo:=DBGMens.SelectedField.FieldName; // .Name;
  if varcampo = 'atitude' Then Begin
    DM_Tabelas.ZQMens_Retorno.Edit;
    if DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString = 'B' then
      DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString := 'R'
    else if DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString = 'R' then
      DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString := 'I'
    else if (DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString = 'I') or (empty(DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString)) then
      DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString := 'B';
    DM_Tabelas.ZQMens_Retorno.Post;
  end;
end;

procedure TFrm_TabelasDeMensagens.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_Tabelas.ZQMens_Motivo.close;
  DM_Tabelas.ZQMens_Retorno.close;
  Frm_TabelasDeMensagens:=nil;
  Action:=CaFree;
end;

procedure TFrm_TabelasDeMensagens.FormShow(Sender: TObject);
begin
  DBGMens.SetFocus;
  DM_Tabelas.ZQMens_Motivo.Open;
  DM_Tabelas.ZQMens_Retorno.Open;
end;

procedure TFrm_TabelasDeMensagens.DBGMensMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var
  varcampo : string;
begin
  varcampo:=DBGMens.SelectedField.FieldName; // .Name;
  if varcampo = 'atitude' Then Begin
    DM_Tabelas.ZQMens_Retorno.Edit;
    if DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString = 'B' then
      DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString := 'R'
    else if DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString = 'R' then
      DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString := 'I'
    else if (DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString = 'I') or (empty(DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString)) then
      DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString := 'B';
    DM_Tabelas.ZQMens_Retorno.Post;
  end;
end;

procedure TFrm_TabelasDeMensagens.DBGMensDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if Column.Field = DM_Tabelas.ZQMens_Retorno.FieldByName('atitude') then begin
    DBGMens.Canvas.FillRect(Rect);
    if DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString = 'B' Then
      DM_TAbelas.ImageList1.draw(DBGMens.Canvas, Rect.Left + 5, Rect.Top + 1, 0)
    else if DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString = 'R' Then
      DM_TAbelas.ImageList1.draw(DBGMens.Canvas, Rect.Left + 5, Rect.Top + 1, 1)
    else if (DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString = 'I') or (empty(DM_Tabelas.ZQMens_Retorno.FieldByName('atitude').AsString)) Then
      DM_TAbelas.ImageList1.draw(DBGMens.Canvas, Rect.Left + 5, Rect.Top + 1, 3);
  end;
end;

procedure TFrm_TabelasDeMensagens.DBGMensKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (DM_Tabelas.ZQMens_Retorno.RecordCount>0) Then Begin
    DM_Tabelas.ZQMens_Motivo.First;
    while not DM_Tabelas.ZQMens_Motivo.Eof do
      DM_Tabelas.ZQMens_Motivo.Delete;
    DM_Tabelas.ZQMens_Retorno.Delete;
    DBGMens.Refresh;
  end;
end;

procedure TFrm_TabelasDeMensagens.DBGMotivoKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (DM_Tabelas.ZQMens_Motivo.RecordCount>0) Then Begin
    DM_Tabelas.ZQMens_Motivo.Delete;
    DBGMotivo.Refresh;
  end;
end;


procedure TFrm_TabelasDeMensagens.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
