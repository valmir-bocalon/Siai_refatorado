unit EscolhaEtiqueta;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, StdCtrls, Grids, DBGrids, XLabel3D, Mask, DBCtrls, XNum;

type
  TFrmEscolhaEtiqueta = class(TForm)
    XBanner1: TXBanner;
    XTit: TXLabel3D;
    DBGrid1: TDBGrid;
    DBECod: TDBEdit;
    Label1: TLabel;
    XNNum: TXNumEdit;
    XNBranco: TXNumEdit;
    procedure FormShow(Sender: TObject);
    procedure DBGrid1KeyPress(Sender: TObject; var Key: Char);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure XNNumKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmEscolhaEtiqueta: TFrmEscolhaEtiqueta;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}


procedure TFrmEscolhaEtiqueta.FormShow(Sender: TObject);
begin
  XNNum.Value := 1;
  XTit.Font.Color := DM_Tabelas.ZQConfiguracaocortopo.Value;
  XBanner1.ColorOf := DM_Tabelas.ZQConfiguracaocorup.Value;
  XBanner1.ColorFor := DM_TAbelas.ZQConfiguracaocordw.Value;
  DM_Tabelas.ZQEtiqueta.Close;
  DM_Tabelas.ZQEtiqueta.Open;
  DBGrid1.SetFocus;

end;

procedure TFrmEscolhaEtiqueta.DBGrid1KeyPress(Sender: TObject;
  var Key: Char);
begin
 if (Key = #27) or (Key = #13) Then
   Close;
end;

procedure TFrmEscolhaEtiqueta.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQEtiquetaidetiqueta.Text Then Begin
    DBGrid1.Canvas.Brush.Color := DM_Tabelas.ZQConfiguracaocorlinha.Value;
    DBGrid1.Canvas.Font.Color := DM_Tabelas.ZQConfiguracaocorFonte.Value;
  end;
  DBGrid1.DefaultDrawDataCell(Rect, DBGrid1.columns[datacol].field, State);

end;

procedure TFrmEscolhaEtiqueta.DBGrid1MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  Close;
end;

procedure TFrmEscolhaEtiqueta.XNNumKeyPress(Sender: TObject;
  var Key: Char);
begin
 if (Key = #27) or (Key = #13) Then
   Close;
end;


procedure TFrmEscolhaEtiqueta.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
