unit alerta;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, Data.DB;

type
  TFrmAlerta = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    DBGrid1: TDBGrid;
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmAlerta: TFrmAlerta;

implementation

uses tabelas, uRuntimeFields;
{$R *.dfm}

procedure TFrmAlerta.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
//  if (TStringGrid(DBGrid1).RowCount-1) < 10 then //Se tiver menos de 10 linhas
    ShowScrollBar(DBGrid1.Handle,SB_VERT,False); //Remove barra Vertical
end;


procedure TFrmAlerta.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
