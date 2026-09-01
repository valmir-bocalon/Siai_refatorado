unit Achavendas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Mask, DBCtrls, Data.DB;

type
  TFrm_Achavendas = class(TForm)
    XBanner2: TXBanner;
    Label2: TLabel;
    Eparticipante: TEdit;
    DBGPart: TDBGrid;
    Label1: TLabel;
    DBECod: TDBEdit;
    procedure EparticipanteChange(Sender: TObject);
    procedure EparticipanteKeyPress(Sender: TObject; var Key: Char);
    procedure DBGPartKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure DBGPartDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Achavendas: TFrm_Achavendas;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Achavendas.EparticipanteChange(Sender: TObject);
begin

  DM_Tabelas.ZQAchavenda.SQL.Clear;
  DM_Tabelas.ZQAchavenda.SQL.Add('Select * from venda join imovel on idimovel=idimovel join loteamento on loteamento_idloteamento=idloteamento');
  DM_Tabelas.ZQAchavenda.SQL.Add(' where nomeloteamento like '+quotedstr('%'+Eparticipante.Text+'%'));
  DM_Tabelas.ZQAchavenda.SQL.Add(' or quadra like '+quotedstr('%'+Eparticipante.Text+'%'));
  DM_Tabelas.ZQAchavenda.SQL.Add(' or lote like '+quotedstr(Eparticipante.Text+'%'));
  DM_Tabelas.ZQAchavenda.SQL.Add(' order by nomeloteamento');
  DM_Tabelas.ZQAchavenda.Open;
end;

procedure TFrm_Achavendas.EparticipanteKeyPress(Sender: TObject;
  var Key: Char);
begin
  Label1.Caption := 'V';
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end;
  if Key = #13 Then Begin
    if DM_Tabelas.ZQAchavenda.RecordCount=1 Then
      Close
    else
      DBGPart.SetFocus;
  end;
end;

procedure TFrm_Achavendas.DBGPartKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end
  else if key = #13 Then
    Close;
end;

procedure TFrm_Achavendas.FormShow(Sender: TObject);
begin
  Eparticipante.SetFocus;
end;

procedure TFrm_Achavendas.DBGPartDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQAchaVenda.FieldByName('idvenda').Text Then Begin
    DBGPart.Canvas.Brush.Color :=$006CFFFF;
    DBGPart.Canvas.Font.Color := $00A80000;
  end;
  DBGPart.DefaultDrawDataCell(Rect, DBGPart.columns[datacol].field, State);

end;


procedure TFrm_Achavendas.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
