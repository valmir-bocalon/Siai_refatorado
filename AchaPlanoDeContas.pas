unit AchaPlanoDeContas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Mask, DBCtrls, Data.DB;

type
  TFrm_AchaPlanoDeContas = class(TForm)
    XBanner2: TXBanner;
    Label2: TLabel;
    Label1: TLabel;
    Eplano: TEdit;
    DBGAchaPlano: TDBGrid;
    DBECodigo: TDBEdit;
    Label3: TLabel;
    LCD: TLabel;
    procedure EplanoChange(Sender: TObject);
    procedure EplanoKeyPress(Sender: TObject; var Key: Char);
    procedure DBGAchaPlanoKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure DBGAchaPlanoDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_AchaPlanoDeContas: TFrm_AchaPlanoDeContas;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_AchaPlanoDeContas.EplanoChange(Sender: TObject);
begin
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Clear;
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add('select * from plano_contas');
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' where (codigo like '+quotedstr(Eplano.Text+'%'));
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' or descricao like '+quotedstr('%'+Eplano.Text+'%'));
  DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' or mascara like '+quotedstr(Eplano.Text+'%')+')');
  if not empty(lcd.Caption) then
    DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' and cred_debi ='+quotedstr(LCD.Caption));
  if not empty(label3.Caption) Then
    DM_Tabelas.ZQAchaPlanoDeContas.SQL.Add(' and classificacao = '+quotedstr(label3.Caption));
  DM_Tabelas.ZQAchaPlanoDeContas.Open;
end;

procedure TFrm_AchaPlanoDeContas.EplanoKeyPress(Sender: TObject;
  var Key: Char);
begin
  Label1.Caption := 'V';
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end;
  if Key = #13 Then Begin
    if DM_Tabelas.ZQAchaPlanoDeContas.RecordCount=1 Then
      Close
    else
      DBGAchaPlano.SetFocus;
  end;
end;

procedure TFrm_AchaPlanoDeContas.DBGAchaPlanoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end
  else if key = #13 Then
    Close;
end;

procedure TFrm_AchaPlanoDeContas.FormShow(Sender: TObject);
begin
  Eplano.SetFocus;
end;

procedure TFrm_AchaPlanoDeContas.DBGAchaPlanoDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECodigo.Text = DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text Then Begin
    DBGAchaPlano.Canvas.Brush.Color :=$006CFFFF;
    DBGAchaPlano.Canvas.Font.Color := $00A80000;
    DBGAchaPlano.Canvas.Font.Style := [FsBold];
  end;
  DBGAchaPlano.DefaultDrawDataCell(Rect, DBGAchaPlano.columns[datacol].field, State);

end;


procedure TFrm_AchaPlanoDeContas.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
