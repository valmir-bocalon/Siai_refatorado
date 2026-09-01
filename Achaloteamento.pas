unit Achaloteamento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Data.DB;

type
  TFrm_Achaloteamento = class(TForm)
    Label2: TLabel;
    Label1: TLabel;
    XBanner2: TXBanner;
    ELoteamento: TEdit;
    DBGAchLot: TDBGrid;
    procedure ELoteamentoChange(Sender: TObject);
    procedure ELoteamentoKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure DBGAchLotKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Achaloteamento: TFrm_Achaloteamento;

implementation

uses Tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Achaloteamento.ELoteamentoChange(Sender: TObject);
begin
  DM_Tabelas.ZQAchaLotea.SQL.Clear;
  DM_Tabelas.ZQAchaLotea.SQL.Add('select * from loteamento');
  DM_Tabelas.ZQAchaLotea.SQL.Add(' where apelido like '+quotedstr('%'+ELoteamento.Text+'%')+'  order by apelido');
  DM_Tabelas.ZQAchaLotea.Open;
end;

procedure TFrm_Achaloteamento.ELoteamentoKeyPress(Sender: TObject;
  var Key: Char);
begin
  Label1.Caption := 'V';
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end;
  if Key = #13 Then Begin
    if DM_Tabelas.ZQAchaLotea.RecordCount=1 Then
      Close
    else
      DBGAchLot.SetFocus;
  end;
end;

procedure TFrm_Achaloteamento.FormShow(Sender: TObject);
begin
  ELoteamento.SetFocus;
end;

procedure TFrm_Achaloteamento.DBGAchLotKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end
  else   if key = #13 Then
    Close;
end;


procedure TFrm_Achaloteamento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
