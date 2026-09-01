unit AchaCidade;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Data.DB;

type
  TFrm_AchaCidade = class(TForm)
    XBanner10: TXBanner;
    Label1: TLabel;
    ECidade: TEdit;
    DBGCid: TDBGrid;
    Label2: TLabel;
    procedure FormShow(Sender: TObject);
    procedure ECidadeChange(Sender: TObject);
    procedure ECidadeKeyPress(Sender: TObject; var Key: Char);
    procedure DBGCidKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_AchaCidade: TFrm_AchaCidade;

implementation

uses Tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_AchaCidade.FormShow(Sender: TObject);
begin
  ECidade.SetFocus;
end;

procedure TFrm_AchaCidade.ECidadeChange(Sender: TObject);
begin
  DM_Tabelas.ZQAchaCidade.SQL.Clear;
  DM_Tabelas.ZQAchaCidade.SQL.Add('select * from cidade');
  DM_tabelas.ZQAchaCidade.SQL.Add(' Where nomecid like '+quotedstr(ECidade.Text+'%')+' order by nomecid');
  DM_tabelas.ZQAchaCidade.Open;
end;

procedure TFrm_AchaCidade.ECidadeKeyPress(Sender: TObject; var Key: Char);
begin
  Label2.Caption := 'V';
  if key = #27 Then  Begin
    Label2.Caption := 'F';
    Close;
  end;
  if Key = #13 Then Begin
    if DM_Tabelas.ZQAchaCidade.RecordCount=1 Then
      Close
    else
      DBGCid.SetFocus;
  end;
end;

procedure TFrm_AchaCidade.DBGCidKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then  Begin
    Label2.Caption := 'F';
    Close;
  end
  else   if key = #13 Then
    Close;

end;


procedure TFrm_AchaCidade.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
