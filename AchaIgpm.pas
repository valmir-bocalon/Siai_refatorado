unit AchaIgpm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Data.DB;

type
  TFrm_AchaIgpm = class(TForm)
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
  Frm_AchaIgpm: TFrm_AchaIgpm;

implementation

uses Tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_AchaIgpm.FormShow(Sender: TObject);
begin
  ECidade.SetFocus;
end;

procedure TFrm_AchaIgpm.ECidadeChange(Sender: TObject);
begin
  DM_Tabelas.ZQAchaIgpm.SQL.Clear;
  DM_Tabelas.ZQAchaIgpm.SQL.Add('select * from Igpm');
  DM_tabelas.ZQAchaIgpm.SQL.Add(' Where mes_ano like '+quotedstr('%'+ECidade.Text+'%')+' order by mes_ano');
  DM_tabelas.ZQAchaIgpm.Open;
end;

procedure TFrm_AchaIgpm.ECidadeKeyPress(Sender: TObject; var Key: Char);
begin
  Label2.Caption := 'V';
  if key = #27 Then  Begin
    Label2.Caption := 'F';
    Close;
  end;
  if Key = #13 Then Begin
    if DM_Tabelas.ZQAchaIgpm.RecordCount=1 Then
      Close
    else
      DBGCid.SetFocus;
  end;
end;

procedure TFrm_AchaIgpm.DBGCidKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then  Begin
    Label2.Caption := 'F';
    Close;
  end
  else   if key = #13 Then
    Close;

end;


procedure TFrm_AchaIgpm.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
