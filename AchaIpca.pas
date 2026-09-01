unit AchaIpca;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Data.DB;

type
  TFrm_AchaIpca = class(TForm)
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
  Frm_AchaIpca: TFrm_AchaIpca;

implementation

uses Tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_AchaIpca.FormShow(Sender: TObject);
begin
  ECidade.SetFocus;
end;

procedure TFrm_AchaIpca.ECidadeChange(Sender: TObject);
begin
  DM_Tabelas.ZQAchaIpca.SQL.Clear;
  DM_Tabelas.ZQAchaIpca.SQL.Add('select * from Ipca');
  DM_tabelas.ZQAchaIpca.SQL.Add(' Where mes_ano like '+quotedstr('%'+ECidade.Text+'%')+' order by mes_ano');
  DM_tabelas.ZQAchaIpca.Open;
end;

procedure TFrm_AchaIpca.ECidadeKeyPress(Sender: TObject; var Key: Char);
begin
  Label2.Caption := 'V';
  if key = #27 Then  Begin
    Label2.Caption := 'F';
    Close;
  end;
  if Key = #13 Then Begin
    if DM_Tabelas.ZQAchaIpca.RecordCount=1 Then
      Close
    else
      DBGCid.SetFocus;
  end;
end;

procedure TFrm_AchaIpca.DBGCidKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then  Begin
    Label2.Caption := 'F';
    Close;
  end
  else   if key = #13 Then
    Close;

end;


procedure TFrm_AchaIpca.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
