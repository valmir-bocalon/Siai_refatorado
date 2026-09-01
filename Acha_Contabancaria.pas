unit Acha_Contabancaria;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Data.DB;

type
  TFrm_Acha_Contabancaria = class(TForm)
    XBanner2: TXBanner;
    Label1: TLabel;
    Label2: TLabel;
    Edit1: TEdit;
    DBEContabancaria: TDBGrid;
    procedure Edit1Change(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
    procedure DBEContabancariaKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Acha_Contabancaria: TFrm_Acha_Contabancaria;

implementation

{$R *.dfm}

uses tabelas, uRuntimeFields;

procedure TFrm_Acha_Contabancaria.Edit1Change(Sender: TObject);
begin
  DM_Tabelas.ZQAchaContaBanc.SQL.Clear;
  DM_Tabelas.ZQAchaContaBanc.SQL.Add('Select * from conta_bancaria join participante on idparticipante=idpaticipante');
  DM_Tabelas.ZQAchaContaBanc.SQL.Add(' where nomebanco like '+quotedstr('%'+Edit1.Text+'%'));
  DM_Tabelas.ZQAchaContaBanc.SQL.Add(' or nometitular like '+quotedstr('%'+Edit1.Text+'%'));
  DM_Tabelas.ZQAchaContaBanc.SQL.Add(' or n_conta like '+quotedstr(Edit1.Text+'%'));
  DM_Tabelas.ZQAchaContaBanc.Open;
end;

procedure TFrm_Acha_Contabancaria.Edit1KeyPress(Sender: TObject;
  var Key: Char);
begin
  Label2.Caption := 'T';
  if key = #27 Then  Begin
    Label2.Caption := 'F';
    Close;
  end;
  if Key = #13 Then Begin
    if DM_Tabelas.ZQAchaContaBanc.RecordCount=1 Then
      Close
    else
      DBEContabancaria.SetFocus;
  end;
end;

procedure TFrm_Acha_Contabancaria.DBEContabancariaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then  Begin
    Label2.Caption := 'F';
    Close;
  end
  else if key = #13 Then
    Close;
end;

procedure TFrm_Acha_Contabancaria.FormShow(Sender: TObject);
begin
  Edit1.SetFocus;
  Label2.Caption := 'T';
end;


procedure TFrm_Acha_Contabancaria.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
