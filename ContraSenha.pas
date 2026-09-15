unit ContraSenha;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DBGrids, XEdit, XBanner;

type
  TFrmContraSenha = class(TForm)
    XBanner1: TXBanner;
    LMasterSenha: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    LMensagem: TLabel;
    XEContraSenha: TXEdit;
    procedure FormShow(Sender: TObject);
    procedure XEContraSenhaKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmContraSenha: TFrmContraSenha;

implementation

uses Tabelas, Funcoes, uRuntimeFields;

{$R *.dfm}


procedure TFrmContraSenha.FormShow(Sender: TObject);
Var
  Vardias : Double;
Begin
  LMasterSenha.Caption := copy(floattostr(random*Time),7,6);
  Vardias := descripto_data(DM_Tabelas.ZQEmpresa.FieldByName('limite').AsString)-date;
  LMensagem.Caption := 'Resta(m) '+floattostr(Vardias)+' dia(s) para o sistema ser Paralisado!!';
end;

procedure TFrmContraSenha.XEContraSenhaKeyPress(Sender: TObject;
  var Key: Char);
begin
if key = #13 Then
    Close;
end;


procedure TFrmContraSenha.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
