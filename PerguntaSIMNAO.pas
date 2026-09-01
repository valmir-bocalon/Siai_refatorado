unit PerguntaSIMNAO;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, StdCtrls, dxButton, dxCore2;

type
  TFrmPerguntaSIMNAO = class(TForm)
    XBanner1: TXBanner;
    DXBNAO: TdxButton;
    DXBSIM: TdxButton;
    Label1: TLabel;
    Label2: TLabel;
    procedure FormShow(Sender: TObject);
    procedure DXBSIMClick(Sender: TObject);
    procedure DXBNAOClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmPerguntaSIMNAO: TFrmPerguntaSIMNAO;

implementation

uses Tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrmPerguntaSIMNAO.FormShow(Sender: TObject);
begin
  DXBNAO.SetFocus;
  if label2.Caption = 'SIM' Then
    DXBSIM.SetFocus;
end;

procedure TFrmPerguntaSIMNAO.DXBSIMClick(Sender: TObject);
begin
  label2.Caption := 'SIM';
  Close;
end;

procedure TFrmPerguntaSIMNAO.DXBNAOClick(Sender: TObject);
begin
  label2.Caption := 'NAO';
  Close;
end;


procedure TFrmPerguntaSIMNAO.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
