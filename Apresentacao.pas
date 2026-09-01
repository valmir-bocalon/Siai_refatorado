unit Apresentacao;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, StdCtrls, ComCtrls, Gauges, JvExControls,
  JvWaitingProgress, JvSpecialProgress, XBanner, acPNG;

type
  TFrmApresentacao = class(TForm)
    Image1: TImage;
    Timer1: TTimer;
    Panel1: TPanel;
    XBanner1: TXBanner;
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmApresentacao: TFrmApresentacao;
  Varcontagem : integer;

implementation
uses
  uRuntimeFields;

{$R *.dfm}

procedure TFrmApresentacao.FormShow(Sender: TObject);
begin
  Varcontagem := 0;
  Panel1.Width := 0;
end;

procedure TFrmApresentacao.Timer1Timer(Sender: TObject);
begin
  Varcontagem := Varcontagem + 1;
  Panel1.Width := varcontagem;
  if Varcontagem = 470 Then Close;
end;

procedure TFrmApresentacao.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if (key = #13) or (key=#27) Then Close;
end;

procedure TFrmApresentacao.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  FrmApresentacao:=nil;
  Action:=CaFree;
end;


procedure TFrmApresentacao.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
