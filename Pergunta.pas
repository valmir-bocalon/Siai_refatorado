unit Pergunta;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, XEdit, XBanner;

type
  TFrmPergunta = class(TForm)
    XBanner1: TXBanner;
    Label5: TLabel;
    XEdit1: TXEdit;
    procedure XEdit1KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmPergunta: TFrmPergunta;

implementation

uses Tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrmPergunta.XEdit1KeyPress(Sender: TObject; var Key: Char);
begin
  if (key = #13) or (key = #27) Then
    close;
end;


procedure TFrmPergunta.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
