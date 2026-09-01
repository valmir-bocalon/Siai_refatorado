unit PerguntaSN;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner,   dxButton;

type
  TFrmPerguntaSN = class(TForm)
    XBanner1: TXBanner;
    Label1: TLabel;
    Label2: TLabel;
    DXBSim: TdxButton;
    DXBNao: TdxButton;
    procedure FormShow(Sender: TObject);
    procedure DXBSimClick(Sender: TObject);
    procedure DXBNaoClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmPerguntaSN: TFrmPerguntaSN;

implementation

uses Tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrmPerguntaSN.FormShow(Sender: TObject);
begin
  XBanner1.ColorOf:=DM_tabelas.ZQConfiguracoescorup.Value;
  XBanner1.ColorFor:=DM_Tabelas.ZQConfiguracoescordw.Value;
  DXBSim.SetFocus;
  if label2.Caption = 'N' Then
    DXBNao.SetFocus;
end;

procedure TFrmPerguntaSN.DXBSimClick(Sender: TObject);
begin
  Label2.Caption := 'S';
  Close;
end;

procedure TFrmPerguntaSN.DXBNaoClick(Sender: TObject);
begin
  Label2.Caption := 'N';
  Close;
end;


procedure TFrmPerguntaSN.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
