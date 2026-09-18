unit PerguntaSIMNAO;
interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, StdCtrls, acPNG, Vcl.ExtCtrls,
  ButtonDxArround;
type
  TFrmPerguntaSIMNAO = class(TForm)
    XBanner1: TXBanner;
    DXBNAO: TdxButtonArround;
    DXBSIM: TdxButtonArround;
    Label1: TLabel;
    Label2: TLabel;
    Image1: TImage;
    procedure FormShow(Sender: TObject);
    procedure DXBSIMClick(Sender: TObject);
    procedure DXBNAOClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;
var
  FrmPerguntaSIMNAO: TFrmPerguntaSIMNAO;
implementation
//uses Tabelas;
{$R *.dfm}
procedure TFrmPerguntaSIMNAO.FormShow(Sender: TObject);
begin
{  XBanner1.ColorOf := DM_Tabelas.ZQConfiguracaocorup.Value;
  XBanner1.ColorFor := DM_TAbelas.ZQConfiguracaocordw.Value;
  label1.Font.Color := DM_TAbelas.ZQConfiguracaocortopo.Value;}
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
end.