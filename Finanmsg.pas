{ Titulo     : Controle Geral de Aplicacoes  v.1.0   VAB Informatica
  Data       : 29/04/98
  Programa   : FINANmsg.PAS
  Comentario : Dialogo de mensagens }
unit FINANmsg;
interface
uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  Forms, Dialogs, ExtCtrls, StdCtrls, Buttons, XP_Form,
  acPNG, XBanner, ButtonDxArround;
type
  TFormMensagem = class(TForm)
    Imagem: TImage;
    btnOK: TdxButtonArround;
    frase: TLabel;
    XBanner1: TXBanner;
    procedure BtnOkClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;
var
  FormMensagem: TFormMensagem;
implementation
{$R *.DFM}
procedure TFormMensagem.BtnOkClick(Sender: TObject);
begin
  Close;
end;
end.