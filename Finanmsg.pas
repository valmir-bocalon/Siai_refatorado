
{ Titulo     : Controle Geral de Aplicacoes  v.1.0   VAB Informatica
  Data       : 29/04/98
  Programa   : FINANmsg.PAS
  Comentario : Dialogo de mensagens }
 
unit FINANmsg;

interface

uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  Forms, Dialogs, ExtCtrls, StdCtrls, Buttons, Label3D;

type
  TFormMensagem = class(TForm)
    BtnOk: TBitBtn;
    Frase: TLabel3D;
    Imagem: TImage;
    procedure BtnOkClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FormMensagem: TFormMensagem;

implementation
uses
  uRuntimeFields;

{$R *.DFM}

procedure TFormMensagem.BtnOkClick(Sender: TObject);
begin
  Close;
end;


procedure TFormMensagem.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
 
{ Final FINANMSG.PAS }