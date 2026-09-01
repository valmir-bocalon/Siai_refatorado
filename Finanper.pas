
{ Titulo     : Controle Geral de Aplicacoes  v.1.0   VAB Informatica
  Data       : 29/04/98
  Programa   : FINANper.PAS
  Comentario : Dialogo de questionamento }
 
unit FINANper;

interface

uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  Forms, Dialogs, ExtCtrls, StdCtrls, Buttons;

type
  TFormPergunta = class(TForm)
    Imagem: TImage;
    Resposta: TLabel;
    Frase: TLabel;
    BtnSim: TBitBtn;
    BtnNao: TBitBtn;
    procedure BtnSimClick(Sender: TObject);
    procedure BtnNaoClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FormPergunta: TFormPergunta;

implementation
uses
  uRuntimeFields;

{$R *.DFM}

procedure TFormPergunta.BtnSimClick(Sender: TObject);
begin
  Resposta.Caption := 'S';
  Close;
end;

procedure TFormPergunta.BtnNaoClick(Sender: TObject);
begin
  Resposta.Caption := 'N';
  Close;
end;

procedure TFormPergunta.FormShow(Sender: TObject);
begin
  BtnSim.SetFocus;
end;


procedure TFormPergunta.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
 
{ Final FINANPER.PAS }