program recebimento_zerar;

uses
  Forms,
  zerar_recebimento in 'zerar_recebimento.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
