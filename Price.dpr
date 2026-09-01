program Price;

uses
  Forms,
  importacao_price in 'importacao_price.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
