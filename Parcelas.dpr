program Parcelas;

uses
  Forms,
  frmparcela in 'frmparcela.pas' {Form1},
  Funcoes in 'Funcoes.pas' {Frm_Funcoes};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TFrm_Funcoes, Frm_Funcoes);
  Application.Run;
end.
