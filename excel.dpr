program excel;

uses
  Forms,
  exportar_excel in 'exportar_excel.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
