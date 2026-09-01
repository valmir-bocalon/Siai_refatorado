program SiAJCopia;

uses
  Forms,
  CopiaAutomatica in 'CopiaAutomatica.pas' {FrmCopiaAutomatica},
  tabelas in 'tabelas.pas' {DM_Tabelas: TDataModule},
  PerguntaSIMNAO in 'PerguntaSIMNAO.pas' {FrmPerguntaSIMNAO};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFrmCopiaAutomatica, FrmCopiaAutomatica);
  Application.CreateForm(TDM_Tabelas, DM_Tabelas);
  Application.CreateForm(TFrmPerguntaSIMNAO, FrmPerguntaSIMNAO);
  Application.Run;
end.
