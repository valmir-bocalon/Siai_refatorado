program limpar;

uses
  Forms,
  integridade_baixa in 'integridade_baixa.pas' {frm_baixa},
  tabelas in 'tabelas.pas' {DM_Tabelas: TDataModule},
  Funcoes in 'Funcoes.pas' {Frm_Funcoes};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(Tfrm_baixa, frm_baixa);
  Application.CreateForm(TDM_Tabelas, DM_Tabelas);
  Application.CreateForm(TFrm_Funcoes, Frm_Funcoes);
  Application.Run;
end.
