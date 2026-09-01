program importa_loteamento;

uses
  Forms,
  importar in 'importar.pas' {frm_baixa},
  tabelas in 'tabelas.pas' {DM_Tabelas: TDataModule},
  Funcoes in 'Funcoes.pas' {Frm_Funcoes},
  Achaloteamento in 'Achaloteamento.pas' {Frm_Achaloteamento};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(Tfrm_baixa, frm_baixa);
  Application.CreateForm(TDM_Tabelas, DM_Tabelas);
  Application.CreateForm(TFrm_Funcoes, Frm_Funcoes);
  Application.CreateForm(TFrm_Achaloteamento, Frm_Achaloteamento);
  Application.Run;
end.
