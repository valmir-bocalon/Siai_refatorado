program acerto;

uses
  Forms,
  acertar_quadralote in 'acertar_quadralote.pas' {Fracerto},
  tabelas in 'tabelas.pas' {DM_Tabelas: TDataModule},
  RelREMESSA in 'RelREMESSA.pas' {Frm_RelREMESSA},
  Imoveis in 'Imoveis.pas' {Frm_Imoveis};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFracerto, Fracerto);
  Application.CreateForm(TDM_Tabelas, DM_Tabelas);
  Application.CreateForm(TFrm_RelREMESSA, Frm_RelREMESSA);
  Application.CreateForm(TFrm_Imoveis, Frm_Imoveis);
  Application.Run;
end.
