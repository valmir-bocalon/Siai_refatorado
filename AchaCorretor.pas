unit AchaCorretor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Data.DB;

type
  TFrm_AchaCorretor = class(TForm)
    Label19: TLabel;
    XBanner4: TXBanner;
    Ecorretor: TEdit;
    DBGAchaCorretor: TDBGrid;
    Label1: TLabel;
    procedure EcorretorChange(Sender: TObject);
    procedure EcorretorKeyPress(Sender: TObject; var Key: Char);
    procedure DBGAchaCorretorKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_AchaCorretor: TFrm_AchaCorretor;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_AchaCorretor.EcorretorChange(Sender: TObject);
begin
  DM_Tabelas.ZQAchaCorretor.SQL.Clear;
  DM_Tabelas.ZQAchacorretor.SQL.Add('select idcorretor,nomecorretor,doc1,creci,telefone1,telefone2,endereco_corr,bairro_corr,cep_corr,Obser_corr,codcidade,PComissao_corretor ');
  DM_Tabelas.ZQAchaCorretor.SQL.Add(' from corretor');
  DM_Tabelas.ZQAchaCorretor.SQL.Add(' where idcorretor='+quotedstr(Ecorretor.Text));
  DM_Tabelas.ZQAchaCorretor.SQL.Add(' or nomecorretor like '+quotedstr('%'+Ecorretor.Text+'%'));
  DM_Tabelas.ZQAchaCorretor.SQL.Add(' or creci like '+quotedstr(Ecorretor.Text));
  DM_Tabelas.ZQAchaCorretor.Open;
end;

procedure TFrm_AchaCorretor.EcorretorKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then Begin
    label1.Caption := 'F';
    close;
  end;
  if (key = #13) Then Begin
    label1.Caption := 'T';
    if (DM_Tabelas.ZQAchaCorretor.RecordCount=1) Then
      close
    else
      DBGAchaCorretor.SetFocus;
  end;

end;

procedure TFrm_AchaCorretor.DBGAchaCorretorKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then
    label1.Caption := 'F';
  if (key = #13) Then
    label1.Caption := 'T';
  if (key = #27) or (key = #13) Then
    close;

end;

procedure TFrm_AchaCorretor.FormShow(Sender: TObject);
begin
  Ecorretor.SetFocus;
end;


procedure TFrm_AchaCorretor.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
