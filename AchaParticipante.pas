unit AchaParticipante;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, XBanner, Mask, DBCtrls, Data.DB;

type
  TFrm_AchaParticipante = class(TForm)
    XBanner2: TXBanner;
    Label2: TLabel;
    Eparticipante: TEdit;
    DBGPart: TDBGrid;
    Label1: TLabel;
    DBECod: TDBEdit;
    procedure EparticipanteChange(Sender: TObject);
    procedure EparticipanteKeyPress(Sender: TObject; var Key: Char);
    procedure DBGPartKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure DBGPartDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_AchaParticipante: TFrm_AchaParticipante;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_AchaParticipante.EparticipanteChange(Sender: TObject);
begin
  DM_Tabelas.ZQAchaParticip.SQL.Clear;
  DM_Tabelas.ZQAchaParticip.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  DM_Tabelas.ZQAchaParticip.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  DM_Tabelas.ZQAchaParticip.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico');
  DM_Tabelas.ZQAchaParticip.SQL.Add(' from participante');
  DM_Tabelas.ZQAchaParticip.SQL.Add(' where nome_parte like '+quotedstr('%'+Eparticipante.Text+'%'));
  DM_Tabelas.ZQAchaParticip.SQL.Add(' or doc1 like '+quotedstr(Eparticipante.Text+'%'));
  DM_Tabelas.ZQAchaParticip.SQL.Add(' or doc2 like '+quotedstr(Eparticipante.Text+'%'));
  DM_Tabelas.ZQAchaParticip.SQL.Add(' or idpaticipante like '+quotedstr(Eparticipante.Text));
  DM_Tabelas.ZQAchaParticip.SQL.Add(' order by nome_parte');
  DM_Tabelas.ZQAchaParticip.Open;
end;

procedure TFrm_AchaParticipante.EparticipanteKeyPress(Sender: TObject;
  var Key: Char);
begin
  Label1.Caption := 'V';
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end;
  if Key = #13 Then Begin
    if DM_Tabelas.ZQAchaParticip.RecordCount=1 Then
      Close
    else
      DBGPart.SetFocus;
  end;
end;

procedure TFrm_AchaParticipante.DBGPartKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then  Begin
    Label1.Caption := 'F';
    Close;
  end
  else if key = #13 Then
    Close;
end;

procedure TFrm_AchaParticipante.FormShow(Sender: TObject);
begin
  Eparticipante.SetFocus;
end;

procedure TFrm_AchaParticipante.DBGPartDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text Then Begin
    DBGPart.Canvas.Brush.Color :=$006CFFFF;
    DBGPart.Canvas.Font.Color := $00A80000;
  end;
  DBGPart.DefaultDrawDataCell(Rect, DBGPart.columns[datacol].field, State);

end;


procedure TFrm_AchaParticipante.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
