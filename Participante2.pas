unit Participante2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,   dxButton, XBanner, ExtCtrls, StdCtrls, Grids, DBGrids,
  XDBNum, DBCtrls, Mask, XDBEdit, wwdbdatetimepicker, ComCtrls, dxCore2, Data.DB;

type
  Tfrmparticipante2 = class(TForm)
    Label2: TLabel;
    Panel2: TPanel;
    XBanner8: TXBanner;
    BTPesquisar: TdxButton;
    BtUltimo: TdxButton;
    BtProximo: TdxButton;
    BTAnterior: TdxButton;
    BtPrimeiro: TdxButton;
    BTFechar: TdxButton;
    XBanner10: TXBanner;
    DBGPart: TDBGrid;
    DBECod: TDBEdit;
    Label1: TLabel;
    Label21: TLabel;
    DBCBPessoa: TDBComboBox;
    Label3: TLabel;
    DBERazao: TDBEdit;
    Label5: TLabel;
    DBEDoc2: TDBEdit;
    Label6: TLabel;
    DBEDoc1: TXDBEdit;
    Label22: TLabel;
    DBEAniversario: TwwDBDateTimePicker;
    wwDBDateTimePicker1: TwwDBDateTimePicker;
    Label57: TLabel;
    DBEFone4: TDBEdit;
    Label4: TLabel;
    DBEFone3: TDBEdit;
    Label19: TLabel;
    DBEFone2: TDBEdit;
    Label17: TLabel;
    DBEFone1: TDBEdit;
    Label15: TLabel;
    DBENacionalidade: TDBEdit;
    Label23: TLabel;
    GroupBox1: TGroupBox;
    XBanner3: TXBanner;
    Label9: TLabel;
    Label10: TLabel;
    Label8: TLabel;
    Label25: TLabel;
    Label7: TLabel;
    Label51: TLabel;
    DBECep: TDBEdit;
    DBEEnde: TDBEdit;
    DBEBairr: TDBEdit;
    DBECompl: TDBEdit;
    dxButton1: TdxButton;
    GroupBox2: TGroupBox;
    XBanner4: TXBanner;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label26: TLabel;
    Label56: TLabel;
    DBEEndeCob: TDBEdit;
    DBEBairrCob: TDBEdit;
    DBECepCob: TDBEdit;
    DBEComplCob: TDBEdit;
    ECidEnder: TDBEdit;
    EEstado: TDBEdit;
    ECidaddeCob: TDBEdit;
    EestadoCob: TDBEdit;
    procedure BTPesquisarClick(Sender: TObject);
    procedure BtPrimeiroClick(Sender: TObject);
    procedure BTAnteriorClick(Sender: TObject);
    procedure BtProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure dxButton1Click(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frmparticipante2: Tfrmparticipante2;

implementation

uses tabelas, funcoes, Umapa, uRuntimeFields;

{$R *.dfm}

procedure Tfrmparticipante2.BTPesquisarClick(Sender: TObject);
begin
  frmparticipante2.FormStyle:=fsNormal;
  AchaCLiente(frmparticipante2.Top+105,frmparticipante2.Left+200,'');
  DM_Tabelas.ZqParticipante.Locate('idpaticipante',DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt,[]);
  DBGPart.SetFocus;

end;

procedure Tfrmparticipante2.BtPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZqParticipante.First;
end;

procedure Tfrmparticipante2.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZqParticipante.Prior;
end;

procedure Tfrmparticipante2.BtProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZqParticipante.Next;
end;

procedure Tfrmparticipante2.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZqParticipante.Last;
end;

procedure Tfrmparticipante2.dxButton1Click(Sender: TObject);
begin
  if frmgoogle=nil then
     frmgoogle:=Tfrmgoogle.Create(Application);
  frmgoogle.Cidade.Text:=ECidEnder.Text;
  frmgoogle.Estado.Text:=EEstado.Text;
  frmgoogle.Rua.Text:=DBEEnde.Text;
  frmgoogle.Cep.Text:=charrem(DBECep.Text);
  frmgoogle.Button1click(Nil);  
  frmgoogle.ShowModal;
  frmgoogle:=nil;
end;

procedure Tfrmparticipante2.BTFecharClick(Sender: TObject);
begin
  close;
end;

procedure Tfrmparticipante2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_Tabelas.ZQresponsavel.close;
  frmparticipante2:=nil;
  Action:=CaFree;
end;

procedure Tfrmparticipante2.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZqParticipante.Close;
  DM_Tabelas.ZqParticipante.SQL.Clear;
  DM_Tabelas.ZqParticipante.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  DM_Tabelas.ZqParticipante.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  DM_Tabelas.ZqParticipante.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico ');
  DM_Tabelas.ZqParticipante.SQL.Add('   from participante order by nome_parte');
  DM_Tabelas.ZqParticipante.open;
end;


procedure Tfrmparticipante2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
