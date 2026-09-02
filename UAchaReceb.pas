unit UAchaReceb;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner, ExtCtrls, DB, Grids, DBGrids;

type
  TFrmAchaReceb = class(TForm)
    EDesc: TEdit;
    RBLan: TRadioButton;
    RBDoc: TRadioButton;
    RBCli: TRadioButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    RBRef: TRadioButton;
    XBanner3: TXBanner;
    procedure EDescKeyPress(Sender: TObject; var Key: Char);
    procedure EDescChange(Sender: TObject);
    procedure RBLanClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    function pesqDoc(texto : string) : boolean;
    function pesqCli(texto : string) : boolean;
    function pesqRef(texto : string) : boolean;
    function pesqLan(texto : string) : boolean;
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmAchaReceb: TFrmAchaReceb;

implementation

uses Tabelas, Funcoes, Cad_Recebimento, uRuntimeFields;

{$R *.dfm}

procedure TFrmAchaReceb.EDescKeyPress(Sender: TObject; var Key: Char);
begin
  if (key = #13) or (key = #27) Then Begin
    Close;
    Exit;
  end;

end;

procedure TFrmAchaReceb.EDescChange(Sender: TObject);
begin
  If RBDoc.Checked Then
    if not pesqdoc(EDesc.Text) Then
      if not pesqlan(EDesc.Text) Then
        if not pesqRef(EDesc.Text) Then
         if not pesqCli(EDesc.Text) Then
            Label3.Caption := 'Nada foi encontrado com esta descrição....';
  if RBLan.Checked Then
    if not pesqlan(EDesc.Text) Then
      if not pesqdoc(EDesc.Text) Then
        if not pesqRef(EDesc.Text) Then
         if not pesqCli(EDesc.Text) Then
            Label3.Caption := 'Nada foi encontrado com esta descrição....';
  if RBRef.Checked Then
    if not pesqRef(EDesc.Text) Then
      if not pesqlan(EDesc.Text) Then
        if not pesqdoc(EDesc.Text) Then
         if not pesqCli(EDesc.Text) Then
            Label3.Caption := 'Nada foi encontrado com esta descrição....';
  if RBCli.Checked Then
    if not pesqCli(EDesc.Text) Then
      if not pesqRef(EDesc.Text) Then
        if not pesqlan(EDesc.Text) Then
          if not pesqdoc(EDesc.Text) Then
            Label3.Caption := 'Nada foi encontrado com esta descrição....';
//  FrmCad_recebimento.botoes_setas;

end;

procedure TFrmAchaReceb.RBLanClick(Sender: TObject);
begin
  EDesc.SetFocus;
end;

procedure TFrmAchaReceb.FormShow(Sender: TObject);
begin
  EDesc.SetFocus;
end;

Function TFrmAchaReceb.pesqDoc(texto : string) : boolean;
Begin
  Label3.Caption := 'Pesquisando por Documento...';
  result := True;
  if DM_Tabelas.ZQRecebimento.Locate('documento',texto,[loPartialKey]) then Begin
    RBDoc.Checked := True;
    Exit;
  end;
  result := False;
End;

Function TFrmAchaReceb.pesqCli(texto : string) : boolean;
Begin
  Label3.Caption := 'Pesquisando por Cliente...';
  result := True;
  DM_Tabelas.ZQRecebimento.sql.clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select * from recebimento order by cliente');
  DM_Tabelas.ZQRecebimento.open;
  if DM_Tabelas.ZqParticipante.Locate('idparticipante',texto,[loPartialKey]) then
    if DM_Tabelas.ZQRecebimento.Locate('cliente',DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt,[]) then Begin
      RBCli.Checked := True;
      exit;
    end;

  if DM_Tabelas.ZqParticipante.Locate('nome',texto,[loPartialKey]) then
    if DM_Tabelas.ZQRecebimento.Locate('cliente',DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt,[]) then Begin
      RBCli.Checked := True;
      exit;
    end;

  result := False;
End;

Function TFrmAchaReceb.pesqRef(texto : string) : boolean;
Begin
  Label3.Caption := 'Pesquisando por Referência...';
  result := True;
  if DM_Tabelas.ZQRecebimento.Locate('refbaixa',Texto,[loPartialKey]) then Begin
    RBRef.Checked := True;
    Exit;
  end;
  result := False;
End;

Function TFrmAchaReceb.pesqLan(texto : string) : boolean;
Begin
  Label3.Caption := 'Pesquisando por Lançamento...';
  result := True;
  if DM_Tabelas.ZQRecebimento.Locate('IdRecebimento',Texto,[loPartialKey]) then Begin
    RBlan.Checked := true;
    Exit;
  end;
  result := False;
End;


procedure TFrmAchaReceb.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
