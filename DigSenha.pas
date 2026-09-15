unit DigSenha;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, XBanner, acPNG;

type
  TFrm_DigSenha = class(TForm)
    Label5: TLabel;
    Label6: TLabel;
    XBanner1: TXBanner;
    Panel1: TPanel;
    imgIcon: TImage;
    imgIcon3: TImage;
    Label1: TLabel;
    Label2: TLabel;
    EUsuario: TEdit;
    ESenha: TEdit;
    Panel2: TPanel;
    XBanner2: TXBanner;
    Label3: TLabel;
    Label4: TLabel;
    procedure FormShow(Sender: TObject);
    procedure EUsuarioExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure ESenhaExit(Sender: TObject);
    procedure FormMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
    procedure WMNCHitTest(var Msg: TWMNCHitTest); message WM_NCHITTEST;

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_DigSenha: TFrm_DigSenha;
  vartentativa : integer;
implementation

{$R *.dfm}

uses principal, tabelas, Funcoes, uRuntimeFields;

procedure TFrm_DigSenha.FormShow(Sender: TObject);
begin
  Label6.Caption := 'F';
  EUsuario.Text := VarUsuario;
  ESenha.Text := '';
  vartentativa := 1;
  if not empty(EUsuario.Text) Then
    ESenha.SetFocus;
end;

procedure TFrm_DigSenha.WMNCHitTest(var Msg: TWMNCHitTest);
begin
inherited;
  if Msg.Result = HTCLIENT then
    Msg.Result := HTCAPTION; // Permite arrastar o formulário ao clicar na Área do cliente
end;

procedure TFrm_DigSenha.EUsuarioExit(Sender: TObject);
begin
  Label6.Caption := 'F';
  VarUsuario := EUsuario.Text;
  DM_Tabelas.ZQUsuTemp.SQL.Clear;
  DM_Tabelas.ZQUsuTemp.SQL.Add('select * from usuario');
  DM_Tabelas.ZQUsuTemp.SQL.Add('where apelido = '+quotedstr(EUsuario.Text));
  DM_Tabelas.ZQUsuTemp.Open;
  if DM_Tabelas.ZQUsuTemp.RecordCount=0 Then Begin
    showmessage('Usuário não cadastrado...');
    EUsuario.Text := '';
    EUsuario.SetFocus;
  end;

end;

procedure TFrm_DigSenha.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then Begin
    Close;
  end;
  if Key = #13 then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_DigSenha.FormMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbLeft then
  begin
    ReleaseCapture; // Libera o foco do mouse
    SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0); // Simula o clique na barra de título
  end;
end;

procedure TFrm_DigSenha.ESenhaExit(Sender: TObject);
begin
  Label6.Caption := 'F';
  if not empty(ESenha.Text) Then Begin
    if DM_Tabelas.ZQUsuTemp.FieldByName('senha').AsString<>ESenha.Text Then Begin
      Showmessage('Senha INCORRETA'+chr(13)+'Tentativa ' +inttostr(vartentativa));
      if vartentativa<4 Then
        ESenha.SetFocus
      else
        close;
      inc(vartentativa);
    end
    else Begin
      if DM_Tabelas.ZQPermissoes.Locate('local_senha,usuario',VarArrayOf([DM_Tabelas.ZQLoc_Senha.FieldByName('idlocais_com_senha').AsLargeInt,DM_Tabelas.ZQUsuTemp.FieldByName('idusuario').AsLargeInt]),[]) Then
      begin
        Frm_principal.xusuario.Caption:=EUsuario.text;
        Label6.Caption := 'V';
      end
      else
      begin
        showmessage('Você não está autorizado a executar este procedimento!!!!');
      end;
      Close;
    end;
  end;
end;


procedure TFrm_DigSenha.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.

