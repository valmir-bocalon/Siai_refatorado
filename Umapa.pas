unit Umapa;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, OleCtrls, SHDocVw, ExtCtrls, XBanner, ShellAPI;

type
  Tfrmgoogle = class(TForm)
    Panel1: TPanel;
    WebBrowser1: TWebBrowser;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Cidade: TEdit;
    Estado: TEdit;
    Rua: TEdit;
    Cep: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    Longitude: TEdit;
    Latitude: TEdit;
    Button1: TButton;
    Button2: TButton;
    Panel2: TPanel;
    Button3: TButton;
    Label7: TLabel;
    XBanner4: TXBanner;
    XBanner1: TXBanner;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    function AbrirNoNavegadorPadrao(const URL: string): Boolean;
    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frmgoogle: Tfrmgoogle;

implementation
uses
  uRuntimeFields;

{$R *.dfm}

function Tfrmgoogle.AbrirNoNavegadorPadrao(const URL: string): Boolean;
begin
  Result := ShellExecute(Handle, 'open', PChar(URL), nil, nil, SW_SHOWNORMAL) > 32;
end;

procedure Tfrmgoogle.Button1Click(Sender: TObject);
Var ConsultaEndereco : string;
begin
   ConsultaEndereco := 'https://www.google.com/maps/search/?api=1&query=';
   If Cidade.Text <> '' Then
      ConsultaEndereco := ConsultaEndereco + Cidade.Text + ',';
   If Estado.Text <> '' Then
      ConsultaEndereco := ConsultaEndereco + Estado.Text + ',';
   If Rua.Text <> ''  Then
      ConsultaEndereco := ConsultaEndereco + Rua.Text + ',';
   If Cep.Text <> ''  Then
      ConsultaEndereco := ConsultaEndereco + Cep.Text;
   Caption := ConsultaEndereco;
   ConsultaEndereco := StringReplace(ConsultaEndereco, ' ', '+', [rfReplaceAll]);
   if AbrirNoNavegadorPadrao(ConsultaEndereco) then
     PostMessage(Handle, WM_CLOSE, 0, 0)
   else
     WebBrowser1.Navigate(ConsultaEndereco);
end;

procedure Tfrmgoogle.Button2Click(Sender: TObject);
Var ConsultaEndereco : string;
begin
   ConsultaEndereco := 'https://www.google.com/maps/search/?api=1&query=';
   If ( Latitude.Text <> '' ) and ( Longitude.Text <> '' ) Then
      ConsultaEndereco := ConsultaEndereco + Latitude.Text + ',' + Longitude.Text ;
   Caption := ConsultaEndereco;
   ConsultaEndereco := StringReplace(ConsultaEndereco, ' ', '+', [rfReplaceAll]);
   if AbrirNoNavegadorPadrao(ConsultaEndereco) then
     PostMessage(Handle, WM_CLOSE, 0, 0)
   else
     WebBrowser1.Navigate(ConsultaEndereco);
end;

procedure Tfrmgoogle.Button3Click(Sender: TObject);
var
  CaminhoPesquisa: string;
begin
  CaminhoPesquisa := ExtractFilePath(Application.ExeName) + 'pesquisa.html';
  if AbrirNoNavegadorPadrao(CaminhoPesquisa) then
    PostMessage(Handle, WM_CLOSE, 0, 0)
  else
    WebBrowser1.Navigate(CaminhoPesquisa);
end;

procedure Tfrmgoogle.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  frmgoogle:=nil;
  Action:=CaFree;
end;


procedure Tfrmgoogle.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
