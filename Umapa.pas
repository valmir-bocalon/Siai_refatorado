unit Umapa;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, OleCtrls, SHDocVw, ExtCtrls, XBanner;

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

procedure Tfrmgoogle.Button1Click(Sender: TObject);
Var ConsultaEndereco : string;
begin
   ConsultaEndereco := 'http://maps.google.com/maps?q=';
   If Cidade.Text <> '' Then
      ConsultaEndereco := ConsultaEndereco + Cidade.Text + ',';
   If Estado.Text <> '' Then
      ConsultaEndereco := ConsultaEndereco + Estado.Text + ',';
   If Rua.Text <> ''  Then
      ConsultaEndereco := ConsultaEndereco + Rua.Text + ',';
   If Cep.Text <> ''  Then
      ConsultaEndereco := ConsultaEndereco + Cep.Text;
   Caption := ConsultaEndereco;
   WebBrowser1.Navigate(ConsultaEndereco);
end;

procedure Tfrmgoogle.Button2Click(Sender: TObject);
Var ConsultaEndereco : string;
begin
   ConsultaEndereco := 'http://maps.google.com/maps?q=';
   If ( Latitude.Text <> '' ) and ( Longitude.Text <> '' ) Then
      ConsultaEndereco := ConsultaEndereco + Latitude.Text + ',' + Longitude.Text ;
   Caption := ConsultaEndereco;
   WebBrowser1.Navigate(ConsultaEndereco);
end;

procedure Tfrmgoogle.Button3Click(Sender: TObject);
begin
     Webbrowser1.Navigate(ExtractFilePath(Application.ExeName)+'pesquisa.html');
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
