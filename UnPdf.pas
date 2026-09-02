unit UnPdf;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.OleCtrls, SHDocVw, Vcl.StdCtrls;

type
  TFrmVisualizarPdf = class(TForm)
    WebBrowser1: TWebBrowser;
    PathPdf: TLabel;
    procedure FormShow(Sender: TObject);
  private
        { Private declarations }

    procedure AfterConstruction; override;
  public
   procedure MostrarPDFContrato(CaminhoPDF: string);
    { Public declarations }
  end;

var
  FrmVisualizarPdf: TFrmVisualizarPdf;

implementation
uses
  uRuntimeFields;

{$R *.dfm}

procedure TFrmVisualizarPdf.MostrarPDFContrato(CaminhoPDF: string);
begin
    if FileExists(CaminhoPDF) then
      WebBrowser1.Navigate(CaminhoPDF)
    else
      ShowMessage('PDF do contrato não encontrado: ' + CaminhoPDF);
end;
procedure TFrmVisualizarPdf.FormShow(Sender: TObject);
begin
    MostrarPDFContrato(PathPdf.Caption);
end;


procedure TFrmVisualizarPdf.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
