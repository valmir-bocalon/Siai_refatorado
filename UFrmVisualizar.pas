unit UFrmVisualizar;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QrExtra, QrExport, qrprntr,  ExtCtrls, StdCtrls, Spin, Buttons, // Label3D,
  XP_Form, XBanner,   dxButton,QRPDFFilt, dxCore2;

type
  TFrmVisualizar = class(TForm)
//    tfXPForm1: TtfXPForm;
    Panel1: TPanel;
    XBanner1: TXBanner;
    BtnPagInteira: TSpeedButton;
    BtnLargura: TSpeedButton;
    SpeedButton2: TSpeedButton;
    BtnImprimir: TSpeedButton;
    LblPagina: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    LblZoom: TLabel;
    btntxt: TSpeedButton;
    btntxt1: TSpeedButton;
    SpnPagina: TSpinButton;
    SpnZoom: TSpinButton;
    QRPreview1: TQRPreview;
    DXBFechar: TdxButton;
    SaveDialog1: TSaveDialog;
    SaveDialog2: TSaveDialog;
    procedure BtnImprimirClick(Sender: TObject);
    procedure SpnPaginaDownClick(Sender: TObject);
    procedure SpnPaginaUpClick(Sender: TObject);
    procedure SpnZoomDownClick(Sender: TObject);
    procedure SpnZoomUpClick(Sender: TObject);
    procedure BtnPagInteiraClick(Sender: TObject);
    procedure BtnLarguraClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure QRPreview1PageAvailable(Sender: TObject; PageNum: Integer);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure SpeedButton2Click(Sender: TObject);
    procedure btntxtClick(Sender: TObject);
    procedure btntxt1Click(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
    procedure MostraPagZoom;
  end;

var
  FrmVisualizar: TFrmVisualizar;

implementation

uses Tabelas, uRuntimeFields;

{$R *.DFM}

procedure TFrmVisualizar.MostraPagZoom;
begin
  LblPagina.Caption := IntToStr(QRPreview1.PageNumber) +
      '/' + IntToStr(QRPreview1.QRPrinter.PageCount);
  LblZoom.Caption := IntToStr(QRPreview1.Zoom) + '%';
end;

procedure TFrmVisualizar.BtnImprimirClick(Sender: TObject);
begin
  try
    frmvisualizar.FormStyle:=fsNormal;
    QRPreview1.ZoomToWidth;
//    if Pergunta('Confirma a Impress�o ?')='S'then
       QRPreview1.QRPrinter.Print;
  finally
    frmvisualizar.FormStyle:=fsStayOnTop;
    QRPreview1.ZoomToWidth;
  end;

end;

procedure TFrmVisualizar.SpnPaginaDownClick(Sender: TObject);
begin
  with QRPreview1 do
    if PageNumber > 1 then
    begin
      PageNumber := PageNumber - 1;
      MostraPagZoom;
    end;
end;

procedure TFrmVisualizar.SpnPaginaUpClick(Sender: TObject);
begin
  with QRPreview1 do
    if PageNumber < QRPrinter.PageCount then
    begin
      PageNumber := PageNumber + 1;
      MostraPagZoom;
    end;
end;

procedure TFrmVisualizar.SpnZoomDownClick(Sender: TObject);
begin
  with QRPreview1 do
    if Zoom >= 20 then
    begin
      Zoom := Zoom - 10;
      MostraPagZoom;
    end;
end;

procedure TFrmVisualizar.SpnZoomUpClick(Sender: TObject);
begin
  with QRPreview1 do
    if Zoom <= 190 then
    begin
      Zoom := Zoom + 10;
      MostraPagZoom;
    end;
end;

procedure TFrmVisualizar.BtnPagInteiraClick(Sender: TObject);
begin
  QRPreview1.ZoomToFit;
  MostraPagZoom;
end;

procedure TFrmVisualizar.BtnLarguraClick(Sender: TObject);
begin
  QRPreview1.ZoomToWidth;
  MostraPagZoom;
end;

procedure TFrmVisualizar.FormShow(Sender: TObject);
begin

  QRPreview1.Zoom := 100;
  MostraPagZoom;
  DXBFechar.SetFocus;
end;

procedure TFrmVisualizar.FormResize(Sender: TObject);
begin
  MostraPagZoom;
end;

procedure TFrmVisualizar.QRPreview1PageAvailable(Sender: TObject;
  PageNum: Integer);
begin
  MostraPagZoom;
end;

procedure TFrmVisualizar.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
  const
  PageDelta=10;
begin
  With VertScrollbar do
  if Key=VK_NEXT then
     Position:=Position+PageDelta
  else if Key = VK_PRIOR then
     position:=Position-PageDelta;
   if key=VK_ADD then
   begin
     with QRPreview1 do
     if Zoom <= 190 then
     begin
      Zoom := Zoom + 10;
      MostraPagZoom;
     end;
   end;
   if key=VK_SUBTRACT then
   begin
    with QRPreview1 do
    if Zoom >= 20 then
    begin
      Zoom := Zoom - 10;
      MostraPagZoom;
    end;
   end;
   if (Shift = ([ssShift])) then
   begin
     if (key=VK_ADD) then
     begin
       with QRPreview1 do
       if PageNumber < QRPrinter.PageCount then
       begin
         PageNumber := PageNumber + 1;
         MostraPagZoom;
       end;
     end;  
   end;
   if (Shift = ([ssShift])) then
   begin
     if(key=VK_SUBTRACT) then
     begin
       with QRPreview1 do
       if PageNumber > 1 then
       begin
         PageNumber := PageNumber - 1;
         MostraPagZoom;
       end;
     end;
   end;
end;

procedure TFrmVisualizar.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  FrmVisualizar:=nil;
  Action:=CaFree;
end;

procedure TFrmVisualizar.FormKeyPress(Sender: TObject; var Key: Char);
begin
   if key=#13 then
   begin
    if FrmVisualizar.WindowState=wsNormal then
    begin
      FrmVisualizar.WindowState:=WsMaximized;
    end
    else
    begin
      FrmVisualizar.WindowState:=WsNormal;
    end;
   end;
end;

procedure TFrmVisualizar.SpeedButton2Click(Sender: TObject);
begin
  try
    frmvisualizar.FormStyle:=fsNormal;
    QRPreview1.ZoomToWidth;
    QRPreview1.QRPrinter.PrintSetup;
  finally
    frmvisualizar.FormStyle:=fsStayOnTop;
    QRPreview1.ZoomToWidth;
  end;    
end;

procedure TFrmVisualizar.btntxtClick(Sender: TObject);
var
//  AExportFilter : TQRHTMLDocumentFilter;
  NomeArq:string;
begin
  // AExportFilter := TQRHTMLDocumentFilter.Create('MASTER.HTM');
     if SaveDialog2.Execute Then Begin
         NomeArq := SaveDialog2.FileName;
         QRPreview1.QRPrinter.ExportToFilter(TQRPDFDocumentFilter.Create(NomeArq));
     end;    
    //  QRPreview1.QRPrinter.ExportToFilter(AExportFilter)
//      AExportFilter.Free;

end;

procedure TFrmVisualizar.btntxt1Click(Sender: TObject);
var
  AExportFilter : TQRAsciiExportFilter;
  varnome : string;
begin
  if SaveDialog1.Execute Then Begin
//    QRPreview1.QRPrinter.Save(savedialog1.filename);
    varnome := SaveDialog1.FileName;
    AExportFilter := TQRAsciiExportFilter.Create(varnome);
    try
      QRPreview1.QRPrinter.ExportToFilter(AExportFilter);
    finally
     AExportFilter.Free;
    end;
  end;
end;

procedure TFrmVisualizar.DXBFecharClick(Sender: TObject);
begin
  Close;
end;


procedure TFrmVisualizar.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.


