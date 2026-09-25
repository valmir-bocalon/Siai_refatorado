unit CopiaAutomatica;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Shellapi, ExtCtrls, FileCtrl, Menus,IniFiles, Registry,
  Mask, XEdit,   dxButton, XBanner, DBGrids;

//  TFlatPanelUnit, TFlatButtonUnit, Buttons, XLabel3D, JvDialogs,
//  JvExControls, JvGradient, JvGradientCaption, JvMenus, MPlayer;

const WM_TRAYICON=WM_USER+1;

type
  TFrmCopiaAutomatica = class(TForm)
    XBanner1: TXBanner;
    Label5: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    LMensagem: TLabel;
    Label15: TLabel;
    DXBGravar: TdxButton;
    DXBFechar: TdxButton;
    ENomeTabela: TEdit;
    EIPServidor: TEdit;
    EUsuarioSQL: TEdit;
    ESenhaSQL: TEdit;
    EDestino: TEdit;
    CBPedesenha: TComboBox;
    ENomeArquivo: TEdit;
    Eporta: TEdit;
    Panel1: TPanel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    XEHora: TXEdit;
    OpenDialog1: TOpenDialog;
    Timer1: TTimer;
    Label16: TLabel;
    XEHora2: TXEdit;
    dxButton1: TdxButton;
    Function SIMNAO(texto, opcao:string) : Boolean; //
    function inttostrZero(numero, tamanho : integer) : string;
    function HexToInt(Hex: string): integer;
    function decodifica(texto : string) : String;
    procedure FormCreate(Sender: TObject);
    procedure EDestinoEnter(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure LerIni;
    procedure Geracopia;
    procedure DXBGravarClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure XEHoraExit(Sender: TObject);
    procedure XEHora2Exit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure dxButton1Click(Sender: TObject);
  private
    { Private declarations }
    Function CreateProcessSimple(cmd: string):boolean;
    procedure WMSysCommand(var Msg: TWMSysCommand); Message WM_SysCommand;
    procedure WMTrayIcon(var Msg: TMessage); message WM_TRAYICON;
    procedure DestroyIcone;
    procedure CriaIcone;

    procedure AfterConstruction; override;
  public
    { Public declarations }
    procedure WMSize(var M : TWMSIZE);Message WM_Size;
  end;

var
  FrmCopiaAutomatica: TFrmCopiaAutomatica;

implementation

uses PerguntaSIMNAO, uRuntimeFields, Funcoes;


{$R *.dfm}

Function TFrmCopiaAutomatica.SIMNAO(texto, opcao:string) : Boolean; //
Begin
  FrmperguntaSIMNAO.Label1.Caption := Texto;
  frmperguntaSIMNAO.Label2.Caption := opcao;
  FrmperguntaSIMNAO.ShowModal;
  result := False;
  if  FrmperguntaSIMNAO.Label2.Caption = 'SIM' Then
    result := True;
End;

function TFrmCopiaAutomatica.inttostrZero(numero, tamanho : integer) : string;
Var
  Varvezes, Xtama : integer;
  Varzeros : String;
Begin
  varzeros := '';
  Xtama := length(inttostr(numero));
  for varvezes := Xtama+1 to tamanho do
    Varzeros := Varzeros + '0';
  result := Varzeros+inttostr(numero);
End;

function TFrmCopiaAutomatica.HexToInt(Hex: string): integer;
begin
  Result := StrToInt('$' + Hex);
end;

function TFrmCopiaAutomatica.decodifica(texto : string) : String;
Var
  VarTamanho,Varconta : integer;
  montalinha1, montalinha2, montalinha3, varletra : string;
  vardivisor : integer;
begin
  vartamanho := length(texto);
  if (vartamanho>0) and (vartamanho mod 3 = 0) then Begin
    vardivisor := strtoint(floattostr(length(texto)/3));
    montalinha1 := copy(texto,1,vardivisor);
    montalinha2 := copy(texto,(vardivisor*1)+1,vardivisor);
    montalinha3 := copy(texto,(vardivisor*2)+1,vardivisor);
    Try
    montalinha1 := inttostrZero(HexToInt(montalinha1),vardivisor);
    except
     exit;
    end;
    Try
      montalinha2 := inttostrZero(HexToInt(montalinha2),vardivisor);
    except
     exit;
    end;
    Try
    montalinha3 := inttostrZero(HexToInt(montalinha3),vardivisor);
    except
     exit;
    end;
    varletra := '';
    for Varconta := 2 to length(montalinha1) do
      varletra := varletra + chr(strtoint(copy(montalinha1,varconta,1)+copy(montalinha2,varconta,1)+copy(montalinha3,varconta,1)));
    Result := varletra;
  end;
end;


procedure TFrmCopiaAutomatica.WMSize(var M:TWMSIZE);
begin
  if M.SizeType=Size_Minimized then
  ShowWindow(Handle,Sw_Hide);
end;

function TFrmCopiaAutomatica.CreateProcessSimple(cmd: string):boolean;
var SUInfo:  TStartupInfo;
   ProcInfo: TProcessInformation;
begin
   FillChar(SUInfo, SizeOf(SUInfo), #0);
   SUInfo.cb      := SizeOf(SUInfo);
   SUInfo.dwFlags := STARTF_USESHOWWINDOW;
   SUInfo.wShowWindow := SW_HIDE;
   result := CreateProcess(nil, PChar(cmd), nil, nil, false, CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil, nil, SUInfo, ProcInfo);
   if result then
   begin
      WaitForSingleObject(ProcInfo.hProcess, INFINITE);
      CloseHandle(ProcInfo.hProcess);
      CloseHandle(ProcInfo.hThread);
   end;
end;

procedure TFrmCopiaAutomatica.CriaIcone;
var
    NotifyIconData: TNotifyIconData;
begin
    with NotifyIconData do
    begin
      cbSize:= SizeOf(TNotifyIconData);
      Wnd:= Self.Handle;
      uID:= 0;
      uCallbackMessage:= WM_TRAYICON;
      uFlags:= NIF_ICON or NIF_TIP or NIF_MESSAGE;
      hIcon:= Application.Icon.Handle;
      szTip:= 'Cópia de dados';
    end;
    Shell_NotifyIcon(NIM_ADD, @NotifyIconData);
//    Application.ShowMainForm:= False;
end;

procedure TFrmCopiaAutomatica.DestroyIcone;
var
    NotifyIconData: TNotifyIconData;
begin
    NotifyIconData.cbSize:= SizeOf(TNotifyIconData);
    NotifyIconData.Wnd:= Self.Handle;
    NotifyIconData.uID:= 0;
    NotifyIconData.uFlags:= 0;
    Shell_NotifyIcon(NIM_DELETE, @NotifyIconData);
end;




    {Captura todas as mensagens enviadas para a aplicação.
     Neste caso, trabalhamos apenas com as mensagens SC_MINIMIZE e SC_MAXIMIZE, que são enviadas
     quando devemos Minimizar ou Maximizar o formulario da aplicação. Capturando estas mensagens
     para esconder o Ícone da aplicação da Barra de Tarefas do Windows e para criar / destruir o
     Ícone que ficará ao lado do relógio.
     Obs: Você pode trabalhar com outras mensagem também. Tipo SC_CLOSE, SC_RESTORE e etc.}
procedure TFrmCopiaAutomatica.WMSysCommand(var Msg: TWMSysCommand);
begin
    case (Msg.CmdType) of
      SC_MINIMIZE:
      begin
        Self.Visible:= False;
        CriaIcone;
      end;
      SC_MAXIMIZE:
      begin
        Self.Visible:= True;
      end
      else
        Inherited
    end;
end;

      {Quando dar um duplo clique sobre o Ícone da aplicação, destruímos este mesmo Ícone e
      exibimos o formulário principal. Isto só funciona se implementarmos um manipulador
      para as Mensagens WMTrayIcon, que é o que faremos abaixo.

      Lembra-se da constante declarada no Início de nossa Unit?
      Pois bem, ela é passada como parâmetro para esta procedure.}
procedure TFrmCopiaAutomatica.WMTrayIcon(var Msg: TMessage);
  var
    Pt: TPoint;
  begin
    if (Msg.LParam = WM_RBUTTONDOWN) then
    begin
      GetCursorPos(Pt);
      PopupMenu.Popup(Pt.X, Pt.Y);
      PopupMenu.AutoPopup:= False;
    end
    else
      if (Msg.LParam = WM_LBUTTONDBLCLK) then
      begin
        Self.Visible:= True;
        DestroyIcone;
      end;
end;




procedure TFrmCopiaAutomatica.FormCreate(Sender: TObject);
type
TRegisterServiceProcess = function (dwProcessID, dwType:DWord) : DWORD; stdcall;

var
Reg: TRegistry;
Handle: THandle;
RegisterServiceProcess: TRegisterServiceProcess;
begin
    LerIni;
    CriaIcone;
    try

      // Inicializa a variavel do tipo TRegistry
      Reg:= TRegistry.Create;

      // Define a hierarquia das pasta, a qual irá trabalhar.
      Reg.RootKey:= HKEY_LOCAL_MACHINE;

      // Cria/entra, dentro da pasta do registro no qual irá aguardar informação
      Reg.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Run', False);

      // Escreve a informação na chave
      // NomeDaChave, pode ser qualquer nome que você queira
      Reg.WriteString('SiAPCopia', Application.ExeName);

      // Fecha a chave do registro
      Reg.CloseKey;
    Finally
      Reg.Free;
    end;

    //*** Nao aparece no Ctr+Alt+Del ***********************************************
{    Handle := LoadLibrary('KERNEL32.DLL');
    RegisterServiceProcess := GetProcAddress(Handle, 'RegisterServiceProcess');
    RegisterServiceProcess(GetCurrentProcessID, 1);
    FreeLibrary(Handle);
    //*** Nao aparece na barra ***********************************************
    SetWindowLong(Application.Handle, GWL_EXSTYLE,
    GetWindowLong(Application.Handle, GWL_EXSTYLE) or
    WS_EX_TOOLWINDOW and not WS_EX_APPWINDOW);
//    timer3.Enabled:=true;
    //esconder
    Application.ShowMainForm := False; }
end;


procedure TFrmCopiaAutomatica.EDestinoEnter(Sender: TObject);
Var
  tamanhonome : integer;
begin
  OpenDialog1.FileName := ENomeArquivo.Text;
  OpenDialog1.InitialDir := EDestino.Text;
  OpenDialog1.Execute;
  for tamanhonome := length(OpenDialog1.FileName) downto 1 do
    if copy(OpenDialog1.FileName,tamanhonome,1) = '\' Then
      Break;
  EDestino.Text := copy(OpenDialog1.FileName,1,tamanhonome);
  ENomeArquivo.Text := copy(OpenDialog1.FileName,tamanhonome+1,length(OpenDialog1.FileName));;
  Perform(WM_NEXTDLGCTL, 0, 0);
  if pos(' ',EDestino.Text)>0 Then Begin
    Mensagem('Não pode existir espaço em branco no meio do nome do DESTINO');
    EDestino.SetFocus;
  end;
end;

procedure TFrmCopiaAutomatica.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmCopiaAutomatica.LerIni;
Var
  ArqIni : tIniFile;
  VarPedeSenha, Varsenha : string;
Begin
  ArqIni := tIniFile.Create('c:\MASTER\backup_Siai.Ini');
  Try
    EIPServidor.Text := ArqIni.ReadString('BANCO_DE_DADOS','HOSTNAME', EIPServidor.Text );
    if length(EIPServidor.Text)=0 then Begin
      ArqIni.WriteString('BANCO_DE_DADOS','HOSTNAME','localhost');
      EIPServidor.Text := 'localhost';
    end;
  Finally
  end;
  Try
  Eporta.Text := ArqIni.ReadString('BANCO_DE_DADOS','PORTA', Eporta.Text );
  if length(Eporta.Text)=0 then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','PORTA','3306');
    Eporta.Text := '3306';
  end;
  Finally
  end;
  Try
    ENomeTabela.Text := ArqIni.ReadString('BANCO_DE_DADOS','SCHEMATA', ENomeTabela.Text );
    if length(ENomeTabela.Text)=0 then Begin
      ArqIni.WriteString('BANCO_DE_DADOS','SCHEMATA','siai');
      ENomeTabela.Text := 'siai';
    end;
  Finally
  end;
  Try
  EUsuarioSQL.Text := ArqIni.ReadString('BANCO_DE_DADOS','USUARIO', EUsuarioSQL.Text );
  if length(EUsuarioSQL.Text)=0 then Begin
    ArqIni.WriteString('BANCO_DE_DADOS','USUARIO','root');
    EUsuarioSQL.Text := 'root';
  end;
  Finally
  end;
  Try
    varsenha := ArqIni.ReadString('BANCO_DE_DADOS','SENHASQL', varsenha );
    if length(varsenha)=0 then Begin
      ArqIni.WriteString('BANCO_DE_DADOS','SENHASQL','000001C70385');
      varsenha := '000001C70385';
    end;
  Finally
  end;
  Try
    EDestino.Text := ArqIni.ReadString('BACKUP','DESTINHO', EDestino.Text );
    if length(EDestino.Text)=0 then Begin
      ArqIni.WriteString('BACKUP','DESTINHO','C:\Master\');
      EDestino.Text := 'C:\Master\';
    end;
  Finally
  end;
  Try
    ENomeArquivo.Text := ArqIni.ReadString('BACKUP','NOME_ARQUIVO', ENomeArquivo.Text );
    if length(ENomeArquivo.Text)=0 then Begin
      ArqIni.WriteString('BACKUP','NOME_ARQUIVO','siai');
      ENomeArquivo.Text := 'siai';
    end;
  Finally
  end;
  Try
    VarPedeSenha := ArqIni.ReadString('BACKUP','PEDESENHA', VarPedeSenha );
    if length(VarPedeSenha)=0 then Begin
      ArqIni.WriteString('BACKUP','PEDESENHA','NºO');
      VarPedeSenha := 'NºO';
    end;
  Finally
  end;
  Try
    XEHora.Text := ArqIni.ReadString('BACKUP','HORA', XEHora.Text );
    if length(XEHora.Text)=0 then Begin
      ArqIni.WriteString('BACKUP','HORA','00:00');
      XEHora.Text := '00:00';
    end;
  Finally
  end;
  Try
    XEHora2.Text := ArqIni.ReadString('BACKUP','HORA2', XEHora2.Text );
    if length(XEHora2.Text)=0 then Begin
      ArqIni.WriteString('BACKUP','HORA','00:00');
      XEHora2.Text := '00:00';
    end;
  Finally
  end;
  ArqIni.Free;
  CBPedesenha.ItemIndex := 0;
  if VarPedeSenha='NºO' Then
    CBPedesenha.ItemIndex := 1;
  ESenhaSQL.Text := decodifica(varsenha);
End;

procedure TFrmCopiaAutomatica.Geracopia;
Var
  comando, novonome, nomevelho : string;
begin
  LerIni;
  nomevelho := EDestino.Text+ENomeArquivo.Text+copy(datetostr(date),7,4)+ copy(datetostr(date),4,2)+ copy(datetostr(date),1,2)+copy(timetostr(time),1,2)+copy(timetostr(time),4,2)+'.sql';
  application.ProcessMessages;
  LMensagem.Caption := 'Iniciando geração da cópia de segurança....';
  application.ProcessMessages;  
  comando:='cmd.exe /c mysqldump.exe '+ ENomeTabela.Text+' -u '+EUsuarioSQL.Text+' -h'+EIPServidor.Text+' -p'+ESenhaSQL.Text+' --databases '+ENomeTabela.text +' > '+ nomevelho;
//  comando:='cmd.exe /c mysqldump.exe '+ ENomeTabela.Text+' -u '+EUsuarioSQL.Text+' -h'+EIPServidor.Text+' -p'+ESenhaSQL.Text+' --opt -v>'+nomevelho;
  if not createProcessSimple(comando) then
    LMensagem.Caption := 'BACKUP NºO FOI PROCESSADO!';
  application.ProcessMessages;  
  LMensagem.Caption := 'Gerei o '+nomevelho;
end;

procedure TFrmCopiaAutomatica.DXBGravarClick(Sender: TObject);
Var
  ArqIni : tIniFile;
begin
  ArqIni := tIniFile.Create('c:\MASTER\backup_Siai.Ini');
  ArqIni.WriteString('BACKUP','DESTINHO',EDestino.Text);
  ArqIni.WriteString('BACKUP','NOME_ARQUIVO',ENomeArquivo.Text);
  Geracopia;
  ArqIni.Free;
end;

procedure TFrmCopiaAutomatica.Timer1Timer(Sender: TObject);
begin
  LerIni;
  if (copy(timetostr(time),1,5) = copy(XEHora.Text,1,5)) or (copy(timetostr(time),1,5) = copy(XEHora2.Text,1,5)) then
    Geracopia;
End;

procedure TFrmCopiaAutomatica.XEHoraExit(Sender: TObject);
Var
  ArqIni : tIniFile;
begin
  ArqIni := tIniFile.Create('c:\MASTER\backup_Siai.Ini');
  ArqIni.WriteString('BACKUP','HORA',XEHora.EditText);
  ArqIni.Free;
end;

procedure TFrmCopiaAutomatica.XEHora2Exit(Sender: TObject);
Var
  ArqIni : tIniFile;
begin
  ArqIni := tIniFile.Create('c:\MASTER\backup_Siai.Ini');
  ArqIni.WriteString('BACKUP','HORA2',XEHora2.EditText);
  ArqIni.Free;
end;

procedure TFrmCopiaAutomatica.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmCopiaAutomatica.dxButton1Click(Sender: TObject);
var
  pat2, pt : string;
  comando, novonome, nomevelho : string;

begin
  LerIni;
  pt:=ExtractFilePath( Application.ExeName );
  Opendialog1.Execute;
  pat2:=Opendialog1.FileName;
  nomevelho := EDestino.Text+ENomeArquivo.Text+copy(datetostr(date),7,4)+ copy(datetostr(date),4,2)+ copy(datetostr(date),1,2)+copy(timetostr(time),1,2)+copy(timetostr(time),4,2)+'.sql';
  if not SIMNAO('Confirma restauração ?','N') Then
    Exit;
  pt:=ExtractFilePath( Application.ExeName );  // informa o diretório de onde o executavel esta...
  comando:='cmd.exe /c mysqladmin.exe -u '+EUsuarioSQL.Text+' -p'+ESenhaSQL.Text+' drop if exists '+ENomeTabela.Text;
  if not createProcessSimple(comando) then
     LMensagem.Caption := 'BASE DE DADOS NºO ATUALIZADA!  ';
  application.ProcessMessages;

  comando:='cmd.exe /c mysqladmin.exe -u'+EUsuarioSQL.Text+' -p'+ESenhaSQL.Text+' create '+ENomeTabela.Text;
  if not createProcessSimple(comando) then
     LMensagem.Caption := 'BASE DE DADOS NÃO CRIADA!  ';
  application.ProcessMessages;
  comando:='cmd.exe /c mysql.exe -u'+EUsuarioSQL.Text+' -p'+ESenhaSQL.Text+' --default-character-set=latin1 '+ENomeTabela.Text+' < '+pat2;
  application.ProcessMessages;
  LMensagem.Caption := 'BASE DE DADOS SENDO RESTAURADA!  ';
  application.ProcessMessages;
  if not createProcessSimple(comando) then
     LMensagem.Caption := 'BASE DE DADOS NºO RESTAURADA!  ';
  application.ProcessMessages;
  LMensagem.Caption := ' BACKUP FOI RESTAURADO! ';
  application.ProcessMessages;
end;


procedure TFrmCopiaAutomatica.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;


end.
