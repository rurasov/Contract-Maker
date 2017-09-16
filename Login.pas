unit Login;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  Vcl.ExtCtrls,
  System.Win.Registry, Winapi.ShellAPI, JvExControls,
  JvGIFCtrl, JvAnimatedImage;

type
  TAuthForm = class(TForm)
    AuthPnl: TPanel;
    lblLogin: TLabel;
    lblPassword: TLabel;
    edtLogin: TEdit;
    edtPassword: TEdit;
    btnAuth: TButton;
    chkShowPassword: TCheckBox;
    pnlWait: TPanel;
    lblStatus: TLabel;
    Gif: TJvGIFAnimator;
    procedure btnAuthClick(Sender: TObject);
    procedure edtPasswordKeyPress(Sender: TObject; var Key: Char);
    procedure chkShowPasswordClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    procedure Auth;
    procedure CheckInstallLocation();
    function IsInstalGDrive: boolean;
    { Private declarations }
  public
    { Public declarations }
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  end;

var
  AuthForm: TAuthForm;
  InstallLocation: string;

implementation

uses
  Main;
{$R *.dfm}

//icon on the taskbar
procedure TAuthForm.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle or WS_EX_APPWINDOW;
  Params.WndParent := GetDesktopWindow;
end;

procedure TAuthForm.CheckInstallLocation;
var
  Reg: TRegistry;
  url: string;
begin
  if not IsInstalGDrive then
    case Application.MessageBox('Приложение Google Drive не установлено. ' +
      ''#13''#10'' + 'Перейти на страницу загрузки Google Drive?',
      'Для продолжения, установите Google Drive', MB_YESNO + MB_ICONSTOP +
      MB_TOPMOST) of
      IDYES:
        begin
          url := 'https://tools.google.com/dlpage/drive/index.html?hl=ru';
          ShellExecute(Handle, 'open', PWideChar(url), nil, nil, SW_NORMAL);
          Abort;
        end;
    end;
end;

function TAuthForm.IsInstalGDrive: boolean;
var
  Reg: TRegistry;
begin
  Reg := TRegIniFile.Create;
  Reg.RootKey := HKEY_LOCAL_MACHINE;
  if not Reg.OpenKey('Software\Google\Drive', false) then
  begin
    Reg.RootKey := HKEY_CURRENT_USER;
    if not Reg.OpenKey('Software\Google\Drive', false) then
      Result := false;
    FreeAndNil(Reg);
    Exit;
  end;
  Result := true;
  FreeAndNil(Reg);
end;


procedure TAuthForm.Auth;
var
  s: string;
begin
  lblStatus.Caption := ('Авторизация...');
  Screen.Cursor := crHourGlass;
  MainForm.idhtp1.Request.Username := edtLogin.Text;
  MainForm.idhtp1.Request.Password := edtPassword.Text;
  try
    s := MainForm.idhtp1.Get('')
  except
    on E: Exception do
    begin
      ShowMessage(E.Message);
    end;
  end;
  if MainForm.idhtp1.ResponseCode = 200 then
  begin
    lblStatus.Caption := ('Авторизация успешна.');
    list.Add('' + edtLogin.Text);
    list.Add('' + edtPassword.Text);
    MainForm.idhtp1.Post
      ('', list);
    list.Clear;
    ModalResult := mrOk;
    MainForm.Visible := true;
  end;
  Gif.Animate := false;
  pnlWait.SendToBack;
  Screen.Cursor := crDefault;
end;

procedure TAuthForm.btnAuthClick(Sender: TObject);
begin
  if Length(edtLogin.Text) < 1 then
  begin
    ShowMessage('Для продолжения введите логин');
    edtLogin.SetFocus;
    Exit;
  end;
  if Length(edtPassword.Text) < 1 then
  begin
    ShowMessage('Для продолжения введите пароль');
    edtPassword.SetFocus;
    Exit;
  end;
  AuthPnl.SendToBack;
  Gif.Animate := true;
  Auth;
end;

procedure TAuthForm.chkShowPasswordClick(Sender: TObject);
begin
  if chkShowPassword.Checked then
    edtPassword.PasswordChar := #0
  else
    edtPassword.PasswordChar := '*';
end;

procedure TAuthForm.edtPasswordKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    btnAuth.Click;
end;

procedure TAuthForm.FormCreate(Sender: TObject);
begin
  CheckInstallLocation();

end;

end.
