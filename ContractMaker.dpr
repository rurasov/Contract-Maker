program ContractMaker;

{$R *.dres}

uses
  Vcl.Forms,
  Main in 'Main.pas' {MainForm},
  Login in 'Login.pas' {AuthForm};

{$R *.res}

begin
  ReportMemoryLeaksOnShutdown := True;
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.ShowMainForm:=False;
  Application.Run;
end.
