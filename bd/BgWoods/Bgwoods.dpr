program Bgwoods;

uses
  Forms,
  bgMain in 'bgMain.pas' {FrmBg};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TFrmBg, FrmBg);
  Application.Run;
end.
