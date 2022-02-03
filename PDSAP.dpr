program PDSAP;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {mainform},
  Unit2 in 'Unit2.pas' {novyform},
  Unit3 in 'Unit3.pas' {workform},
  Vcl.Themes,
  Vcl.Styles,
  Unit4 in 'Unit4.pas' {Form4};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'PD SAP';
  Application.CreateForm(Tmainform, mainform);
  Application.CreateForm(Tnovyform, novyform);
  Application.CreateForm(Tworkform, workform);
  Application.CreateForm(TForm4, Form4);
  Application.Run;
end.
