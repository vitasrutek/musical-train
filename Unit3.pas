unit Unit3;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.WinXCtrls, Vcl.StdCtrls,
  Vcl.ExtCtrls;

type
  Tworkform = class(TForm)
    Panel1: TPanel;
    Label2: TLabel;
    Label1: TLabel;
    ActivityIndicator1: TActivityIndicator;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  workform: Tworkform;

implementation

{$R *.dfm}

uses Unit1;

procedure Tworkform.Button1Click(Sender: TObject);
begin
workform.Close;
end;

end.
