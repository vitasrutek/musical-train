unit Unit4;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, INIFiles;

type
  TForm4 = class(TForm)
    ListBox1: TListBox;
    procedure FormShow(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ListBox1MouseLeave(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;

implementation

{$R *.dfm}

uses Unit1;

procedure TForm4.FormCreate(Sender: TObject);
var
  Reg1,Reg2:  THandle;
begin
  Reg1 := CreateRoundRectRgn(10, 10, self.Width, self.Width, 10, 10);  //Create a round-corned sentence area at the form specified by the form
  Reg2 := CreateRectRgn(10, 100, self.Width, self.Width);     //Create a rectangular area at the location specified by the form
  CombineRgn(Reg1, Reg1, Reg2, RGN_OR);   //Merged in two regions of REG1, Reg2, and REG1 gathered in Reg1, REG2 collection and collections!
  SetwindowRgn(handle, Reg1, True);
end;

procedure TForm4.FormShow(Sender: TObject);
var i: integer;
begin
ListBox1.Clear;
for I := 0 to mainform.ComboBox1.Items.Count do
 ListBox1.Items.Add(mainform.ComboBox1.Items[i]);
top := screen.Height - height;
left := screen.Width - width;
end;

procedure TForm4.ListBox1Click(Sender: TObject);
begin
mainform.TrayIcon1DblClick(self);
mainform.ComboBox1.ItemIndex := ListBox1.ItemIndex;
mainform.ComboBox1.OnSelect(self);
end;

procedure TForm4.ListBox1MouseLeave(Sender: TObject);
begin
close;
end;

end.
