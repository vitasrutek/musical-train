unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, INIFiles, ActiveX, ComObj;

type
  Tnovyform = class(TForm)
    GroupBox2: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    GroupBox3: TGroupBox;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    Button1: TButton;
    Button3: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button3Click(Sender: TObject);
    procedure Edit2Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private

    { Private declarations }
  public
    procedure Wedit(soubor, co, kam: String);
    procedure Eedit(soubor, co: String; k1, k2: integer);

    { Public declarations }
  end;

var
  novyform: Tnovyform;
   cesta: string;

const
  wdFindContinue = 1;
  wdReplaceOne = 1;
  wdReplaceAll = 2;

implementation

{$R *.dfm}

uses Unit3, Unit1;

procedure Tnovyform.Button3Click(Sender: TObject);
begin
Edit1.Text := '';
Edit2.Text := '';
Edit3.Text := '';
Edit4.Text := '';
Edit5.Text := '';
novyform.Close;
end;

procedure Tnovyform.FormClose(Sender: TObject; var Action: TCloseAction);
begin
workform.Close;
end;

procedure Tnovyform.FormShow(Sender: TObject);
begin
Edit1.Text := '';
Edit2.Text := '';
Edit3.Text := '';
Edit4.Text := '';
Edit5.Text := '';
end;

procedure Tnovyform.Wedit(soubor, co, kam: String);
var
WordApp: Variant;
begin
WordApp := CreateOLEObject('Word.Application');
WordApp.Documents.Open(cesta + '\' + soubor);
WordApp.Selection.Find.Text := co;
WordApp.Selection.Find.Replacement.Text := kam;
WordApp.Selection.Find.Forward := True;
WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
WordApp.ActiveDocument.SaveAs(cesta + '\' + soubor);
WordApp.Quit;
WordApp := Unassigned;
end;

procedure Tnovyform.Edit2Change(Sender: TObject);
var i:integer;
b: Boolean;
NewChar: string;
begin
b := false;
  for i := 1 to length(edit2.Text) do
    begin
    if (Edit2.Text[i] = '.') or (Edit2.Text[i] = '/') then
        b:= true;
    end;
    if b = true then
      begin
        NewChar := StringReplace( Edit2.Text, '.', ' ' , [rfReplaceAll, rfIgnoreCase]);
        edit2.Text := Newchar;
        NewChar := StringReplace( Edit2.Text, '/', '-' , [rfReplaceAll, rfIgnoreCase]);
        edit2.Text := Newchar;
        ShowMessage('Byla provedena úprava textu - změna "." a "/".');
        b := false;
      end;
end;

procedure Tnovyform.Eedit(soubor, co: String; k1, k2: integer);
var xls, xlw: Variant;
begin
xls := CreateOLEObject('Excel.Application');
xlw := xls.WorkBooks.Open(cesta + '\' + soubor);
xls.ActiveSheet.Cells[k1, k2].Value := co;
xlw.Save;
xlw.Close;
xlw := UnAssigned;
xls.Quit;
xls := UnAssigned;
end;

procedure Tnovyform.Button1Click(Sender: TObject);
begin
           begin
            workform.show;
            workform.Label2.Caption := 'Vytvářím adresář....';
            cesta := (mainform.Path + '\' + Edit1.Text + ' ' + Edit2.Text);
            CreateDir(cesta);
            CreateDir (cesta + '\ZL');
            CreateDir (cesta + '\DGN');
            CreateDir (cesta + '\RDF');
            CreateDir (cesta + '\VYKESY');
            CreateDir (cesta + '\GEO');
            workform.Label2.Caption := 'Kopíruji položky....';
            if CheckBox1.Checked = true then
              begin
                CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\textova cast - DUR.docx'), PChar(cesta + '\' + 'textova cast - DUR.docx'), False);
              end;
            if CheckBox2.Checked = true then
              begin
                CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\textova cast - DPS.docx'), PChar(cesta + '\' + 'textova cast - DPS.docx'), False);
              end;
            workform.Label2.Caption := 'Přepisuji dokumenty DUR....';
            if CheckBox1.Checked = true then
              begin
               Wedit('textova cast - DUR.docx', '_CISLO', Edit1.Text);
               Wedit('textova cast - DUR.docx', '_NAZEV', Edit2.Text);
               Wedit('textova cast - DUR.docx', '_OBEC', Edit3.Text);
               Wedit('textova cast - DUR.docx', '_KATASTR', Edit4.Text);
              end;
            workform.Label2.Caption := 'Přepisuji dokumenty DPS....';
            if CheckBox2.Checked = true then
              begin
               Wedit('textova cast - DPS.docx', '_CISLO', Edit1.Text);
               Wedit('textova cast - DPS.docx', '_NAZEV', Edit2.Text);
               Wedit('textova cast - DPS.docx', '_OBEC', Edit3.Text);
               Wedit('textova cast - DPS.docx', '_KATASTR', Edit4.Text);
              end;
             workform.Close;
            mainform.ComboBox1.Update;
            novyform.close;
          end;
end;
end.
