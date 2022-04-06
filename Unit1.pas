unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Inifiles, Vcl.StdCtrls, ShellApi, Comobj,
  Vcl.ExtDlgs, excel2000, Vcl.OleServer, Excel2010, Vcl.ExtCtrls, Vcl.AppEvnts,
  ExcelXP, System.Win.TaskbarCore, Vcl.Taskbar, Vcl.Menus, Vcl.Buttons, OleCtrls, System.Zip;

type
  Tmainform = class(TForm)
    GroupBox1: TGroupBox;
    ComboBox1: TComboBox;
    GroupBox2: TGroupBox;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    GroupBox3: TGroupBox;
    Button1: TButton;
    Button2: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    Button11: TButton;
    Button12: TButton;
    Button13: TButton;
    FileOpenDialog1_old: TFileOpenDialog;
    Button16: TButton;
    Label5: TLabel;
    ExcelApplication1: TExcelApplication;
    TrayIcon1: TTrayIcon;
    ApplicationEvents1: TApplicationEvents;
    Label6: TLabel;
    Edit5: TEdit;
    FileOpenDialog1: TOpenDialog;
    Button3: TButton;
    Label7: TLabel;
    Edit6: TEdit;
    Taskbar1: TTaskbar;
    Button14: TButton;
    Button15: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    Konec1: TMenuItem;
    Zobrazitaplikaci1: TMenuItem;
    Zobrazitseznamstaveb1: TMenuItem;
    BitBtn1: TBitBtn;
    Label8: TLabel;
    Edit7: TEdit;
    Timer1: TTimer;
    Timer2: TTimer;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Edit9: TEdit;
    Button17: TButton;
    Button18: TButton;
    GroupBox4: TGroupBox;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn5: TBitBtn;
    BitBtn6: TBitBtn;
    BitBtn7: TBitBtn;
    BitBtn8: TBitBtn;
    BitBtn9: TBitBtn;
    BitBtn10: TBitBtn;
    BitBtn11: TBitBtn;
    BitBtn12: TBitBtn;
    CheckBox1: TCheckBox;
    Label13: TLabel;
    Edit8: TEdit;
    CheckBox2: TCheckBox;
    Label14: TLabel;
    Button7: TButton;
    procedure FormCreate(Sender: TObject);
    procedure ComboBox1Select(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure CopyFolder(CopyFrom, CopyTo: String);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Clic(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button15Click(Sender: TObject);
    procedure Button16Click(Sender: TObject);
    procedure Edit1Click(Sender: TObject);
    procedure Edit2Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure TrayIcon1DblClick(Sender: TObject);
    procedure ApplicationEvents1Minimize(Sender: TObject);
    procedure Edit5Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure TrayIcon1Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure Konec1Click(Sender: TObject);
    procedure Zobrazitaplikaci1Click(Sender: TObject);
    procedure Zobrazitseznamstaveb1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure Edit7Change(Sender: TObject);
    procedure Button18Click(Sender: TObject);
    procedure Button17Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn9Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure BitBtn10Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn11Click(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure BitBtn12Click(Sender: TObject);

  private
    { Private declarations }
    LCID: Cardinal;                     // XLS
    ExcelApp: TExcelApplication;        // XLS
    myxlBook: TExcelWorkbook;           // XLS
  public
    Path: String;
    PDsoubor: string;
    procedure Comboupdate;
    procedure Xedit(soubor, co: String; k1, k2: integer);
    procedure Xedit2(soubor, co: String; k1, k2: integer);
    procedure OpenAndModifyAnExistingExcelFileAndSaveAs(cislo: string);
    function RemoveDiacritics(const S:string):string;
    { Public declarations }
  end;

var
  mainform: Tmainform;
  WordApp: Variant;
  xls, xlw: Variant;
  Excl, Wrkbk: OleVariant;
  searchResult : TSearchRec;

implementation

{$R *.dfm}

uses Unit2, Unit3, Unit4, Unit5;

procedure Tmainform.ApplicationEvents1Minimize(Sender: TObject);
begin
Hide();
WindowState := wsMinimized;
TrayIcon1.Visible := True;
TrayIcon1.Animate := True;
TrayIcon1.ShowBalloonHint;
end;

procedure Tmainform.BitBtn10Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\Formular pro SJZ.xlsx')) then
ShellExecute(Application.Handle, 'open', Pchar(cesta + '\Formular pro SJZ.xlsx'), nil, nil, SW_SHOWNORMAL)
else begin
  CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\Formular pro SJZ.xlsx'), PChar(cesta + '\' + 'Formular pro SJZ.xlsx'), False);
  BitBtn10.Kind := bkYes;
  BitBtn10.Caption := 'SJZ';
end;
end;

procedure Tmainform.BitBtn11Click(Sender: TObject);
begin
workform.show;
workform.Label2.Caption := 'Připravuji NN Geoportal....';
CopyFolder ((extractfilepath(Application.exename)+'\VZORY\_CD'), (cesta + '\_CD ' + edit1.text ));
workform.Close;
BitBtn11.Kind := bkYes;
BitBtn11.Caption := 'CD';
end;

procedure Tmainform.BitBtn12Click(Sender: TObject);
begin
workform.show;
workform.Label2.Caption := 'Připravuji VN Geoportal....';
CopyFolder ((extractfilepath(Application.exename)+'\VZORY\DOK_PDVN_x'), (cesta + '\' + 'DOK_PDVN_' + edit1.text ));
    begin
    FileOpenDialog1.Filename := '';
    FileOpenDialog1.InitialDir := (cesta + '\');
    FileOpenDialog1.Filter := 'PDF|textova*.pdf';
    FileOpenDialog1.Title := 'Vyber textovou zprávu - PDF';
    if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\1_ZPRAVA\TECHNICKA_ZPRAVA_' + Edit1.Text + '_t.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF|*situace*.pdf';
      FileOpenDialog1.Title := 'Vyber situaci - PDF';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\2_SITUACE\SITUACE_' + Edit1.Text + '_s.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF|*schema*.pdf';
      FileOpenDialog1.Title := 'Vyber jednopólové schéma zapojení - PDF';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\3_JEDNOPOLOVE_SCHEMA\JEDNOPOLOVE_SCHEMA_' + Edit1.Text + '_j.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'XLSX|pov*.xlsx';
      FileOpenDialog1.Title := 'Vyber POV - XLSX';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\4_POV\POV_' + Edit1.Text + '_p.xlsx'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'XLS|'+edit1.text+'*.xls';
      FileOpenDialog1.Title := 'Vyber rozpočet - XLS';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\5_ROZPOCET\ROZPOCET_STAVBY_' + Edit1.Text + '_r.xls'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF - plan*|plan*.pdf';
      FileOpenDialog1.Title := 'Vyber BOZP';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\7_SOUVISEJICI_DOKUMENTY\BOZP_' + Edit1.Text + '.pdf'), False);
        end;
    end;

TZipFile.ZipDirectoryContents(cesta + '\DOK_PDVN_' + edit1.text + '.zip', cesta + '\DOK_PDVN_' + edit1.text);
BitBtn12.Kind := bkYes;
BitBtn12.Caption := 'VN';
workform.Close;
end;

procedure Tmainform.BitBtn1Click(Sender: TObject);
begin
Comboupdate;
ComboBox1.ItemIndex := 1;
ComboBox1.OnSelect(self);
end;

procedure Tmainform.BitBtn2Click(Sender: TObject);
begin
if (edit3.Text = '') or (Edit4.Text = '') then
  Application.MessageBox('Doplň obec a katastrální území', 'Chyba', MB_ICONERROR)
  else begin
    if FileExists (PChar(cesta + '\textova cast - DUR.docx')) then
      ShellExecute(Application.Handle, 'open', Pchar(cesta + '\textova cast - DUR.docx'), nil, nil, SW_SHOWNORMAL)
        else begin
          workform.show;
          CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\textova cast - DUR.docx'), PChar(cesta + '\' + 'textova cast - DUR.docx'), False);
          workform.Label2.Caption := 'Přepisuji dokument DUR....';
          WordApp := CreateOLEObject('Word.Application');
          WordApp.Documents.Open(cesta + '\textova cast - DUR.docx');
          WordApp.Selection.Find.Text := '_CISLO';
          WordApp.Selection.Find.Replacement.Text := Edit1.Text;
          WordApp.Selection.Find.Forward := True;
          WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
          WordApp.Selection.Find.Text := '_NAZEV';
          WordApp.Selection.Find.Replacement.Text := Edit2.Text;
          WordApp.Selection.Find.Forward := True;
          WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
          WordApp.Selection.Find.Text := '_OBEC';
          WordApp.Selection.Find.Replacement.Text := Edit3.Text;
          WordApp.Selection.Find.Forward := True;
          WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
          WordApp.Selection.Find.Text := '_KATASTR';
          WordApp.Selection.Find.Replacement.Text := Edit4.Text;
          WordApp.Selection.Find.Forward := True;
          WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
          WordApp.ActiveDocument.SaveAs(cesta + '\textova cast - DUR.docx');
          WordApp.Quit;
          WordApp := Unassigned;
          BitBtn2.Kind := bkYes;
          BitBtn2.Caption := 'DUR';
          workform.close;
        end;
  end;
end;

procedure Tmainform.BitBtn3Click(Sender: TObject);
begin
if (edit3.Text = '') or (Edit4.Text = '') then
  Application.MessageBox('Doplň obec a katastrální území', 'Chyba', MB_ICONERROR)
  else begin
    if FileExists (PChar(cesta + '\textova cast - DPS.docx')) then
      ShellExecute(Application.Handle, 'open', Pchar(cesta + '\textova cast - DPS.docx'), nil, nil, SW_SHOWNORMAL)
else begin
        workform.show;
        CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\textova cast - DPS.docx'), PChar(cesta + '\' + 'textova cast - DPS.docx'), False);
          workform.Label2.Caption := 'Přepisuji dokument DPS....';
        WordApp := CreateOLEObject('Word.Application');
        WordApp.Documents.Open(cesta + '\textova cast - DPS.docx');
        WordApp.Selection.Find.Text := '_CISLO';
        WordApp.Selection.Find.Replacement.Text := Edit1.Text;
        WordApp.Selection.Find.Forward := True;
        WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
        WordApp.Selection.Find.Text := '_NAZEV';
        WordApp.Selection.Find.Replacement.Text := Edit2.Text;
        WordApp.Selection.Find.Forward := True;
        WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
        WordApp.Selection.Find.Text := '_OBEC';
        WordApp.Selection.Find.Replacement.Text := Edit3.Text;
        WordApp.Selection.Find.Forward := True;
        WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
        WordApp.Selection.Find.Text := '_KATASTR';
        WordApp.Selection.Find.Replacement.Text := Edit4.Text;
        WordApp.Selection.Find.Forward := True;
        WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
        WordApp.ActiveDocument.SaveAs(cesta + '\textova cast - DPS.docx');
        WordApp.Quit;
        WordApp := Unassigned;
        BitBtn3.Kind := bkYes;
        BitBtn3.Caption := 'DPS';
        workform.close;
      end;
end;
end;

procedure Tmainform.BitBtn4Click(Sender: TObject);
begin
if (edit3.Text = '') or (Edit4.Text = '') then
  Application.MessageBox('Doplň obec a katastrální území', 'Chyba', MB_ICONERROR)
    else begin
      if FileExists (PChar(cesta + '\textova cast - DUR - bourani.docx')) then
        ShellExecute(Application.Handle, 'open', Pchar(cesta + '\textova cast - DUR - bourani.docx'), nil, nil, SW_SHOWNORMAL)
          else begin
            workform.show;
            CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\textova cast - DUR - bourani.docx'), PChar(cesta + '\' + 'textova cast - DUR - bourani.docx'), False);
            workform.Label2.Caption := 'Přepisuji dokument DUR....';
            WordApp := CreateOLEObject('Word.Application');
            WordApp.Documents.Open(cesta + '\textova cast - DUR - bourani.docx');
            WordApp.Selection.Find.Text := '_CISLO';
            WordApp.Selection.Find.Replacement.Text := Edit1.Text;
            WordApp.Selection.Find.Forward := True;
            WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
            WordApp.Selection.Find.Text := '_NAZEV';
            WordApp.Selection.Find.Replacement.Text := Edit2.Text;
            WordApp.Selection.Find.Forward := True;
            WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
            WordApp.Selection.Find.Text := '_OBEC';
            WordApp.Selection.Find.Replacement.Text := Edit3.Text;
            WordApp.Selection.Find.Forward := True;
            WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
            WordApp.Selection.Find.Text := '_KATASTR';
            WordApp.Selection.Find.Replacement.Text := Edit4.Text;
            WordApp.Selection.Find.Forward := True;
            WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
            WordApp.ActiveDocument.SaveAs(cesta + '\textova cast - DUR - bourani.docx');
            WordApp.Quit;
            WordApp := Unassigned;
            BitBtn4.Kind := bkYes;
            BitBtn4.Caption := 'DUR odstr';
            workform.close;
          end;
    end;
end;

procedure Tmainform.BitBtn5Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\POV.xlsx')) then
  ShellExecute(Application.Handle, 'open', Pchar(cesta + '\POV.xlsx'), nil, nil, SW_SHOWNORMAL)
else begin
  workform.show;
  CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\POV.XLSX'), PChar(cesta + '\POV.xlsx'), False);
  workform.Label2.Caption := 'Přepisuji dokument POV....';
  Xedit2((cesta + '\POV.xlsx'),Edit1.Text, 1, 1);
  Xedit2((cesta + '\POV.xlsx'),Edit2.Text, 1, 4);
  workform.close;
  BitBtn5.Kind := bkYes;
  BitBtn5.Caption := 'POV';
end;
end;

procedure Tmainform.BitBtn6Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\Desky.xlsx')) then
  ShellExecute(Application.Handle, 'open', Pchar(cesta + '\Desky.xlsx'), nil, nil, SW_SHOWNORMAL)
    else begin
      workform.show;
      CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\Desky.xlsx'), PChar(cesta + '\' + 'Desky.xlsx'), False);
      workform.Label2.Caption := 'Přepisuji dokument Desky....';
      Xedit2((cesta + '\Desky.xlsx'),Edit1.Text, 4, 3); //cislo
      Xedit2((cesta + '\Desky.xlsx'),Edit2.Text, 3, 3); //nazev
      BitBtn6.Kind := bkYes;
      BitBtn6.Caption := 'Desky';
      workform.close;
    end;
end;

procedure Tmainform.BitBtn7Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\uvodni listy.docx')) then
  ShellExecute(Application.Handle, 'open', Pchar(cesta + '\uvodni listy.docx'), nil, nil, SW_SHOWNORMAL)
    else begin
      workform.show;
      CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\uvodni listy.docx'), PChar(cesta + '\' + 'uvodni listy.docx'), False);
      workform.Label2.Caption := 'Přepisuji dokument Listy....';
      WordApp := CreateOLEObject('Word.Application');
      WordApp.Documents.Open(cesta + '\uvodni listy.docx');
      WordApp.Selection.Find.Text := '_CISLO';
      WordApp.Selection.Find.Replacement.Text := Edit1.Text;
      WordApp.Selection.Find.Forward := True;
      WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
      WordApp.Selection.Find.Text := '_NAZEV';
      WordApp.Selection.Find.Replacement.Text := Edit2.Text;
      WordApp.Selection.Find.Forward := True;
      WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
      WordApp.ActiveDocument.SaveAs(cesta + '\uvodni listy.docx');
      WordApp.Quit;
      WordApp := Unassigned;
      BitBtn7.Kind := bkYes;
      BitBtn7.Caption := 'Listy';
      workform.close;
    end;
end;

procedure Tmainform.BitBtn8Click(Sender: TObject);
begin
workform.show;
workform.Label2.Caption := 'Připravuji NN Geoportal....';
  CopyFolder ((extractfilepath(Application.exename)+'\VZORY\DOK_PDNN_x'), (cesta + '\DOK_PDNN_' + edit1.text ));
    begin
    FileOpenDialog1.InitialDir := (cesta + '\');
    FileOpenDialog1.Filename := '';
    FileOpenDialog1.Filter := 'PDF - textova*.pdf|textova*.pdf';
    FileOpenDialog1.Title := 'Vyber textovou zprávu - PDF';
    if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\1_ZPRAVA\TECHNICKA_ZPRAVA_' + Edit1.Text + '_t.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF - *situace*.pdf|*situace*.pdf';
      FileOpenDialog1.Title := 'Vyber situaci - PDF';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\2_SITUACE\SITUACE_' + Edit1.Text + '_s.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF - *schema*.pdf|*schema*.pdf';
      FileOpenDialog1.Title := 'Vyber jednopólové schéma zapojení - PDF';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\3_JEDNOPOLOVE_SCHEMA\JEDNOPOLOVE_SCHEMA_' + Edit1.Text + '_j.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'XLSX|pov*.xlsx';
      FileOpenDialog1.Title := 'Vyber POV - XLSX';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\4_POV\POV_' + Edit1.Text + '_p.xlsx'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'XLS|'+edit1.text+'*.xls';
      FileOpenDialog1.Title := 'Vyber rozpočet - XLS';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\5_ROZPOCET\ROZPOCET_STAVBY_' + Edit1.Text + '_r.xls'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF - plan*|plan*.pdf';
      FileOpenDialog1.Title := 'Vyber BOZP';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\7_SOUVISEJICI_DOKUMENTY\BOZP_' + Edit1.Text + '.pdf'), False);
        end;
    end;

  TZipFile.ZipDirectoryContents(cesta + '\DOK_PDNN_' + edit1.text + '.zip', cesta + '\DOK_PDNN_' + edit1.text);
  BitBtn8.Kind := bkYes;
  BitBtn8.Caption := 'NN';
  workform.Close;
end;

procedure Tmainform.BitBtn9Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\Odpady.xlsx')) then
ShellExecute(Application.Handle, 'open', Pchar(cesta + '\Odpady.xlsx'), nil, nil, SW_SHOWNORMAL)
else begin
      workform.show;
      CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\Odpady.xlsx'), PChar(cesta + '\' + 'Odpady.xlsx'), False);
      workform.Label2.Caption := 'Přepisuji dokument Odpady....';
      Xedit2((cesta + '\Odpady.xlsx'),Edit1.Text, 4, 2);
      Xedit2((cesta + '\Odpady.xlsx'),Edit2.Text, 3, 2);
      BitBtn9.Kind := bkYes;
      BitBtn9.Caption := 'Odpady';
      workform.close;
     end;
end;

procedure Tmainform.Button10Clic(Sender: TObject);
begin
workform.show;
workform.Label2.Caption := 'Připravuji VN Geoportal....';
CopyFolder ((extractfilepath(Application.exename)+'\VZORY\DOK_PDVN_x'), (cesta + '\' + 'DOK_PDVN_' + edit1.text ));
    begin
    FileOpenDialog1.Filename := '';
    FileOpenDialog1.InitialDir := (cesta + '\');
    FileOpenDialog1.Filter := 'PDF|textova*.pdf';
    FileOpenDialog1.Title := 'Vyber textovou zprávu - PDF';
    if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\1_ZPRAVA\TECHNICKA_ZPRAVA_' + Edit1.Text + '_t.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF|*situace*.pdf';
      FileOpenDialog1.Title := 'Vyber situaci - PDF';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\2_SITUACE\SITUACE_' + Edit1.Text + '_s.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF|*schema*.pdf';
      FileOpenDialog1.Title := 'Vyber jednopólové schéma zapojení - PDF';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\3_JEDNOPOLOVE_SCHEMA\JEDNOPOLOVE_SCHEMA_' + Edit1.Text + '_j.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'XLSX|pov*.xlsx';
      FileOpenDialog1.Title := 'Vyber POV - XLSX';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\4_POV\POV_' + Edit1.Text + '_p.xlsx'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'XLS|'+edit1.text+'*.xls';
      FileOpenDialog1.Title := 'Vyber rozpočet - XLS';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\5_ROZPOCET\ROZPOCET_STAVBY_' + Edit1.Text + '_r.xls'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF - plan*|plan*.pdf';
      FileOpenDialog1.Title := 'Vyber BOZP';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDVN_' + edit1.text + '\7_SOUVISEJICI_DOKUMENTY\BOZP_' + Edit1.Text + '.pdf'), False);
        end;
    end;
workform.Close;
end;

procedure Tmainform.Button11Click(Sender: TObject);
begin
workform.show;
workform.Label2.Caption := 'Připravuji NN Geoportal....';
CopyFolder ((extractfilepath(Application.exename)+'\VZORY\_CD'), (cesta + '\_CD ' + edit1.text ));
workform.Close;
end;

procedure Tmainform.Button12Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\Desky.xlsx')) then
  Application.MessageBox('Soubor již existuje..', 'Chyba', MB_ICONERROR)
    else begin
      workform.show;
      CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\Desky.xlsx'), PChar(cesta + '\' + 'Desky.xlsx'), False);
      workform.Label2.Caption := 'Přepisuji dokument Desky....';
      Xedit2((cesta + '\Desky.xlsx'),Edit1.Text, 4, 3); //cislo
      Xedit2((cesta + '\Desky.xlsx'),Edit2.Text, 3, 3); //nazev
      workform.close;
    end;
end;

procedure Tmainform.Button13Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\uvodni listy.docx')) then
  Application.MessageBox('Soubor již existuje..', 'Chyba', MB_ICONERROR)
    else begin
      workform.show;
      CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\uvodni listy.docx'), PChar(cesta + '\' + 'uvodni listy.docx'), False);
      workform.Label2.Caption := 'Přepisuji dokument Listy....';
      WordApp := CreateOLEObject('Word.Application');
      WordApp.Documents.Open(cesta + '\uvodni listy.docx');
      WordApp.Selection.Find.Text := '_CISLO';
      WordApp.Selection.Find.Replacement.Text := Edit1.Text;
      WordApp.Selection.Find.Forward := True;
      WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
      WordApp.Selection.Find.Text := '_NAZEV';
      WordApp.Selection.Find.Replacement.Text := Edit2.Text;
      WordApp.Selection.Find.Forward := True;
      WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
      WordApp.ActiveDocument.SaveAs(cesta + '\uvodni listy.docx');
      WordApp.Quit;
      WordApp := Unassigned;
      workform.close;
    end;
end;

procedure Tmainform.Button14Click(Sender: TObject);
begin
       ShellExecute(Application.Handle, 'open', Pchar(PDsoubor), nil, nil, SW_SHOWNORMAL);
end;

procedure Tmainform.Button15Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\Odpady.xlsx')) then
Application.MessageBox('Soubor již existuje..', 'Chyba', MB_ICONERROR)
else begin
      workform.show;
      CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\Odpady.xlsx'), PChar(cesta + '\' + 'Odpady.xlsx'), False);
      workform.Label2.Caption := 'Přepisuji dokument Odpady....';
      Xedit2((cesta + '\Odpady.xlsx'),Edit1.Text, 4, 2);
      Xedit2((cesta + '\Odpady.xlsx'),Edit2.Text, 3, 2);
      workform.close;
     end;
end;

procedure Tmainform.Button16Click(Sender: TObject);
begin
if (edit3.Text = '') or (Edit4.Text = '') then
  Application.MessageBox('Doplň obec a katastrální území', 'Chyba', MB_ICONERROR)
    else begin
      if FileExists (PChar(cesta + '\textova cast - DUR - bourani.docx')) then
        Application.MessageBox('Soubor již existuje..', 'Chyba', MB_ICONERROR)
          else begin
            workform.show;
            CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\textova cast - DUR - bourani.docx'), PChar(cesta + '\' + 'textova cast - DUR - bourani.docx'), False);
            workform.Label2.Caption := 'Přepisuji dokument DUR....';
            WordApp := CreateOLEObject('Word.Application');
            WordApp.Documents.Open(cesta + '\textova cast - DUR - bourani.docx');
            WordApp.Selection.Find.Text := '_CISLO';
            WordApp.Selection.Find.Replacement.Text := Edit1.Text;
            WordApp.Selection.Find.Forward := True;
            WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
            WordApp.Selection.Find.Text := '_NAZEV';
            WordApp.Selection.Find.Replacement.Text := Edit2.Text;
            WordApp.Selection.Find.Forward := True;
            WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
            WordApp.Selection.Find.Text := '_OBEC';
            WordApp.Selection.Find.Replacement.Text := Edit3.Text;
            WordApp.Selection.Find.Forward := True;
            WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
            WordApp.Selection.Find.Text := '_KATASTR';
            WordApp.Selection.Find.Replacement.Text := Edit4.Text;
            WordApp.Selection.Find.Forward := True;
            WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
            WordApp.ActiveDocument.SaveAs(cesta + '\textova cast - DUR - bourani.docx');
            WordApp.Quit;
            WordApp := Unassigned;
              workform.close;
          end;
    end;
end;

procedure Tmainform.Button17Click(Sender: TObject);
const
  olMailItem = 0;
var
  Outlook: OLEVariant;
  MailItem: Variant;
  MailInspector : Variant;
//  stringlist : TStringList;

  myline: string;
  separatorpos: integer;
  partfirst: string;
  partsecond: string;
begin
  myline:=RemoveDiacritics(Edit9.text);             //odstranit diakritiku od technika
  if myline.Contains(' ') then
  begin
    separatorpos:=pos(' ',myline);
    partfirst:=trim(copy(myline,1,separatorpos-1));          // prijmeni
    partsecond:=trim(copy(myline,separatorpos+1));           // jmeno
  end;

  try
   Outlook:=GetActiveOleObject('Outlook.Application') ;
  except
   Outlook:=CreateOleObject('Outlook.Application') ;
  end;
  try
  //  Stringlist := TStringList.Create;
    MailItem := Outlook.CreateItem(olMailItem) ;
    MailItem.Subject := Edit5.Text;
    MailItem.Recipients.Add(partsecond + '.' + partfirst + '@cezdistribuce.cz');
    //MailItem.Attachments.Add('c:\FILE1.txt');
    //Stringlist := TStringList.Create;
    //StringList.Add('body here');
    //MailItem.Body := 'ahoj';
   // MailItem.Send;   //SENDS A MAIL WITH OUT OUTLOOK WINDOW. USE "SAVE" FOR DRAFT

  // TO SHOW OUTLOOK DIALOG. BUT YOU HAVE SET MAILITEM.SEND AS COMMENT
    MailInspector := MailItem.GetInspector;
    MailInspector.display(false); //true means modal
//    MailInspector.Send;

  finally
    Outlook := Unassigned;
  //  StringList.Free;
  end;


end;

procedure Tmainform.Button18Click(Sender: TObject);
const
  olMailItem = 0;
var
  Outlook: OLEVariant;
  MailItem: Variant;
  MailInspector : Variant;
//  stringlist : TStringList;
begin
 try
   Outlook:=GetActiveOleObject('Outlook.Application') ;
  except
   Outlook:=CreateOleObject('Outlook.Application') ;
  end;
  try
  //  Stringlist := TStringList.Create;
    MailItem := Outlook.CreateItem(olMailItem) ;
    MailItem.Subject := Edit5.Text;
    if Edit7.text = 'Petra' then MailItem.Recipients.Add('michkova@elmos.cz');
    if Edit7.text = 'Verča' then MailItem.Recipients.Add('cabanova@elmos.cz');
    if Edit7.text = 'Sylva' then MailItem.Recipients.Add('cicvarkova@elmos.cz');
    //MailItem.Attachments.Add('c:\FILE1.txt');
    //Stringlist := TStringList.Create;
    //StringList.Add('body here');
    //MailItem.Body := 'ahoj';
   // MailItem.Send;   //SENDS A MAIL WITH OUT OUTLOOK WINDOW. USE "SAVE" FOR DRAFT

  // TO SHOW OUTLOOK DIALOG. BUT YOU HAVE SET MAILITEM.SEND AS COMMENT
    MailInspector := MailItem.GetInspector;
    MailInspector.display(false); //true means modal
//    MailInspector.Send;

  finally
    Outlook := Unassigned;
  //  StringList.Free;
  end;

end;

function Tmainform.RemoveDiacritics(const S:string):string;
var SLength,ResultLength:Integer;
begin
  SLength:=Length(S);
  ResultLength := LCMapString(LANG_ENGLISH, NORM_IGNORENONSPACE, PChar(S), SLength, nil, 0);
  SetLength(Result, ResultLength);
  if ResultLength > 0 then LCMapString(LANG_ENGLISH, NORM_IGNORENONSPACE, PChar(S), SLength, PChar(Result), ResultLength);
end;

procedure Tmainform.Button1Click(Sender: TObject);
begin
novyform.show;
end;

procedure Tmainform.Button2Click(Sender: TObject);
var dir: string;
begin
dir := (mainform.Path + '\' + Edit1.Text + ' ' + Edit2.Text);
ShellExecute(Application.Handle, PChar('explore'), PChar(dir), nil, nil, SW_SHOWNORMAL);
end;

procedure Tmainform.Button3Click(Sender: TObject);
//var searchResult : TSearchRec;
begin
//if findfirst('\\192.168.1.100\pd\_Vítězslav\'+'*'+edit1.text+'*', faDirectory, searchResult) = 0 then
  begin
 //   if (searchResult.attr and faDirectory) = faDirectory
 //     then begin
        //caption := searchResult.Name;
        ShellExecute(Application.Handle, PChar('explore'), PChar('\\192.168.1.100\pd\_Vítězslav\' +searchResult.Name), nil, nil, SW_SHOWNORMAL);
        //showmessage (PChar('\\192.168.1.100\PD\'+ searchResult.Name));
  //      FindClose(searchResult);
  //    end;
end;
end;

procedure Tmainform.Button4Click(Sender: TObject);
begin
if (edit3.Text = '') or (Edit4.Text = '') then
  Application.MessageBox('Doplň obec a katastrální území', 'Chyba', MB_ICONERROR)
  else begin
    if FileExists (PChar(cesta + '\textova cast - DUR.docx')) then
      Application.MessageBox('Soubor již existuje..', 'Chyba', MB_ICONERROR)
        else begin
          workform.show;
          CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\textova cast - DUR.docx'), PChar(cesta + '\' + 'textova cast - DUR.docx'), False);
          workform.Label2.Caption := 'Přepisuji dokument DUR....';
          WordApp := CreateOLEObject('Word.Application');
          WordApp.Documents.Open(cesta + '\textova cast - DUR.docx');
          WordApp.Selection.Find.Text := '_CISLO';
          WordApp.Selection.Find.Replacement.Text := Edit1.Text;
          WordApp.Selection.Find.Forward := True;
          WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
          WordApp.Selection.Find.Text := '_NAZEV';
          WordApp.Selection.Find.Replacement.Text := Edit2.Text;
          WordApp.Selection.Find.Forward := True;
          WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
          WordApp.Selection.Find.Text := '_OBEC';
          WordApp.Selection.Find.Replacement.Text := Edit3.Text;
          WordApp.Selection.Find.Forward := True;
          WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
          WordApp.Selection.Find.Text := '_KATASTR';
          WordApp.Selection.Find.Replacement.Text := Edit4.Text;
          WordApp.Selection.Find.Forward := True;
          WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
          WordApp.ActiveDocument.SaveAs(cesta + '\textova cast - DUR.docx');
          WordApp.Quit;
          WordApp := Unassigned;
          workform.close;
        end;
  end;
end;

procedure Tmainform.Button5Click(Sender: TObject);
begin
if (edit3.Text = '') or (Edit4.Text = '') then
  Application.MessageBox('Doplň obec a katastrální území', 'Chyba', MB_ICONERROR)
  else begin
if FileExists (PChar(cesta + '\textova cast - DPS.docx')) then
Application.MessageBox('Soubor již existuje..', 'Chyba', MB_ICONERROR)
else begin
workform.show;
CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\textova cast - DPS.docx'), PChar(cesta + '\' + 'textova cast - DPS.docx'), False);
  workform.Label2.Caption := 'Přepisuji dokument DPS....';
WordApp := CreateOLEObject('Word.Application');
WordApp.Documents.Open(cesta + '\textova cast - DPS.docx');
WordApp.Selection.Find.Text := '_CISLO';
WordApp.Selection.Find.Replacement.Text := Edit1.Text;
WordApp.Selection.Find.Forward := True;
WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
WordApp.Selection.Find.Text := '_NAZEV';
WordApp.Selection.Find.Replacement.Text := Edit2.Text;
WordApp.Selection.Find.Forward := True;
WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
WordApp.Selection.Find.Text := '_OBEC';
WordApp.Selection.Find.Replacement.Text := Edit3.Text;
WordApp.Selection.Find.Forward := True;
WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
WordApp.Selection.Find.Text := '_KATASTR';
WordApp.Selection.Find.Replacement.Text := Edit4.Text;
WordApp.Selection.Find.Forward := True;
WordApp.Selection.Find.Execute(Replace := wdReplaceAll);
WordApp.ActiveDocument.SaveAs(cesta + '\textova cast - DPS.docx');
WordApp.Quit;
WordApp := Unassigned;
  workform.close;
end;
end;
end;

procedure Tmainform.Button6Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\POV.xlsx')) then
Application.MessageBox('Soubor již existuje..', 'Chyba', MB_ICONERROR)
else begin
workform.show;
  CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\POV.XLSX'), PChar(cesta + '\POV.xlsx'), False);
  workform.Label2.Caption := 'Přepisuji dokument POV....';
Xedit2((cesta + '\POV.xlsx'),Edit1.Text, 1, 1);
Xedit2((cesta + '\POV.xlsx'),Edit2.Text, 1, 4);
  workform.close;
end;
end;

procedure Tmainform.Button7Click(Sender: TObject);
begin
form5.show;
end;

procedure Tmainform.Button8Click(Sender: TObject);
begin
if FileExists (PChar(cesta + '\Formular pro SJZ.xlsx')) then
Application.MessageBox('Soubor již existuje..', 'Chyba', MB_ICONERROR)
else begin
  CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\Formular pro SJZ.xlsx'), PChar(cesta + '\' + 'Formular pro SJZ.xlsx'), False);
end;
end;

procedure Tmainform.Button9Click(Sender: TObject);
begin
workform.show;
workform.Label2.Caption := 'Připravuji NN Geoportal....';
  CopyFolder ((extractfilepath(Application.exename)+'\VZORY\DOK_PDNN_x'), (cesta + '\DOK_PDNN_' + edit1.text ));
    begin
    FileOpenDialog1.InitialDir := (cesta + '\');
    FileOpenDialog1.Filename := '';
    FileOpenDialog1.Filter := 'PDF - textova*.pdf|textova*.pdf';
    FileOpenDialog1.Title := 'Vyber textovou zprávu - PDF';
    if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\1_ZPRAVA\TECHNICKA_ZPRAVA_' + Edit1.Text + '_t.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF - *situace*.pdf|*situace*.pdf';
      FileOpenDialog1.Title := 'Vyber situaci - PDF';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\2_SITUACE\SITUACE_' + Edit1.Text + '_s.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF - *schema*.pdf|*schema*.pdf';
      FileOpenDialog1.Title := 'Vyber jednopólové schéma zapojení - PDF';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\3_JEDNOPOLOVE_SCHEMA\JEDNOPOLOVE_SCHEMA_' + Edit1.Text + '_j.pdf'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'XLSX|pov*.xlsx';
      FileOpenDialog1.Title := 'Vyber POV - XLSX';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\4_POV\POV_' + Edit1.Text + '_p.xlsx'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'XLS|'+edit1.text+'*.xls';
      FileOpenDialog1.Title := 'Vyber rozpočet - XLS';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\5_ROZPOCET\ROZPOCET_STAVBY_' + Edit1.Text + '_r.xls'), False);
        end;
      FileOpenDialog1.Filename := '';
      FileOpenDialog1.InitialDir := (cesta + '\');
      FileOpenDialog1.Filter := 'PDF - plan*|plan*.pdf';
      FileOpenDialog1.Title := 'Vyber BOZP';
      if FileOpenDialog1.Execute = true then
        begin
          FileOpenDialog1.InitialDir := (cesta + '\');
          CopyFile(PChar(FileOpenDialog1.FileName), PChar(cesta + '\' + 'DOK_PDNN_' + edit1.text + '\7_SOUVISEJICI_DOKUMENTY\BOZP_' + Edit1.Text + '.pdf'), False);
        end;
    end;
  workform.Close;
end;

procedure Tmainform.CopyFolder(CopyFrom, CopyTo: String);
var
FO: TSHFileOpStruct;
begin
If Not DirectoryExists(CopyFrom) Then Exit;
FO.Wnd := Application.Handle;
FO.wFunc := FO_COPY;
FO.pFrom := PChar(IncludeTrailingBackSlash(CopyFrom) + '*.*' + #0#0);
FO.pTo := PChar(IncludeTrailingBackSlash(CopyTo) + #0#0);
FO.fFlags := FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or FOF_SILENT Or FOF_ALLOWUNDO;
ShFileOperation(FO);
end;

procedure Tmainform.Edit1Click(Sender: TObject);
begin
Edit1.SelectAll;
Edit1.CopyToClipboard;
end;

procedure Tmainform.Edit2Click(Sender: TObject);
begin
edit2.SelectAll;
Edit2.CopyToClipboard;
end;

procedure Tmainform.Edit5Click(Sender: TObject);
begin
Edit5.SelectAll;
Edit5.CopyToClipboard;
end;

procedure Tmainform.Edit7Change(Sender: TObject);
begin
if Edit7.text = 'P ' then Edit7.text := 'Petra';
if Edit7.text = 'Č ' then Edit7.text := 'Verča';
if Edit7.text = 'S ' then Edit7.text := 'Sylva';
end;

procedure TMainform.OpenAndModifyAnExistingExcelFileAndSaveAs(cislo: string);
var
   ExcelFileName: String;
   ExcelApplication, ExcelWorkbook, ExcelWorksheet: Variant;
   obsah: variant;
   datum: integer;
   i: integer;
begin
     ExcelFileName := PDsoubor;        
     ExcelApplication := Null;
     ExcelWorkbook := Null;
     ExcelWorksheet := Null;
     try
           ExcelApplication := CreateOleObject('Excel.Application');
     except
           ExcelApplication := Null;
     end;

     If VarIsNull(ExcelApplication) = False then
        begin
             try
                ExcelApplication.Visible := False; 
                ExcelApplication.DisplayAlerts := False; 

                try
                   ExcelWorkbook := ExcelApplication.Workbooks.Open(ExcelFileName);
                except
                      ExcelWorkbook := Null;
                end;

                If VarIsNull(ExcelWorkbook) = False then
                   begin
                        try
                           ExcelWorksheet := ExcelWorkbook.WorkSheets[1]; //[1]
                        except
                              ExcelWorksheet := Null;
                         end;

                     If VarIsNull(ExcelWorksheet) = False then
                           begin
                                for i := 1 to 300 do
                                  begin
                                    obsah := ExcelWorksheet.Cells[i, 1];
                                    if obsah.value = cislo
                                     then
                                      begin
                                        datum := ExcelWorksheet.Cells[i, 4];
                                          if datum = 0 then
                                            begin
                                              datum := ExcelWorksheet.Cells[i, 3];
                                              Edit6.Text := datetostr(datum);
                                            end else
                                              if datum <> 0 then
                                                begin
                                                  datum := ExcelWorksheet.Cells[i, 4];
                                                  Edit6.Text := datetostr(datum);
                                                end;
                                      end;
                                  end;
                             end;
                   end;
             finally
                    ExcelApplication.Workbooks.Close;
                    ExcelApplication.DisplayAlerts := True;
                    ExcelApplication.Quit;
                    ExcelWorksheet := Unassigned;
                    ExcelWorkbook := Unassigned;
                    ExcelApplication := Unassigned;
             end;
        end;
end;

procedure Tmainform.ComboBox1Select(Sender: TObject);
var xyz: string;
begin
button2.Enabled := true;
button3.Enabled := true;
edit2.Text := copy(ComboBox1.Text, 15,100);
edit1.Text := copy(ComboBox1.Text, 0,13);

if findfirst('\\192.168.1.100\pd\_Vítězslav\'+'*'+edit1.text+'*', faDirectory, searchResult) = 0 then
  begin
    if (searchResult.attr and faDirectory) = faDirectory
      then
        begin
        xyz := Copy(searchResult.Name, 2,2);
        FindClose(searchResult);
        edit7.Text := xyz;
        end;
      end;

cesta := (mainform.Path + '\' + Edit1.Text + ' ' + Edit2.Text);
edit5.Text := ComboBox1.Text;
//Edit6.Color := clRed;
//Edit6.Text := '... zjištuji datum ...';
//OpenAndModifyAnExistingExcelFileAndSaveAs(Edit1.Text);
//Edit6.Color := clWindow;
Form5.Button2.OnClick(self);
caption := Edit5.Text;
TrayIcon1.Hint := Edit5.Text;
edit3.Text := '';
edit4.Text := '';

if FileExists (PChar(cesta + '\textova cast - DUR.docx')) then
  begin
    BitBtn2.Kind := bkYes;
    BitBtn2.Caption := 'DUR';
  end else
    begin
      BitBtn2.Kind := bkNo;
      BitBtn2.Caption := 'DUR';
    end;

if FileExists (PChar(cesta + '\textova cast - DUR - bourani.docx')) then
  begin
    BitBtn4.Kind := bkYes;
    BitBtn4.Caption := 'DUR odstr';
  end else
    begin
      BitBtn4.Kind := bkNo;
      BitBtn4.Caption := 'DUR odstr';
    end;

if FileExists (PChar(cesta + '\textova cast - DPS.docx')) then
  begin
    BitBtn3.Kind := bkYes;
    BitBtn3.Caption := 'DPS';
  end else
    begin
      BitBtn3.Kind := bkNo;
      BitBtn3.Caption := 'DPS';
    end;

if FileExists (PChar(cesta + '\POV.xlsx')) then
  begin
    BitBtn5.Kind := bkYes;
    BitBtn5.Caption := 'POV';
  end else
    begin
      BitBtn5.Kind := bkNo;
      BitBtn5.Caption := 'POV';
    end;

if FileExists (PChar(cesta + '\Odpady.xlsx')) then
  begin
    BitBtn9.Kind := bkYes;
    BitBtn9.Caption := 'Odpady';
  end else
    begin
      BitBtn9.Kind := bkNo;
      BitBtn9.Caption := 'Odpady';
    end;

if FileExists (PChar(cesta + '\Desky.xlsx')) then
  begin
    BitBtn6.Kind := bkYes;
    BitBtn6.Caption := 'Desky';
  end else
    begin
      BitBtn6.Kind := bkNo;
      BitBtn6.Caption := 'Desky';
    end;

if FileExists (PChar(cesta + '\Formular pro SJZ.xlsx')) then
  begin
    BitBtn10.Kind := bkYes;
    BitBtn10.Caption := 'SJZ';
  end else
    begin
      BitBtn10.Kind := bkNo;
      BitBtn10.Caption := 'SJZ';
    end;

if FileExists (PChar(cesta + '\uvodni listy.docx')) then
  begin
    BitBtn7.Kind := bkYes;
    BitBtn7.Caption := 'Listy';
  end else
    begin
      BitBtn7.Kind := bkNo;
      BitBtn7.Caption := 'Listy';
    end;

if DirectoryExists (PChar(cesta + '\_CD ' + edit1.text)) then
  begin
    BitBtn11.Kind := bkYes;
    BitBtn11.Caption := 'CD';
  end else
    begin
      BitBtn11.Kind := bkNo;
      BitBtn11.Caption := 'CD';
    end;


if FileExists (PChar(cesta + '\DOK_PDNN_' + edit1.text + '.zip')) then
  begin
    BitBtn8.Kind := bkYes;
    BitBtn8.Caption := 'NN';
  end else
    begin
      BitBtn8.Kind := bkNo;
      BitBtn8.Caption := 'NN';
    end;

if FileExists (PChar(cesta + '\DOK_PDVN_' + edit1.text + '.zip')) then
  begin
    BitBtn12.Kind := bkYes;
    BitBtn12.Caption := 'VN';
  end else
    begin
      BitBtn12.Kind := bkNo;
      BitBtn12.Caption := 'VN';
    end;


if FileExists (PChar(cesta + '\__souhrnko.pdf')) then
  begin
    CheckBox1.Checked := true;
    Label13.Caption := DateToStr(FileDateToDateTime(FileAge (cesta + '\__souhrnko.pdf')))
  end else
    begin
      CheckBox1.Checked := false;
      Label13.Caption := '-';
    end;

if FileExists (PChar(cesta + '\plan_bozp_' + Edit1.Text + '.pdf')) then
  begin
    CheckBox2.Checked := true;
    Label14.Caption := DateToStr(FileDateToDateTime(FileAge (cesta + '\plan_bozp_' + Edit1.Text + '.pdf')))
  end else
    begin
      CheckBox2.Checked := false;
      Label14.Caption := '-';
end;
end;

procedure Tmainform.Xedit(soubor, co: String; k1, k2: integer);
var
myxlApp: TExcelApplication;
myxlBook: TExcelWorkbook;
begin
myxlApp := TExcelApplication.Create(Nil);
myxlApp.Connect;
myxlApp.Visible[LCID] := true;
myxlApp.Workbooks.Open ( soubor,
 EmptyParam , EmptyParam , EmptyParam , EmptyParam ,
 EmptyParam , EmptyParam , EmptyParam , EmptyParam ,
 EmptyParam , EmptyParam , EmptyParam , EmptyParam , EmptyParam , EmptyParam , 0 );
myxlBook := TExcelWorkbook.Create(myxlApp);
myxlBook.ConnectTo(myxlApp.ActiveWorkbook);
myxlApp.Cells.Item[k1,k2] := co;
myxlBook.Save;
myxlBook.Close(True,soubor); 
myxlBook.Disconnect;
FreeAndNil(myxlBook);
myxlApp.Disconnect;
myxlApp.Quit;
FreeAndNil(myxlApp);
end;

procedure Tmainform.Xedit2(soubor, co: String; k1, k2: integer);
var
   ExcelFileName, ExcelFileNameNew: String;
   ExcelApplication, ExcelWorkbook, ExcelWorksheet: Variant;
begin
     ExcelFileName := soubor;        
     ExcelApplication := Null;
     ExcelWorkbook := Null;
     ExcelWorksheet := Null;
      try
        ExcelApplication := CreateOleObject('Excel.Application');
     except
           ExcelApplication := Null;
     end;
     If VarIsNull(ExcelApplication) = False then
        begin
             try
                ExcelApplication.Visible := False;
                ExcelApplication.DisplayAlerts := False; 
                try
                   ExcelWorkbook := ExcelApplication.Workbooks.Open(ExcelFileName);
                except
                      ExcelWorkbook := Null;
                end;
                If VarIsNull(ExcelWorkbook) = False then
                   begin
                        try
                           ExcelWorksheet := ExcelWorkbook.WorkSheets[1]; //[1]
                        except
                              ExcelWorksheet := Null;
                        end;

                        If VarIsNull(ExcelWorksheet) = False then
                           begin
                                ExcelWorksheet.Select;
                                ExcelWorksheet.Cells[k1,k2] := co;                
                                ExcelWorksheet.Cells[1,1].Select;                                 
                                ExcelWorkbook.SaveAs(ExcelFileName);
                           end;
                   end;
             finally
                    ExcelApplication.Workbooks.Close;
                    ExcelApplication.DisplayAlerts := True;
                    ExcelApplication.Quit;
                    ExcelWorksheet := Unassigned;
                    ExcelWorkbook := Unassigned;
                    ExcelApplication := Unassigned;
             end;
        end;
end;

procedure Tmainform.Zobrazitaplikaci1Click(Sender: TObject);
begin
TrayIcon1DblClick(self)
end;

procedure Tmainform.Zobrazitseznamstaveb1Click(Sender: TObject);
begin
TrayIcon1Click(self)
end;

procedure Tmainform.FormCreate(Sender: TObject);
begin
Comboupdate;
LCID := GetUserDefaultLCID;       // XLS
//Timer1.enabled := true;
end;

procedure Tmainform.Konec1Click(Sender: TObject);
begin
Application.Terminate
end;

procedure Tmainform.Timer1Timer(Sender: TObject);
begin
workform.show;
workform.Label2.Caption := 'Aktualizace dat....';
Form5.Button1.OnClick(self);
label9.Caption := 'Aktualizace: ' + TimeToStr(Time) + ' - ' + DateToStr(date);
timer1.Enabled := false;
workform.close;
end;

procedure Tmainform.Timer2Timer(Sender: TObject);
begin
workform.show;
workform.Label2.Caption := 'Aktualizace dat....';
Form5.Button1.OnClick(self);
label9.Caption := 'Aktualizace: ' + TimeToStr(Time) + ' - ' + DateToStr(date);
workform.close;
end;

procedure Tmainform.TrayIcon1Click(Sender: TObject);
begin
  Form4.Show();
  WindowState := wsNormal;
  Application.BringToFront();
end;

procedure Tmainform.TrayIcon1DblClick(Sender: TObject);
begin
  TrayIcon1.Visible := False;
  Show();
  WindowState := wsNormal;
  Application.BringToFront();
  Form4.close;
end;

procedure Tmainform.Comboupdate;
var INI: TIniFile;
    SR, searchResult: TSearchRec;
    soubor_PD, cesta_PD: string;
begin
INI := TIniFile.Create((extractfilepath(Application.exename))+'nastaveni.ini');
Path := INI.ReadString('Nastaveni','cesta','');
soubor_PD := INI.ReadString('Nastaveni','soubor_PD_tabulka','');
cesta_PD := INI.ReadString('Nastaveni','cesta_PD_tabulka','');
begin
      if findfirst(cesta_PD + soubor_PD + '*', faAnyFile, searchResult) = 0 then
        begin
          if (searchResult.attr {and faAnyfile}) = faAnyfile
            then else begin
             PDSoubor := cesta_PD + searchResult.Name;
             FindClose(searchResult);
            end;
        end;

      ComboBox1.Clear;
       if FindFirst(Path+'\*', faDirectory, SR) = 0 then
         try
            repeat
              if ExtractFileExt(sr.FindData.cFileName) = '' then
                if (Copy(SR.Name, 1, 1) = 'I') or (Copy(SR.Name, 1, 1) = 'O') then
                  begin
                    combobox1.Items.Add(sr.FindData.cFileName);
                  end;
            until FindNext(SR) <> 0;
         finally
            FindClose(SR);
         end;
       ComboBox1.ItemIndex := 0;
      INI.Free;
end;

end;
end.
