unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Inifiles, Vcl.StdCtrls, ShellApi, Comobj,
  Vcl.ExtDlgs, excel2000, Vcl.OleServer, Excel2010, Vcl.ExtCtrls, Vcl.AppEvnts,
  ExcelXP, System.Win.TaskbarCore, Vcl.Taskbar, Vcl.Menus, Vcl.Buttons;

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
    procedure Button10Click(Sender: TObject);
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
    { Public declarations }
  end;

var
  mainform: Tmainform;
  WordApp: Variant;
  xls, xlw: Variant;
  Excl, Wrkbk: OleVariant;

implementation

{$R *.dfm}

uses Unit2, Unit3, Unit4;

procedure Tmainform.ApplicationEvents1Minimize(Sender: TObject);
begin
Hide();
WindowState := wsMinimized;
TrayIcon1.Visible := True;
TrayIcon1.Animate := True;
TrayIcon1.ShowBalloonHint;
end;

procedure Tmainform.BitBtn1Click(Sender: TObject);
begin
Comboupdate;
ComboBox1.ItemIndex := 1;
ComboBox1.OnSelect(self);
end;

procedure Tmainform.Button10Click(Sender: TObject);
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
      CopyFile(PChar(extractfilepath(Application.exename)+'\VZORY\Desky - 2022-01.xlsx'), PChar(cesta + '\' + 'Desky.xlsx'), False);
      workform.Label2.Caption := 'Přepisuji dokument Desky....';
      Xedit2((cesta + '\Desky.xlsx'),Edit1.Text, 3, 12);
      Xedit2((cesta + '\Desky.xlsx'),Edit2.Text, 2, 12);
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
var searchResult : TSearchRec;
begin
if findfirst('\\192.168.1.100\pd\_Vítězslav\'+'*'+edit1.text+'*', faDirectory, searchResult) = 0 then
  begin
    if (searchResult.attr and faDirectory) = faDirectory
      then begin
        caption := searchResult.Name;
        ShellExecute(Application.Handle, PChar('explore'), PChar('\\192.168.1.100\pd\_Vítězslav\' +searchResult.Name), nil, nil, SW_SHOWNORMAL);
        //showmessage (PChar('\\192.168.1.100\PD\'+ searchResult.Name));
        FindClose(searchResult);
      end;
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
begin
button2.Enabled := true;
button3.Enabled := true;
edit2.Text := copy(ComboBox1.Text, 15,100);
edit1.Text := copy(ComboBox1.Text, 0,13);
cesta := (mainform.Path + '\' + Edit1.Text + ' ' + Edit2.Text);
edit5.Text := ComboBox1.Text;
Edit6.Color := clRed;
Edit6.Text := '... zjištuji datum ...';
OpenAndModifyAnExistingExcelFileAndSaveAs(Edit1.Text);
Edit6.Color := clWindow;
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
end;

procedure Tmainform.Konec1Click(Sender: TObject);
begin
Application.Terminate
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
                if Copy(SR.Name, 1, 1) = 'I' then
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
