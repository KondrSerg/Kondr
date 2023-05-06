unit Redakt;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ExtCtrls, StdCtrls, ComCtrls, Menus, AdvGrid, AsgMemo,
  asgcheck, AdvObj, BaseGrid, DB, ADODB, asgprev, AsgFindDialog, Mask,
  AdvDropDown, AdvCustomGridDropDown, AdvGridDropDown, AdvProgr, ActnList,
  XPMan, System.Actions;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    MainMenu1: TMainMenu;
    C1: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    PageControl2: TPageControl;
    TabSheet2: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    TabSheet7: TTabSheet;
    PageControl3: TPageControl;
    TabSheet8: TTabSheet;
    TabSheet10: TTabSheet;
    CapitalCheck1: TCapitalCheck;
    MemoEditLink1: TMemoEditLink;
    SG: TAdvStringGrid;
    Button1: TButton;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    Memo1: TMemo;
    Memo2: TMemo;
    Button2: TButton;
    AdvPreviewDialog1: TAdvPreviewDialog;
    PopupMenu1: TPopupMenu;
    N5: TMenuItem;
    PB: TAdvProgress;
    R25PB1: TAdvProgress;
    R25Sg1: TAdvStringGrid;
    Button5: TButton;
    Button6: TButton;
    R25SG2: TAdvStringGrid;
    R25PB2: TAdvProgress;
    GermSG: TAdvStringGrid;
    GermPB: TAdvProgress;
    Button7: TButton;
    Button8: TButton;
    Label1: TLabel;
    SGVir: TAdvStringGrid;
    PBVir: TAdvProgress;
    Button9: TButton;
    Edit1: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Edit2: TEdit;
    Label7: TLabel;
    Label8: TLabel;
    SGGibka: TAdvStringGrid;
    Button3: TButton;
    Button4: TButton;
    SGR25: TAdvStringGrid;
    Button12: TButton;
    Button13: TButton;
    PBGib: TProgressBar;
    Edit3: TEdit;
    Label9: TLabel;
    Label10: TLabel;
    PG25: TProgressBar;
    Label11: TLabel;
    TabSheet3: TTabSheet;
    PageControl4: TPageControl;
    TabSheet4: TTabSheet;
    Label12: TLabel;
    Label13: TLabel;
    SGSbor1: TAdvStringGrid;
    Button10: TButton;
    Button11: TButton;
    PBSbor1f: TProgressBar;
    TabSheet9: TTabSheet;
    Label14: TLabel;
    Button14: TButton;
    Button15: TButton;
    PBSborMRP: TProgressBar;
    SGSbor2: TAdvStringGrid;
    PBSbor1fD: TProgressBar;
    SGMRP: TAdvStringGrid;
    Label15: TLabel;
    Label16: TLabel;
    TabSheet11: TTabSheet;
    Label17: TLabel;
    SGR25Sbor: TAdvStringGrid;
    PBR25Sbor: TProgressBar;
    Button16: TButton;
    Button17: TButton;
    Edit4: TEdit;
    Edit5: TEdit;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    SGTSbor: TAdvStringGrid;
    PBSborStoj: TProgressBar;
    TabSheet12: TTabSheet;
    SGSbor1f: TAdvStringGrid;
    SGSbor1fEl: TAdvStringGrid;
    Label21: TLabel;
    Label22: TLabel;
    PB1f: TProgressBar;
    PB1fEl: TProgressBar;
    Button18: TButton;
    Button19: TButton;
    N6: TMenuItem;
    PageVozd: TPageControl;
    ts1: TTabSheet;
    ts5: TTabSheet;
    ts8: TTabSheet;
    SGReg: TAdvStringGrid;
    SGRegL: TAdvStringGrid;
    lbl1: TLabel;
    lbl2: TLabel;
    N7: TMenuItem;
    btn1: TButton;
    ts2: TTabSheet;
    xpmnfst1: TXPManifest;
    actlst1: TActionList;
    act1: TAction;
    act2: TAction;
    SGRegL1: TAdvStringGrid;
    lbl3: TLabel;
    SGGermS: TAdvStringGrid;
    SGGermP: TAdvStringGrid;
    btn3: TButton;
    lbl4: TLabel;
    lbl5: TLabel;
    SGGerm: TAdvStringGrid;
    lbl6: TLabel;
    SGTulp1: TAdvStringGrid;
    SGTulp2: TAdvStringGrid;
    SGTulp3: TAdvStringGrid;
    lbl7: TLabel;
    lbl8: TLabel;
    lbl9: TLabel;
    btn5: TButton;
    SGTulp: TAdvStringGrid;
    lbl10: TLabel;
    SGKlara: TAdvStringGrid;
    lbl11: TLabel;
    btn6: TButton;
    ts3: TTabSheet;
    SGUVK: TAdvStringGrid;
    btn7: TButton;
    lbl12: TLabel;
    procedure Button1Click(Sender: TObject);
    function SetConnect1(var AQ: TADOQuery): Boolean;
    function SetConnect(var AQ: TADOQuery;
        var AC: TAdoConnection): Boolean;
    function mkQuerySelect(var AAQ: TADOQuery; const sQ: WideString; Args:
        array of const): boolean;
    function mkQueryUpdate(var AAQ: TADOQuery; const sQ: WideString; Args:
        array of const): boolean;
    function mkQueryInsert(var AAQ: TADOQuery; const sQ: WideString;
                        Args: array
                        of const): boolean;
    function mkQueryDelete(var AAQ: TADOQuery; const sQ: WideString;
                        Args: array
                        of const): boolean;
    procedure Check(Module: string; Str: variant; Col: Integer);
    function Float(var Str:String): String;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure SGGetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure SGGetEditorProp(Sender: TObject; ACol, ARow: Integer;
      AEditLink: TEditLink);
    procedure N5Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure PageControl2MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Button9Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure PageControl3MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Button12Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure Button16Click(Sender: TObject);
    procedure Button18Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button15Click(Sender: TObject);
    procedure Button17Click(Sender: TObject);
    procedure Button19Click(Sender: TObject);
    procedure PageControl4MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure N3Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure act1Execute(Sender: TObject);
    procedure act2Execute(Sender: TObject);
    procedure SGRegEditingDone(Sender: TObject);
    procedure SGR_Save(Tabl,str,Str1,Str2,Str3,str4,Str5,Str6,Znak,Znak1,Znak2,Znak3,znak4,znak5:string);
    procedure SGRegLEditingDone(Sender: TObject);
    procedure SGRegL1EditingDone(Sender: TObject);
    procedure btn3Click(Sender: TObject);
    procedure SGGermSEditingDone(Sender: TObject);
    procedure SGGermPEditingDone(Sender: TObject);
    procedure SGGermEditingDone(Sender: TObject);
    procedure btn5Click(Sender: TObject);
    procedure SGTulp1EditingDone(Sender: TObject);
    procedure SGTulp2EditingDone(Sender: TObject);
    procedure SGTulp3EditingDone(Sender: TObject);
    procedure SGTulpEditingDone(Sender: TObject);
    procedure btn6Click(Sender: TObject);
    procedure SGKlaraEditingDone(Sender: TObject);
    procedure btn7Click(Sender: TObject);
    procedure SGUVKEditingDone(Sender: TObject);
    procedure PageVozdMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
    DEF_CatalogName, MSSQL_CONN_STR: string;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
function TForm1.Float(var Str:String): String;
Var PosTrud:Integer;
Begin
        Result:='';
  PosTrud := Pos(',', Str); //Трудоемкость FLOAT
  if PosTrud <> 0 then
  begin
    Delete(Str, PosTrud, 1);
    Insert('.', Str, PosTrud);
  end;
  Result:=Str;

end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 procedure TForm1.SGR_Save(Tabl,str,Str1,Str2,Str3,str4,Str5,Str6,Znak,Znak1,Znak2,Znak3,znak4,znak5:string);
 var i:Integer;
 begin

 end;  
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
function TForm1.SetConnect1(var AQ: TADOQuery): Boolean;
begin
        Result := SetConnect(AQ, Form1.ADOConnection1);
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function TForm1.SetConnect(var AQ: TADOQuery;
        var AC: TAdoConnection): Boolean;
var
        MSSQL_CONN_STR, S: string;
begin
        Result := false;
        S := ExtractFileDir(ParamStr(0));
        Memo1.Lines.Clear;
        Memo1.Lines.LoadFromFile(S + '\Conn.ini');
        DEF_CatalogName := Memo1.Lines[0];

        MSSQL_CONN_STR :=  Memo1.Lines[4];

        {MSSQL_CONN_STR :=
                'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
                'Persist Security Info=True;User Id=testuser;' +
                'Initial Catalog=' + DEF_CatalogName +
                ';Data Source=DINAMO\SQLEXPRESS';  }
        if ADOConnection1 = nil then
                ADOConnection1 := TADOConnection.Create(Self);
        ADOConnection1.LoginPrompt := False;
        ADOConnection1.Connected := False;
        ADOConnection1.ConnectionString := MSSQL_CONN_STR;
        try
                ADOConnection1.Connected := True;
        except
                MessageDlg('Не удалось подключиться к базе данных!', mtError,
                        [mbOk], 0);
                Close;
        end;
        try
                AQ.Connection := AC;
                AQ.Close;
                AQ.SQL.Clear;
                AQ.SQL.Text := 'Select * From Klapana';
                AQ.Open;
        except
                AQ.Close;
                Check('SetConnect', 'Ошибка коннекта', 0);
                exit;
        end;
        Result := true;
end;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
function TForm1.mkQuerySelect(var AAQ: TADOQuery; const sQ: WideString; Args:
        array of const): boolean;
begin
        Result := false;
        if not SetConnect1(AAQ) then
                exit;
        try
                AAQ.Close;
                AAQ.SQL.Clear;
                AAQ.SQL.Text := Format(sQ, Args);
                AAQ.Active := true;
                check('query', AAQ.SQL.Text, AAQ.RecordCount);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                AAQ.Close;
                exit;
        end;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//Лог
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function TForm1.mkQueryUpdate(var AAQ: TADOQuery; const sQ: WideString; Args:
        array of const): boolean;
var
        Str: string;
begin
        Result := false;
        if not SetConnect1(AAQ) then
                exit;
        try
                AAQ.Close;
                AAQ.SQL.Clear;
                Str := Format(sQ, Args);
                AAQ.SQL.Text := Format(sQ, Args);
                AAQ.ExecSQL;
                check('query', AAQ.SQL.Text, 1);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                AAQ.Close;
                MessageDlg('Не удалось обновить запись!', mtError, [mbOk], 0);
                exit;
        end;
        AAQ.Close;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
function TForm1.mkQueryInsert(var AAQ: TADOQuery; const sQ: WideString; Args:
        array of const): boolean;
var
        Str: string;
begin
        Result := false;
        if not SetConnect1(AAQ) then
                exit;
        ADOConnection1.BeginTrans;
        try
                AAQ.Close;
                AAQ.SQL.Clear;
                Str := Format(sQ, Args);
                AAQ.SQL.Text := Format(sQ, Args);
                AAQ.ExecSQL;
                ADOConnection1.CommitTrans();
                check('query', AAQ.SQL.Text, 1);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                MessageDlg('Не удалось вставить запись!', mtError, [mbOk], 0);
                AAQ.Close;
                ADOConnection1.RollbackTrans();
                exit;
        end;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function TForm1.mkQueryDelete(var AAQ: TADOQuery; const sQ: WideString; Args:
        array of const): boolean;
var
        Str: string;
begin
        Result := false;
        if not SetConnect1(AAQ) then
                exit;
        try
                AAQ.Close;
                AAQ.SQL.Clear;
                Str := Format(sQ, Args);
                AAQ.SQL.Text := Format(sQ, Args);

                AAQ.ExecSQL;
                check('query', AAQ.SQL.Text, 1);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                MessageDlg('Не удалось удалить запись!', mtError, [mbOk], 0);
                AAQ.Close;
                exit;
        end;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
procedure TForm1.Check(Module: string; Str: variant; Col: Integer);
var
        myFile: TextFile;
        S, Str1: string;
begin
        S := ExtractFileDir(ParamStr(0)) + '\Log.ini';
        AssignFile(myFile, S);
        Str1 := Module + '  ' + Str + '  ' + IntToStr(Col);
        try
                Append(myFile);
                WriteLn(myFile, Str1);
                CloseFile(myFile);
        except
                ReWrite(myFile);
        end;
        Form1.Memo2.Lines.Add(Module + '  ' + Str + '  ' + IntToStr(Col));
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
procedure TForm1.Button1Click(Sender: TObject);
Var I:Integer;
begin                       //Резка
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=1)',
                        ['[Резка Гибка]']) then
                        exit;
  Sg.RowCount:=ADOQuery1.RecordCount+1;
  Sg.ColCount:=6;
  sg.Cells[0,0]:='Толщина';
    sg.Cells[1,0]:='Высота';
    sg.Cells[2,0]:=('Длина');
    sg.Cells[3,0]:=('Норма');
    sg.Cells[4,0]:=('№');
    sg.Cells[5,0]:=('Таблица');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    sg.Cells[0,I+1]:=ADOQuery1.FieldByName('Толщина').AsString;
    sg.Cells[1,I+1]:=ADOQuery1.FieldByName('Высота').AsString;
    sg.Cells[2,I+1]:=ADOQuery1.FieldByName('Длина').AsString;
    sg.Cells[3,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    sg.Cells[4,I+1]:=ADOQuery1.FieldByName('№').AsString;
    sg.Cells[5,I+1]:=ADOQuery1.FieldByName('Таблица').AsString;
    Sg.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  Sg.LoadVisualProps('12.ini');
  Sg.LoadColSizes;
  Sg.SortByColumn(0);
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  Выруб
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=11)',
                        ['[Резка Гибка]']) then
                        exit;
  SgVir.RowCount:=ADOQuery1.RecordCount+1;
  SgVir.ColCount:=4;
  SgVir.Cells[0,0]:='Кол Углов';
    SgVir.Cells[1,0]:=('Норма');
    SgVir.Cells[2,0]:=('№');
    SgVir.Cells[3,0]:=('Таблица');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SgVir.Cells[0,I+1]:=ADOQuery1.FieldByName('Высота').AsString;
    SgVir.Cells[1,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SgVir.Cells[2,I+1]:=ADOQuery1.FieldByName('№').AsString;
    SgVir.Cells[3,I+1]:=ADOQuery1.FieldByName('Таблица').AsString;
    SgVir.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SgVir.LoadVisualProps('12.ini');
  SgVir.LoadColSizes;
  SgVir.SortByColumn(1);
  //Sg.AutoSize:=True;
 //+++++++++++++++++++++++++++++++++++++++++++++++
   if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=12)',
                        ['[Резка Гибка]']) then
                        exit;
   Edit1.Text:=ADOQuery1.FieldByName('Норма').AsString;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Sg.SaveVisualProps('12.ini');
  Sg.SaveColSizes;
end;

procedure TForm1.Button2Click(Sender: TObject);
Var I:Integer;
        Tol,Dlin,Vis,Norm:String;
begin
   if
  MessageDlg('Сохранить изменения?',
  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  Exit
  Else
  Begin
        PB.Min:=0;
        PB.Max:= SG.RowCount;
        For I:=1 To SG.RowCount Do
        Begin
                Tol:=Sg.Cells[1,I];
                Tol:=Float(tol);

                Dlin:=Sg.Cells[3,I];
                Dlin:=Float(Dlin);

                Vis:=Sg.Cells[2,I];
                Vis:=Float(Vis);

                Norm:=Sg.Cells[4,I];
                Norm:=Float(Norm);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Норма]=' +
                #39 +Norm + #39 +
                ' WHERE ([Толщина]=' + #39 + Tol + #39+') AND  ([Высота]=' + #39 + Vis + #39+') AND ([Длина]=' + #39 + Dlin + #39+') AND ([№]=' + #39 + '1' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
                PB.Position:=I;
        End;
        PB.Position:=0;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Выруб
        PBVir.Min:=0;
        PBVir.Max:= SGVir.RowCount;
        For I:=1 To SGVir.RowCount Do
        Begin

                Vis:=SGVir.Cells[0,I];
                Vis:=Float(Vis);

                Norm:=SGVir.Cells[1,I];
                Norm:=Float(Norm);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Норма]=' +
                #39 +Norm + #39 +
                ' WHERE ([Высота]=' + #39 + Vis + #39+') AND ([№]=' + #39 + '11' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
                PBVir.Position:=I;
        End;
        PBVir.Position:=0;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
               Norm:=Edit1.Text;
                Norm:=Float(Norm);
        if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Норма]=' +
                #39 +Norm + #39 +
                ' WHERE ([№]=' + #39 + '12' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
   Button1.Click;
  end;
     {   if FileExists('cars.dat') then
      Begin
        SG.Combobox.Items.LoadFromFile('cars.dat');
        SG.Combobox.DropDownCount:=3;
      end; }

end;

procedure TForm1.SGGetEditorType(Sender: TObject; ACol, ARow: Integer;
  var AEditor: TEditorType);
begin
    {with advstringgrid1 do
    case acol of
    1,2: aEditor:=edComboList;
    end;}
end;

procedure TForm1.SGGetEditorProp(Sender: TObject; ACol, ARow: Integer;
  AEditLink: TEditLink);
  Var Str,Str1:String;
begin
  {  with advstringgrid1 do
    case acol of
    1:begin
      ClearComboString;
      if FileExists('cars.dat') then
      Begin
        Combobox.Items.LoadFromFile('cars.dat');
        Combobox.DropDownCount:=3;

        Str:=Combobox.Items.Strings[0];
        Str1:=Combobox.Items.Strings[1];
        end;
    end;
    2:begin
      ClearComboString;
      if (cells[1,arow]<>'') then  
      if FileExists(cells[1,arow]+'.dat') then
        Combobox.Items.LoadFromFile(cells[1,arow]+'.dat');
    end; 
   end;  }
end;

procedure TForm1.N5Click(Sender: TObject);
Var Row1:Integer;
begin
  SG.AddRow;
  SG.Cells[5,SG.RowCount-1]:='1';
  SG.Cells[6,SG.RowCount-1]:='Резка на ножницах';
end;

procedure TForm1.FormShow(Sender: TObject);
Var S:String;
begin
        S := ExtractFileDir(ParamStr(0));
        Memo1.Lines.Clear;
        Memo1.Lines.LoadFromFile(S + '\Conn.ini');
        DEF_CatalogName := Memo1.Lines[0];
        Caption:=DEF_CatalogName;
        PageControl2.ActivePageIndex:=0;
        PageControl3.ActivePageIndex:=0;
        Button1.Click;
        Btn1.Click;
end;

procedure TForm1.Button5Click(Sender: TObject);
Var I:Integer;
begin
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=7)',
                        ['[Резка Гибка]']) then
                        exit;    //
  R25Sg1.RowCount:=ADOQuery1.RecordCount+1;
  R25Sg1.ColCount:=4;
    R25Sg1.Cells[0,0]:=('Длина');
    R25Sg1.Cells[1,0]:=('Норма');
    R25Sg1.Cells[2,0]:=('№');
    R25Sg1.Cells[3,0]:=('Таблица');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    R25Sg1.Cells[0,I+1]:=ADOQuery1.FieldByName('Длина').AsString;
    R25Sg1.Cells[1,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    R25Sg1.Cells[2,I+1]:=ADOQuery1.FieldByName('№').AsString;
    R25Sg1.Cells[3,I+1]:=ADOQuery1.FieldByName('Таблица').AsString;
    R25Sg1.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  R25Sg1.LoadVisualProps('12.ini');
  R25Sg1.LoadColSizes;
  R25Sg1.SortByColumn(0);
//++++++++++++++++++++++++++++++++
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=8)',
                        ['[Резка Гибка]']) then
                        exit;
  R25SG2.RowCount:=ADOQuery1.RecordCount+1;
  R25Sg2.ColCount:=4;
    R25SG2.Cells[0,0]:=('Длина');
    R25SG2.Cells[1,0]:=('Норма');
    R25SG2.Cells[2,0]:=('№');
    R25SG2.Cells[3,0]:=('Таблица');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    R25SG2.Cells[0,I+1]:=ADOQuery1.FieldByName('Длина').AsString;
    R25SG2.Cells[1,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    R25SG2.Cells[2,I+1]:=ADOQuery1.FieldByName('№').AsString;
    R25SG2.Cells[3,I+1]:=ADOQuery1.FieldByName('Таблица').AsString;
    R25SG2.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  R25SG2.LoadVisualProps('12.ini');
  R25SG2.LoadColSizes;
  R25SG2.SortByColumn(0);
end;

procedure TForm1.Button6Click(Sender: TObject);
Var I:Integer;
        Tol,Dlin,Vis,Norm:String;
begin
   if
  MessageDlg('Сохранить изменения?',
  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  Exit
  Else
  Begin
        R25PB1.Min:=0;
        R25PB1.Max:= R25Sg1.RowCount;
        For I:=1 To R25Sg1.RowCount Do
        Begin
                Dlin:=R25Sg1.Cells[1,I];
                Dlin:=Float(Dlin);

                Norm:=R25Sg1.Cells[2,I];
                Norm:=Float(Norm);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Норма]=' +
                #39 +Norm + #39 +
                ' WHERE ([Длина]=' + #39 + Dlin + #39+') AND ([№]=' + #39 + '7' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
                R25PB1.Position:=I;
        End;
        R25PB1.Position:=0;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        R25PB2.Min:=0;
        R25PB2.Max:= R25Sg2.RowCount;
        For I:=1 To R25Sg2.RowCount Do
        Begin
                Dlin:=R25Sg2.Cells[1,I];
                Dlin:=Float(Dlin);

                Norm:=R25Sg2.Cells[2,I];
                Norm:=Float(Norm);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Норма]=' +
                #39 +Norm + #39 +
                ' WHERE ([Длина]=' + #39 + Dlin + #39+') AND ([№]=' + #39 + '8' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
                R25PB2.Position:=I;
        End;
        R25PB2.Position:=0;
        Button5.Click;
  End;
end;

procedure TForm1.Button7Click(Sender: TObject);
Var I:Integer;
begin
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=6) AND (Длина<>1)',
                        ['[Резка Гибка]']) then
                        exit;
  GermSG.RowCount:=ADOQuery1.RecordCount+1;
  GermSG.ColCount:=4;
    GermSG.Cells[0,0]:=('Длина');
    GermSG.Cells[1,0]:=('Норма');
    GermSG.Cells[2,0]:=('№');
    GermSG.Cells[3,0]:=('Таблица');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    GermSG.Cells[0,I+1]:=ADOQuery1.FieldByName('Длина').AsString;
    GermSG.Cells[1,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    GermSG.Cells[2,I+1]:=ADOQuery1.FieldByName('№').AsString;
    GermSG.Cells[3,I+1]:=ADOQuery1.FieldByName('Таблица').AsString;
    GermSG.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  GermSG.LoadVisualProps('12.ini');
  GermSG.LoadColSizes;
  GermSG.SortByColumn(1);
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Фаска
if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=6) AND (Длина=1)',
                        ['[Резка Гибка]']) then
                        exit;
        Edit2.Text:=ADOQuery1.FieldByName('Норма').AsString;
end;

procedure TForm1.Button8Click(Sender: TObject);
Var I:Integer;
        Tol,Dlin,Vis,Norm:String;
begin
   if
  MessageDlg('Сохранить изменения?',
  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  Exit
  Else
  Begin
        GermPB.Min:=0;
        GermPB.Max:= GermSG.RowCount;
        For I:=1 To GermSG.RowCount Do
        Begin
                Dlin:=GermSG.Cells[1,I];
                Dlin:=Float(Dlin);

                Norm:=GermSG.Cells[2,I];
                Norm:=Float(Norm);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Норма]=' +
                #39 +Norm + #39 +
                ' WHERE ([Длина]=' + #39 + Dlin + #39+') AND ([№]=' + #39 + '6' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
                GermPB.Position:=I;
        End;
        GermPB.Position:=0;
        //_+++++++++++++++++++++++++++++++++++++++++++++++++++++++Фаска

        Norm:=Edit2.Text;
        Norm:=Float(Norm);
        if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Норма]=' +
                #39 +Norm + #39 +
                ' WHERE ([Длина]=' + #39 + '1' + #39+') AND ([№]=' + #39 + '6' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
        Button7.Click;
  End;
end;

procedure TForm1.PageControl2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
        Case PageControl2.ActivePageIndex  of
        0:Button1.Click;
        1:Button7.Click;
        2:Button5.Click;
        End;
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
if not Form1.mkQueryDelete( Form1.ADOQuery1, 'DELETE FROM %s  WHERE (Высота=' + #39 +'4'+#39+') AND ([№]='+#39+'11'+#39+')' ,['[Резка Гибка]'] )
        Then
        Exit;
end;

procedure TForm1.Button3Click(Sender: TObject);
Var I:Integer;
begin
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=3)',
                        ['[Резка Гибка]']) then
                        exit;
  SGGibka.RowCount:=ADOQuery1.RecordCount+1;
  SGGibka.ColCount:=6;
  SGGibka.Cells[0,0]:='Толщина';
    SGGibka.Cells[1,0]:='КолГибов';
    SGGibka.Cells[2,0]:=('Длина');
    SGGibka.Cells[3,0]:=('Норма');
    SGGibka.Cells[4,0]:=('№');
    SGGibka.Cells[5,0]:=('Таблица');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGGibka.Cells[0,I+1]:=ADOQuery1.FieldByName('Толщина').AsString;
    SGGibka.Cells[1,I+1]:=ADOQuery1.FieldByName('Высота').AsString;
    SGGibka.Cells[2,I+1]:=ADOQuery1.FieldByName('Длина').AsString;
    SGGibka.Cells[3,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGGibka.Cells[4,I+1]:=ADOQuery1.FieldByName('№').AsString;
    SGGibka.Cells[5,I+1]:=ADOQuery1.FieldByName('Таблица').AsString;
    SGGibka.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGGibka.LoadVisualProps('12.ini');
  SGGibka.LoadColSizes;
  SGGibka.SortByColumn(0);
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=4)',
                        ['[Резка Гибка]']) then
                        exit;
  Edit3.Text:=ADOQuery1.FieldByName('Норма').AsString;
end;

procedure TForm1.PageControl3MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
        Case PageControl3.ActivePageIndex  of
        0:Button3.Click;
        1:Button12.Click;
        End;
end;

procedure TForm1.Button12Click(Sender: TObject);
Var I:Integer;
begin
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=9)',
                        ['[Резка Гибка]']) then
                        exit;
  SGR25.RowCount:=ADOQuery1.RecordCount+1;
  SGR25.ColCount:=5;
  SGR25.Cells[0,0]:='Кол Створок';
    SGR25.Cells[1,0]:='Норма';
    SGR25.Cells[2,0]:=('Время на коррект. программы');
    SGR25.Cells[3,0]:=('№');
    SGR25.Cells[4,0]:=('Таблица');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGR25.Cells[0,I+1]:=ADOQuery1.FieldByName('Толщина').AsString;
    SGR25.Cells[1,I+1]:=ADOQuery1.FieldByName('Высота').AsString;
    SGR25.Cells[2,I+1]:=ADOQuery1.FieldByName('Длина').AsString;
    SGR25.Cells[3,I+1]:=ADOQuery1.FieldByName('№').AsString;
    SGR25.Cells[4,I+1]:=ADOQuery1.FieldByName('Таблица').AsString;
    SGR25.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGR25.LoadVisualProps('12.ini');
  SGR25.LoadColSizes;
  SGR25.SortByColumn(0);
end;

procedure TForm1.Button4Click(Sender: TObject);
Var I:Integer;
        Tol,Dlin,Vis,Norm:String;
begin
   if
  MessageDlg('Сохранить изменения?',
  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  Exit
  Else
  Begin

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Выруб
        PBGib.Min:=0;
        PBGib.Max:= SGGibka.RowCount;
        For I:=1 To SGGibka.RowCount Do
        Begin
                TOl:=SGGibka.Cells[0,I];
                Tol:=Float(Tol);

                Vis:=SGGibka.Cells[1,I];
                Vis:=Float(Vis);

                Dlin:=SGGibka.Cells[2,I];
                Dlin:=Float(Dlin);

                Norm:=SGGibka.Cells[3,I];
                Norm:=Float(Norm);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Норма]=' +
                #39 +Norm + #39 +
                ' WHERE ([Длина]=' + #39 + Dlin + #39+') AND ([Толщина]=' + #39 + Tol + #39+') AND ([Высота]=' + #39 + Vis + #39+') AND ([№]=' + #39 + '3' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
                PBGib.Position:=I;
        End;
        PBGib.Position:=0;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
               Norm:=Edit3.Text;
                Norm:=Float(Norm);
        if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Норма]=' +
                #39 +Norm + #39 +
                ' WHERE ([№]=' + #39 + '4' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
      Button3.Click;
  end;
end;

procedure TForm1.Button13Click(Sender: TObject);
Var I:Integer;
        Tol,Dlin,Vis,Norm:String;
begin
   if
  MessageDlg('Сохранить изменения?',
  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  Exit
  Else
  Begin

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Выруб
        PG25.Min:=0;
        PG25.Max:= SGR25.RowCount;
        For I:=1 To SGR25.RowCount Do
        Begin
                TOl:=SGR25.Cells[0,I];
                Tol:=Float(Tol);

                Vis:=SGR25.Cells[1,I];
                Vis:=Float(Vis);

                Dlin:=SGR25.Cells[2,I];
                Dlin:=Float(Dlin);

                Norm:=SGR25.Cells[3,I];
                Norm:=Float(Norm);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Высота]=' +
                #39 +Vis + #39 + ', Длина='+
                #39 +Dlin + #39 + ' WHERE  ([Толщина]=' + #39 + Tol + #39+')  AND ([№]=' + #39 + '9' + #39+')' ,
                ['[Резка Гибка]']) then
                Exit;
                PG25.Position:=I;
        End;
        PG25.Position:=0;
        Button12.Click;
end;
End;

procedure TForm1.Button10Click(Sender: TObject);
Var I:Integer;
begin
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=1)',
                        ['[Сборка корпуса]']) then
                        exit;
  SGSbor1.RowCount:=ADOQuery1.RecordCount+1;
  SGSbor1.ColCount:=6;
  SGSbor1.Cells[0,0]:='Полупериметр';
    SGSbor1.Cells[1,0]:='1 Лопатка';
    SGSbor1.Cells[2,0]:=('2 Лопатка');
    SGSbor1.Cells[3,0]:=('4 Лопатка');
    SGSbor1.Cells[4,0]:=('№');
    SGSbor1.Cells[5,0]:=('Таблица');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGSbor1.Cells[0,I+1]:=ADOQuery1.FieldByName('Размер_Клапана').AsString;
    SGSbor1.Cells[1,I+1]:=ADOQuery1.FieldByName('Лопатка1').AsString;
    SGSbor1.Cells[2,I+1]:=ADOQuery1.FieldByName('Лопатка2').AsString;
    SGSbor1.Cells[3,I+1]:=ADOQuery1.FieldByName('Лопатка4').AsString;
    SGSbor1.Cells[4,I+1]:=ADOQuery1.FieldByName('№').AsString;
    SGSbor1.Cells[5,I+1]:=ADOQuery1.FieldByName('Конст').AsString;
    SGSbor1.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGSbor1.LoadVisualProps('12.ini');
  SGSbor1.LoadColSizes;
  SGSbor1.SortByColumn(0);

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=2)',
                        ['[Сборка корпуса]']) then
                        exit;
  SGSbor2.RowCount:=ADOQuery1.RecordCount+1;
  SGSbor2.ColCount:=6;
  SGSbor2.Cells[0,0]:='Полупериметр';
    SGSbor2.Cells[1,0]:='1 Лопатка';
    SGSbor2.Cells[2,0]:=('2 Лопатка');
    SGSbor2.Cells[3,0]:=('4 Лопатка');
    SGSbor2.Cells[4,0]:=('№');
    SGSbor2.Cells[5,0]:=('Таблица');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGSbor2.Cells[0,I+1]:=ADOQuery1.FieldByName('Размер_Клапана').AsString;
    SGSbor2.Cells[1,I+1]:=ADOQuery1.FieldByName('Лопатка1').AsString;
    SGSbor2.Cells[2,I+1]:=ADOQuery1.FieldByName('Лопатка2').AsString;
    SGSbor2.Cells[3,I+1]:=ADOQuery1.FieldByName('Лопатка4').AsString;
    SGSbor2.Cells[4,I+1]:=ADOQuery1.FieldByName('№').AsString;
    SGSbor2.Cells[5,I+1]:=ADOQuery1.FieldByName('Конст').AsString;
    SGSbor2.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGSbor2.LoadVisualProps('12.ini');
  SGSbor2.LoadColSizes;
  SGSbor2.SortByColumn(0);
end;

procedure TForm1.Button14Click(Sender: TObject);
Var I:Integer;
begin

        if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=3)',
                        ['[Сборка корпуса]']) then
                        exit;
        SGMRP.RowCount:=ADOQuery1.RecordCount+2;
        SGMRP.ColCount:=6;
        //+++++++++++++++++++++++++++++++++++++++++++++++++
        SGMRP.MergeCells(0,0,1,2);
        SGMRP.Cells[0,0]:='Полупериметр';
        //SGMRP.WordWraps[0,0]:=True;
        //+++++++++++++++++++++++++++++++++++++++++++++++++
        SGMRP.MergeCells(1,0,3,1);
        SGMRP.Cells[1,0] := 'Установка холодного корпуса';
        //+++++++++++++++++++++++++++++++++++++++++++++++++
        SGMRP.Cells[1,1] := 'Вермикулит(сверление отверстий,установка)';
        //+++++++++++++++++++++++++++++++++++++++++++++++++
        SGMRP.Cells[2,1] := 'Установка на клапан';
        //+++++++++++++++++++++++++++++++++++++++++++++++++
        SGMRP.Cells[3,1] := 'Итого';
        //+++++++++++++++++++++++++++++++++++++++++++++++++
        SGMRP.MergeCells(4,0,1,2);
        SGMRP.Cells[4,0] := 'Установка сетки, жалюзийной решетки';
        //+++++++++++++++++++++++++++++++++++++++++++++++++
        SGMRP.MergeCells(5,0,1,2);
        SGMRP.Cells[5,0]:=('№');
  for I:=0 to ADOQuery1.RecordCount Do
  Begin
    SGMRP.Cells[0,I+2]:=ADOQuery1.FieldByName('Размер_Клапана').AsString;
    SGMRP.Cells[1,I+2]:=ADOQuery1.FieldByName('Лопатка1').AsString;
    SGMRP.Cells[2,I+2]:=ADOQuery1.FieldByName('Лопатка2').AsString;
    SGMRP.Cells[3,I+2]:=ADOQuery1.FieldByName('Лопатка3').AsString;
    SGMRP.Cells[4,I+2]:=ADOQuery1.FieldByName('Лопатка4').AsString;
    SGMRP.Cells[5,I+2]:=ADOQuery1.FieldByName('№').AsString;
    SGMRP.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGMRP.SortByColumn(0);
//+++++++++++++++++++++++++++++++++++++++++++++++++
   if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([Конст]='+#39+'мрп'+#39+')',
                        ['[Сборка корпуса]']) then
                        exit;
   Edit4.Text:=ADOQuery1.FieldByName('Лопатка1').AsString;
//+++++++++++++++++++++++++++++++++++++++++++++++++
   if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([Конст]='+#39+'р25'+#39+')',
                        ['[Сборка корпуса]']) then
                        exit;
   Edit5.Text:=ADOQuery1.FieldByName('Лопатка1').AsString;
end;

procedure TForm1.Button16Click(Sender: TObject);
Var I:Integer;
begin
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=4)',
                        ['[Сборка корпуса]']) then
                        exit;
    SGR25Sbor.RowCount:=ADOQuery1.RecordCount+1;
    SGR25Sbor.ColCount:=4;
    SGR25Sbor.Cells[0,0]:='Высота В,мм';
    SGR25Sbor.Cells[1,0]:='Ширина А,мм';
    SGR25Sbor.Cells[2,0]:=('Норма');
    SGR25Sbor.Cells[3,0]:=('№');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGR25Sbor.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка1').AsString;
    SGR25Sbor.Cells[1,I+1]:=ADOQuery1.FieldByName('Размер_Клапана').AsString;
    SGR25Sbor.Cells[2,I+1]:=ADOQuery1.FieldByName('Лопатка2').AsString;
    SGR25Sbor.Cells[3,I+1]:=ADOQuery1.FieldByName('№').AsString;
    SGR25Sbor.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGR25Sbor.LoadVisualProps('12.ini');
  SGR25Sbor.LoadColSizes;
  SGR25Sbor.SortByColumn(0);
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=7)',
                        ['[Сборка корпуса]']) then
                        exit;
    SGTSbor.RowCount:=ADOQuery1.RecordCount+1;
    SGTSbor.ColCount:=3;
    SGTSbor.Cells[0,0]:='Кол-во лопаток';
    SGTSbor.Cells[1,0]:=('Норма');
    SGTSbor.Cells[2,0]:=('№');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGTSbor.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка1').AsString;
    SGTSbor.Cells[1,I+1]:=ADOQuery1.FieldByName('Лопатка2').AsString;
    SGTSbor.Cells[2,I+1]:=ADOQuery1.FieldByName('№').AsString;
    SGTSbor.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGTSbor.LoadVisualProps('12.ini');
  SGTSbor.LoadColSizes;
  SGTSbor.SortByColumn(0);
end;

procedure TForm1.Button18Click(Sender: TObject);
Var I:Integer;
begin
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=5)',
                        ['[Сборка корпуса]']) then
                        exit;
  SGSbor1f.RowCount:=ADOQuery1.RecordCount+1;
  SGSbor1f.ColCount:=5;
  SGSbor1f.Cells[0,0]:='Полупериметр';
    SGSbor1f.Cells[1,0]:='Кол-во лопаток';
    SGSbor1f.Cells[2,0]:=('Кол-во приводов');
    SGSbor1f.Cells[3,0]:=('Норма');
    SGSbor1f.Cells[4,0]:=('№');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGSbor1f.Cells[0,I+1]:=ADOQuery1.FieldByName('Размер_Клапана').AsString;
    SGSbor1f.Cells[1,I+1]:=ADOQuery1.FieldByName('Лопатка2').AsString;
    SGSbor1f.Cells[2,I+1]:=ADOQuery1.FieldByName('Лопатка3').AsString;
    SGSbor1f.Cells[3,I+1]:=ADOQuery1.FieldByName('Лопатка1').AsString;
    SGSbor1f.Cells[4,I+1]:=ADOQuery1.FieldByName('№').AsString;
    SGSbor1f.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGSbor1f.LoadVisualProps('12.ini');
  SGSbor1f.LoadColSizes;
  SGSbor1f.SortByColumn(0);

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s Where ([№]=6)',
                        ['[Сборка корпуса]']) then
                        exit;
  SGSbor1fEl.RowCount:=ADOQuery1.RecordCount+1;
  SGSbor1fEl.ColCount:=5;
  SGSbor1fEl.Cells[0,0]:='Полупериметр';
    SGSbor1fEl.Cells[1,0]:='Кол-во лопаток';
    SGSbor1fEl.Cells[2,0]:=('Кол-во приводов');
    SGSbor1fEl.Cells[3,0]:=('Норма');
    SGSbor1fEl.Cells[4,0]:=('№');
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGSbor1fEl.Cells[0,I+1]:=ADOQuery1.FieldByName('Размер_Клапана').AsString;
    SGSbor1fEl.Cells[1,I+1]:=ADOQuery1.FieldByName('Лопатка2').AsString;
    SGSbor1fEl.Cells[2,I+1]:=ADOQuery1.FieldByName('Лопатка3').AsString;
    SGSbor1fEl.Cells[3,I+1]:=ADOQuery1.FieldByName('Лопатка1').AsString;
    SGSbor1fEl.Cells[4,I+1]:=ADOQuery1.FieldByName('№').AsString;
    SGSbor1fEl.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGSbor1fEl.LoadVisualProps('12.ini');
  SGSbor1fEl.LoadColSizes;
  SGSbor1fEl.SortByColumn(0);
end;

procedure TForm1.Button11Click(Sender: TObject);
Var I:Integer;
        Lop1,Lop2,Lop3,Lop4,Nomer,Razmer:String;
begin
   if
  MessageDlg('Сохранить изменения?',
  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  Exit
  Else
  Begin

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Сборка1
        PBSbor1f.Min:=0;
        PBSbor1f.Max:= SGSbor1.RowCount;
        For I:=1 To SGSbor1.RowCount Do
        Begin
                Razmer:=SGSbor1.Cells[0,I];

                Lop1:=SGSbor1.Cells[1,I];
                Lop1:=Float(Lop1);

                Lop2:=SGSbor1.Cells[2,I];
                Lop2:=Float(Lop2);

                Lop4:=SGSbor1.Cells[3,I];
                Lop4:=Float(Lop4);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Лопатка1]=' +#39 +Lop1 + #39 +',[Лопатка2]=' +#39 +Lop2 + #39+',[Лопатка4]=' +#39 +Lop4 + #39 +
                ' WHERE ([Размер_Клапана]=' + #39 + Razmer + #39+') AND ([№]=' + #39 + '1' + #39+')' ,
                ['[Сборка корпуса]']) then
                Exit;
                PBSbor1f.Position:=I;
        End;
        PBSbor1f.Position:=0;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++= Сборк2
        PBSbor1fD.Min:=0;
        PBSbor1fD.Max:= SGSbor2.RowCount;
        For I:=1 To SGSbor2.RowCount Do
        Begin
                Razmer:=SGSbor2.Cells[0,I];

                Lop1:=SGSbor2.Cells[1,I];
                Lop1:=Float(Lop1);

                Lop2:=SGSbor2.Cells[2,I];
                Lop2:=Float(Lop2);

                Lop4:=SGSbor2.Cells[3,I];
                Lop4:=Float(Lop4);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Лопатка1]=' +#39 +Lop1 + #39 +',[Лопатка2]=' +#39 +Lop2 + #39+',[Лопатка4]=' +#39 +Lop4 + #39 +
                ' WHERE ([Размер_Клапана]=' + #39 + Razmer + #39+') AND ([№]=' + #39 + '2' + #39+')' ,
                ['[Сборка корпуса]']) then
                Exit;
                PBSbor1fD.Position:=I;
        End;
        PBSbor1fD.Position:=0;
        Button10.Click;
  End;
end;

procedure TForm1.Button15Click(Sender: TObject);
Var I:Integer;
        Lop1,Lop2,Lop3,Lop4,Nomer,Razmer:String;
begin
   if
  MessageDlg('Сохранить изменения?',
  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  Exit
  Else
  Begin

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Сборка1
        PBSborMRP.Min:=0;
        PBSborMRP.Max:= SGMRP.RowCount;
        For I:=1 To SGMRP.RowCount Do
        Begin
                Razmer:=SGMRP.Cells[0,I+1];

                Lop1:=SGMRP.Cells[1,I+1];
                Lop1:=Float(Lop1);

                Lop2:=SGMRP.Cells[2,I+1];
                Lop2:=Float(Lop2);

                Lop3:=SGMRP.Cells[3,I+1];
                Lop3:=Float(Lop3);

                Lop4:=SGMRP.Cells[4,I+1];
                Lop4:=Float(Lop4);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Лопатка1]=' +#39 +Lop1 + #39 +',[Лопатка2]=' +#39 +Lop2 + #39+',[Лопатка3]=' +#39 +Lop3 + #39 +',[Лопатка4]=' +#39 +Lop4 + #39 +
                ' WHERE ([Размер_Клапана]=' + #39 + Razmer + #39+') AND ([№]=' + #39 + '3' + #39+')' ,
                ['[Сборка корпуса]']) then
                Exit;
                PBSborMRP.Position:=I;
        End;
        PBSborMRP.Position:=0;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++= Сборк2
                Lop1:=Edit4.Text;
                Lop1:=Float(Lop1);
                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Лопатка1]=' +#39 +Lop1 + #39 +
                ' WHERE ([Конст]=' + #39 + 'мрз' + #39+') OR ([Конст]=' + #39 + 'мрп' + #39+') ' ,
                ['[Сборка корпуса]']) then
                Exit;
       //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=
                Lop1:=Edit5.Text;
                Lop1:=Float(Lop1);
                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Лопатка1]=' +#39 +Lop1 + #39 +
                ' WHERE ([Конст]=' + #39 + 'р25' + #39+')' ,
                ['[Сборка корпуса]']) then
                Exit;
                Button14.Click;
  End;
end;

procedure TForm1.Button17Click(Sender: TObject);
Var I:Integer;
        Lop1,Lop2,Lop3,Lop4,Nomer,Razmer:String;
begin
   if
  MessageDlg('Сохранить изменения?',
  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  Exit
  Else
  Begin

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Сборка1
        PBR25Sbor.Min:=0;
        PBR25Sbor.Max:= SGR25Sbor.RowCount;
        For I:=1 To SGR25Sbor.RowCount Do
        Begin
                Razmer:=SGR25Sbor.Cells[1,I];

                Lop1:=SGR25Sbor.Cells[0,I];
                Lop1:=Float(Lop1);

                Lop2:=SGR25Sbor.Cells[2,I];
                Lop2:=Float(Lop2);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Лопатка2]=' +#39 +Lop2 + #39+
                ' WHERE ([Размер_Клапана]=' + #39 + Razmer + #39+') AND ([№]=' + #39 + '4' + #39+') AND ([Лопатка1]=' + #39 + Lop1 + #39+')' ,
                ['[Сборка корпуса]']) then
                Exit;
                PBR25Sbor.Position:=I;
        End;
        PBR25Sbor.Position:=0;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++= Сборк2
        PBSborStoj.Min:=0;
        PBSborStoj.Max:= SGTSbor.RowCount;
        For I:=1 To SGTSbor.RowCount Do
        Begin

                Lop1:=SGTSbor.Cells[0,I];
                Lop1:=Float(Lop1);

                Lop2:=SGTSbor.Cells[1,I];
                Lop2:=Float(Lop2);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Лопатка2]=' +#39 +Lop2 + #39 +
                ' WHERE ([Лопатка1]=' + #39 + Lop1 + #39+') AND ([№]=' + #39 + '7' + #39+')' ,
                ['[Сборка корпуса]']) then
                Exit;
                PBSborStoj.Position:=I;
        End;
        PBSborStoj.Position:=0;
        Button16.Click;
  End;
end;

procedure TForm1.Button19Click(Sender: TObject);
Var I:Integer;
        Lop1,Lop2,Lop3,Lop4,Nomer,Razmer:String;
begin
   if
  MessageDlg('Сохранить изменения?',
  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  Exit
  Else
  Begin

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Сборка1
        PB1f.Min:=0;
        PB1f.Max:= SGSbor1f.RowCount;
        For I:=1 To SGSbor1f.RowCount Do
        Begin
                Razmer:=SGSbor1f.Cells[0,I];

                Lop2:=SGSbor1f.Cells[1,I];  //kol lop
                Lop2:=Float(Lop2);

                Lop3:=SGSbor1f.Cells[2,I];//Kol privod
                Lop3:=Float(Lop3);

                Lop1:=SGSbor1f.Cells[3,I]; //Norma
                Lop1:=Float(Lop1);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Лопатка1]=' +#39 +Lop1 + #39 +
                ' WHERE ([Размер_Клапана]=' + #39 + Razmer + #39+') AND ([Лопатка2]=' +#39 +Lop2 + #39 +') AND ([Лопатка3]=' +#39 +Lop3 + #39+') AND ([№]=' + #39 + '5' + #39+')' ,
                ['[Сборка корпуса]']) then
                Exit;
                PB1f.Position:=I;
        End;
        PB1f.Position:=0;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++= Сборк2
        PB1fEl.Min:=0;
        PB1fEl.Max:= SGSbor1fEl.RowCount;
        For I:=1 To SGSbor1fEl.RowCount Do
        Begin
                Razmer:=SGSbor1fEl.Cells[0,I];

                Lop2:=SGSbor1fEl.Cells[1,I];
                Lop2:=Float(Lop2);

                Lop3:=SGSbor1fEl.Cells[2,I];
                Lop3:=Float(Lop3);

                Lop1:=SGSbor1fEl.Cells[3,I];
                Lop1:=Float(Lop1);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Лопатка1]=' +#39 +Lop1 + #39 +
                ' WHERE ([Размер_Клапана]=' + #39 + Razmer + #39+') AND ([№]=' + #39 + '6' + #39+') AND ([Лопатка2]=' +#39 +Lop2 + #39+') AND ([Лопатка3]=' +#39 +Lop3 + #39 +')' ,
                ['[Сборка корпуса]']) then
                Exit;
                PB1fEl.Position:=I;
        End;
        PB1fEl.Position:=0;
        Button18.Click;
  End;
end;

procedure TForm1.PageControl4MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
        Case PageControl4.ActivePageIndex  of
        0:Button10.Click;
        1:Button14.Click;
        2:Button16.Click;
        3:Button18.Click;
        End;
end;

procedure TForm1.N3Click(Sender: TObject);
begin
      PageControl1.Visible:=True;
        PageVozd.Visible:=False;
end;

procedure TForm1.N7Click(Sender: TObject);
begin
        PageControl1.Visible:=False;
        PageVozd.Visible:=True;
end;

procedure TForm1.btn1Click(Sender: TObject);
Var I:Integer;
begin
    SGReg.ColCount:=7;
    SGReg.Cells[0,0]:=('Кол Лопаток');
    SGReg.Cells[1,0]:=('При ширине <=300');
    SGReg.Cells[2,0]:=('При ширине от 301 до 850');
    SGReg.Cells[3,0]:=('При ширине от 851 до 1500');
    SGReg.Cells[4,0]:=('Норма');
    SGReg.Cells[5,0]:=('');
    SGReg.ColWidths[5]:=0;
    SGReg.Cells[6,0]:=('Имя');

  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Сборка Регуляр]']) then
                        exit;
  SGReg.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGReg.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка').AsString;
    SGReg.Cells[1,I+1]:=ADOQuery1.FieldByName('300').AsString;
    SGReg.Cells[2,I+1]:=ADOQuery1.FieldByName('850').AsString;
    SGReg.Cells[3,I+1]:=ADOQuery1.FieldByName('1500').AsString;
    SGReg.Cells[4,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGReg.Cells[5,I+1]:=ADOQuery1.FieldByName('Номер').AsString;
    SGReg.Cells[6,I+1]:=ADOQuery1.FieldByName('Конст').AsString;
    SGReg.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  //-----------------------------------------------------------
    SGRegL.ColCount:=7;
    SGRegL.Cells[0,0]:=('Кол Лопаток');
    SGRegL.Cells[1,0]:=('При ширине <=300');
    SGRegL.Cells[2,0]:=('При ширине от 301 до 850');
    SGRegL.Cells[3,0]:=('При ширине от 851 до 1500');
    SGRegL.Cells[4,0]:=('Норма');
    SGRegL.Cells[5,0]:=('');
    SGRegL.ColWidths[5]:=0;
    SGRegL.Cells[6,0]:=('Имя');
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Сборка Регуляр Л]']) then
                        exit;
  SGRegL.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGRegL.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка').AsString;
    SGRegL.Cells[1,I+1]:=ADOQuery1.FieldByName('300').AsString;
    SGRegL.Cells[2,I+1]:=ADOQuery1.FieldByName('850').AsString;
    SGRegL.Cells[3,I+1]:=ADOQuery1.FieldByName('1500').AsString;
    SGRegL.Cells[4,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGRegL.Cells[5,I+1]:=ADOQuery1.FieldByName('Номер').AsString;
    SGRegL.Cells[6,I+1]:=ADOQuery1.FieldByName('Конст').AsString;
    SGRegL.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  //-----------------------------------------------------------
    SGRegL1.ColCount:=6;
    SGRegL1.Cells[0,0]:=('Кол Лопаток');
    SGRegL1.Cells[1,0]:=('При ширине до 1000');
    SGRegL1.Cells[2,0]:=('При ширине больше 1000');
    SGRegL1.Cells[3,0]:=('Норма');
    SGRegL1.Cells[4,0]:=('');
    SGRegL1.ColWidths[4]:=0;
    SGRegL1.Cells[5,0]:=('Имя');
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Подсборка Регуляр]']) then
                        exit;
  SGRegL1.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGRegL1.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка').AsString;
    SGRegL1.Cells[1,I+1]:=ADOQuery1.FieldByName('1000').AsString;
    SGRegL1.Cells[2,I+1]:=ADOQuery1.FieldByName('1500').AsString;
    SGRegL1.Cells[3,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGRegL1.Cells[4,I+1]:=ADOQuery1.FieldByName('Номер').AsString;
    SGRegL1.Cells[5,I+1]:=ADOQuery1.FieldByName('Конст').AsString;
    SGRegL1.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  SGReg.LoadVisualProps('12.ini');
  SGReg.LoadColSizes;
  SGReg.SortByColumn(1);
  //
  SGRegL.LoadVisualProps('12.ini');
  SGRegL.LoadColSizes;
  SGRegL.SortByColumn(1);
  //

  SGRegL1.SortByColumn(1);
end;

procedure TForm1.act1Execute(Sender: TObject);
begin
        Memo2.Visible:=True;
end;

procedure TForm1.act2Execute(Sender: TObject);
begin
      Memo2.Visible:=False;
end;

procedure TForm1.SGRegEditingDone(Sender: TObject);
Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4:String;
begin
                Kol:=SGReg.Cells[0,SGReg.Row];
                Konst:=SGReg.Cells[6,SGReg.Row];

                Norm1:=SGReg.Cells[1,SGReg.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGReg.Cells[2,SGReg.Row];
                Norm2:=Float(Norm2);
                //---------------------
                Norm3:=SGReg.Cells[3,SGReg.Row];
                Norm3:=Float(Norm3);
                //---------------------
                Norm4:=SGReg.Cells[4,SGReg.Row];
                Norm4:=Float(Norm4);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm4 + #39 +','+
                '[300]='+#39 +Norm1 + #39 +','+
                '[850]='+#39 +Norm2 + #39 +','+
                '[1500]='+#39 +Norm3 + #39 +' '+
                ' WHERE ([Лопатка]=' + #39 + Kol + #39+') AND ([Конст]=' + #39 + Konst + #39+')' ,
                ['[Сборка Регуляр]']) then
                Exit;
end;

procedure TForm1.SGRegLEditingDone(Sender: TObject);
Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4:String;
begin
                Kol:=SGRegL.Cells[0,SGRegL.Row];
                Konst:=SGRegL.Cells[6,SGRegL.Row];

                Norm1:=SGRegL.Cells[1,SGRegL.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGRegL.Cells[2,SGRegL.Row];
                Norm2:=Float(Norm2);
                //---------------------
                Norm3:=SGRegL.Cells[3,SGRegL.Row];
                Norm3:=Float(Norm3);
                //---------------------
                Norm4:=SGRegL.Cells[4,SGRegL.Row];
                Norm4:=Float(Norm4);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm4 + #39 +','+
                '[300]='+#39 +Norm1 + #39 +','+
                '[850]='+#39 +Norm2 + #39 +','+
                '[1500]='+#39 +Norm3 + #39 +' '+
                ' WHERE ([Лопатка]=' + #39 + Kol + #39+') AND ([Конст]=' + #39 + Konst + #39+')' ,
                ['[Сборка Регуляр Л]']) then
                Exit;
end;

procedure TForm1.SGRegL1EditingDone(Sender: TObject);
Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4:String;
begin
                Kol:=SGRegL1.Cells[0,SGRegL1.Row];
                Konst:=SGRegL1.Cells[5,SGRegL1.Row];

                Norm1:=SGRegL1.Cells[1,SGRegL1.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGRegL1.Cells[2,SGRegL1.Row];
                Norm2:=Float(Norm2);
                //---------------------
                Norm3:=SGRegL1.Cells[3,SGRegL1.Row];
                Norm3:=Float(Norm3);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm3 + #39 +','+
                '[1000]='+#39 +Norm1 + #39 +','+
                '[1500]='+#39 +Norm2 + #39 +
                ' WHERE ([Лопатка]=' + #39 + Kol + #39+') AND ([Конст]=' + #39 + Konst + #39+')' ,
                ['[Подсборка Регуляр]']) then
                Exit;
end;

procedure TForm1.btn3Click(Sender: TObject);
Var I:Integer;
begin
    SGGermS.ColCount:=8;
    SGGermS.Cells[0,0]:=('Кол Лопаток');
    SGGermS.Cells[1,0]:=('При ширине <=550');
    SGGermS.Cells[2,0]:=('При ширине от 551 до 1100');
    SGGermS.Cells[3,0]:=('При ширине от 1101 до 1530');
    SGGermS.Cells[4,0]:=('При ширине >1530');
    SGGermS.Cells[5,0]:=('Норма');
    SGGermS.Cells[6,0]:=('');
    SGGermS.ColWidths[6]:=0;
    SGGermS.Cells[7,0]:=('Имя');

  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Сборка ГермикС]']) then
                        exit;
  SGGermS.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGGermS.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка').AsString;
    SGGermS.Cells[1,I+1]:=ADOQuery1.FieldByName('550').AsString;
    SGGermS.Cells[2,I+1]:=ADOQuery1.FieldByName('1100').AsString;
    SGGermS.Cells[3,I+1]:=ADOQuery1.FieldByName('1530').AsString;
    SGGermS.Cells[4,I+1]:=ADOQuery1.FieldByName('2066').AsString;
    SGGermS.Cells[5,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGGermS.Cells[6,I+1]:=ADOQuery1.FieldByName('Номер').AsString;
    SGGermS.Cells[7,I+1]:=ADOQuery1.FieldByName('Конст').AsString;
    SGGermS.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  //-----------------------------------------------------------
    SGGermP.ColCount:=8;
    SGGermP.Cells[0,0]:=('Кол Лопаток');
    SGGermP.Cells[1,0]:=('При ширине <=550');
    SGGermP.Cells[2,0]:=('При ширине от 551 до 1100');
    SGGermP.Cells[3,0]:=('При ширине от 1101 до 1530');
    SGGermP.Cells[4,0]:=('При ширине >1530');
    SGGermP.Cells[5,0]:=('Норма');
    SGGermP.Cells[6,0]:=('');
    SGGermP.ColWidths[6]:=0;
    SGGermP.Cells[7,0]:=('Имя');

  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Сборка ГермикП]']) then
                        exit;
  SGGermP.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGGermP.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка').AsString;
    SGGermP.Cells[1,I+1]:=ADOQuery1.FieldByName('550').AsString;
    SGGermP.Cells[2,I+1]:=ADOQuery1.FieldByName('1100').AsString;
    SGGermP.Cells[3,I+1]:=ADOQuery1.FieldByName('1530').AsString;
    SGGermP.Cells[4,I+1]:=ADOQuery1.FieldByName('2066').AsString;
    SGGermP.Cells[5,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGGermP.Cells[6,I+1]:=ADOQuery1.FieldByName('Номер').AsString;
    SGGermP.Cells[7,I+1]:=ADOQuery1.FieldByName('Конст').AsString;
    SGGermP.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
    //-----------------------------------------------------------
    SGGerm.ColCount:=8;
    SGGerm.Cells[0,0]:=('Кол Лопаток');
    SGGerm.Cells[1,0]:=('При ширине <=550');
    SGGerm.Cells[2,0]:=('При ширине от 551 до 1100');
    SGGerm.Cells[3,0]:=('При ширине от 1101 до 1530');
    SGGerm.Cells[4,0]:=('При ширине >1530');
    SGGerm.Cells[5,0]:=('Норма');
    SGGerm.Cells[6,0]:=('');
    SGGerm.ColWidths[6]:=0;
    SGGerm.Cells[7,0]:=('Имя');

  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[ПодСборка ГермикС_П]']) then
                        exit;
  SGGerm.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGGerm.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка').AsString;
    SGGerm.Cells[1,I+1]:=ADOQuery1.FieldByName('550').AsString;
    SGGerm.Cells[2,I+1]:=ADOQuery1.FieldByName('1100').AsString;
    SGGerm.Cells[3,I+1]:=ADOQuery1.FieldByName('1530').AsString;
    SGGerm.Cells[4,I+1]:=ADOQuery1.FieldByName('2066').AsString;
    SGGerm.Cells[5,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGGerm.Cells[6,I+1]:=ADOQuery1.FieldByName('Номер').AsString;
    SGGerm.Cells[7,I+1]:=ADOQuery1.FieldByName('Конст').AsString;
    SGGerm.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
SGGerm.SortByColumn(1);
SGGermS.SortByColumn(1);
SGGermP.SortByColumn(1);
end;

procedure TForm1.SGGermSEditingDone(Sender: TObject);
Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4 ,Norm5:String;
begin
                Kol:=SGGermS.Cells[0,SGGermS.Row];
                Konst:=SGGermS.Cells[7,SGGermS.Row];

                Norm1:=SGGermS.Cells[1,SGGermS.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGGermS.Cells[2,SGGermS.Row];
                Norm2:=Float(Norm2);
                //---------------------
                Norm3:=SGGermS.Cells[3,SGGermS.Row];
                Norm3:=Float(Norm3);
                //---------------------
                Norm4:=SGGermS.Cells[4,SGGermS.Row];
                Norm4:=Float(Norm4);
                //---------------------
                Norm5:=SGGermS.Cells[5,SGGermS.Row];
                Norm5:=Float(Norm5);
                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm5 + #39 +','+
                '[550]='+#39 +Norm1 + #39 +','+
                '[1100]='+#39 +Norm2 + #39 +','+
                '[1530]='+#39 +Norm3 + #39 +','+
                '[2066]='+#39 +Norm4 + #39 +
                ' WHERE ([Лопатка]=' + #39 + Kol + #39+') AND ([Конст]=' + #39 + Konst + #39+')' ,
                ['[Сборка ГермикС]']) then
                Exit;
end;

procedure TForm1.SGGermPEditingDone(Sender: TObject);
Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4 ,Norm5:String;
begin
                Kol:=SGGermP.Cells[0,SGGermP.Row];
                Konst:=SGGermP.Cells[7,SGGermP.Row];

                Norm1:=SGGermP.Cells[1,SGGermP.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGGermP.Cells[2,SGGermP.Row];
                Norm2:=Float(Norm2);
                //---------------------
                Norm3:=SGGermP.Cells[3,SGGermP.Row];
                Norm3:=Float(Norm3);
                //---------------------
                Norm4:=SGGermP.Cells[4,SGGermP.Row];
                Norm4:=Float(Norm4);
                //---------------------
                Norm5:=SGGermP.Cells[5,SGGermP.Row];
                Norm5:=Float(Norm5);
                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm5 + #39 +','+
                '[550]='+#39 +Norm1 + #39 +','+
                '[1100]='+#39 +Norm2 + #39 +','+
                '[1530]='+#39 +Norm3 + #39 +','+
                '[2066]='+#39 +Norm4 + #39 +
                ' WHERE ([Лопатка]=' + #39 + Kol + #39+') AND ([Конст]=' + #39 + Konst + #39+')' ,
                ['[Сборка ГермикП]']) then
                Exit;
end;

procedure TForm1.SGGermEditingDone(Sender: TObject);
Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4 ,Norm5:String;
begin
                Kol:=SGGerm.Cells[0,SGGerm.Row];
                Konst:=SGGerm.Cells[7,SGGerm.Row];

                Norm1:=SGGerm.Cells[1,SGGerm.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGGerm.Cells[2,SGGerm.Row];
                Norm2:=Float(Norm2);
                //---------------------
                Norm3:=SGGerm.Cells[3,SGGerm.Row];
                Norm3:=Float(Norm3);
                //---------------------
                Norm4:=SGGerm.Cells[4,SGGerm.Row];
                Norm4:=Float(Norm4);
                //---------------------
                Norm5:=SGGerm.Cells[5,SGGerm.Row];
                Norm5:=Float(Norm5);
                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm5 + #39 +','+
                '[550]='+#39 +Norm1 + #39 +','+
                '[1100]='+#39 +Norm2 + #39 +','+
                '[1530]='+#39 +Norm3 + #39 +','+
                '[2066]='+#39 +Norm4 + #39 +
                ' WHERE ([Лопатка]=' + #39 + Kol + #39+') AND ([Конст]=' + #39 + Konst + #39+')' ,
                ['[ПодСборка ГермикС_П]']) then
                Exit;
end;

procedure TForm1.btn5Click(Sender: TObject);
Var I:Integer;
begin
    SGTulp1.ColCount:=3;
    SGTulp1.Cells[0,0]:=('Кол Лопаток');
    SGTulp1.Cells[1,0]:=('Тюльпан1-В');
    SGTulp1.Cells[2,0]:=('Тюльпан1-Н');
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Тюльпан1]']) then
                        exit;
  SGTulp1.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGTulp1.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка').AsString;
    SGTulp1.Cells[1,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGTulp1.Cells[2,I+1]:=ADOQuery1.FieldByName('НормаН').AsString;
    SGTulp1.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  //-----------------------------------------------------------
    SGTulp2.ColCount:=3;
    SGTulp2.Cells[0,0]:=('Кол Лопаток');
    SGTulp2.Cells[1,0]:=('Тюльпан2-В');
    SGTulp2.Cells[2,0]:=('Тюльпан2-Н');
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Тюльпан2]']) then
                        exit;
  SGTulp2.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGTulp2.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка').AsString;
    SGTulp2.Cells[1,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGTulp2.Cells[2,I+1]:=ADOQuery1.FieldByName('НормаН').AsString;
    SGTulp2.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
    //-----------------------------------------------------------
    SGTulp3.ColCount:=5;
    SGTulp3.Cells[0,0]:=('Кол Лопаток');
    SGTulp3.Cells[1,0]:=('Тюльпан3-Н До 400');
    SGTulp3.Cells[2,0]:=('Тюльпан3-Н Свыше 400');
    SGTulp3.Cells[3,0]:=('Тюльпан3-В До 400');
    SGTulp3.Cells[4,0]:=('Тюльпан3-В Свыше 400');
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Тюльпан3]']) then
                        exit;
  SGTulp3.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGTulp3.Cells[0,I+1]:=ADOQuery1.FieldByName('Лопатка').AsString;
    SGTulp3.Cells[1,I+1]:=ADOQuery1.FieldByName('Норма400').AsString;
    SGTulp3.Cells[2,I+1]:=ADOQuery1.FieldByName('Норма401').AsString;
    SGTulp3.Cells[3,I+1]:=ADOQuery1.FieldByName('Норма400В').AsString;
    SGTulp3.Cells[4,I+1]:=ADOQuery1.FieldByName('Норма401В').AsString;
    SGTulp3.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  //-----------------------------------------------------------
    SGTulp.ColCount:=2;
    SGTulp.Cells[0,0]:=('Размер');
    SGTulp.Cells[1,0]:=('Норма');
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Тюльпан_Переходник]']) then
                        exit;
  SGTulp.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGTulp.Cells[0,I+1]:=ADOQuery1.FieldByName('Размер').AsString;
    SGTulp.Cells[1,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGTulp.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
    //-----------------------------------------------------------
SGTulp.SortByColumn(1);
SGTulp1.SortByColumn(1);
SGTulp2.SortByColumn(1);
SGTulp3.SortByColumn(1);
end;

procedure TForm1.SGTulp1EditingDone(Sender: TObject);

Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4 ,Norm5:String;
begin
                Kol:=SGTulp1.Cells[0,SGTulp1.Row];

                Norm1:=SGTulp1.Cells[1,SGTulp1.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGTulp1.Cells[2,SGTulp1.Row];
                Norm2:=Float(Norm2);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm1 + #39 +','+
                '[НормаН]='+#39 +Norm2 + #39 +
                ' WHERE ([Лопатка]=' + #39 + Kol + #39+') ' ,
                ['[Тюльпан1]']) then
                Exit;
end;

procedure TForm1.SGTulp2EditingDone(Sender: TObject);
Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4 ,Norm5:String;
begin
                Kol:=SGTulp2.Cells[0,SGTulp2.Row];

                Norm1:=SGTulp2.Cells[1,SGTulp2.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGTulp2.Cells[2,SGTulp2.Row];
                Norm2:=Float(Norm2);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm1 + #39 +','+
                '[НормаН]='+#39 +Norm2 + #39 +
                ' WHERE ([Лопатка]=' + #39 + Kol + #39+') ' ,
                ['[Тюльпан2]']) then
                Exit;
end;

procedure TForm1.SGTulp3EditingDone(Sender: TObject);
Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4 ,Norm5:String;
begin
                Kol:=SGTulp3.Cells[0,SGTulp3.Row];

                Norm1:=SGTulp3.Cells[1,SGTulp3.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGTulp3.Cells[2,SGTulp3.Row];
                Norm2:=Float(Norm2);
                //---------------------
                Norm3:=SGTulp3.Cells[3,SGTulp3.Row];
                Norm3:=Float(Norm3);
                //---------------------
                Norm4:=SGTulp3.Cells[4,SGTulp3.Row];
                Norm4:=Float(Norm4);
                //---------------------

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма400]='+#39 +Norm1 + #39 +','+
                '[Норма401]='+#39 +Norm2 + #39 +','+
                '[Норма400В]='+#39 +Norm3 + #39 +','+
                '[Норма401В]='+#39 +Norm4 + #39 +
                ' WHERE ([Лопатка]=' + #39 + Kol + #39+')' ,
                ['[Тюльпан3]']) then
                Exit;
end;

procedure TForm1.SGTulpEditingDone(Sender: TObject);
Var I:Integer;
        Kol,Konst,Norm1,Norm2,Norm3,Norm4 ,Norm5:String;
begin
                Kol:=SGTulp.Cells[0,SGTulp.Row];

                Norm1:=SGTulp.Cells[1,SGTulp.Row];
                Norm1:=Float(Norm1);
                //---------------------
                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm1 + #39 +
                ' WHERE ([Размер]=' + #39 + Kol + #39+') ' ,
                ['[Тюльпан_Переходник]']) then
                Exit;
end;

procedure TForm1.btn6Click(Sender: TObject);
Var I:Integer;
begin
    SGKlara.ColCount:=4;
    SGKlara.Cells[0,0]:=('Высота');
    SGKlara.Cells[1,0]:=('Длина');
    SGKlara.Cells[2,0]:=('Норма');
    SGKlara.Cells[3,0]:=('Примечание');
  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[Клара]']) then
                        exit;
  SGKlara.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGKlara.Cells[0,I+1]:=ADOQuery1.FieldByName('Высота').AsString;
    SGKlara.Cells[1,I+1]:=ADOQuery1.FieldByName('Длина').AsString;
    SGKlara.Cells[2,I+1]:=ADOQuery1.FieldByName('Норма').AsString;
    SGKlara.Cells[3,I+1]:=ADOQuery1.FieldByName('Примечание').AsString;
    SGKlara.AutoSizeRow(i,0);
    ADOQuery1.Next;
  End;
  //-----------------------------------------------------------
  SGKlara.SortByColumn(0);
end;

procedure TForm1.SGKlaraEditingDone(Sender: TObject);
Var I:Integer;
        A,B,Norm1:String;
begin
                A:=SGKlara.Cells[0,SGKlara.Row];

                B:=SGKlara.Cells[1,SGKlara.Row];
                //---------------------
                Norm1:=SGKlara.Cells[2,SGKlara.Row];
                Norm1:=Float(Norm1);

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[Норма]='+#39 +Norm1 + #39 +
                ' WHERE ([Высота]=' + #39 + A + #39+') AND ([Длина]=' + #39 + B + #39+') ' ,
                ['[Клара]']) then
                Exit;
end;

procedure TForm1.btn7Click(Sender: TObject);
Var I:Integer;
begin
    SGUVK.ColCount:=11;
    SGUVK.Cells[0,0]:=('Высота');

    SGUVK.Cells[1,0]:=('<600 р.Прив');
    SGUVK.Cells[2,0]:=('<600 Эл.Прив');

    SGUVK.Cells[3,0]:=('<1000 р.Прив');
    SGUVK.Cells[4,0]:=('<1000 Эл.Прив');

    SGUVK.Cells[5,0]:=('<1500 р.Прив');
    SGUVK.Cells[6,0]:=('<1500 Эл.Прив');

    SGUVK.Cells[7,0]:=('<2000 р.Прив');
    SGUVK.Cells[8,0]:=('<2000 Эл.Прив');

    SGUVK.Cells[9,0]:=('<2100 р.Прив');
    SGUVK.Cells[10,0]:=('<2100 Эл.Прив');


  if not mkQuerySelect(ADOQuery1,
                        'Select * from %s ',
                        ['[УВК_Сборка]']) then
                        exit;
  SGUVK.RowCount:=ADOQuery1.RecordCount+1;
  for I:=0 to ADOQuery1.RecordCount-1 Do
  Begin
    SGUVK.Cells[0,I+1]:=ADOQuery1.FieldByName('Высота').AsString;

    SGUVK.Cells[1,I+1]:=ADOQuery1.FieldByName('600р').AsString;
    SGUVK.Cells[2,I+1]:=ADOQuery1.FieldByName('600э').AsString;

    SGUVK.Cells[3,I+1]:=ADOQuery1.FieldByName('1000р').AsString;
    SGUVK.Cells[4,I+1]:=ADOQuery1.FieldByName('1000э').AsString;

    SGUVK.Cells[5,I+1]:=ADOQuery1.FieldByName('1500р').AsString;
    SGUVK.Cells[6,I+1]:=ADOQuery1.FieldByName('1500э').AsString;

    SGUVK.Cells[7,I+1]:=ADOQuery1.FieldByName('2000р').AsString;
    SGUVK.Cells[8,I+1]:=ADOQuery1.FieldByName('2000э').AsString;

    SGUVK.Cells[9,I+1]:=ADOQuery1.FieldByName('2100р').AsString;
    SGUVK.Cells[10,I+1]:=ADOQuery1.FieldByName('2100э').AsString;
    SGUVK.AutoSizeRow(i,0);
    ADOQuery1.Next;
end;
end;

procedure TForm1.SGUVKEditingDone(Sender: TObject);
Var I:Integer;
        A,B,Norm1,Norm2,Norm3,Norm4,Norm5,Norm6 ,Norm7,Norm8,Norm9,Norm10,
        Kol1,Kol2,Kol3,Kol4,Kol5,Kol6,Kol7,Kol8,Kol9,Kol10:String;
begin
                A:=SGUVK.Cells[0,SGUVK.Row];
                //---------------------
                Norm1:=SGUVK.Cells[1,SGUVK.Row];
                Norm1:=Float(Norm1);
                //---------------------
                Norm2:=SGUVK.Cells[2,SGUVK.Row];
                Norm2:=Float(Norm2);
                //---------------------
                Norm3:=SGUVK.Cells[3,SGUVK.Row];
                Norm3:=Float(Norm3);
                //---------------------
                Norm4:=SGUVK.Cells[4,SGUVK.Row];
                Norm4:=Float(Norm4);
                //---------------------
                Norm5:=SGUVK.Cells[5,SGUVK.Row];
                Norm5:=Float(Norm5);
                //---------------------

                Norm6:=SGUVK.Cells[6,SGUVK.Row];
                Norm6:=Float(Norm6);
                //---------------------
                Norm7:=SGUVK.Cells[7,SGUVK.Row];
                Norm7:=Float(Norm7);
                //---------------------
                Norm8:=SGUVK.Cells[8,SGUVK.Row];
                Norm8:=Float(Norm8);
                //---------------------
                Norm9:=SGUVK.Cells[9,SGUVK.Row];
                Norm9:=Float(Norm9);
                //---------------------
                Norm10:=SGUVK.Cells[10,SGUVK.Row];
                Norm10:=Float(Norm10);
                //---------------------

                if not mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                '[600р]='+#39 +Norm1 + #39 +',[600э]='+#39 +Norm2 + #39 +','+
                '[1000р]='+#39 +Norm3 + #39 +',[1000э]='+#39 +Norm4 + #39 +','+
                '[1500р]='+#39 +Norm5 + #39 +',[1500э]='+#39 +Norm6 + #39 +','+
                '[2000р]='+#39 +Norm7 + #39 +',[2000э]='+#39 +Norm8 + #39 +','+
                '[2100р]='+#39 +Norm9 + #39 +',[2100э]='+#39 +Norm10 + #39 +
                ' WHERE ([Высота]=' + #39 + A + #39+') ' ,
                ['[УВК_Сборка]']) then
                Exit;
end;

procedure TForm1.PageVozdMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
        Case PageVozd.ActivePageIndex  of
        0:Btn1.Click;
        1:Btn3.Click;
        2:Btn5.Click;
        3:Btn6.Click;
        4:Btn7.Click;
        End;
end;

end.
