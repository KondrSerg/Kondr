unit USpec;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, Grids, ExtCtrls, StdCtrls, Buttons, Menus,Clipbrd, DB,
  ADODB,System.Masks,ShellAPI;

type
        TFSpec = class(TForm)
                Panel1: TPanel;
                Panel2: TPanel;
    SpecGrid: TStringGrid;
    Label1: TLabel;
    Label2: TLabel;
    SpeedButton1: TSpeedButton;
    ComboBox1: TComboBox;
    Label3: TLabel;
    Memo1: TMemo;
    Edit1: TEdit;
    Button1: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    Memo2: TMemo;
    N3: TMenuItem;
    N4: TMenuItem;
    lbl1: TLabel;
    btn1: TButton;
    ADOQuery3: TADOQuery;
    ADOConnection: TADOConnection;
    Memo3: TMemo;
    N2: TMenuItem;
    GP: TLabel;
    KO: TLabel;
                procedure FormShow(Sender: TObject);
                procedure SpecGridDblClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpecGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure ComboBox1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpecGridMouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure SpecGridMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure SpecGridDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure Button1Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure SpecGridKeyPress(Sender: TObject; var Key: Char);
    procedure btn1Click(Sender: TObject);
    procedure Memo3DblClick(Sender: TObject);
        private
                { Private declarations }
                TmpCol,TmpRow:Integer;
        public
                { Public declarations }
                AColG,ARowG,Fl,FL2,Flex:Integer;
        end;

var
        FSpec: TFSpec;

implementation

uses Main, UDetal, UNewNakl, Unit4, mainBRAK;

{$R *.dfm}

procedure TFSpec.FormShow(Sender: TObject);
var
        I,Y,j: Integer;
        S,GP,KLAP:String;
begin
FNewNakl.Clear_StringGrid(SpecGrid);
S := ExtractFileDir(ParamStr(0));
        Memo3.Lines.Clear;
        Memo1.Lines.Clear;
        Memo1.Lines.LoadFromFile(S + '\SpecGrid.ini');
        SpecGrid.ColCount := Memo1.Lines.Count;
        for I := 0 to Memo1.Lines.Count - 1 do
                SpecGrid.ColWidths[i] := StrToInt(Memo1.Lines.Strings[i]);
        Fl:=StrToInt(Label2.Caption);
        SpecGrid.ColCount := I_FN_NOM1 + 25;
        SpecGrid.Cells[0, 0] := '№';

        SpecGrid.Cells[I_FN_ZAK1, 0] := FN_ZAK1;
        SpecGrid.Cells[I_FN_POS_KOMP, 0] := 'ВидЭлемента';
        SpecGrid.Cells[I_FN_POSS, 0] := 'Элемент';
        SpecGrid.Cells[I_FN_OBOZN, 0] := 'КолНаЕд';

        SpecGrid.Cells[I_FN_TYP_S, 0] := 'Количество';
        SpecGrid.Cells[I_FN_NOM_NOM, 0] := 'ЕИ';
        SpecGrid.Cells[I_FN_ZBOR_ED, 0] :='Обозначение';
        SpecGrid.Cells[I_FN_ZBOR_ED, 0] := 'Обозначение';
        SpecGrid.Cells[I_FN_MODEL, 0] := 'Диаметр';
        SpecGrid.Cells[I_FN_KATEG, 0] := 'Масса';
        SpecGrid.Cells[I_FN_KOL_S, 0] := 'Длина';
        SpecGrid.Cells[I_FN_SPRAM, 0] := 'Ширина';
        SpecGrid.Cells[I_FN_UKRUG, 0] := 'ДлинаРазв';
        SpecGrid.Cells[I_FN_KPD, 0] := 'ШиринаРазв';
        SpecGrid.Cells[I_FN_KOMPL_W, 0] := 'Объем';
        SpecGrid.Cells[I_FN_NOM_GUR, 0] := 'КолГибов';
        SpecGrid.Cells[I_FN_NOM1, 0] := 'ТИП';
        SpecGrid.Cells[I_FN_NOM1+1, 0] := 'Материал';
        SpecGrid.Cells[I_FN_NOM1+2, 0] := 'IdГП';
        SpecGrid.Cells[I_FN_NOM1 + 4, 0] := 'Н\ч Гибка';
        SpecGrid.Cells[I_FN_NOM1 + 5, 0] := 'Н\ч Ножницы';
        SpecGrid.Cells[I_FN_NOM1 + 6, 0] := 'Н\ч Пила';
        SpecGrid.Cells[I_FN_NOM1 + 7, 0] := 'Дата';
        SpecGrid.Cells[I_FN_NOM1 + 8, 0] := 'idКо';
        SpecGrid.Cells[I_FN_NOM1 + 9, 0] := 'ШиринаПолки1';
        SpecGrid.Cells[I_FN_NOM1 + 10, 0] := 'ШиринаПолки2';
        SpecGrid.Cells[I_FN_NOM1 + 11, 0] := 'Покрытие';
                        if FL=0 Then
                        Begin
                        if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where [IdГП]=' +
                        #39+Label1.Caption+#39+' AND ([IdКО]=' +
                          #39+Lbl1.Caption+#39+')  Order By ВидЭлемента ', ['Специф']) then
                        exit;
                        End;
                        if Fl=1 Then
                        Begin
                          if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where ([IdГП]=' +
                          #39+Label1.Caption+#39+') AND ([IdКО]=' +
                          #39+Lbl1.Caption+#39+') Order By ВидЭлемента  ', ['СпецифВозд']) then
                          exit;
                        end;
                        if Fl=3 Then
                        Begin
                          if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where ([IdГП]=' +
                          #39+Label1.Caption+#39+') Order By ВидЭлемента  ', ['Специф750']) then
                          exit;
                        end;
                        if Fl=4 Then
                        Begin
                          if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where ([IdГП]=' +
                          #39+Label1.Caption+#39+') Order By ВидЭлемента  ', ['Специф710']) then
                          exit;
                        end;
                          I:=0;
                  if Form1.ADOQuery1.RecordCount = 0 then
                  begin
                    SpecGrid.RowCount := 2;
                    Form1.Clear_StringGrid(SpecGrid);
                    exit;
                  end;

                              SpecGrid.RowCount := SpecGrid.RowCount+Form1.ADOQuery1.RecordCount + 1;
                                Form1.ADOQuery1.First;
                                for J := 0 to Form1.ADOQuery1.RecordCount - 1 do
                                begin
                                  SpecGrid.Cells[0, I + 1] := Form1.ADOQuery1.FieldByName('id').AsString ;//IntToStr(I + 1);
                                SpecGrid.Cells[I_FN_ZAK1, I + 1] :=
                                Form1.ADOQuery1.FieldByName('Технолог').AsString;
                                SpecGrid.Cells[I_FN_POS_KOMP, I + 1] :=
                                Form1.ADOQuery1.FieldByName('ВидЭлемента').AsString;
                                SpecGrid.Cells[I_FN_POSS, I + 1] :=
                                Form1.ADOQuery1.FieldByName('Элемент').AsString;
                                SpecGrid.Cells[I_FN_OBOZN, I + 1] :=
                                Form1.ADOQuery1.FieldByName('КолНаЕд').AsString;

                                SpecGrid.Cells[I_FN_TYP_S, I + 1] :=
                                Form1.ADOQuery1.FieldByName('Количество').AsString;
                                SpecGrid.Cells[I_FN_NOM_NOM, I + 1] :=
                                Form1.ADOQuery1.FieldByName('ЕИ').AsString;
                                  SpecGrid.Cells[I_FN_ZBOR_ED, I + 1] :=
                                    Form1.ADOQuery1.FieldByName('Обозначение').AsString;

                                  SpecGrid.Cells[I_FN_MODEL, I + 1] :=
                                Form1.ADOQuery1.FieldByName('Диаметр').AsString;
                                SpecGrid.Cells[I_FN_KATEG, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('Масса').AsString;
                                SpecGrid.Cells[I_FN_KOL_S, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('Длина').AsString;
                                SpecGrid.Cells[I_FN_SPRAM, I + 1] :=
                                Form1.ADOQuery1.FieldByName('Ширина').AsString;
                                  SpecGrid.Cells[I_FN_UKRUG, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('ДлинаРазв').AsString;
                                  SpecGrid.Cells[I_FN_KPD, I + 1] :=
                                Form1.ADOQuery1.FieldByName('ШиринаРазв').AsString;
                                  SpecGrid.Cells[I_FN_KOMPL_W, I + 1] :=
                                Form1.ADOQuery1.FieldByName('Объем').AsString;
                                SpecGrid.Cells[I_FN_NOM_GUR, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('КолГибов').AsString;
                                SpecGrid.Cells[I_FN_NOM1, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('Тип').AsString;
                                SpecGrid.Cells[I_FN_NOM1 + 1, I + 1] :=
                                    Form1.ADOQuery1.FieldByName('Материал').AsString;
                                SpecGrid.Cells[I_FN_NOM1 + 2, I + 1] :=
                                    Form1.ADOQuery1.FieldByName('IdГП').AsString;
                                if Fl<>3 then
                                Begin
                                SpecGrid.Cells[I_FN_NOM1 + 4, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('Н\ч Гибка').AsString;
                                SpecGrid.Cells[I_FN_NOM1 + 5, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('Н\ч Ножницы').AsString;
                                  if Fl<>4 then

                                SpecGrid.Cells[I_FN_NOM1 + 6, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('Н\ч Пила').AsString;
                                End;
                                SpecGrid.Cells[I_FN_NOM1 + 7, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('Дата').AsString;
                                SpecGrid.Cells[I_FN_NOM1 + 8, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('idКо').AsString;
                                SpecGrid.Cells[I_FN_NOM1 + 9, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('ШиринаПолки1').AsString;
                                SpecGrid.Cells[I_FN_NOM1 + 10, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('ШиринаПолки2').AsString;
                                SpecGrid.Cells[I_FN_NOM1 + 11, I + 1] :=
                                  Form1.ADOQuery1.FieldByName('Покрытие').AsString;

                                  Inc(I);
                                Form1.ADOQuery1.Next;
                                end;



        //---------------------------------------------------------------



       { SpecGrid.RowCount := Form1.ADOQuery1.RecordCount + 1;
        Form1.ADOQuery1.First;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                SpecGrid.Cells[0, I + 1] :=  Form1.ADOQuery1.FieldByName('id').AsString ;//IntToStr(I + 1);
                //if Form1.Spec=1 Then

                SpecGrid.Cells[I_FN_ZAK1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Технолог').AsString ;
               // Else
                //SpecGrid.Cells[I_FN_ZAK1, I + 1] :=
                //        Form1.ADOQuery1.FieldByName(FN_ZAK1).AsString;
                SpecGrid.Cells[I_FN_POS_KOMP, I + 1] :=
                        Form1.ADOQuery1.FieldByName('ВидЭлемента').AsString;
                SpecGrid.Cells[I_FN_POSS, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Элемент').AsString;
                SpecGrid.Cells[I_FN_OBOZN, I + 1] :=
                        Form1.ADOQuery1.FieldByName('КолНаЕд').AsString;

                SpecGrid.Cells[I_FN_TYP_S, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Количество').AsString;
                SpecGrid.Cells[I_FN_NOM_NOM, I + 1] :=
                        Form1.ADOQuery1.FieldByName('ЕИ').AsString;
               if Form1.Spec=1 Then
               Begin
                        SpecGrid.Cells[I_FN_ZBOR_ED, I + 1] :=
                        Form1.ADOQuery1.FieldByName('ОбозначениеШ').AsString;
               end
               else
                    SpecGrid.Cells[I_FN_ZBOR_ED, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Обозначение').AsString;

                SpecGrid.Cells[I_FN_MODEL, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Диаметр').AsString;
                SpecGrid.Cells[I_FN_KATEG, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Масса').AsString;
                SpecGrid.Cells[I_FN_KOL_S, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Длина').AsString;
                SpecGrid.Cells[I_FN_SPRAM, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Ширина').AsString;
                SpecGrid.Cells[I_FN_UKRUG, I + 1] :=
                        Form1.ADOQuery1.FieldByName('ДлинаРазв').AsString;
                SpecGrid.Cells[I_FN_KPD, I + 1] :=
                        Form1.ADOQuery1.FieldByName('ШиринаРазв').AsString;
                SpecGrid.Cells[I_FN_KOMPL_W, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Объем').AsString;
                SpecGrid.Cells[I_FN_NOM_GUR, I + 1] :=
                        Form1.ADOQuery1.FieldByName('КолГибов').AsString;
                SpecGrid.Cells[I_FN_NOM1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Тип').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Материал').AsString;
               { SpecGrid.Cells[I_FN_NOM1 + 2, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Trumph').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 3, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Ножницы').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 4, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Углоруб').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 5, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Угол').AsString;

                SpecGrid.Cells[I_FN_NOM1 + 6, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Гибка').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 7, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Прокатка').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 8, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Пила').AsString;
                //SpecGrid.Cells[I_FN_NOM1 + 9, I + 1] :=
                //        Form1.ADOQuery1.FieldByName('Пила ленточная').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 10, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Сварка').AsString;
                //SpecGrid.Cells[I_FN_NOM1 + 11, I + 1] :=
                //        Form1.ADOQuery1.FieldByName('IdГП').AsString;

                //SpecGrid.Cells[I_FN_NOM1 + 12, I + 1] :=
                //        Form1.ADOQuery1.FieldByName('HACO').AsString;
                {SpecGrid.Cells[I_FN_NOM1 + 13, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Вальцы').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 14, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Кромкогиб').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 15, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Токарный').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 16, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Ротор').AsString;  }

                //SpecGrid.Cells[I_FN_NOM1 + 20, I + 1] :=
                //        Form1.ADOQuery1.FieldByName('Дата').AsString;
                //SpecGrid.Cells[I_FN_NOM1 + 21, I + 1] :=
                //        Form1.ADOQuery1.FieldByName('IdКо').AsString;
                {SpecGrid.Cells[I_FN_NOM1 + 22, I + 1] :=
                        Form1.ADOQuery1.FieldByName('ШиринаПолки1').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 23, I + 1] :=
                        Form1.ADOQuery1.FieldByName('ШиринаПолки2').AsString;  }
               // Form1.ADOQuery1.Next;
       // end;
end;

procedure TFSpec.Memo3DblClick(Sender: TObject);
var S:string;
FileHandle,i,y: Integer;
F:file;
begin
 // I:= SendMessage(Memo3.Handle, EM_LINEINDEX, word(-1), 0);
  Y := SendMessage(Memo3.Handle, EM_LINEFROMCHAR, Memo3.SelStart,0);
  i := Memo3.SelStart - SendMessage(Memo3.Handle, EM_LINEINDEX,Y, 0);
  S:=Memo3.Lines.Strings[y];
  AssignFile(f,S);
  //Reset(f, S);
 // ShellExecute(Form1.Handle, nil, PChar(S), nil, nil, SW_SHOWNORMAL);
  FileHandle:= ShellExecute(0,'open', PChar(S), nil, nil, SW_SHOW);
end;

function FindFiles(StartFolder, Mask,D: string;ii:Integer; List: TStrings; ScanSubFolders: Boolean = True):String;
var
SearchRec: TSearchRec;
FindResult: Integer;
fullFilePath,ss,S1,Dir:String;
begin
    Dir:=D;
    List.BeginUpdate;
    try
      StartFolder := IncludeTrailingBackslash(StartFolder);
      FindResult := FindFirst(StartFolder + '*.*', faAnyFile, SearchRec);
      try
      while FindResult = 0 do
        with SearchRec do
        begin
          if (Attr and faDirectory) <> 0 then
          begin
            if ScanSubFolders and (Name <> '.') and (Name <> '..') then
              FindFiles(StartFolder + Name, Mask,D,ii, List, ScanSubFolders);
          end
          else
          begin
            if MatchesMask(Name, Mask) then  //StartFolder +
            begin
              List.Add(StartFolder +Name);
              fullFilePath:=StartFolder + Name;
              S1:=Name;
             { if fullFilePath<>'' then
              begin
              ss:=FormatDateTime('ss:zzz', Now);
              SS:=StringReplace(ss, ':', '_',[rfReplaceAll, rfIgnoreCase]);
              // Теперь проверяем существует ли файл
              if FileExists(D + '\Оси\'+S1) then
              Begin
              List.Add(fullFilePath);
                //CopyFile(PWideChar(fullFilePath),
                PWideChar(D + '\Оси\'+S1+'_'+SS), True) //существует'
                End
              else
                CopyFile(PWideChar(fullFilePath),
                PWideChar(D + '\Оси\'+S1), True);
              end; }
              //CopyFile(PWideChar(fullFilePath), PWideChar(D + Name), False);
            end;
          end;
          FindResult := FindNext(SearchRec);
        end;
      finally
        FindClose(SearchRec);
      end;
    finally
      List.EndUpdate;
    end;
end;

procedure TFSpec.Button1Click(Sender: TObject);
Var I,F:Integer;
        Nom_pos,Nom_pos2,Fl:Integer;
        Str,Dir_Ser,Put,fileName,Dir,fullFilePath:string;
begin
    //Ось ТЕКИ 269.15.00.009
{    Dir_Ser :=  '\\Mss-dc06\v\КТО\Кутдусов\';
    Dir :=  '\\Mss-dc06\v\ОИТ\Cklapana2';
    Str:=SpecGrid.Cells[SpecGrid.Col,SpecGrid.Row ];
    Nom_pos:=Pos('ТЕКИ',Str);
    if Nom_pos<>0 then
    begin
      Delete(Str,1,Nom_pos+3);
      Str:='*Ось*';
    end;
//
    Nom_pos:=Pos('ВГ',Str);
    if Nom_pos<>0 then
    begin
      Delete(Str,1,Nom_pos+1);
      Str:='*Ось*';
    end;
    Str:='*Ось*.*';
//Str:='FilesIndex*';
    if Str<> '' then
    Begin
      fileName := Str ;
      fullFilePath := FindFiles(Dir_Ser,fileName ,Dir,i, Memo3.Lines, True); }



end;


procedure TFSpec.SpecGridDblClick(Sender: TObject);
var
        Nom_pos,Nom_pos2,Fl,i:Integer;
        Str,Dir_Ser,Put,fileName,Dir,fullFilePath:string;
begin
    //Ось ТЕКИ 269.15.00.009
    Dir_Ser :=  '\\Mss-dc06\v\КТО\';
    Str:=SpecGrid.Cells[SpecGrid.Col,SpecGrid.Row ];
    //
    Nom_pos:=Pos('-',Str);
    if Nom_pos<>0 then
    begin
      Delete(Str,Nom_pos,300);
    end;

    Nom_pos:=Pos('ТЕКИ',Str);
    if Nom_pos<>0 then
    begin
      Delete(Str,1,Nom_pos+3);
      Str:=Str+'*';
    end;
//
    Nom_pos:=Pos('ВГ',Str);
    if Nom_pos<>0 then
    begin
      Delete(Str,1,Nom_pos+1);
      Str:='*'+Str+'*';
    end;

//Str:='FilesIndex*';
    if Str<> '' then
    Begin
      fileName := Str ;
      Memo3.Lines.Add(Str);
      fullFilePath := FindFiles(Dir_Ser,fileName ,Dir,i, Memo3.Lines, True);
      Memo3.Lines.Add('Ok');
    end;

        if (Form1.FlagDolg=1) OR (Form1.FlagDolg=2) Then
        Begin

        End;

end;

procedure TFSpec.SpeedButton1Click(Sender: TObject);
begin
Form1.ExportGridtoExcel1(SpecGrid);
end;

procedure TFSpec.SpecGridSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
  Var R,Rect1:TRect;
begin
 TmpCol:=ACol;
        TmpRow:=ARow;
end;

procedure TFSpec.ComboBox1Click(Sender: TObject);
Var Fl:Integer;
begin
        Fl:=StrToInt(Label2.Caption);
        if Form1.Spec=0 Then
        Begin
                if Fl=0 Then
                Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' + Label3.Caption + ']=' +
                #39 + ComboBox1.Text+ #39 +
                ' WHERE ([Элемент]=' + #39 + SpecGrid.Cells[I_FN_POSS, ARowG] + #39+') and ([IdГП]=' + #39 + SpecGrid.Cells[I_FN_NOM1+11, ARowG] + #39+' )',
                ['Специф']) then
                Exit;

                end;
                if Fl=1 Then
                Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' + Label3.Caption + ']=' +
                #39 + ComboBox1.Text+ #39 +
                ' WHERE ([Элемент]=' + #39 + SpecGrid.Cells[I_FN_POSS, ARowG] + #39+') and ([IdГП]=' + #39 + SpecGrid.Cells[I_FN_NOM1+11, ARowG] + #39+') and ([IdКО]=' + #39 + SpecGrid.Cells[I_FN_NOM1+15, ARowG] + #39+' )',
                ['СпецифВозд']) then
                Exit;

                end;
        end;
        if Form1.Spec=1 Then
        Begin
                if Fl=0 Then
                Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' + Label3.Caption + ']=' +
                #39 + ComboBox1.Text+ #39 +
                ' WHERE ([ОбозначениеШ]=' + #39 + SpecGrid.Cells[I_FN_ZBOR_ED,ARowG] + #39+')',
                ['Шаблон']) then
                Exit;

                end;
                if Fl=1 Then
                Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' + Label3.Caption + ']=' +
                #39 + ComboBox1.Text+ #39 +
                ' WHERE ([ОбозначениеШ]=' + #39 + SpecGrid.Cells[I_FN_ZBOR_ED,ARowG] + #39+')',
                ['ШаблонВоз']) then
                Exit;

                end;

        end;
                SpecGrid.Cells[AColG,SpecGrid.Row]:=ComboBox1.Items[ComboBox1.ItemIndex];
end;

procedure TFSpec.FormClose(Sender: TObject; var Action: TCloseAction);
Var
I:Integer;
S:String;
begin
S := ExtractFileDir(ParamStr(0));
Memo1.Lines.Clear;
        Form1.Spec:=0;
        FSpec.fl2:=0;
        for I := 0 to SpecGrid.ColCount - 2 do
                Memo1.Lines.Add(IntToStr(SpecGrid.ColWidths[i]));
        Memo1.Lines.SaveToFile(S + '\SpecGrid.ini');
end;

procedure TFSpec.SpecGridMouseWheelDown(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
        Handled := true;
        SpecGrid.Perform(WM_VSCROLL, SB_LINEDOWN, 0);
end;

procedure TFSpec.SpecGridMouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
        Handled := true;
        SpecGrid.Perform(WM_VSCROLL, SB_LINEUP, 0);
end;

procedure TFSpec.SpecGridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
  Var Str:String;
  res:Integer;
begin
        Str:=SpecGrid.Cells[I_FN_ZAK1, ARow];
        Res:=AnsiCompareStr('TRUE',Str);
        if  Res=0 then
        begin
                                SpecGrid.canvas.brush.Color := RGB(100, 149,
                                        237); //Отгруженные
                                SpecGrid.Canvas.Font.Color := clBlack;
                                SpecGrid.Canvas.FillRect(Rect);
                                SpecGrid.Canvas.TextOut(Rect.Left, Rect.Top,
                                SpecGrid.Cells[ACol, ARow]);
        end;
end;


procedure TFSpec.N1Click(Sender: TObject);
Var Str1:String;
begin
        Str1 := SpecGrid.Cells[TmpCol, tmpRow];
        Memo2.Lines.Clear;
        Memo2.Lines.Add(Str1);
        Clipboard.AsText := Memo2.Lines.Strings[0];
end;

procedure TFSpec.N3Click(Sender: TObject);
begin
       // SpecGrid.RowCount := SpecGrid.RowCount + 1;
       // Form1.FSelectRow(SpecGrid, SpecGrid.RowCount - 1);
        Form4.SpecGrid.Cells[I_FN_NOM1+10, 1]:=SpecGrid.Cells[I_FN_NOM1+10, 1];
        Form4.SpecGrid.Cells[I_FN_ZAK1, 1]:=SpecGrid.Cells[I_FN_ZAK1, 1];
        Form4.SpecGrid.Cells[I_FN_NOM1 + 13, 1]:=SpecGrid.Cells[I_FN_NOM1 + 13, 1];
        Form4.Show;
end;

procedure TFSpec.N4Click(Sender: TObject);
begin
  //Form1.DeleteARow(SpecGrid, SpecGrid.Row);
 { Fl:=StrToInt(Label2.Caption);
   if (MessageDlg('Ты собираешься удалить!'+SpecGrid.Cells[I_FN_POSS,SpecGrid.Row]+'  '+SpecGrid.Cells[I_FN_ZBOR_ED,SpecGrid.Row],mtInformation,[mbYes,MbNo],0)=mrYes) Then
  Begin
  if Fl=0 Then
    Begin

   if not Form1.mkQueryDelete( Form1.ADOQuery1, 'DELETE FROM %s Where  (ВидЭлемента= '+#39+SpecGrid.Cells[I_FN_POS_KOMP,SpecGrid.Row]+#39+
  ') AND (Элемент= '+#39+SpecGrid.Cells[I_FN_POSS,SpecGrid.Row]+#39+
  ') AND (Обозначение= '+#39+SpecGrid.Cells[I_FN_ZBOR_ED,SpecGrid.Row]+#39+
  ') AND (Длина= '+#39+SpecGrid.Cells[I_FN_KOL_S,SpecGrid.Row]+#39+
  ') AND (Ширина= '+#39+SpecGrid.Cells[I_FN_SPRAM,SpecGrid.Row]+#39+
  ') AND (ДлинаРазв= '+#39+SpecGrid.Cells[I_FN_UKRUG,SpecGrid.Row]+#39+')'
   ,
         ['Специф'] )
         Then
         Exit;
    end;
   if Fl=1 Then
    Begin

  if not Form1.mkQueryDelete( Form1.ADOQuery1, 'DELETE FROM %s Where '+
  ' (ВидЭлемента= '+#39+SpecGrid.Cells[I_FN_POS_KOMP,SpecGrid.Row]+#39+
  ') AND (Элемент= '+#39+SpecGrid.Cells[I_FN_POSS,SpecGrid.Row]+#39+
  ') AND (Обозначение= '+#39+SpecGrid.Cells[I_FN_ZBOR_ED,SpecGrid.Row]+#39+
  ') AND (Длина= '+#39+SpecGrid.Cells[I_FN_KOL_S,SpecGrid.Row]+#39+
  ') AND (Ширина= '+#39+SpecGrid.Cells[I_FN_SPRAM,SpecGrid.Row]+#39+
  ') AND (ДлинаРазв= '+#39+SpecGrid.Cells[I_FN_UKRUG,SpecGrid.Row]+#39+')'
   ,
         ['ШаблонВоз'] )
         Then
         Exit;

    end;
  end;    }

end;

procedure TFSpec.N2Click(Sender: TObject);
Var id,i,x,IZ_ID,Kol,Kol_kl,n,r,R1,Res:Integer;
  Zak,BZ,Dat,Izd,Det,Oboz,GP,KO, Nom_G,Detali,Izdel,Zav,Fam,Opis,Plan1,
  str,Uch,Odobren,Vid:string;
begin

         if not FBRAK.mkQuerySelect(FBRAK.ADOQuery1,
                'Select  max(ID) as S from %s' ,
                ['БРАК']) then
                exit;
        id:= FBRAK.ADOQuery1.FieldByName('S').AsInteger+1;
        N:=0;
        R:=Form1.zclrstrngrd1.Row;
        R1:=SpecGrid.Row;
        GP :=FSpec.Gp.Caption;
        KO :=FSpec.ko.Caption;
        Izdel:=Form1.zclrstrngrd1.Cells[I_FN_NAM,r];//FSpec.Caption;
        IZ_ID:=7;
        Izd :='750';
         { zclrstrngrd1.Cells[I_FN_DAT, 0] := FN_DAT;
  zclrstrngrd1.Cells[I_FN_ZAK, 0] := FN_ZAK;

  zclrstrngrd1.Cells[I_FN_NAM, 0] := FN_NAM;
  zclrstrngrd1.Cells[I_FN_KOL, 0] := FN_KOL;
  zclrstrngrd1.Cells[I_FN_KOL_ZAP, 0] := FN_KOL_ZAP;  }
        Zak:= Form1.zclrstrngrd1.Cells[I_FN_ZAK,r];
        BZ :='';// Form1.zclrstrngrd1.Cells[N+3,r];
        Zav:= Form1.zclrstrngrd1.Cells[I_FN_KOL_ZAP + 23,r];
        Fam:=Form1.UserName1;
        Dat:= FormatDateTime('mm.dd.YYYY', StrToDate(Form1.zclrstrngrd1.Cells[I_FN_DAT,r]));
        Kol_kl:= StrToInt(Form1.zclrstrngrd1.Cells[I_FN_KOL_ZAP,r]);
        {        SpecGrid.Cells[I_FN_POSS, 0] := 'Элемент';
        SpecGrid.Cells[I_FN_OBOZN, 0] := 'КолНаЕд';

        SpecGrid.Cells[I_FN_TYP_S, 0] := 'Количество';
        SpecGrid.Cells[I_FN_NOM_NOM, 0] := 'ЕИ';
        SpecGrid.Cells[I_FN_ZBOR_ED, 0] :='Обозначение';
        SpecGrid.Cells[I_FN_POS_KOMP, 0] := 'ВидЭлемента';}
      Vid:=SpecGrid.Cells[ I_FN_POS_KOMP, r1 ];
      res:=Pos('Детал',Vid);
      if Res=0 then
      Begin
       MessageDlg('Только Детали!', mtError, [mbOk], 0);
       Exit;
      End;
        Det :=SpecGrid.Cells[ I_FN_POSS, r1 ];
        Oboz:=SpecGrid.Cells[ I_FN_ZBOR_ED, r1 ];
        Kol :=StrToInt(SpecGrid.Cells[ I_FN_TYP_S, r1 ]);
        Str:='Детали&'+Det +'||'+Oboz+'~'+IntToStr(Kol*Kol_kl)+';';
        Uch:='Сборка';
        Opis:='';

        Nom_G:= Form1.zclrstrngrd1.Cells[N,r];

        //Детали&Стенка горизонт.под привод||ТЕКИ 269.15.01.001-850~1
//;Детали&Стенка вертикальная||ТЕКИ 269.15.01.007-450~1

//;
        Odobren:='1';
        Plan1:=FormatDateTime('mm.dd.YYYY', StrToDate( Form1.zclrstrngrd1.Cells[I_FN_KOL_ZAP + 3,r]));
        Odobren:='1';

        if not FBRAK.mkQuerySelect(FBRAK.ADOQuery1,
                'Select  * from %s where (idgp=%s) AND  (Номер=%s) AND  (Планирование=%s)' ,
                ['БРАК',#39 + Gp + #39,#39 + Nom_G + #39,#39 + Plan1 + #39]) then
                exit;
        if FBRAK.ADOQuery1.RecordCount<>0  then
        begin
              Str:=Str+#10#13+FBRAK.ADOQuery1.FieldByName('КТО').AsString;
              if not FBRAK.mkQueryUpdate(FBRAK.ADOQuery1, 'UPDATE %s SET [КТО]='+
                        #39+'%s'+ #39+
                        ' WHERE (idgp=%s) AND  (Номер=%s) AND  (Планирование=%s)',
                        ['Брак',Str,#39 + Gp + #39,#39 + Nom_G + #39,#39 + Plan1 + #39])
                        then Exit;
        end
        Else

        if not FBRAK.mkQueryInsert(FBRAK.ADOQuery1, 'Insert Into %s ' +
        '(Заказ,БЗ,ЗавНом,Фамилия,Дата,[Тип продукции],Деталь ,'+
        'Чертеж ,Кол ,Участок ,Описание,idgp,idko,Издел,ФамилияОТК,Tab,КТО,Планирование,Номер,Одобрено,Flag)' +
        'Values (%s,%s,%s,%s,%s,%s,%s,'+
        '%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',
        ['Брак', #39 + Zak + #39, #39 + BZ + #39, #39 + Zav
         + #39, #39 + Fam + #39, #39 + Dat + #39, #39 + Izd
        + #39, #39 + Det + #39, #39 + Oboz + #39, #39 + IntToSTr(Kol) + #39, #39 + Uch + #39
        ,  #39 + Opis+ #39, #39 + GP + #39,#39 + KO + #39,#39 + Izdel + #39,#39 + '*' + #39,#39 +
         IntToSTr(IZ_ID)+ #39,#39 + Str + #39,#39 + Plan1 + #39,#39 + Nom_G + #39,#39 + Odobren + #39,#39 + '1' + #39],'1') then
        exit;

end;

procedure TFSpec.SpecGridKeyPress(Sender: TObject; var Key: Char);
Var Str,GP,KO:String;
begin
       { if (Key = #13) Then
        Begin
          Edit1.Text:=IntToStr(SpecGrid.Col)+'  '+ IntToStr(SpecGrid.Row)+'  '+ SpecGrid.Cells[SpecGrid.Col,SpecGrid.Row];
          Fl:=StrToInt(Label2.Caption);
                if Form1.Spec=0 Then
                Begin
                GP:=Label1.Caption;
                KO:=Lbl1.Caption;
                        if Fl=0 Then
                        Begin
                          //StringGrid10.Cells[I_FN_SGP + 2, 0] := 'IdГП';
                          //SpecGrid.Cells[I_FN_ZBOR_ED, 0] :='Обозначение';
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' + SpecGrid.Cells[SpecGrid.Col, 0] + ']=' +
                        #39 +SpecGrid.Cells[SpecGrid.Col,SpecGrid.row ] + #39 +
                        ' WHERE ([Обозначение]=' + #39 + SpecGrid.Cells[I_FN_ZBOR_ED, SpecGrid.row] + #39+') and ([IdГП]=' + #39 + GP + #39+' )',
                        ['Специф']) then
                        Exit;

                        end;
                        if Fl=1 Then
                        Begin
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET ['+SpecGrid.Cells[SpecGrid.Col, 0]+ ']=' +
                        #39 +SpecGrid.Cells[SpecGrid.Col,SpecGrid.row ]+ #39 +
                        ' WHERE ([Обозначение]=' + #39 + SpecGrid.Cells[I_FN_ZBOR_ED,SpecGrid.row]+#39+') and ([IdГП]=' + #39 + GP + #39+') and ([IdКО]=' + #39 + KO + #39+' )',
                        ['СпецифВозд']) then
                        Exit;

                        end;
                end;
        end;}
end;

procedure TFSpec.btn1Click(Sender: TObject);
Var I,F,Kol_Vo,Dl:Integer;
S1,S2,S3,Err,GP,Oboz,Elem,D:String;
Temp:TADOQuery;
begin
if ADOConnection = nil then
        ADOConnection := TADOConnection.Create(Self);
        ADOConnection.LoginPrompt := False;
        ADOConnection.Connected := False;
        ADOConnection.ConnectionString := Form1.MSSQL_CONN_STR_Glob;
        try
                ADOConnection.Connected := True;

                Temp:=TADOQuery.Create(nil);
                Temp.Connection:=ADOConnection;
        except
                MessageDlg('Не удалось подключиться к базе данных1!', mtError,
                        [mbOk], 0);
                Close;
        end;

        Temp.SQL.Text:='CREATE TABLE #Client (Kol int,'+
        'Klapan nvarchar(250),GP nvarchar(250),Dl int)';
        try
        Temp.ExecSQL;
        except
                Err:='1222';
        end;
                                      // AND (Элемент=' +#39+'Лопатка'+#39')
        Form1.Clear_StringGrid(SpecGrid);
        F:=StrToInt(Label2.Caption);
        S1:='(Изделие Like ' +#39+'%ГЕРМИК-ДУ-'+'%'+#39+')';
        S2:='(ВидЭлемента= ' +#39+'Сборочные единицы'+#39')';
        S3:='(Элемент =' +#39+'Лопатка'+#39+')';
        if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where ( Дата> '+
        #39+'12.31.2016'+#39+') AND %s  '+
        ' AND %s AND %s ', ['Специф',S1,S2,S3]) then  //Order By Технолог
        exit;
        SpecGrid.ColCount := I_FN_NOM1 + 24;
        SpecGrid.Cells[0, 0] := '№';

        if Form1.ADOQuery1.RecordCount = 0 then
        begin
                SpecGrid.RowCount := 1;
                Form1.Clear_StringGrid(SpecGrid);
                exit;
        end;

         SpecGrid.RowCount := Form1.ADOQuery1.RecordCount + 1;
        Form1.ADOQuery1.First;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
               if not Form1.mkQuerySelect1(Form1.ADOQuery2, 'Select * from %s Where ( IdГП= '+
        #39+Form1.ADOQuery1.FieldByName('IdГП').AsString+#39')', ['Klapana']) then  //Order By Технолог
        exit;


               F:=StrToInt( Form1.ADOQuery1.FieldByName('Количество').AsString);
               D:= Form1.ADOQuery1.FieldByName('Длина').AsString ;
               if D<>'' Then
                Dl:=StrToInt(D );
               Kol_Vo:=StrToInt( Form1.ADOQuery2.FieldByName('Кол во').AsString);

               Oboz:= Form1.ADOQuery1.FieldByName('Обозначение').AsString;
               Elem:=Form1.ADOQuery1.FieldByName('IdГП').AsString ;
                Temp.SQL.Text:='INSERT INTO #Client '+
                '(Kol,Klapan,GP,Dl'+
                ') Values ('
                +IntToStr(F*Kol_Vo)+
                ','+#39+Oboz+#39+
                ','+#39+Elem+#39+
                ','+#39+IntToStr(Dl*F*Kol_Vo)+#39+
                ')';

                try
                Temp.ExecSQL;
                except
                Err:='1222';

                end;

                Form1.ADOQuery1.Next;
        end;
        Temp.Close;                                      // (Kol>' +#39+'1'+#39')
        Temp.SQL.Text:=
          'Select Klapan,Sum(Kol) AS SS,Sum(Dl) AS DD from #Client Where (Kol>' +#39+'1'+#39')  Group By Klapan';
        Temp.Open;
       SpecGrid.RowCount := Temp.RecordCount + 1;
        Temp.First;
        for I := 0 to Temp.RecordCount - 1 do
        begin
                                SpecGrid.Cells[0, I + 1] := IntToStr(I + 1);

                SpecGrid.Cells[I_FN_TYP_S, I + 1] :=
                        Temp.FieldByName('SS').AsString;

               SpecGrid.Cells[I_FN_ZBOR_ED, I + 1] :=
                       Temp.FieldByName('Klapan').AsString;
               SpecGrid.Cells[I_FN_ZBOR_ED+1, I + 1] :=
                       Temp.FieldByName('DD').AsString;
                Temp.Next;
        End;


end;

end.
