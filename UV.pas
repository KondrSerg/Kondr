unit UV;

interface

uses
        Windows, Messages, SysUtils, Variants,ExcelXP, Classes, Graphics, Controls,
        Forms,
        Dialogs, Grids, ExtCtrls, StdCtrls, ComCtrls, Math, ComObj, Menus,UConnCeh,UConnKlap,UOsnova,
  AdvObj, BaseGrid, AdvGrid;

type
        TFV = class(TForm)
                Panel1: TPanel;
                Panel2: TPanel;
                Panel3: TPanel;
                SGS: TStringGrid;
                Button1: TButton;
                Button2: TButton;
                Label1: TLabel;
                Label2: TLabel;
                Label3: TLabel;
                Label4: TLabel;
                Label5: TLabel;
                Label6: TLabel;
                DateTimePicker1: TDateTimePicker;
                Label7: TLabel;
                Memo1: TMemo;
                Button3: TButton;
                ListBox1: TListBox;
                Memo2: TMemo;
    Edit1: TEdit;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    Image2: TImage;
    Label8: TLabel;
    Image5: TImage;
    Label61: TLabel;
    Image1: TImage;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    btn1: TButton;
    Memo3: TMemo;
    SG: TAdvStringGrid;
    Memommo1: TMemo;
    Label16: TLabel;
    Label17: TLabel;
    Edit2: TEdit;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
                procedure FormShow(Sender: TObject);
                procedure SGSSelectCell(Sender: TObject; ACol, ARow: Integer;
                        var CanSelect: Boolean);
                procedure SGSSetEditText(Sender: TObject; ACol, ARow: Integer;
                        const Value: string);
                procedure Button2Click(Sender: TObject);
                procedure Button1Click(Sender: TObject);
                procedure FormClose(Sender: TObject; var Action: TCloseAction);
                procedure DateTimePicker1Enter(Sender: TObject);
                procedure Button3Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure SGSDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure DateTimePicker1Change(Sender: TObject);
        private
                { Private declarations }
                Enabled_Ok1: Integer;
                Enabled_Ok2: Integer;
        public
                { Public declarations }
                SUT:Integer;
                Puty:String;
                Conn_Ceh:Connect_Miass_Ceh;
                Conn_Klap:Connect_Miass_Klap;
                UOsnova_Main:Osnova_Main ;
                Flag_Update, //0  Запустили ВСЕ,  1 Запустили ЧАСТЬ  2 Запустили ПОСЛЕДНИЕ
                Flag_Insert,Cvet,Cvet1:Integer;
        end;

var
        FV: TFV;

implementation

uses Main;

{$R *.dfm}

procedure TFV.FormShow(Sender: TObject);
var
        I, Kol, Kol_Zap, Max: Integer;
        S,Tab1,Zapusk: string;
begin
        Tab1:=Label23.Caption; // Klapana
        Zapusk:=Label22.Caption;
        SyStem.SysUtils.FormatSettings.DecimalSeparator :=(',');
        if Form1.FlagDolg=1 Then
        Begin
            Btn1.Visible := True;
            Edit1.Visible:=True;
        end;
        DateTimePicker1.DateTime:=Now;
        Button1.Enabled := False;
        Enabled_Ok1 := 0;
        Enabled_Ok2 := 0;
        Label2.Caption := '0';
        Label4.Caption := '0';
        Label6.Caption := '0';
        S := ExtractFileDir(ParamStr(0));
        Memo1.Lines.Clear;
        Memo1.Lines.LoadFromFile(S + '\SGS.ini');
        SGS.ColCount := Memo1.Lines.Count;
        for I := 0 to Memo1.Lines.Count - 2 do
        begin
                SGS.ColWidths[i] := StrToInt(Memo1.Lines.Strings[i]);
        end;
        for i:=0 To SGS.RowCount Do
        Begin
            SGS.Rows[i].Clear;
        End;
        SGS.ColCount := 46;
        SGS.Cells[0, 0] := '№';
        SGS.Cells[1, 0] := FN_PLAN_DATA;
        SGS.Cells[2, 0] := FN_ZAK;
        SGS.Cells[3, 0] := FN_KOL;
        //SGS.Cells[4,0]:=FN_KOL_ZAP;
        SGS.Cells[4, 0] := 'Свободно';
        SGS.Cells[5, 0] := 'Кол-во на запуск';
        SGS.Cells[6, 0] := FN_SVARKA_NC + ' на единицу';
        SGS.Cells[7, 0] := FN_SBOR_KLAP_NC + ' на единицу';
        SGS.Cells[8, 0] := 'Всего ' + FN_SVARKA_NC;
        SGS.Cells[9, 0] := 'Всего ' + FN_SBOR_KLAP_NC;
        SGS.Cells[10, 0] := 'ID';
        SGS.Cells[11, 0] := 'Кол запущенных';
        SGS.Cells[12, 0] := 'Кол Раз';
        SGS.Cells[13, 0] := 'Номер Раз';
        SGS.Cells[14, 0] := 'Privod';
        SGS.Cells[15, 0] := 'Дата';
        SGS.Cells[16, 0] := 'Изделие';
        SGS.Cells[17, 0] := 'A';
        SGS.Cells[18, 0] := 'B';
        SGS.Cells[19, 0] := FN_TEHNOLOG;
        SGS.Cells[20, 0] := 'НачНомер';
        SGS.Cells[21, 0] := 'КонНомер';
        SGS.Cells[22, 0] := 'НачКод';
        SGS.Cells[23, 0] := 'КонКод';//Заводск номер из 20 ячейкй
        SGS.Cells[24, 0] := 'Кол приводов(общее)';
        SGS.Cells[25, 0] := 'IdГП';
        SGS.Cells[26, 0] := 'НомерПос';
        SGS.Cells[27, 0] := 'КодПос';

        SGS.Cells[28, 0] := 'ПДО';
        SGS.Cells[29, 0] := 'КТО';

        SGS.Cells[30, 0] := 'Сборка лопаток';
        SGS.Cells[31, 0] := 'Сборка тяг';
        SGS.Cells[32, 0] := 'Всего Сборка лопаток';
        SGS.Cells[33, 0] := 'Всего Сборка тяг';
        SGS.Cells[34, 0] := 'Исполнение';

        SGS.Cells[35, 0] := 'BZ';
        SGS.Cells[36, 0] := 'Brik';
        SGS.Cells[37, 0] := 'легион';
        SGS.Cells[38, 0] := 'bz';
        SGS.Cells[39, 0] := 'Технолог Прим';
        SGS.Cells[40, 0] := 'Расчетная дата готовности';
         SGS.Cells[41, 0] := 'IdКО';
         SGS.Cells[42, 0] := 'СборкаРЕШ';
         SGS.Cells[44, 0] := 'СборкаРЕШ Всего';
         SGS.Cells[45, 0] := 'Электро';
        SGS.ColWidths[3] := 0;

        SGS.ColWidths[10] := 0;
        SGS.ColWidths[11] := 0;
        SGS.ColWidths[12] := 0;
       // SGS.ColWidths[14] := 40;
        SGS.ColWidths[17] := 0;
        SGS.ColWidths[18] := 0;

        SGS.ColWidths[19] := 0;
        SGS.ColWidths[20] := 0;
        SGS.ColWidths[21] := 0;
        SGS.ColWidths[22] := 0;
        SGS.ColWidths[23] := 0;
        SGS.ColWidths[24] := 0;
        SGS.ColWidths[25] := 0;
        SGS.ColWidths[26] := 0;
        SGS.ColWidths[27] := 0;

        SGS.ColWidths[29] := 0;
        SGS.ColWidths[30] := 0;
        SGS.ColWidths[31] := 0;
        SGS.ColWidths[32] := 0;
        SGS.ColWidths[33] := 0;
        SGS.ColWidths[34] := 100;

        if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from [%s] Where   (NOT [' + FN_TEHNOLOG
        + '] IS NULL) AND (['  +   FN_KOL_ZAP + ']<[' + FN_KOL + ']) AND'+
        ' ([' + FN_SBOR_KLAP_NC +
        ']<>0)  AND ([Отмена] IS NULL) AND ((Х IS NULL) or (Х=0)) Order By Заказ ', [Tab1]) then
                exit;
       //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
       //ЗАГОТОВКА
        {if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where  (NOT [' + FN_Mod_Privod +
                '] IS NULL) AND ([' + FN_SBOR_KLAP_NC +
                        ']<>0)  AND ([Отмена] IS NULL) AND (Х IS NULL) AND (NOT [' + FN_TEHNOLOG
                        + '] IS NULL) Order By Заказ ', ['Klapana']) then
                exit; }
        SGS.RowCount := Form1.ADOQuery1.RecordCount + 1;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol := Form1.ADOQuery1.FieldByName(FN_KOL).AsInteger;
                Kol_Zap := Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsInteger;
                SGS.Cells[0, i + 1] := IntToStr(I + 1);
                SGS.Cells[1, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_PLAN_DATA).AsString;
                SGS.Cells[2, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZAK).AsString;
                SGS.Cells[3, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL).AsString;

                Kol := Kol - Kol_Zap;
                SGS.Cells[4, i + 1] := IntToStr(Kol);
                SGS.Cells[6, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SVARKA_NC).AsString;
                SGS.Cells[7, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SBOR_KLAP_NC).AsString;
                SGS.Cells[10, i + 1] :=
                        Form1.ADOQuery1.FieldByName('ID').AsString;
                SGS.Cells[11, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsString;
                SGS.Cells[12, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_RAZ).AsString;
                SGS.Cells[13, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOMER_RAZ).AsString;
                SGS.Cells[14, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_MOD_PRIVOD).AsString;
                SGS.Cells[15, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_DAT).AsString;
                SGS.Cells[16, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NAM).AsString;

                SGS.Cells[17, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_A).AsString;
                SGS.Cells[18, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_B).AsString;
                SGS.Cells[19, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_TEHNOLOG).AsString;
                SGS.Cells[20, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НачНомер').AsString;
                SGS.Cells[21, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КонНомер').AsString;
                SGS.Cells[22, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НачКод').AsString;
                SGS.Cells[23, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КонКод').AsString;
                SGS.Cells[24, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Привод').AsString;
                SGS.Cells[25, i + 1] :=
                        Form1.ADOQuery1.FieldByName('IdГП').AsString;
                SGS.Cells[26, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НомерПос').AsString;
                SGS.Cells[27, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КодПос').AsString;

                SGS.Cells[28, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SGP).AsString;
                SGS.Cells[29 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('№Поз').AsString;

                SGS.Cells[30 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Сборка лопаток').AsString;
                SGS.Cells[31 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Сборка тяг').AsString;
               { SGS.Cells[32 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Гибка').AsString;
                SGS.Cells[33 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Трумпф').AsString; }
                SGS.Cells[34 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Исполнение').AsString;
                SGS.Cells[35 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('БЗ').AsString;
                SGS.Cells[36 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Брикет').AsString;
                SGS.Cells[37 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Заказчик').AsString;
                SGS.Cells[38 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('bz').AsString;
                SGS.Cells[39 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Технолог Прим').AsString+' '+
                        Form1.ADOQuery1.FieldByName(FN_SBOR_PRIM).AsString;
                SGS.Cells[40 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Расчетная дата готовности').AsString;
                SGS.Cells[41 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('IdКО').AsString;
                SGS.Cells[42 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('СборкаРЕШ').AsString;
               SGS.Cells[45 , I + 1] :=
                        Form1.ADOQuery1.FieldByName('Электро').AsString;
                Form1.ADOQuery1.Next;

        end;
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select MAX(Номер) as a from [%s] ',
                [Tab1]) then
                exit;
        Max := Form1.ADOQuery1.FieldValues['a'];    //
        Label4.Caption := IntToStr(Max+1); //+ZZZZZZZZZZZZZZZZZZ
end;

procedure TFV.SGSSelectCell(Sender: TObject; ACol, ARow: Integer;
        var CanSelect: Boolean);
begin
        if ACol = 5 then
        begin
                SGS.Options := SGS.Options + [goEditing];

        end
        else
                SGS.Options := SGS.Options - [goEditing];
end;

procedure TFV.SGSSetEditText(Sender: TObject; ACol, ARow: Integer;
        const Value: string);
var
        I, Kol_Zap, Kol_Zap_R, Kol: Integer;
        Svar, Sbor, Svar_o, Sbor_o,
        Sb_Lop,Sb_Tyag,Sb_lop_O,Sb_Tyag_O,Resh,Resh_O,
        Nog,Pila,Gib,Trumpf,Kanban,
        Nog_o,Pila_o,Gib_O,Trumpf_o,Kanban_O: Double;
        Pdo,Kto:String;
begin
        I := ACol;
        if SGS.Cells[5, ARow] = '' then
                Exit;

        SyStem.SysUtils.FormatSettings.DecimalSeparator :=(',');
        PDO:=SGS.Cells[28, ARow];
        KTO:=SGS.Cells[29, ARow];
        Kol_Zap := StrToInt(SGS.Cells[5, ARow]);
        Kol_Zap_R := StrToInt(SGS.Cells[11, ARow]);
        Kol := StrToInt(SGS.Cells[3, ARow]);
        if SGS.Cells[6, ARow] = '' then
                Svar := 0
        else
                Svar := StrToFloat(SGS.Cells[6, ARow]);
        if SGS.Cells[7, ARow] = '' then
                Sbor := 0
        else
                Sbor := StrToFloat(SGS.Cells[7, ARow]);
        //---------------------
        if SGS.Cells[30, ARow] = '' then
                Sb_Lop := 0
        else
                Sb_Lop := StrToFloat(SGS.Cells[30, ARow]);

        if SGS.Cells[31, ARow] = '' then
                Sb_Tyag := 0
        else
                Sb_Tyag := StrToFloat(SGS.Cells[31, ARow]);

        if SGS.Cells[42, ARow] = '' then
                Resh := 0
        else
                Resh := StrToFloat(SGS.Cells[42, ARow]);

        Sb_Lop:= RoundTo(Sb_Lop * Kol_Zap, -3);
        Sb_Tyag:= RoundTo(Sb_Tyag * Kol_Zap, -3);
        Resh   := RoundTo(Resh * Kol_Zap, -3);
        SGS.Cells[32, ARow] := FloatToStr(Sb_Lop);
        SGS.Cells[33, ARow] := FloatToStr(Sb_Tyag);
        SGS.Cells[44, ARow] := FloatToStr(Resh);
        //SGS.Cells[42, 0] := 'СборкаРЕШ';
       //  SGS.Cells[44, 0] := 'СборкаРЕШ Всего';


        Svar := RoundTo(Svar * Kol_Zap, -3);
        Sbor := RoundTo(Sbor * Kol_Zap, -3);
        SGS.Cells[8, ARow] := FloatToStr(Svar);
        SGS.Cells[9, ARow] := FloatToStr(Sbor);

        Svar_o := 0;
        Sbor_o := 0;
        Sb_lop_O := 0;
        Sb_Tyag_o := 0;

        for I := 1 to SGS.RowCount do
        begin
                if SGS.Cells[8, I] <> '' then
                begin
                        Kol_Zap := StrToInt(SGS.Cells[5, ARow]);
                        Kol_Zap_R := StrToInt(SGS.Cells[4, ARow]);
                        //++++++++++++++++++++++++++++++++++++++++
                       { if ((PDO<>'') AND (KTO='')) Then
                        Begin
                                 If Kol_Zap>1 then
                                 Begin
                                 MessageDlg('Можно запустить только 1 клапан (Первый Прогон) !',
                                        mtError, [mbOk], 0);
                                SGS.Cells[5, ARow] := '';
                                Exit;

                                end;
                        end;
                        //++++++++++++++++++++++++++++++++++++++++
                        if ((PDO<>'') AND (KTO<>'')) Then
                        Begin
                                 If Kol_Zap>1 then
                                 Begin
                                 MessageDlg('Можно запустить только 1 клапан (Второй Прогон) !',
                                        mtError, [mbOk], 0);
                                 SGS.Cells[5, ARow] := '';
                                Exit;

                                end;
                        end;
                        //++++++++++++++++++++++++++++++++++++++++
                        if ((PDO='') AND (KTO<>'')) Then
                        Begin
                                 MessageDlg('Спроси мастера, клапан косячный!',
                                        mtError, [mbOk], 0);
                                SGS.Cells[5, ARow] := '';
                                Exit;
                        end; }
                        Kol := Kol_Zap_R - Kol_Zap;
                        //++++++++++++++++++++++++++++++++++++++++
                        if Kol < 0 then
                        begin
                                MessageDlg('Можно запустить только ' +
                                        SGS.Cells[4, ARow] + ' !',
                                        mtError, [mbOk], 0);
                                SGS.Cells[5, ARow] := '';
                                Exit;
                        end;
                        if SGS.Cells[5, i] = '' then
                                Kol_Zap := 0
                        else
                                Kol_Zap := StrToInt(SGS.Cells[5, i]);
                        if Kol_Zap = 0 then
                        begin
                                Svar_o := Svar_o - StrToFloat(SGS.Cells[8, I]);
                                Sbor_o := Sbor_o - StrToFloat(SGS.Cells[9, I]);

                                Sb_Lop_o := Sb_Lop_o - StrToFloat(SGS.Cells[32, I]);
                                Sb_Tyag_o := Sb_Tyag_o - StrToFloat(SGS.Cells[33, I]);
                                Resh := Resh - StrToFloat(SGS.Cells[44, I]);
                        end;
        //SGS.Cells[44, ARow] := FloatToStr(Resh);
        //SGS.Cells[42, 0] := 'СборкаРЕШ';
       //  SGS.Cells[44, 0] := 'СборкаРЕШ Всего';

                        if Kol_Zap <> 0 then
                        begin
                                Svar_o := RoundTo(Svar_o +
                                        StrToFloat(SGS.Cells[8, I]), -3);
                                Sbor_o := RoundTo(Sbor_o +
                                        StrToFloat(SGS.Cells[9, I]), -3);
                                Sb_Lop_o := Sb_Lop_o + StrToFloat(SGS.Cells[32, I]);
                                Sb_Tyag_o := Sb_Tyag_o + StrToFloat(SGS.Cells[33, I]);
                                Resh := Resh + StrToFloat(SGS.Cells[44, I]);
                        end;
                end;
        end;
        Label2.Caption := FloatToStr(Svar_o);
        Label6.Caption := FloatToStr(Sbor_o);

        Label20.Caption := FloatToStr(Sb_Lop_o);
        Label21.Caption := FloatToStr(Sb_Tyag_o);
        Label26.Caption := FloatToStr(Resh);
end;

procedure TFV.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFV.Button1Click(Sender: TObject);
var
        I, J,U,uu,ff, ID, Pos1, Kol_Zap, Kol_Zap_R, Kol, Rec_Count, Kol_Ob, Res, e:
        Integer;
        Svar, Sbor, Svar_o, Sbor_o,Sb_lop_O,Sb_Tyag_O,Sb_Resh_O,El_O,EL: Double;
        Dat, Dat1, Kol_Raz, Nomer_Raz, IdGP, Zak, S_Svar, S_Sbor, Dat_zak,
        Dat_Z, ss,Privod, Vn_Dat,D, Dir, S, Teh_Gotov, Teh_Obrab, Plan_Dat,
        God,Mes,Shtrih_Str,Shtrih2_Str,Priv_Str,PDO,KTO,fileName,A,B,
        Elem,Vid,Oboz,Mat,Kol_S,Stan,D1,Dir1,BZ,Brik,bz1,Sb_Lop,Sb_Tyag,Sb_Resh,Raskl,R_Dat,Up,Tab1,Zapusk,Spec,EL_S: string;
        XL2: Variant;
        Ii,Kol_Sh,Kol_pr,NachNom,KonNom,NomPos,Shtrih2,Shtrih,ff1,ff2: Integer;
        NachKod,KonKOd,KodPos:Int64;
        Res_Nog,Res_Trumph,Res_Gib,Res_Pila,Kanban,Val,Svark,Pokr,Rotor,Kol_Kl:Integer;
Doc: OleVariant;
Appl, wdDoc, wdRang,Sheet1,Sheet2,Rang:variant;
Leg,KO:String;
      {  FV.Label22.Caption := 'Запуск';
  FV.Label23.Caption := 'Klapana';
  FV.Label24.Caption := 'Специф';
  FV.Label25.Caption := '0';}
begin
        e := 1;
        U:=0;
        Tab1:='['+Label23.Caption+']';
        Zapusk:='['+Label22.Caption+']';
        Spec:='['+Label24.Caption+']';
        SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
        Form1.Clear_StringGrid(SG);
        //Label4.Caption:='11587';
        if (MessageDlg('Проверь, все ли клапана? Потом не добавишь!',mtInformation,[mbYes,MbNo],0)=mrYes) Then
        Begin
        //Label4.Caption:=Edit2.Text;
        Kol_Ob := 0;
        Vn_Dat := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
        D := FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);

        for I := 1 to SGS.RowCount do
        begin
                if SGS.Cells[10, i] <> '' then
                begin
                        ID := StrToInt(SGS.Cells[10, I]);
                        Leg:= (SGS.Cells[37, I]);
                        bz1:=(SGS.Cells[38, I]);
                        if (SGS.Cells[7, I] <> '') and (ID <> 0) and
                                (SGS.Cells[5, i] <> '') then
                        begin
                                PDO:=SGS.Cells[28, SGS.Row];
                                KTO:=SGS.Cells[29, SGS.Row];

                                A:=SGS.Cells[17, i];
                                B:=SGS.Cells[18, i];
                                BZ :=SGS.Cells[35, i];
                                Brik:=SGS.Cells[36, i];
                                UP:= SGS.Cells[34, i];
                                KO:= SGS.Cells[41, i];
                                Nomer_Raz := '';
                                Kol_Raz := '';
                                Kol_Zap := StrToInt(SGS.Cells[5, i]);
                                if  SGS.Cells[24, i]='' Then
                                Kol_Pr:=0
                                else
                                Kol_pr := StrToInt(SGS.Cells[24, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[22, i]='' Then
                                NachKod:=0
                                Else
                                NachKod:= StrToInt64(SGS.Cells[22, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[23, i]='' Then
                                KonKod:=0
                                Else
                                KonKod:= StrToInt64(SGS.Cells[23, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[27, i]='' Then
                                KodPos:=0
                                Else
                                KodPos:= StrToInt64(SGS.Cells[27, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[20, i]='' Then
                                NachNom:=0
                                Else
                                NachNom:= StrToInt(SGS.Cells[20, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++

                                if  SGS.Cells[21, i]='' Then
                                KonNom:=0
                                Else
                                KonNom:= StrToInt(SGS.Cells[21, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[26, i]='' Then
                                NomPos:=0
                                Else
                                NomPos:= StrToInt(SGS.Cells[26, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                idgp:=SGS.Cells[25, i];
                                Plan_Dat := Form1.ConvertDat1(SGS.Cells[1, i]);
                                if SGS.Cells[40, i]='' then
                                R_Dat:='NULL'
                                else
                                R_Dat := #39+Form1.ConvertDat1(SGS.Cells[40, i])+#39;
                                Teh_Gotov := Form1.ConvertDat1(SGS.Cells[19,
                                        i]);
                                if Kol_Zap = 0 then
                                        Continue;
                                Kol_Ob := Kol_Ob + Kol_Zap;
                                Kol_Zap_R := StrToInt(SGS.Cells[11, i]);
                                Kol := StrToInt(SGS.Cells[3, i]);
                                Privod := (SGS.Cells[14, i]);

                                Dat_Z := SGS.Cells[15, i];
                                res := Pos('.', Dat_z);
                                if Res <> 0 then
                                begin
                                        SS := Copy(Dat_Z, 1, Res);
                                        Delete(Dat_Z, 1, Res);
                                        Insert(SS, Dat_z, 4);
                                end;
                                Zak := SGS.Cells[2, i];
                                Dat_Zak := Dat_z;
                                Dat := FormatDateTime('mm.dd.yyyy', Now);
                                Dat1 := FormatDateTime('mm.dd.yyyy',
                                        DateTimePicker1.Date);

                                if Kol <> Kol_Zap then
                                begin
                                        if SGS.Cells[12, i] = '' then
                                                Kol_Raz := IntToStr(Kol_Zap)
                                        else
                                                Kol_Raz := SGS.Cells[12, i] + ','
                                                        + IntToStr(Kol_Zap);
                                        if SGS.Cells[13, i] = '' then
                                                Nomer_Raz := Label4.Caption
                                        else
                                                Nomer_Raz := SGS.Cells[13, i] +
                                                        ',' + Label4.Caption;
                                end;
                                if Kol = Kol_Zap then
                                begin
                                        Kol_Raz := IntToStr(Kol_Zap);
                                        Nomer_Raz := Label4.Caption;
                                end;
                                if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                        'Select * from %s Where  (Заказ= ' + #39
                                        + Zak + #39 + ') AND ([IdГП]= ' +
                                        #39 + (IDgp) + #39 + ') AND ([IdКО]= ' +
                                        #39 + (KO) + #39 + ')',
                                        [Zapusk]) then
                                        exit;
                                S_Svar := SGS.Cells[8, I];
                                S_Sbor := SGS.Cells[9, I];
                                Res := Pos(',', S_Svar);
                                Delete(S_Svar, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Svar, Res);

                                Res := Pos(',', S_Sbor);
                                Delete(S_Sbor, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Sbor, Res);
                                //
                                Sb_LOp := SGS.Cells[32, I];
                                Sb_Tyag := SGS.Cells[33, I];
                                Sb_Resh:= SGS.Cells[44, I];

                                Res := Pos(',', Sb_LOp);
                                Delete(Sb_LOp, Res, 1);
                                if Res <> 0 then
                                        Insert('.', Sb_LOp, Res);

                                Res := Pos(',', Sb_Tyag);
                                Delete(Sb_Tyag, Res, 1);
                                if Res <> 0 then
                                        Insert('.', Sb_Tyag, Res);

                               Res := Pos(',', Sb_Resh);
                                Delete(Sb_Resh, Res, 1);
                                if Res <> 0 then
                                Insert('.', Sb_Resh, Res);

                                if SGS.Cells[45, I]='' then
                                  el_S:='0'
                                Else
                                el_S:=SGS.Cells[45, I];
                                Res := Pos(',', el_S);
                                Delete(el_S, Res, 1);
                                if Res <> 0 then
                                Insert('.', el_S, Res);
                                EL:=StrToFloat(EL_S);
                                EL_O:=EL*Kol_ZAP;

                                Rec_Count := Form1.ADOQuery2.RecordCount;
                                Priv_Str:=IntToStr(Kol_pr*Kol_Zap);
                                if (Rec_Count = 0) and (Kol_Zap = Kol) then
                                        //Запустили все
                                begin
                                  NomPos:=KonNom;
                                  KodPos:=KonKod;
                                end;
                        //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                                if (Rec_Count = 0) and (Kol_Zap <> Kol) then //Запустили не все
                                begin
                                  //+++++++++++++++++++++++++++++++++++++++++++++
                                  NomPos:=NachNom +Kol_Zap-1;
                                  KodPos:=NachKod +Kol_Zap-1;
                                  //----------------------------------------------
                                  KonNom:=NachNom+Kol_Zap-1;
                                  KonKod:=NachKod+Kol_Zap-1;
                                end;
                                if (Rec_Count <> 0)  then
                                        //Запустили последнии (в запуске уже были)
                                begin

                                  NachKod:= KodPos+1;
                                  //+++++++++++++++++++++++++++++++++++++++++++++
                                  KonKod:= NachKod+Kol_Zap-1;
                                  //+++++++++++++++++++++++++++++++++++++++++++++
                                  NachNom:= NomPos+1;
                                  //+++++++++++++++++++++++++++++++++++++++++++++
                                  KonNom:= NachNom+Kol_Zap-1;
                                  //+++++++++++++++++++++++++++++++++++++++++++++
                                  NomPos:= NachNom+Kol_Zap-1;
                                  KodPos:= NachKod+Kol_Zap-1;
                                end;
                                  if not Form1.mkQueryInsert(Form1.ADOQuery2,
                                  'Insert Into %s ' +
                                  '([Дата],[план Дата],Заказ,Изделие,[Кол во],'+
                                  '[Кол во запущенных],[' + FN_RAS_DATA_GOTOVN+ '],Номер,[Н\ч Сварка],'+
                                  '[Н\ч Сборка Клапана],'+
                                  '[МодПривода],НачНомер,КонНомер,НачКод,КонКод,Привод,IdГП,IdКО,ПДО,'+
                                  'КТО,Планирование,A,B,Cvet,ФлагЗаготовки,БЗ,Брикет,Заказчик,bz,'+
                                  '[Сборка лопаток],[Сборка тяг],[СборкаРЕШ],[Сборка Примечание],Исполнение,Электро) ' +
                                  'Values (%s,%s,%s,%s,%s ,%s,' +
                                  '%s,%s,%s,%s,%s ,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s ,%s,%s,%s,%s)',
                                  [Zapusk,
                                  #39 + Dat_Zak + #39,
                                  #39 + Plan_dat + #39,
                                  #39 + SGS.Cells[2, i] +#39,
                                  #39 + SGS.Cells[16, i] +#39,
                                  #39 + SGS.Cells[3, i] +#39,
                                  #39 + IntToStr(Kol_Zap)+ #39,
                                   R_Dat ,
                                  #39 + Label4.Caption + #39,
                                  #39 + S_Svar + #39,
                                  #39 + S_Sbor + #39,
                                  #39 + Privod + #39,
                                  #39 + IntToStr(NachNom) +#39,
                                  #39 + IntToStr(KonNom) +#39,
                                  #39 + IntToStr(NachKod) +#39,
                                  #39 + IntToStr(KonKod) +#39,
                                  #39 + Priv_Str +#39,
                                  #39 + IdGP +#39,
                                  #39 + KO +#39,
                                  #39 + PDO +#39,
                                  #39 + KTO +#39,
                                  #39 + Dat1 + #39,#39 + (A) +#39,
                                  #39 + (B) +#39,
                                  #39 + IntToStr(Cvet1) +#39,
                                  #39 + '1' +#39,
                                  #39 + BZ +#39,
                                  #39 + Brik +#39,
                                  #39 + (Leg) +#39,#39 + (bz1) +#39,
                                  #39 + Sb_Lop + #39,
                                  #39 + Sb_Tyag + #39,
                                  #39 + Sb_Resh + #39,
                                  #39 + SGS.Cells[39, i] +#39  ,
                                  #39 + UP +#39,
                                  #39 + FloatToStr(el_o) +#39
                                  ])
                                  then  exit;

                                Form1.Memo2.Lines.Add(S_Sbor);
                                Inc(e);
                                //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        end;
                end;
        end;
        //ЗАГОТОВКА

                UOsnova_Main:=Osnova_Main.Create() ;
                fileName := '';//ExtractFileDir(ParamStr(0));//+'\Klapan.EXE';
                Vn_Dat := FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
                UOsnova_Main.Flag_Error :=0;
                Form1.Memo2.Lines.Add('Osnova');
                SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
                 UOsnova_Main.Osnova( Vn_Dat,Zapusk,Spec,#39+Label4.Caption+#39,fileName, 1,Cvet1);
                //++++++++++++++++++++++++++++++++++++++++++++++++++
                if  UOsnova_Main.Flag_Error =2 Then //ДЕТАЛИ НЕ ВСЕ
                //ЗАНЕСТИ В ЗАПУСК И Поставить флаг ФлагЗаготовки и убрать Дату Планирования
                Begin
                        //++++++++++++++++++++++++++++++++++++++++++++++++++++
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [ФлагЗаготовки]=' + #39
                        + '0' + #39+ ',Планирование=NULL WHERE [Номер]=' + #39 +
                        (Label4.Caption) + #39, [Zapusk])
                        then
                        Exit;
                        Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДЕТАЛИ НЕ ВСЕ,Запуск,Специф,'+Label4.Caption+','+fileName+' )UOsnova_Main.Flag_Error =2') ;
                end;
                //++++++++++++++++++++++++++++++++++++++++++++++++++
                if  UOsnova_Main.Flag_Error =1 Then //ДАТА ЛЕВАЯ УДАЛИТЬ Из ЗАПУСКА
                                                //И UPDATE КЛАПАНА НЕ ДЕЛАТЬ
                Begin
                        if not Form1.mkQueryDelete( Form1.ADOQuery1, 'DELETE FROM %s Where (Номер= '
                        +#39+Label4.Caption+#39+') '  ,
                        [Zapusk] )
                        Then
                        Exit;
                        Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДАТА НЕ ТА - '+Vn_Dat+',Запуск,Специф,'+Label4.Caption+','+fileName+' )UOsnova_Main.Flag_Error =1') ;
                        //Exit;
                end;
                Dir1:='Суточные задания';
                pdo:='';
                Form1.SUT1(D,Label4.Caption,Dir1,StrToInt(Label25.Caption),Cvet1,pdo);//ZZZZZZZZZZZZZZZZZZZZZZ
                //++++++++++++++++++++++++++++++++++++++++++++++++++
                if  (UOsnova_Main.Flag_Error =0) OR (UOsnova_Main.Flag_Error =2) Then
                Begin
                  for I := 1 to SGS.RowCount do
                  begin
                    if SGS.Cells[10, i] <> '' then
                    begin
                      ID := StrToInt(SGS.Cells[10, I]);
                      if (SGS.Cells[7, I] <> '') and (ID <> 0) and
                      (SGS.Cells[5, i] <> '') then
                      begin
                                PDO:=SGS.Cells[28, SGS.Row];
                                KTO:=SGS.Cells[29, SGS.Row];
                                Nomer_Raz := '';
                                Kol_Raz := '';
                                Kol_Zap := StrToInt(SGS.Cells[5, i]);

                                //+++++++++++++++++++++++++++++++++++++++++++++
                                idgp:=SGS.Cells[25, i];
                                KO  :=SGS.Cells[41, i];
                                Plan_Dat := Form1.ConvertDat1(SGS.Cells[1, i]);
                                                                if SGS.Cells[40, i]='' then
                                R_Dat:='NULL'
                                else
                                R_Dat := #39+Form1.ConvertDat1(SGS.Cells[40, i])+#39;
                                Teh_Gotov := Form1.ConvertDat1(SGS.Cells[19,
                                        i]);
                                //Teh_Obrab := Form1.ConvertDat1(SGS.Cells[20,
                                //        i]);
                                if Kol_Zap = 0 then
                                        Continue;
                                Kol_Ob := Kol_Ob + Kol_Zap;
                                Kol_Zap_R := StrToInt(SGS.Cells[11, i]);
                                Kol := StrToInt(SGS.Cells[3, i]);
                                Privod := (SGS.Cells[14, i]);
                                SGS.Cells[11, i] := IntToStr(Kol_Zap +
                                        Kol_Zap_R);
                                SGS.Cells[4, i] := IntToStr(Kol - (Kol_Zap +
                                        Kol_Zap_R));
                                Dat_Z := SGS.Cells[15, i];
                                res := Pos('.', Dat_z);
                                if Res <> 0 then
                                begin
                                        SS := Copy(Dat_Z, 1, Res);
                                        Delete(Dat_Z, 1, Res);
                                        Insert(SS, Dat_z, 4);
                                end;
                                Zak := SGS.Cells[2, i];
                                Dat_Zak := Dat_z;
                                Dat := FormatDateTime('mm.dd.yyyy', Now);
                                Dat1 := FormatDateTime('mm.dd.yyyy',
                                        DateTimePicker1.Date);

                                if Kol <> Kol_Zap then
                                begin
                                        if SGS.Cells[12, i] = '' then
                                                Kol_Raz := IntToStr(Kol_Zap)
                                        else
                                                Kol_Raz := SGS.Cells[12, i] + ','
                                                        + IntToStr(Kol_Zap);
                                        if SGS.Cells[13, i] = '' then
                                                Nomer_Raz := Label4.Caption
                                        else
                                                Nomer_Raz := SGS.Cells[13, i] +
                                                        ',' + Label4.Caption;
                                end;
                                if Kol = Kol_Zap then
                                begin
                                        Kol_Raz := IntToStr(Kol_Zap);
                                        Nomer_Raz := Label4.Caption;
                                end;
                                if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                        'Select * from %s Where  (Заказ= ' + #39
                                        + Zak + #39 + ') AND ([IdГП]= ' +
                                        #39 + (IDgp) + #39 + ') AND ([IdКо]= ' +
                                        #39 + (KO) + #39 + ')',
                                        [Zapusk]) then
                                        exit;
                                S_Svar := SGS.Cells[8, I];
                                S_Sbor := SGS.Cells[9, I];
                                Res := Pos(',', S_Svar);
                                Delete(S_Svar, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Svar, Res);

                                Res := Pos(',', S_Sbor);
                                Delete(S_Sbor, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Sbor, Res);
                                //
                                Sb_LOp := SGS.Cells[32, I];
                                Sb_Tyag := SGS.Cells[33, I];
                                Sb_Resh := SGS.Cells[44, I];
                                Res := Pos(',', Sb_LOp);
                                Delete(Sb_LOp, Res, 1);
                                if Res <> 0 then
                                        Insert('.', Sb_LOp, Res);

                                Res := Pos(',', Sb_Tyag);
                                Delete(Sb_Tyag, Res, 1);
                                if Res <> 0 then
                                        Insert('.', Sb_Tyag, Res);

                                Res := Pos(',', Sb_Resh);
                                Delete(Sb_Resh, Res, 1);
                                if Res <> 0 then
                                        Insert('.', Sb_Resh, Res);
                                //
                                Form1.Memo2.Lines.Add('UP '+S_Sbor);

                                Rec_Count := Form1.ADOQuery2.RecordCount;
                                Priv_Str:=IntToStr(Kol_pr*Kol_Zap);
                                if (Rec_Count = 0) and (Kol_Zap = Kol) then
                                        //Запустили все
                                begin
                                        NomPos:=KonNom;
                                        KodPos:=KonKod;
                                        //++++++++++++++++++++++++++++++++++++++++++++++++++++
                                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                                'UPDATE %s SET [Кол во запущенных]=' + #39
                                                + IntToStr(Kol_Zap +
                                                Kol_Zap_R) + #39
                                                + ',[Кол во запущенных пос]=' +
                                                #39 + IntToStr(Kol_Zap) + #39 +
                                                ',[ПланЗапуск]=' + #39
                                                + Dat1 + #39 + ',Номер=' + #39 +
                                                Label4.Caption + #39 +
                                                ',[' + FN_KOL_RAZ + ']='
                                                +
                                                #39 + Kol_Raz + #39 + ',[' +
                                                FN_NOMER_RAZ + ']=' + #39 +
                                                Nomer_Raz +#39 +
                                                ',[' + 'НомерПос' + ']='
                                                +#39 + IntToStr(NomPos) + #39 +
                                                ',[' + 'КодПос' + ']='
                                                +#39 + IntToStr(KodPos) + #39 + ' WHERE ([IDГП]=' + #39 +
                                                (IDGP) + #39+') AND ([IDКо]=' + #39 +
                                                (KO) + #39+')', [Tab1])
                                                        then
                                                Exit;
                                end;
                        //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                                if (Rec_Count = 0) and (Kol_Zap <> Kol) then //Запустили не все
                                begin
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                NomPos:=NachNom +Kol_Zap-1;
                                KodPos:=NachKod +Kol_Zap-1;
                                //----------------------------------------------
                                KonNom:=NachNom+Kol_Zap-1;
                                KonKod:=NachKod+Kol_Zap-1;
                                        //++++++++++++++++++++++++++++++++++++++++++++++++++++
                                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                                'UPDATE %s SET [Кол во запущенных]=' + #39
                                                + IntToStr(Kol_Zap +
                                                Kol_Zap_R) + #39
                                                + ',[Кол во запущенных пос]=' +
                                                #39 + IntToStr(Kol_Zap) + #39 +
                                                ',[ПланЗапуск]=' + #39
                                                + Dat1 + #39 + ',Номер=' + #39 +
                                                Label4.Caption + #39 +
                                                ',[' + FN_KOL_RAZ + ']='
                                                +
                                                #39 + Kol_Raz + #39 +
                                                ',[' +FN_NOMER_RAZ + ']=' + #39 +Nomer_Raz +#39+
                                                ',[НомерПос]=' + #39 +IntToStr(NomPos) +#39+
                                                ',[КодПос]=' + #39 +IntToStr(KodPos) +#39+
                                                 ' WHERE ([IDГП]=' + #39 +
                                                (IDGP) + #39+') AND ([IDКо]=' + #39 +
                                                (KO) + #39+')', [Tab1])
                                        then
                                        Exit;
                                end;
                                if (Rec_Count <> 0)  then
                                        //Запустили последнии (в запуске уже были)
                                begin

                                NachKod:= KodPos+1;
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                KonKod:= NachKod+Kol_Zap-1;
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                NachNom:= NomPos+1;
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                KonNom:= NachNom+Kol_Zap-1;
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                NomPos:= NachNom+Kol_Zap-1;
                                KodPos:= NachKod+Kol_Zap-1;
                                        //++++++++++++++++++++++++++++++++++++++++++++++++++++
                                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                                'UPDATE %s SET [Кол во запущенных]=' + #39
                                                + IntToStr(Kol_Zap +
                                                Kol_Zap_R) + #39
                                                + ',[Кол во запущенных пос]=' +
                                                #39 + IntToStr(Kol_Zap) + #39 +
                                                ',[ПланЗапуск]=' + #39
                                                + Dat1 + #39 + ',Номер=' + #39 +
                                                Label4.Caption + #39 +
                                                ',[' + FN_KOL_RAZ + ']='+#39 + Kol_Raz + #39 +
                                                ',[' +FN_NOMER_RAZ + ']=' + #39 +Nomer_Raz +#39+
                                                ',[НомерПос]=' + #39 +IntToStr(NomPos) +#39+
                                                ',[КодПос]=' + #39 +IntToStr(KodPos) +#39+
                                                 ' WHERE ([IDГП]=' + #39 +
                                                (IDGP) + #39+') AND ([IDКо]=' + #39 +
                                                (KO) + #39+')', [Tab1])
                                        then
                                        Exit;
                                end;
                      End;

                    end;
                  End;

                End;
        Close;
      End;
end;

procedure TFV.FormClose(Sender: TObject; var Action: TCloseAction);
var
        I: Integer;
        S: string;
begin
        S := ExtractFileDir(ParamStr(0));
        Memo1.Lines.Clear;
        for I := 0 to SGS.ColCount - 1 do
                Memo1.Lines.Add(IntToStr(SGS.ColWidths[i]));
        Memo1.Lines.SaveToFile(S + '\SGS.ini');
        for i:=0 To SGS.RowCount Do
        Begin
            SGS.Rows[i].Clear;
        End;
end;

procedure TFV.DateTimePicker1Change(Sender: TObject);
var S,S1,Tab1,Zapusk,Spec:string;
D,D1:TDateTime;
I:Integer;
begin
                 Tab1:='['+Label23.Caption+']';
        Zapusk:='['+Label22.Caption+']';
        Spec:='['+Label24.Caption+']';
         S:=FormatDateTime('mm.dd.YYYY',DateTimePicker1.Date);
       { if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  max(планирование) as S from %s Where планирование<'+#39+S+#39 ,
                ['Запуск']) then
                exit;
        D:= Form1.ADOQuery1.FieldByName('S').AsDateTime;
        S1:=FormatDateTime('mm.dd.YYYY',D);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where планирование='+#39+S1+#39 ,
                ['Запуск']) then
                exit;  }
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  max(планирование) as S from %s Where планирование<'+#39+S+#39 ,
                [Zapusk]) then
                exit;
        D:= Form1.ADOQuery1.FieldByName('S').AsDateTime;
//+++++++++++++++++++++++++++++++
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  max(планирование) as S from %s Where планирование<'+#39+S+#39 ,
                [Zapusk]) then
                exit;
        D1:= Form1.ADOQuery1.FieldByName('S').AsDateTime;

        if D>D1 then
        Begin
        S1:=FormatDateTime('mm.dd.YYYY',D);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where планирование='+#39+S1+#39 ,
                [Zapusk]) then
                exit;
        End
        Else
        Begin
        S1:=FormatDateTime('mm.dd.YYYY',D1);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where планирование='+#39+S1+#39 ,
                [Zapusk]) then
                exit;
        End;

        I:=Form1.ADOQuery1.RecordCount;
        Cvet:= Form1.ADOQuery1.FieldByName('Cvet').AsInteger;
        case Cvet of
            0:
            begin
              Label16.Font.Color:=RGB(255,255,255);//Белый
              Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
            1:
            begin
              Label16.Font.Color:=RGB(234,239,45); //Желтый
              Label17.Font.Color:=RGB(41,243,104);//Зеленый
              Cvet1:=2;
            End;
            2:
            begin
              Label16.Font.Color:=RGB(41,243,104);//Зеленый
              Label17.Font.Color:=RGB(122,156,250); //Голубой
              Cvet1:=3;
            End;
            3:
            begin
              Label16.Font.Color:=RGB(122,156,250); //Голубой
              Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
        end;
end;

procedure TFV.DateTimePicker1Enter(Sender: TObject);

begin
        Button1.Enabled := True;

end;

procedure TFV.Button3Click(Sender: TObject);

var
        I, J,u, ID, Pos1, Kol_Zap, Kol_Zap_R, Kol, Rec_Count, Kol_Ob, Res, e:
        Integer;
        Svar, Sbor, Svar_o, Sbor_o: Double;
        Dat, Dat1, Kol_Raz, Nomer_Raz, IdGP, Zak, S_Svar, S_Sbor, Dat_zak,
        Dat_Z, ss,Privod, Vn_Dat, Dir, S, Teh_Gotov, Teh_Obrab, Plan_Dat,
        God,Mes,Shtrih_Str,Shtrih2_Str,Priv_Str,PDO,KTO,fileName,A,B: string;
        XL2: Variant;
        Ii,Kol_Sh,Kol_pr,NachNom,KonNom,NomPos,Shtrih2,Shtrih: Integer;
        NachKod,KonKOd,KodPos:Int64;
Doc: OleVariant;
Appl, wdDoc, wdRang,Sheet1,Sheet2,Rang:variant;


begin
        e := 1;
        U:=0;
        if (MessageDlg('Проверь, все ли клапана? Потом не добавишь!',mtInformation,[mbYes,MbNo],0)=mrYes) Then
        Begin
        Kol_Ob := 0;
        Vn_Dat := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);
        Dir :=Form1.Put_KTO+'\CKlapana\Суточные задания\';
        CreateDir(Dir);

        Dir :=Form1.Put_KTO+'\CKlapana\Суточные задания\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\Суточные задания\' + God + '\' + Mes+'\';
        CreateDir(Dir);

        XL2 := CreateOleObject('Excel.Application');
        CopyFile(PWideChar(Form1.Put_KTO+'\CKlapana\2013\SUT.xlsx'),
        PWideChar(Dir + '\'+Vn_Dat +' СутЗадан_КПУ_КПД № ' + Label4.Caption + '.xlsx'), False);
        XL2.Workbooks.Open(Dir + '\' + Vn_Dat + ' СутЗадан_КПУ_КПД № ' +
                Label4.Caption + '.xlsx');

        XL2.Application.EnableEvents := false;
        for I := 1 to SGS.RowCount do
        begin
               // if SGS.Cells[10, i] <> '' then
                begin
                       // ID := StrToInt(SGS.Cells[10, I]);
                      //  if (SGS.Cells[5, I] <> '') and (SGS.Cells[7, I] <> '') and (ID <> 0)
                       // then
                        begin
                                PDO:=SGS.Cells[28, SGS.Row];
                                KTO:=SGS.Cells[29, SGS.Row];

                                A:=SGS.Cells[17, i];
                                B:=SGS.Cells[18, i];
                               { Nomer_Raz := '';
                                Kol_Raz := '';
                                Kol_Zap := 0;//StrToInt(SGS.Cells[5, i]);
                                if  SGS.Cells[24, i]='' Then
                                Kol_Pr:=0
                                else
                                Kol_pr := StrToInt(SGS.Cells[24, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[22, i]='' Then
                                NachKod:=0
                                Else
                                NachKod:= StrToInt64(SGS.Cells[22, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[23, i]='' Then
                                KonKod:=0
                                Else
                                KonKod:= StrToInt64(SGS.Cells[23, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[27, i]='' Then
                                KodPos:=0
                                Else
                                KodPos:= StrToInt64(SGS.Cells[27, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[20, i]='' Then
                                NachNom:=0
                                Else
                                NachNom:= StrToInt(SGS.Cells[20, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++

                                if  SGS.Cells[21, i]='' Then
                                KonNom:=0
                                Else
                                KonNom:= StrToInt(SGS.Cells[21, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                if  SGS.Cells[26, i]='' Then
                                NomPos:=0
                                Else
                                NomPos:= StrToInt(SGS.Cells[26, i]);
                                //+++++++++++++++++++++++++++++++++++++++++++++
                                idgp:=SGS.Cells[25, i];
                                Plan_Dat := Form1.ConvertDat1(SGS.Cells[1, i]);
                                Teh_Gotov := Form1.ConvertDat1(SGS.Cells[19,
                                        i]);
                                //if Kol_Zap = 0 then
                               //         Continue;
                                Kol_Ob := Kol_Ob + Kol_Zap;
                                Kol_Zap_R := StrToInt(SGS.Cells[11, i]);
                                Kol := StrToInt(SGS.Cells[3, i]);
                                Privod := (SGS.Cells[14, i]);

                                Dat_Z := SGS.Cells[15, i];
                                res := Pos('.', Dat_z);
                                if Res <> 0 then
                                begin
                                        SS := Copy(Dat_Z, 1, Res);
                                        Delete(Dat_Z, 1, Res);
                                        Insert(SS, Dat_z, 4);
                                end;
                                Zak := SGS.Cells[2, i];
                                Dat_Zak := Dat_z;
                                Dat := FormatDateTime('mm.dd.yyyy', Now);
                                Dat1 := FormatDateTime('mm.dd.yyyy',
                                        DateTimePicker1.Date);

                                if Kol <> Kol_Zap then
                                begin
                                        if SGS.Cells[12, i] = '' then
                                                Kol_Raz := IntToStr(Kol_Zap)
                                        else
                                                Kol_Raz := SGS.Cells[12, i] + ','
                                                        + IntToStr(Kol_Zap);
                                        if SGS.Cells[13, i] = '' then
                                                Nomer_Raz := Label4.Caption
                                        else
                                                Nomer_Raz := SGS.Cells[13, i] +
                                                        ',' + Label4.Caption;
                                end;
                                if Kol = Kol_Zap then
                                begin
                                        Kol_Raz := IntToStr(Kol_Zap);
                                        Nomer_Raz := Label4.Caption;
                                end;
                                if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                        'Select * from %s Where  (Заказ= ' + #39
                                        + Zak + #39 + ') AND ([IdГП]= ' +
                                        #39 + (IDgp) + #39 + ')',
                                        ['Запуск']) then
                                        exit;
                                S_Svar := SGS.Cells[8, I];
                                S_Sbor := SGS.Cells[9, I];
                                Res := Pos(',', S_Svar);
                                Delete(S_Svar, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Svar, Res);

                                Res := Pos(',', S_Sbor);
                                Delete(S_Sbor, Res, 1);
                                if Res <> 0 then
                                        Insert('.', S_Sbor, Res);

                                Rec_Count := Form1.ADOQuery2.RecordCount;
                                Priv_Str:=IntToStr(Kol_pr*Kol_Zap); }
                                //+ZZZZZZ
                                Sheet1 := XL2.WorkSheets.Item[3+U];
                                Inc(U);
                                Sheet2 := XL2.WorkSheets.Add;
                                Memo2.Lines.Add('6');
                                Sheet2.Activate;
                                Sheet2.Name := Label4.Caption+' '+IntToStr(I);

                                Sheet1.Range['A1:Q80', EmptyParam].Copy(EmptyParam);
                                Sheet2.Range['A1', EmptyParam].PasteSpecial(OleVariant(xlPasteColumnWidths), OleVariant(xlPasteSpecialOperationNone), False, False);
                                Sheet2.Range['A1', EmptyParam].PasteSpecial(OleVariant(xlPasteValues), OleVariant(xlPasteSpecialOperationNone), False, False);
                                Sheet2.Range['A1', EmptyParam].PasteSpecial(OleVariant(xlPasteFormats), OleVariant(xlPasteSpecialOperationNone), False, False);
                                XL2.CutCopyMode := False;

                                rang:=Sheet2.Range['B7',emptyparam];
                                rang.value:=Label4.Caption;
                                rang:=Sheet2.Range['D7',emptyparam];
                                rang.value:=Zak;
                                rang:=Sheet2.Range['E7',emptyparam];
                                rang.value:='';
                                rang:=Sheet2.Range['F7',emptyparam];
                                rang.value:=SGS.Cells[16, i];
                                rang:=Sheet2.Range['K7',emptyparam];
                                rang.value:=Kol_Zap;
                                //+ZZZZZ

                        end;
                end;
        end;
        XL2.Visible := True;
        XL2 := UnAssigned;
  End;
end;

procedure TFV.N1Click(Sender: TObject);
begin
      SGS.Cells[5, SGS.Row]:='0';
end;

procedure TFV.SGSDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
        if (ACol<>0) AND (ARow<>0) Then
        Begin
        if (SGS.Cells[28, ARow]<>'') AND (SGS.Cells[29, ARow]='') Then
        Begin
                SGS.canvas.brush.Color := RGB(247, 247,
                                        26);
                //Желтый (1 Прогон
                SGS.Canvas.FillRect(Rect);
                SGS.Canvas.TextOut(Rect.Left, Rect.Top,
                SGS.Cells[ACol, ARow]);
        end;

        if (SGS.Cells[28, ARow]<>'') AND (SGS.Cells[29, ARow]<>'') Then
        Begin
                SGS.canvas.brush.Color := RGB(248, 82,
                                        111); //Розовый
                // (2 Прогон
                SGS.Canvas.FillRect(Rect);
                SGS.Canvas.TextOut(Rect.Left, Rect.Top,
                SGS.Cells[ACol, ARow]);
        end;

        if (SGS.Cells[28, ARow]='') AND (SGS.Cells[29, ARow]<>'') Then
        Begin
                SGS.canvas.brush.Color := RGB(242, 12,
                                        20); //Красный
                // (Капец КТО
                SGS.Canvas.FillRect(Rect);
                SGS.Canvas.TextOut(Rect.Left, Rect.Top,
                SGS.Cells[ACol, ARow]);
        end;

        end;
end;

end.

