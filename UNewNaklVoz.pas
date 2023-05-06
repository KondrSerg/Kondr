unit UNewNaklVoz;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Grids, ExtCtrls, ComObj,Math, DB, ADODB;

type
  //Второй
  TFNAklVoz = class(TForm)
    Panel1: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    SG1: TStringGrid;
    Button1: TButton;
    Button2: TButton;
    Memo1: TMemo;
    ProgressBar1: TProgressBar;
    SD1: TStringGrid;
    Button4: TButton;
    Button5: TButton;
    Panel3: TPanel;
    Label4: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    SG2: TStringGrid;
    Button3: TButton;
    Edit1: TEdit;
    Edit2: TEdit;
    ComboBox1: TComboBox;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Panel2: TPanel;
    Label5: TLabel;
    Label6: TLabel;
    Label1: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label15: TLabel;
    SD: TStringGrid;
    SD2: TStringGrid;
    SD9: TStringGrid;
    SD5: TStringGrid;
    SD6: TStringGrid;
    SD7: TStringGrid;
    Stenki: TStringGrid;
    SD8: TStringGrid;
    SD4: TStringGrid;
    SD3: TStringGrid;
    SSvarka: TStringGrid;
    Memo2: TMemo;
    btn1: TButton;
    mmo1: TMemo;
    mmo2: TMemo;
    chk1: TCheckBox;
    RSG1: TStringGrid;
    RSG2: TStringGrid;
    RSG3: TStringGrid;
    RSG4: TStringGrid;
    SQLConnector1: TADOConnection;
    ADOQuery2: TADOQuery;
    btn2: TButton;
    btn3: TButton;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure SG2DblClick(Sender: TObject);
    procedure Clear_StringGrid(SG:TStringGrid);
    function Float(var Str:String): String;
    procedure Button2Click(Sender: TObject);
    procedure Clear_StringGrid1(StringGrid: TStringGrid);
    procedure SG2Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure btn3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Puty:String;
  end;

var
  FNAklVoz: TFNAklVoz;

implementation

uses Main;

{$R *.dfm}
function TFNAklVoz.Float(var Str:String): String;
Var PosTrud:Integer;
Begin
        Result:='';
  PosTrud := Pos('.', Str); //Трудоемкость FLOAT
  if PosTrud <> 0 then
  begin
    Delete(Str, PosTrud, 1);
    Insert(',', Str, PosTrud);
  end;
  Result:=Str;

end;
procedure TFNAklVoz.FormShow(Sender: TObject);
var
        I, Kol, J, y, Res, Nom1, Nom2: Integer;
        Nom_Poz, Zak, IDGP, Kol_Zap, Kol_Raz: string;
begin

        //Button1.Enabled := False;

        SG1.Enabled := False;
        SG1.RowCount := 1;
        SG2.ColCount := 10;
        Form1.Clear_StringGrid(SG1);
        Form1.Clear_StringGrid(SG2);
        Form1.Clear_StringGrid(SD1);
        SG2.Cells[0, 0] := 'Заказ';
        SG2.Cells[1, 0] := 'Б\з';
        SG2.Cells[2, 0] := 'Расчетная дата';
        SG2.Cells[3, 0] := 'Кол во запущенных';
        SG2.Cells[4, 0] := 'Номер';
        SG2.Cells[5, 0] := 'IDГП';
        SG2.Cells[6, 0] := 'Тип';
        SG2.Cells[7, 0] := 'Исполнение';
        SG2.Cells[8, 0] := 'IDКО';
        SG2.Cells[9, 0] := 'КолЛопаток';
       { SG2.ColWidths[0] := 0;
        SG2.ColWidths[6] := 0;
        SG2.ColWidths[7] := 0;
        SG2.ColWidths[8] := 0;
        SG2.ColWidths[5] := 0; }
       // SQLConnector1.ConnectionString := Form1.MSSQL_CONN_STR_Glob;
        Edit1.Text := '0';
        Memo1.Lines.Clear;
        Nom2 := 0;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++   AND ([Заготовка Запуск] IS NULL)674   AND ([Файл] IS NULL)                                                                     //AND ([Технолог Готов] IS NULL)
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where  ( [Номер]<>0) AND ([Отмена] IS NULL)  AND ( [Х] IS NULL) Order By Номер',
                ['ЗапускВозд']) then
                exit;
        Y := 1;
        j := 1;
        SG2.RowCount := Form1.ADOQuery2.RecordCount + 1;
        Memo1.Lines.Clear;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                for I := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        Nom1 := Form1.ADOQuery2.FieldByName('Номер').AsInteger;
                        if j > 1 then
                                Nom2 := StrToInt(SG2.Cells[4, j - 1]);
                        if Nom1 = Nom2 then
                        begin
                                Form1.ADOQuery2.Next;
                                Continue;
                        end;
                        SG2.Cells[0, j] :=
                                Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SG2.Cells[1, j] :=
                                Form1.ADOQuery2.FieldByName('БЗ').AsString;
                        SG2.Cells[2, j] :=
                                Form1.ADOQuery2.FieldByName('Расчетная дата готовности').AsString;
                        SG2.Cells[3, j] :=
                                Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        SG2.Cells[4, j] :=
                                Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SG2.Cells[5, j] :=
                                Form1.ADOQuery2.FieldByName('IDГП').AsString;
                        SG2.Cells[6, j] :=
                                Form1.ADOQuery2.FieldByName('Тип').AsString;
                        SG2.Cells[7, j] :=
                                Form1.ADOQuery2.FieldByName('Исполнение').AsString;
                        SG2.Cells[8, j] :=
                                Form1.ADOQuery2.FieldByName('IDКО').AsString;
                        SG2.Cells[9, j] :=
                                Form1.ADOQuery2.FieldByName('КолЛопаток').AsString;
                        //Nom_Poz := Form1.ADOQuery2.FieldByName(FN_POS).AsString;
                        Inc(y);
                        Inc(j);
                        Form1.ADOQuery2.Next;
                end;
        end;
        SG2.RowCount := Y + 2;
end;

procedure TFNAklVoz.FormClose(Sender: TObject; var Action: TCloseAction);
begin

        Label11.Caption := '0';
        Label9.Caption := '0';
        Label1.Caption := '0';
end;

//Очистка Грида

procedure TFNAklVoz.Clear_StringGrid1(StringGrid: TStringGrid);
var
        i, j: integer;
begin
        for i := 0 to StringGrid.ColCount - 1 do
                for j := 0 to StringGrid.RowCount - 1 do
                begin
                        //StringGrid.Objects[i, j] := nil;
                        StringGrid.Cells[I, j] := '';

                end;
for j := 0 to StringGrid.RowCount  do
                begin
                        //StringGrid.Objects[i, j] := nil;
                        Memo2.Lines.Add( StringGrid.Cells[0, j]);

                end;

end;
procedure TFNAklVoz.Button3Click(Sender: TObject);
var
        I, Kol, Y, j, Pos, Pos1, x: Integer;
        Nom_Poz, Zak, IDGP, Briket, Briket1, Briket2, Br, Br1, Br2, Kol_Zap:
        string;
begin
        SG1.Enabled:=True;
        Clear_StringGrid1(SG1);

        SG1.ColCount := 15;
        SG1.Cells[0, 0] := 'Заказ';
        SG1.Cells[1, 0] := 'БЗ';
        SG1.Cells[2, 0] := 'Кол во запущенных';
        SG1.Cells[3, 0] := 'Номер';
        SG1.Cells[4, 0] := 'IDГП';
        SG1.Cells[5, 0] := 'Клапан';
        SG1.Cells[6, 0] := 'Сборка';
        SG1.Cells[7, 0] := 'Сварка';
        SG1.Cells[8, 0] := 'Привод';
        SG1.Cells[9, 0] := 'Тип';
        SG1.Cells[10, 0] := 'Исполнение';
        SG1.Cells[11, 0] := 'IDКО';
        SG1.Cells[12, 0] := 'A';
        SG1.Cells[13, 0] := 'B';
        SG1.Cells[14, 0] := 'КолЛопаток';
        //  ([Заказ]= '+#39+Zak+#39+')AND
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where   (Номер='
                + #39 + Label15.Caption + #39 + ') AND (Отмена IS NULL) Order By Заказ ', ['ЗапускВозд']) then
                exit;
        SG1.RowCount:=1;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                SG1.RowCount := Form1.ADOQuery2.RecordCount + 2;
                SG1.FixedRows:=1;
                for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        SG1.Cells[0, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SG1.Cells[1, i + 1] :=
                                Form1.ADOQuery2.FieldByName('БЗ').AsString;
                        SG1.Cells[2, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        SG1.Cells[3, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SG1.Cells[4, i + 1] :=
                                Form1.ADOQuery2.FieldByName('IDГП').AsString;
                        SG1.Cells[5, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Изделие').AsString;
                        SG1.Cells[6, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сборка Клапана').AsString;
                        SG1.Cells[7, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сварка').AsString;
                        SG1.Cells[8, i + 1] :=
                                Form1.ADOQuery2.FieldByName('МодПривода').AsString;
                        SG1.Cells[9, i+1] :=
                                Form1.ADOQuery2.FieldByName('Тип').AsString;
                        SG1.Cells[10, i+1] :=
                                Form1.ADOQuery2.FieldByName('Исполнение').AsString;
                        SG1.Cells[11, i+1] :=
                                Form1.ADOQuery2.FieldByName('IDКО').AsString;
                        SG1.Cells[12, i+1] :=
                                Form1.ADOQuery2.FieldByName('A').AsString;
                        SG1.Cells[13, i+1] :=
                                Form1.ADOQuery2.FieldByName('B').AsString;
                        SG1.Cells[14, i+1] :=
                                Form1.ADOQuery2.FieldByName('КолЛопаток').AsString;


                        Form1.ADOQuery2.Next;
                end;
        end;
end;
procedure TFNAklVoz.Clear_StringGrid(SG:TStringGrid);
Var I:Integer;
Begin
        for i:=0 To SG.RowCount Do
        Begin
            SG.Rows[i].Clear;
        End;

end;
procedure TFNAklVoz.Button1Click(Sender: TObject);
const xlPasteFormats = $FFFFEFE6;
        xlNone = $FFFFEFD2;
var
        Kol_ZAP, Kol, Kol_Zap_Ranee, Res, Nomer, Zak_Int, i, A, B, D, y, F_KPU,
        F_KPD,xy, Res1, e, Zag, Zap, j, g, k,Q,Res_Mat,Res_Proch,Res_Stand,qq,Res_Sbor,Res_Klap,Podstavka,Srav,Srav1:Integer;
        Res_VElem,Res_Elem,Res_Oboz,Flag,o,p,s,h,f,m,n,v,w,KolGib,AA,bb,Res_Tol,F2,F1,Res_KPU,Res_KPD,Res_list,Res_Ger:Integer;
        Res_Lop,Sten_V,Sten_G,hh,pp,Res_Kog_Priv,GG,Res_Kompl_priv,ww,www,Res_Detal,Os,R25,Ugol,Lenta,Kol_Otv,ID,Dl_Slova,Kol_Lop:Integer;
        Zak, Dat, Plan_Dat, Vn_Dat, Nom, Pos_Vst, Pos_Ml, R, Pereh, Privod,
                Zag_S,
                Zap_S, Dir_main, God, Mes, Fil,Dir,DlinRaz,ShirRaz,Mat,
                Dlina,Shir,Kol_Gib,Razm,Tol_S,Pos_Flan1,Pos_Flan,Pos_Dop,
                Pos_Sn,Pos_Privod,Pos_Isp,Pos_Ram,SvarkaS,Kol_Lops,BZ: string;
        Svar, Sbor, Izdel,VElem,Elem,Oboz, IDGP,IDKO,Tip,EI,NC_S,Pos1,Pos2,Pos3,Pos4,Pos5,OboznSh,NC_S2,N2,Diam,N_S: string;
       Res_Tol55,Kol_Ed2, Kol_Ed,Kol_Oboz,Kol_Ed1,Kol_Oboz1,NC,NC2,ND22,C,DlinI,ShirI,Tol,Razmet_KPU,DlinD,A_Lenta,NC_OB,T,NCB,Diam_i: Double;
        Dat1, Dat2, Dat3: TDate;
        Kanban,Trumph,Nog,Ugl,Gib,Prok,Pila,Svarka,Tehnolog,VALCI,HAco,Kromka,Tokarn,Rotor:Boolean;
        XL, XL1, XL2, XL_Temp, V1: Variant;
        Res_Tol1,Res_Tol2,Res_Resh,Res_N2,Res_Svarka,PosTrud,DD,NERG,
        S1,S2,S3,S4:Integer;
        Nog_List:TStringList;
        RecNom,RecCol,Dvoin_klap:Integer;
begin
        Nog_List :=TStringList.Create;
        F_KPU := 0;
        F_KPD := 0;
        Dvoin_klap:=0;
        y := 1;
        Xy := 1;
        g:=1;
        K:=1;
        O:=1;
        S:=1;
        P:=1;
        H:=1;
        N:=1;
        M:=1;
        v:=1;
        w:=1;
        ww:=1;
        www:=1;
        Q:=1;
        qq:=1;
        AA:=1;
        PP:=1;

        S1:=1;
        S2:=1;
        S3:=1;
        S4:=1;
        e := 30;

        //Vn_Dat := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);

        Clear_StringGrid(SD1);
        Clear_StringGrid(SD2);
        Clear_StringGrid(SD3);
        Clear_StringGrid(SD4);
        Clear_StringGrid(SD5);
        Clear_StringGrid(SD6);
        Clear_StringGrid(SD7);
        Clear_StringGrid(SD8);
        Clear_StringGrid(SD9);
        Clear_StringGrid(SD);
        Clear_StringGrid(Stenki);

        Clear_StringGrid(RSG1);
        Clear_StringGrid(RSG2);
        Clear_StringGrid(RSG3);
        Clear_StringGrid(RSG4);

        Memo2.Lines.Clear;
        SD1.Cells[2,0]:='Kanban True';
        SD2.Cells[2,0]:='Kanban False';
        SD3.Cells[2,0]:='Trumph';
        SD4.Cells[2,0]:='Nognicy';
        SD5.Cells[2,0]:='Gibka';
        SD6.Cells[2,0]:='Prokat';
        SD7.Cells[2,0]:='Pila';
        SD8.Cells[2,0]:='Uglorub';


        for i := 1 to SG1.RowCount do
        begin
                if SG1.Cells[2, i] <> '' then
                begin
                        Kol_Oboz:=0;
                        Kol_Zap:=0;
                        NC_OB:=0;
                        NC:=0;
                        NC_S:='';
                        NC_S2:='';
                        Tip:='';
                        BZ:=SG1.Cells[1, 0];
                        Nom := SG1.Cells[3, i];
                        Izdel := SG1.Cells[5, i];
                        if  SG1.Cells[14, i]='' then
                        Kol_Lop:=0
                        Else
                        Kol_Lop:= StrToInt(SG1.Cells[14, i]);
                        Kol_Zap := StrToInt(SG1.Cells[2, i]);
                        Tip:=Izdel;
                        Izdel := SG1.Cells[5, i];
                        IDGP:=SG1.Cells[4, i];
                        IDKO:=SG1.Cells[11, i];
                        //++++++++++++++++++++++
                        SQLConnector1.Connected:=True;
                        ADOQuery2.Close;
                        ADOQuery2.SQL.Clear;
                        ADOQuery2.SQL.Text :='Select * from СпецифВозд Where   (IdГП=' + #39 +
                                SG1.Cells[4, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[11, i] + #39 +') AND (Элемент=' + #39 +
                                'Клапан воздушный' + #39 +') ' ;
                        ADOQuery2.Active := true;
                        if   ADOQuery2.recorDCount>1 Then //Клапан разбит на два
                        Begin
                            Dvoin_klap:=1;
                        end
                        Else
                             Dvoin_klap:=0;
                        //++++++++++++++++++++++
                        SQLConnector1.Connected:=True;
                        ADOQuery2.Close;
                        ADOQuery2.SQL.Clear;
                        ADOQuery2.SQL.Text :='Select * from СпецифВозд Where   (IdГП=' + #39 +
                                SG1.Cells[4, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[11, i] + #39 +') AND (Элемент=' + #39 +
                                'Лопатка' + #39 +') AND (ВидЭлемента='+
                             #39 +'Детали'+ #39 +') ' ;
                        ADOQuery2.Active := true;

                       { if not Form1.mkQuerySelect(Form1.ADOQuery2, //Вытащить кол лопаток
                                , ['СпецифВозд']) then
                                exit;   }
                        Kol_lopS:=ADOQuery2.FieldByName('Количество').AsString;
                        if Kol_Lops='' then
                        Kol_Lop:=0
                        else
                        Kol_Lop:= ADOQuery2.FieldByName('Количество').AsInteger;
                        if Dvoin_klap=1  Then
                        Begin
                            {if not Form1.mkQueryDelete(Form1.ADOQuery1, 'DELETE FROM %s WHERE (IdГП=' + #39 +
                                SG1.Cells[4, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[11, i] + #39 +') AND (ВидЭлемента='+
                             #39 +'Деталь'+ #39 +') AND (Элемент='+
                             #39 +'Лопатка'+ #39 +')'
                                , ['СпецифВозд']) then
                                exit; }

                             Kol_Lop:=Kol_Lop DIV 2;

                        end;
                       { if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where   (IdГП=' + #39 +
                                SG1.Cells[4, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[11, i] + #39 +') ', ['СпецифВозд']) then
                                exit;  }
                       SQLConnector1.Connected:=True;
                        ADOQuery2.Close;
                        ADOQuery2.SQL.Clear;
                        ADOQuery2.SQL.Text :='Select * from СпецифВозд Where   (IdГП=' + #39 +
                                SG1.Cells[4, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[11, i] + #39 +') ' ;
                        ADOQuery2.Active := true;
                 for J:=0 To  ADOQuery2.RecordCount-1 Do
                 Begin
                      Flag:=0;
                      KolGib:=0;
                      VElem:= ADOQuery2.FieldByName('ВидЭлемента').AsString;
                      Elem:= ADOQuery2.FieldByName('Элемент').AsString;
                      Oboz:= ADOQuery2.FieldByName('Обозначение').AsString;
                      Kol_Ed2:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Ed:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Oboz:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      //Trumph,Nog,Ugl,Gib,Prok,Pila
                      Dlina:=ADOQuery2.FieldByName('Длина').AsString;
                      Diam:=ADOQuery2.FieldByName('Диаметр').AsString;
                      Shir:= ADOQuery2.FieldByName('Ширина').AsString;
                      DlinRaz:= StringReplace(ADOQuery2.FieldByName('ДлинаРазв').AsString,'.',',',[rfReplaceAll]);
                      ShirRaz:= StringReplace(ADOQuery2.FieldByName('ШиринаРазв').AsString,'.',',',[rfReplaceAll]);

                      if DlinRaz ='' Then
                            DlinRaz:='0';
                      if ShirRaz ='' Then
                            ShirRaz:='0';
                      Kol_Gib:= ADOQuery2.FieldByName('КолГибов').AsString;
                      if Kol_Gib=''  Then
                      KolGib:=0
                      Else
                      KolGib:=StrToInt(Kol_Gib);

                      Mat:= ADOQuery2.FieldByName('Материал').AsString;
                      if Mat='' Then
                                        Mat:='0';
                      EI:= ADOQuery2.FieldByName('ЕИ').AsString;

                      Kanban:= ADOQuery2.FieldByName('Канбан').AsBoolean;
                      Trumph:= ADOQuery2.FieldByName('Trumph').AsBoolean;
                      Nog:= ADOQuery2.FieldByName('Ножницы').AsBoolean;
                      Ugl:= ADOQuery2.FieldByName('Углоруб').AsBoolean;
                      Gib:= ADOQuery2.FieldByName('Гибка').AsBoolean;
                      Prok:= ADOQuery2.FieldByName('Прокатка').AsBoolean;
                      Pila:= ADOQuery2.FieldByName('Пила').AsBoolean;

                      HACO:= ADOQuery2.FieldByName('HACO').AsBoolean;
                      VALCI:= ADOQuery2.FieldByName('Вальцы').AsBoolean;
                      KROMKA:=ADOQuery2.FieldByName('Кромкогиб').AsBoolean;
                      Res:=Pos('Обечайка',Elem); // (Res<>0) or
                      Res1:=Pos('Корпус выкатной',Elem);
                     if (Diam<>'') AND ((Res1<>0)) then
                     begin
                      Diam_i :=StrToFloat(Diam);
                      If Diam_I<400 then
                      begin
                         TOKARN:= True;
                         ROTOR:=  False;
                      end
                      else
                      begin
                         TOKARN:=False ;
                         ROTOR:=True  ;
                      end;
                      //TOKARN:= ADOQuery2.FieldByName('Токарный').AsBoolean;
                      //ROTOR:= ADOQuery2.FieldByName('Ротор').AsBoolean;
                     end;
                      //==================================================
                      Res_Mat:=Pos('Материалы',VELem);
                     Res_Proch:=Pos('Прочие изделия',VELem);
                     Res_Stand:=Pos('Стандартные изделия',VELem);
                     Res_Sbor:=Pos('Сборочные единицы',VELem);
                     Res_Klap:=Pos('Клапан',ELem);
                     Res_Kompl_priv:=Pos('Комплект привода',ELem);
                      //XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                      Res_Detal:=AnsiCompareStr('Детали',VElem);
                      OboznSh:=Oboz;
                      Res:=Pos('-',OboznSh);
                      if Res<>0 Then                      //   AND ((Sten_G<>0) ANd (Sten_V<>0))
                        Delete(OboznSh,Res,30);
                      if (Res_Detal=0) Then
                      Begin
                        if not Form1.mkQuerySelect66(Form1.ADOQuery4,
                                'Select * from %s Where   (ОбозначениеШ=' + #39 +
                                OboznSh + #39 +') ', ['ШаблонВоз']) then
                                exit;
                        Tehnolog:=Form1.ADOQuery4.FieldByName('Технолог').AsBoolean;
                        if Tehnolog=False then
                        begin
                                Kanban:=False ;
                                Trumph:=False ;
                                Nog:=False;
                                Ugl:= False;
                                Gib:= False;
                                Prok:= False;
                                Pila:= False;
                                Svarka:=False;
                        end
                        Else
                        begin
                        Kanban:= Form1.ADOQuery4.FieldByName('Канбан').AsBoolean;
                        Trumph:= Form1.ADOQuery4.FieldByName('Trumph').AsBoolean;
                        Nog:= Form1.ADOQuery4.FieldByName('Ножницы').AsBoolean;

                        HACO:= Form1.ADOQuery4.FieldByName('HACO').AsBoolean;
                        VALCI:= Form1.ADOQuery4.FieldByName('Вальцы').AsBoolean;
                        KROMKA:= Form1.ADOQuery4.FieldByName('Кромкогиб').AsBoolean;
                        TOKARN:= Form1.ADOQuery4.FieldByName('Токарный').AsBoolean;
                        ROTOR:= Form1.ADOQuery4.FieldByName('Ротор').AsBoolean;
                        Res:=Pos('Обечайка',Elem);   //(Res<>0) or
                     Res1:=Pos('Корпус выкатной',Elem);
                     if (Diam<>'') AND ((Res1<>0)) then
                     begin
                      Diam_i :=StrToFloat(Diam);
                      If Diam_I<400 then
                      begin
                         TOKARN:= True;
                         ROTOR:=  False;
                      end
                      else
                      begin
                         TOKARN:=False ;
                         ROTOR:=True ;
                      end;
                      //TOKARN:= ADOQuery2.FieldByName('Токарный').AsBoolean;
                      //ROTOR:= ADOQuery2.FieldByName('Ротор').AsBoolean;
                     end;
                        if (chk1.Checked=True) then
                        Begin
                                if Nog=True then
                                begin
                                Trumph:= True;
                                Nog:=False;
                                end;
                        end;
                        Ugl:= Form1.ADOQuery4.FieldByName('Углоруб').AsBoolean;
                        Gib:= Form1.ADOQuery4.FieldByName('Гибка').AsBoolean;
                        Prok:= Form1.ADOQuery4.FieldByName('Прокатка').AsBoolean;
                        Pila:= Form1.ADOQuery4.FieldByName('Пила').AsBoolean;
                        Svarka:=Form1.ADOQuery4.FieldByName('Сварка').AsBoolean;
                        end;
                        Ugol:= Form1.ADOQuery4.FieldByName('Угол').AsInteger;
                        Kol_Gib:= Form1.ADOQuery4.FieldByName('КолГибов').AsString;
                        if Kol_Gib=''  Then
                        KolGib:=0
                        Else
                        KolGib:=StrToInt(Kol_Gib);

                      //XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                      If (Svarka=True) Then
                      Begin
                            Dl_Slova:=Length(OboznSh);
                            Delete(OboznSh,Dl_Slova,1);
                      end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++++
                      NC:=0;
                      NC_S:='';
                      Tol:=0;
                      Tol_S:='';
                      //==============================================
                      Res:=Pos('Тяга',Elem);
                      NERG:=Pos('НЕРЖ',Oboz);
                      if (Res<>0) and (Nerg<>0) then
                      begin
                        Trumph:=True;
                        Gib:=True;
                        Kanban:=False;
                        if not Form1.mkQueryUpdate2(Form1.ADOQuery1,
                                        'UPDATE %s SET [Trumph]=' +
                                        #39+ 'True' + #39 +
                                        ', Гибка='+#39+ 'True' + #39 +
                                        ', Канбан='+#39+ 'False' + #39 +
                                        ' WHERE (Обозначение=' + #39 + Oboz + #39 +') AND ([IdГП]=' + #39 + IDGP + #39 +
                                                ') ', ['СпецифВозд']) then
                                        Exit;
                        {DD:=1;
                        Continue;}
                      end;
                      //==============================
                      if VALCI =True then
                        begin
                                RSG1.Cells[0,s1]:=IntToStr(s1);
                                RSG1.Cells[1,s1]:=Tip;
                                RSG1.Cells[2,s1]:=Elem;
                                RSG1.Cells[3,s1]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG1.Cells[4,s1]:=Diam;
                                RSG1.Cells[5,s1]:=Oboz;
                                //RSG1.Cells[6,s1]:=Dlina;
                                RSG1.Cells[6,s1]:=Kol_Gib;
                                RSG1.Cells[7,s1]:=Mat;
                                RSG1.Cells[8,s1]:=N_S;//Н\ч
                                Inc(S1);
                                RSG1.RowCount:= RSG1.RowCount+1;
                        end;
                        //==============================
                        if KROMKA =True then
                        begin
                                RSG2.Cells[0,s2]:=IntToStr(s2);
                                RSG2.Cells[1,s2]:=Tip;
                                RSG2.Cells[2,s2]:=Elem;
                                RSG2.Cells[3,s2]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG2.Cells[4,s2]:=Diam;
                                RSG2.Cells[5,s2]:=Oboz;
                                //RSG2.Cells[6,s2]:=Dlina;
                                RSG2.Cells[6,s2]:=Kol_Gib;
                                RSG2.Cells[7,s2]:=Mat;
                                RSG2.Cells[8,s2]:=N_S;//Н\ч
                                Inc(S2);
                                RSG2.RowCount:= RSG2.RowCount+1;
                        end;
                        //==============================
                        if ROTOR =True then
                        begin
                                RSG3.Cells[0,s3]:=IntToStr(s3);
                                RSG3.Cells[1,s3]:=Tip;
                                RSG3.Cells[2,s3]:=Elem;
                                RSG3.Cells[3,s3]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG3.Cells[4,s3]:=Diam;
                                RSG3.Cells[5,s3]:=Oboz;
                                //RSG3.Cells[6,s3]:=Dlina;
                                RSG3.Cells[6,s3]:=Kol_Gib;
                                RSG3.Cells[7,s3]:=Mat;
                                RSG3.Cells[8,s3]:=N_S;//Н\ч
                                Inc(S3);
                                RSG3.RowCount:= RSG3.RowCount+1;
                        end;
                        //==============================
                        if TOKARN =True then
                        begin
                                RSG4.Cells[0,s4]:=IntToStr(s4);
                                RSG4.Cells[1,s4]:=Tip;
                                RSG4.Cells[2,s4]:=Elem;
                                RSG4.Cells[3,s4]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG4.Cells[4,s4]:=Diam;
                                RSG4.Cells[5,s4]:=Oboz;
                                //RSG4.Cells[6,s]:=Dlina;
                                RSG4.Cells[6,s4]:=Kol_Gib;
                                RSG4.Cells[7,s4]:=Mat;
                                RSG4.Cells[8,s4]:=N_S;//Н\ч
                                Inc(S4);
                                RSG4.RowCount:= RSG4.RowCount+1;
                        end;

                      end;
                        //======================================================
                      Res_Stand:=Pos('Стандартные изделия',VELem);
                      if (Kanban=True) AND (REs_Stand=0) AND (Res_Sbor=0) AND (Res_Mat=0)   Then
                      Begin
                        For y:=1 To SD1.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD1.Cells[3,y]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD1.Cells[6,y]);

                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin

                               Kol_Ed:=Kol_Ed+StrToFloat(SD1.Cells[4,y]);
                               SD1.Cells[4,y]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD1.Cells[5,y]);
                               SD1.Cells[5,y]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                        end;
                        If Flag=0 Then
                        Begin
                                SD1.Cells[2,k]:=Tip;
                                SD1.Cells[3,k]:=Elem;
                                SD1.Cells[4,k]:=FloatToStr(Kol_Ed);
                                SD1.Cells[5,k]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD1.Cells[6,k]:=Oboz;
                                SD1.Cells[0,k]:=IntToStr(k);
                                SD1.Cells[1,k]:=EI;
                                SD1.Cells[7,k]:=Dlina;
                                SD1.Cells[8,k]:=Shir;
                                SD1.Cells[9,k]:=DlinRaz;
                                SD1.Cells[10,k]:=ShirRaz;
                                SD1.Cells[11,k]:=Kol_Gib;
                                SD1.Cells[12,k]:=Mat;
                                SD1.Cells[13,k]:=VElem;
                                Inc(k);
                                SD1.RowCount:= SD1.RowCount+1;
                                //Break;
                        end;
                      end;

                     //Заготовка (Kanban=False) AND (Res_Mat=0) AND (Res_Proch=0) AND (Res_Stand=0) AND (Res_Sbor=0) AND (Res_Kompl_priv=0) AND (Res_Klap=0)
                      if (Kanban=False) AND (Res_Detal=0) AND (Tehnolog=True) Then
                      //For y:=1 To SD1.RowCount Do Заготовка
                      Begin
                      //Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      //Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Ed:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Oboz:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                        For O:=1 To SD2.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD2.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD2.Cells[6,o]);

                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(SD2.Cells[4,o]);
                               SD2.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD2.Cells[5,o]);
                               SD2.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                        end;
                        If Flag=0 Then
                        Begin
                                SD2.Cells[2,p]:=Tip;
                                SD2.Cells[3,p]:=Elem;
                                SD2.Cells[4,p]:=FloatToStr(Kol_Ed);
                                SD2.Cells[5,p]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD2.Cells[6,p]:=Oboz;
                                SD2.Cells[0,p]:=IntToStr(p);
                                SD2.Cells[1,p]:=EI;
                                SD2.Cells[7,p]:=Dlina;
                                SD2.Cells[8,p]:=Shir;
                                SD2.Cells[9,p]:=DlinRaz;
                                SD2.Cells[10,p]:=ShirRaz;
                                SD2.Cells[11,p]:=Kol_Gib;
                                SD2.Cells[12,p]:=Mat;
                                SD2.Cells[13,p]:=VElem;
                                Inc(p);
                                SD2.RowCount:= SD2.RowCount+1;
                                //Break;
                        end;
                      end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++ (Kanban=False)  AND
                      if (Res_Sbor<>0)   Then       //Сборка
                      Begin
                      Kol_Ed:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Oboz:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));

                        For O:=1 To SD.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(SD.Cells[4,o]);
                               SD.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD.Cells[5,o]);
                               SD.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                        end;
                        If Flag=0 Then
                        Begin
                                SD.Cells[2,aa]:=Tip;
                                SD.Cells[3,aa]:=Elem;
                                SD.Cells[4,aa]:=FloatToStr(Kol_Ed);
                                SD.Cells[5,aa]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD.Cells[6,aa]:=Oboz;
                                SD.Cells[0,aa]:=IntToStr(aa);
                                SD.Cells[1,aa]:=EI;
                                SD.Cells[7,aa]:=Dlina;
                                SD.Cells[8,aa]:=Shir;
                                SD.Cells[9,aa]:=DlinRaz;
                                SD.Cells[10,aa]:=ShirRaz;
                                SD.Cells[11,aa]:=Kol_Gib;
                                SD.Cells[12,aa]:=Mat;
                                SD.Cells[13,aa]:=VElem;
                                Inc(aa);
                                SD.RowCount:= SD.RowCount+1;
                                //Break;
                        end;
                      end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND ((Res_Stand<>0) OR (Res_Mat<>0) OR (Res_Proch<>0) OR  (Res_Kompl_priv<>0)) Then       //СCklad
                      Begin
                      Kol_Ed:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Oboz:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));

                        For O:=1 To SD9.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD9.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD9.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(SD9.Cells[4,o]);
                               SD9.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD9.Cells[5,o]);
                               SD9.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                        end;
                        If Flag=0 Then
                        Begin
                                SD9.Cells[2,qq]:=Tip;
                                SD9.Cells[3,qq]:=Elem;
                                SD9.Cells[4,qq]:=FloatToStr(Kol_Ed);
                                SD9.Cells[5,qq]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD9.Cells[6,qq]:=Oboz;
                                SD9.Cells[0,qq]:=IntToStr(qq);
                                SD9.Cells[1,qq]:=EI;
                                SD9.Cells[7,qq]:=Dlina;
                                SD9.Cells[8,qq]:=Shir;
                                SD9.Cells[9,qq]:=DlinRaz;
                                SD9.Cells[10,qq]:=ShirRaz;
                                SD9.Cells[11,qq]:=Kol_Gib;
                                SD9.Cells[12,qq]:=Mat;
                                SD9.Cells[13,qq]:=VElem;
                                Inc(qq);
                                SD9.RowCount:= SD9.RowCount+1;
                                //Break;
                        end;
                      end;
                     Res_N2:=AnsiCompareStr('2Н',N2);
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Trumph=True) AND (Res_Detal=0) Then
                      Begin
                      Kol_Ed:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Oboz:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                                SD3.Cells[2,s]:=Tip;
                                SD3.Cells[3,s]:=Elem;
                                SD3.Cells[4,s]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                //SD3.Cells[5,s]:=FloatToStr(Kol_Oboz);
                                SD3.Cells[6,s]:=Oboz;
                                SD3.Cells[0,s]:=IntToStr(s);
                                SD3.Cells[1,s]:=EI;
                                SD3.Cells[7,s]:=Dlina;
                                SD3.Cells[8,s]:=Shir;
                                SD3.Cells[9,s]:=DlinRaz;
                                SD3.Cells[10,s]:=ShirRaz;
                                SD3.Cells[11,s]:=Kol_Gib;
                                SD3.Cells[12,s]:=Mat;
                                if DlinRaz='' Then
                                DlinRaz:='0';
                                SD3.Cells[14,s]:=FloatToStr(C);
                                SD3.Cells[13,s]:=VElem;
                                Inc(s);
                                SD3.RowCount:= SD3.RowCount+1;
                     end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Nog=True) AND (Res_Detal=0) Then
                     Begin

                                Kol_Ed:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Oboz:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                                SD4.Cells[2,h]:=Tip;
                                SD4.Cells[3,h]:=Elem;
                                SD4.Cells[4,h]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                SD4.Cells[6,h]:=Oboz;
                                SD4.Cells[0,h]:=IntToStr(h);
                                SD4.Cells[1,h]:=EI;
                                SD4.Cells[7,h]:=Shir;
                                SD4.Cells[8,h]:=Dlina;
                                SD4.Cells[9,h]:=ShirRaz;
                                SD4.Cells[10,h]:=DlinRaz;
                                SD4.Cells[11,h]:=Kol_Gib;
                                SD4.Cells[12,h]:=Mat;
                                if DlinRaz='' Then
                                DlinRaz:='0';
                                //+ ++++++++++++++++++++++++++++++++++++++++++
                                //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                //Лист ДПРНМ 1,2х600х1500 Л63 ГОСТ 931-90
                                        Res_List:=Pos('ДПРНМ',Mat);
                                        if Res_List<>0 Then
                                        mat:='1,2';
                                 //Лента на шестьдесят НЕРЖ 0,3
                                 Res_List:=Pos('НЕРЖ',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                Res_List:=Pos('Л',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// Лента
                                                Delete(Mat,1,Res_Tol);
                                                //
                                                Res_Tol:=Pos(' ',Mat);//ОЦ
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('х',Mat);//x
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos('*',Mat);//*
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);
                                        end
                                        Else
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);
                                                Delete(Mat,1,Res_Tol);
                                        end;
                                        Res_List:=Pos('ОЦ',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);

                                //+++++++++++++++++++++++++++++++++++++++++++++
                                Tol:=StrToFloat(Mat);
                                Srav:=CompareValue(Tol,0.55);
                                if (Srav=0) OR (Srav<0) Then
                                Begin
                                        Tol_S:='0.55';
                                        Res_Tol55:=0.55;
                                End;
                                Srav:=CompareValue(Tol,0.7);
                                if ((Srav>0) OR (Srav=0)) AND (Tol<=1) Then
                                Begin
                                        Tol_S:='0.7';
                                        Res_Tol55:=7;
                                End;

                                Srav:=CompareValue(Tol,1.2);
                                //Srav1:= CompareValue(Tol,0.7);
                                if ((Srav=0) OR (Srav>0)) AND (Tol<=2) Then
                                Begin
                                        Tol_S:='1.2';
                                        Res_Tol55:=2;
                                End;
                                DlinI:=StrToFloat(Shir);
                                ShirI:=StrToFloat(Dlina);

                                if DlinI<=500 Then
                                        DlinI:=500;
                                //---------------
                                if (DlinI>500) AND (DlinI<=1000) Then
                                        DlinI:=1000;
                                //---------------
                                if (DlinI>1000) AND (DlinI<=1500) Then
                                        DlinI:=1500;
                                //---------------
                                if (DlinI>1500) AND (DlinI<=2000) Then
                                        DlinI:=2000;
                                //---------------
                                if (DlinI>2000) AND (DlinI<=2500) Then
                                        DlinI:=2500;
                                //+====================================
                                if Res_Tol55=0.55 Then
                                Begin
                                if ShirI<=50 Then
                                        ShirI:=50;
                                //---------------
                                if (ShirI>50) AND (ShirI<=100) Then
                                        ShirI:=100;
                                //---------------
                                if (ShirI>100) AND (ShirI<=200) Then
                                        ShirI:=200;
                                //---------------
                                if (ShirI>200) AND (ShirI<=500) Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) AND (ShirI<=600) Then
                                        ShirI:=600;
                                //---------------
                                if (ShirI>600) AND (ShirI<=1250) Then
                                        ShirI:=1250;
                                End;
                                //+====================================
                                if Res_Tol55=7 Then
                                Begin
                                if ShirI<=200 Then
                                        ShirI:=200;
                                //---------------
                                if (ShirI>200) AND (ShirI<=500) Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) Then
                                        ShirI:=1250;

                                End;
                                //+====================================
                                if Res_Tol55=2 Then
                                Begin
                                if ShirI<=200 Then
                                        ShirI:=200;
                                //---------------
                                if (ShirI>200) AND (ShirI<=500) Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) Then
                                        ShirI:=1250;

                                End;

                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'1'+ #39 +') And ([Толщина]='+Tol_S+') And ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([Высота]=' + #39 +FloatToStr(ShirI)+ #39 +')', ['[Резка Гибка]']) then
                                exit;
                                NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                if NC_S='' Then
                                        NC_S:='0';
                                //SD4.Cells[13,h]:=NC_S;
                                NC:=StrToFloat(NC_S)*Kol_Oboz*Kol_Zap;
//+++++++++++++++++++++++++++++++++++++++++++++

                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'11'+ #39 +') AND ([Высота]=' + #39 +IntToStr(Ugol)+ #39 +') ', ['[Резка Гибка]']) then
                                exit;
                                NC_S2:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                if NC_S2='' Then
                                        NC_S2:='0';

                                Razmet_KPU:=Form1.ADOQuery1.FieldByName('Норма').AsFloat;
                                SD4.Cells[16,h]:=FloatToStr(StrToFloat(NC_S)+Razmet_KPU);
                                NC:=NC+(Razmet_KPU*Kol_Oboz*Kol_Zap);
                                //end;
                                SD4.Cells[13,h]:=FloatToStr(C);
                                SD4.Cells[14,h]:=VElem;
                                SD4.Cells[15,h]:=FloatToStr(NC);
                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                'UPDATE %s SET [Н\ч Ножницы]=' + #39
                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                ') AND (IDКО=' + #39 + IDKO + #39 +
                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['СпецифВозд']) then
                                Exit;
                                //===========================================
                                Nog_List.Values['Typ']:=Tip;
                                Nog_List.Values['Elem']:=Elem;
                                Nog_List.Values['Kol']:=FloatToStr(Kol_Oboz*Kol_Zap);
                                Nog_List.Values['Oboz']:=Oboz;
                                Nog_List.Values['h']:=IntToStr(h);
                                Nog_List.Values['EI']:=EI;
                                Nog_List.Values['Shir']:=Shir;
                                Nog_List.Values['Dlina']:=Dlina;
                                Nog_List.Values['ShirRaz']:=ShirRaz;
                                Nog_List.Values['DlinRaz']:=DlinRaz;
                                Nog_List.Values['Kol_Gib']:=Kol_Gib;
                                Nog_List.Values['Mat']:=Mat;
                                Nog_List.Values['C']:=FloatToStr(C);
                                Nog_List.Values['VElem']:=VElem;
                                Nog_List.Values['NC']:=FloatToStr(NC);
                                //========================================
                                Inc(h);
                                SD4.RowCount:= SD4.RowCount+1;
                                //Break;
                       // End;
                     end;
                     //end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Gib=True) AND (Res_Detal=0) Then

                     Begin
                                Kol_Ed:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Oboz:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                                SD5.Cells[2,n]:=Tip;
                                SD5.Cells[3,n]:=Elem;
                                SD5.Cells[4,n]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD5.Cells[6,n]:=Oboz;
                                SD5.Cells[0,n]:=IntToStr(n);
                                SD5.Cells[1,n]:=EI;
                                SD5.Cells[7,n]:=Dlina;
                                SD5.Cells[8,n]:=Shir;
                                SD5.Cells[9,n]:=DlinRaz;
                                SD5.Cells[10,n]:=ShirRaz;
                                SD5.Cells[11,n]:=IntToStr(KolGib);
                                SD5.Cells[12,n]:=Mat;
                                SD5.Cells[13,n]:=VElem;
                                if  SG1.Cells[14, i]='' then
                                        Kol_Lop:=0
                                Else
                                        Kol_Lop:= StrToInt(SG1.Cells[14, i]);
                                Res_Lop:=AnsiCompareStr('Корпус лопатки',Elem);

                                if Res_Lop=0 Then
                                Begin
                                        SD5.Cells[14,n]:=FloatToStr(0.034*Kol_Ed*Kol_Zap);
                                        SD5.Cells[19,n]:=FloatToStr(0.034);
                                End
                                Else
                                Begin
                                        ShirI:=StrToFloat(ShirRaz);
                                        DlinI:=StrToFloat(DlinRaz);
                                        //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                        Res_List:=Pos('ДПРНМ',Mat);
                                        if Res_List<>0 Then
                                        mat:='1,2';
                                        //Лента Оц 1,5х264
                                        //Лист ОЦ 08пс 1,5х1250х2500
                                        //ОЦ 1,5
                                        //Лента на шестьдесят НЕРЖ 0,3

                                 Res_List:=Pos('НЕРЖ',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                        Res_List:=Pos('Л',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// Лента
                                                Delete(Mat,1,Res_Tol);
                                                //
                                                Res_Tol:=Pos(' ',Mat);//ОЦ
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('х',Mat);//x
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos('*',Mat);//*
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);
                                        end
                                        Else
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);
                                                Delete(Mat,1,Res_Tol);
                                        end;
                                        Res_List:=Pos('ОЦ',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);



                                        if Mat='' Then
                                        Mat:='0';
                                        Tol:=StrToFloat(Mat);
                                        if Tol<=1 Then
                                        Tol_S:='0.55';
                                        if (Tol>1) AND (Tol<=2) Then
                                        Tol_S:='2';
                                        if Tol>2 Then
                                        Tol_S:='3';
                                        if DlinI<=2000 Then
                                        DlinI:=2000
                                        Else
                                        DlinI:=3000;
                                        if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'3'+ #39 +') And ([Толщина]=' + #39 +(Tol_s)+ #39 +
                                ') And ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([Высота]=' + #39 +IntToStr(KolGib)+ #39 +')', ['[Резка Гибка]']) then
                                exit;
                                NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                If NC_S='' Then
                                NC:=0 Else
                                NC:=StrToFloat(NC_S);
                                SD5.Cells[14,n]:=FloatToStr(NC*Kol_Ed*Kol_Zap);
                                SD5.Cells[19,n]:=NC_S;
                                end;
                                //======================
                                Res_Resh:=Pos('Решетка жалюзийная',Elem);
                                if Res_Resh<>0 Then
                                Begin
                                       if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                        'Select * from %s Where   ([№]=' + #39 +'9'+ #39 +') And ([Толщина]=' + #39 +IntToStr(KolGib)+ #39 +
                                        ')', ['[Резка Гибка]']) then
                                        exit;
                                        NC_S:= (Form1.ADOQuery1.FieldByName('Высота').AsString);
                                        If NC_S='' Then
                                                NC:=0 Else
                                                NC:=StrToFloat(NC_S);
                                        //++++++++++++
                                        NC_S2:= (Form1.ADOQuery1.FieldByName('Длина').AsString);
                                        If NC_S2='' Then
                                                NC2:=0 Else
                                                NC2:=StrToFloat(NC_S2);
                                        SD5.Cells[19,n]:=FloatToStr(StrToFloat(NC_S)+StrToFloat(NC_S2));
                                        SD5.Cells[14,n]:=FloatToStr((NC*Kol_Zap)+(NC2/Kol_Zap));
                                end;
                                Res_Lop:=AnsiCompareStr('Стенка лопатки',Elem);

                                if Res_Lop=0 Then
                                Begin
                                     Kol_Lop:=ADOQuery2.FieldByName('Количество').AsInteger;;
                                end;
                                Lenta:=Pos('Лента уплотнительная',Elem);
                                if Lenta <>0 Then
                                Begin
                                      DlinD:=StrToInt(Dlina)-1.5;
                                      B:=StrToInt(Dlina);
                                      if B<335 Then
                                        DlinI:=170;
                                      if (B>=335) AND (B<500) Then
                                      Begin
                                        DlinI:=335;
                                      end;

                                      if (B>=500) AND (B<665) Then
                                      Begin
                                        DlinI:=500;
                                      end;

                                      if (B>=665) AND (B<830) Then
                                      Begin
                                        DlinI:=665;
                                      end;

                                      if (B>=830) AND (B<995) Then
                                      Begin
                                        DlinI:=830;
                                      end;

                                      if (B>=995) AND (B<1160) Then
                                      Begin
                                        DlinI:=995;
                                      end;

                                      if (B>=1160) AND (B<1325) Then
                                      Begin
                                        DlinI:=1160;
                                      end;

                                      if (B>=1325) AND (B<1490) Then
                                      Begin
                                        DlinI:=1325;
                                      end;

                                      if (B>=1490) AND (B<1655) Then
                                      Begin
                                        DlinI:=1490;
                                      end;

                                      if (B>=1655) AND (B<1820) Then
                                      Begin
                                        DlinI:=1655;
                                      end;

                                      if (B>=1820) AND (B<1985) Then
                                      Begin
                                        DlinI:=1820;
                                      end;

                                      if (B>=1985) AND (B<2150) Then
                                      Begin
                                        DlinI:=1985;
                                      end;

                                      if (B>=2150) AND (B<2315) Then
                                      Begin
                                        DlinI:=2150;
                                      end;

                                      if (B>=2315) AND (B<2440) Then
                                      Begin
                                        DlinI:=2315;
                                      end; 

                                      A_Lenta:=(StrToInt(Dlina)-1.5-(Kol_Lop-1)*150)/2;
                                      SD5.Cells[15,n]:=FloatToStr(DlinD);
                                      SD5.Cells[16,n]:= FloatToStr(A_Lenta);
                                      SD5.Cells[17,n]:=IntToStr(Kol_Lop);
                                      SD5.Cells[18,n]:='150';
                                      T:=0.008*Kol_Lop*Kol_Ed*Kol_Zap+0.008*Kol_Ed*Kol_Zap; //Пробивка отверстий в одной ленте
                                      SD5.Cells[20,n]:= FloatToStr(T);
                                      if SD5.Cells[20,n]='' Then
                                        NCB:=(StrToFloat(SD5.Cells[14,n]))
                                      Else
                                      Begin
                                        NCB:=(StrToFloat(SD5.Cells[14,n])+StrToFloat(SD5.Cells[20,n]));
                                        SD5.Cells[14,n]:=FloatToStr(NCB);
                                      End;

                                end;
                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                'UPDATE %s SET [Н\ч Гибка]=' + #39
                                + Form1.ConvertFloat(FloatToStr(NCB)) + #39 +
                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                ') AND (IDКО=' + #39 + IDKO + #39 +
                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['СпецифВозд']) then
                                Exit;
                                Inc(n);
                                SD5.RowCount:= SD5.RowCount+1;
                                //Break;
                     end;
                     //end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Pila=True) AND (Res_Detal=0) Then
                      Begin
                      Kol_Ed:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Oboz:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));

                                SD6.Cells[2,w]:=Tip;
                                SD6.Cells[3,w]:=Elem;
                                SD6.Cells[4,w]:=FloatToStr(Kol_Ed*Kol_Zap);
                               // SD7.Cells[5,w]:=FloatToStr(Kol_Oboz);
                                SD6.Cells[6,w]:=Oboz;
                                SD6.Cells[0,w]:=IntToStr(w);
                                SD6.Cells[1,w]:=EI;
                                SD6.Cells[7,w]:=Dlina;
                                SD6.Cells[8,w]:=Shir;
                                SD6.Cells[9,w]:=DlinRaz;
                                SD6.Cells[10,w]:=ShirRaz;
                                //SD6.Cells[11,w]:=Kol_Gib;
                                SD6.Cells[12,w]:=Mat;
                                SD6.Cells[13,w]:=VElem;
                                //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                Res_List:=Pos('ДПРНМ',Mat);
                                        if Res_List<>0 Then
                                        mat:='1,2';
                                        //Лента Оц 1,5х264
                                        //Лист ОЦ 08пс 1,5х1250х2500
                                        //ОЦ 1,5
                                        //Лента на шестьдесят НЕРЖ 0,3
                                 Res_List:=Pos('НЕРЖ',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                        Res_List:=Pos('Л',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// Лента
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//ОЦ
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('х',Mat);//x
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos('*',Mat);//*
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);
                                        end
                                        Else
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);
                                                Delete(Mat,1,Res_Tol);
                                        end;
                                        Res_List:=Pos('ОЦ',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);



                                        if Mat='' Then
                                        Mat:='0';
                                        //++++++++++++++++++++++++++++++++==
                                          DlinI:=StrToFloat(Dlina);
                                          if DlinI<=250 Then
                                          DlinI:=250;
                                          if (DlinI>250) AND (DlinI<=500) Then
                                          DlinI:=500;
                                          if (DlinI>500) AND (DlinI<=750) Then
                                          DlinI:=750;
                                          if (DlinI>750) AND (DlinI<=1000) Then
                                          DlinI:=1000;
                                          if (DlinI>1000) AND (DlinI<=1250) Then
                                          DlinI:=1250;
                                          if (DlinI>1250) AND (DlinI<=2500) Then
                                          DlinI:=2495;
                                          if (DlinI>2500)  Then
                                          DlinI:=2500;
                                          if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                          'Select * from %s Where   ([№]=' + #39 +'13'+ #39 +') And  ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[Резка Гибка]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S);            // 
                                          SD6.Cells[14,w]:=FloatToStr(NC*Kol_Ed*Kol_Zap);
                                       // End;
                                 if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                'UPDATE %s SET [Н\ч Пила]=' + #39
                                + Form1.ConvertFloat(SD6.Cells[14,w]) + #39 +
                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                ') AND (IDКО=' + #39 + IDKO + #39 +
                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['СпецифВозд']) then
                                Exit;
                                Inc(w);
                                SD6.RowCount:= SD6.RowCount+1;
                                //Break;
                     //end;
                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Pila=True) AND (Res_Detal=0) Then
                      Begin
                      Kol_Ed:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));
                      Kol_Oboz:= StrToFloat(StringReplace(ADOQuery2.FieldByName('Количество').AsString,'.',',',[rfReplaceAll]));

                                SD7.Cells[2,www]:=Tip;
                                SD7.Cells[3,www]:=Elem;
                                SD7.Cells[4,www]:=FloatToStr(Kol_Ed*Kol_Zap);
                               // SD7.Cells[5,w]:=FloatToStr(Kol_Oboz);
                                SD7.Cells[6,www]:=Oboz;
                                SD7.Cells[0,www]:=IntToStr(www);
                                SD7.Cells[1,www]:=EI;
                                SD7.Cells[7,www]:=Dlina;
                                SD7.Cells[8,www]:=Shir;
                                SD7.Cells[9,www]:=DlinRaz;
                                SD7.Cells[10,www]:=ShirRaz;
                                //SD7.Cells[11,w]:=Kol_Gib;
                                SD7.Cells[12,www]:=Mat;
                                SD7.Cells[13,www]:=VElem;
                                Inc(www);
                                SD7.RowCount:= SD7.RowCount+1;
                     end;
                       RecCol:=ADOQuery2.RecordCount;
                       RecNom:=ADOQuery2.RecNo ;
                       ADOQuery2.Next;
                   end;
                End;
        End;

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Dir_main := ExtractFileDir(ParamStr(0));
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);
        Dir :=Form1.Put_KTO+'\CKlapana\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\' + God + '\' + Mes+ '\';
        CreateDir(Dir);
        Puty:=(Dir +'\№ '+Label15.Caption+'.xlsx');
        XL := CreateOleObject('Excel.Application');
        CopyFile(PWideChar(Form1.Put_KTO+'\CKlapana\2013\Воздушные.xlsx'), PWideChar(Dir +'\№ '+Label15.Caption+'.xlsx'), False);
        XL.Workbooks.Open(Dir + '\№ '+Label15.Caption+'.xlsx');
        XL.Application.EnableEvents := false;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG1.RowCount-1 Do //Вальцы
        Begin

            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 2] := RSG1.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 3] := RSG1.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 4] := RSG1.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 5] := RSG1.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 6] := RSG1.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 7] := RSG1.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 8] := RSG1.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 9] := RSG1.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 10] := RSG1.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 11] := RSG1.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[10].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[10].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[10].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG1.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[10].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[10].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[10].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[10].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[10].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[10].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[10].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[10].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 3] :='Вальцовщик';
         XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[10].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[10].Rows[q+7].Font.Size := 12;
                //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG2.RowCount-1 Do //Кромкогиб
        Begin

            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 2] := RSG2.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 3] := RSG2.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 4] := RSG2.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 5] := RSG2.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(5), 6] := RSG2.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(5), 7] := RSG2.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 8] := RSG2.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 9] := RSG2.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 10] := RSG2.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 11] := RSG2.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[11].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[11].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[11].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG2.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[11].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[11].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[11].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[11].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[11].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[11].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[11].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[11].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 3] :='Кромкогибец';
         XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[11].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[11].Rows[q+7].Font.Size := 12;
                         //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG4.RowCount-1 Do //Токарный
        Begin

            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 2] := RSG4.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 3] := RSG4.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 4] := RSG4.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 5] := RSG4.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 6] := RSG4.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 7] := RSG4.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 8] := RSG4.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 9] := RSG4.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 10] := RSG4.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 11] := RSG4.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[12].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[12].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[12].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG4.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[12].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[12].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[12].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[12].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[12].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[12].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[12].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[12].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 3] :='Токарь';
         XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[12].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[12].Rows[q+7].Font.Size := 12;
                         //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG3.RowCount-1 Do //Ротор
        Begin

            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 2] := RSG3.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 3] := RSG3.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 4] := RSG3.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 5] := RSG3.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 6] := RSG3.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 7] := RSG3.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 8] := RSG3.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 9] := RSG3.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 10] := RSG3.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 11] := RSG3.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[13].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[13].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[13].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG3.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[13].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[13].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[13].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[13].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[13].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[13].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[13].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[13].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 3] :='Роторист';
         XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[13].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[13].Rows[q+7].Font.Size := 12;

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  Kanban
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(I+4)+':E'+IntToStr(I+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[1].Rows[i+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[1].Rows[i+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        Q:=Q+2;   //Kanban
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+2].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Select;

        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 10] := 'Мат';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+2, 4] := 'Kanban';
        For I:=0 To SD1.RowCount-1 Do
        Begin
            Razm:=SD1.Cells[7,i+1]+'*'+SD1.Cells[7,i+1];
            if SD1.Cells[0,i+1]='' Then
            Continue;


            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := SD1.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := SD1.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := SD1.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;//Razmer
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := SD1.Cells[4,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := SD1.Cells[5,i+1];

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := SD1.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := SD1.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := SD1.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+5), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[1].Cells[2, 6] :=Label15.Caption;
            Inc(Q);
        end;
        Q:=Q+2;

        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Cells[(Q)+(4), 4] :='Заготовка';
        For I:=0 To SD2.RowCount-1 Do
        Begin
            Razm:=SD2.Cells[7,i+1]+'*'+SD2.Cells[8,i+1];
            if SD2.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := SD2.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := SD2.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := SD2.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := SD2.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := SD2.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := SD2.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := SD2.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := SD2.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+3), 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Cells[(Q)+(4), 4] :='Сбор. Единицы';
        For I:=0 To SD.RowCount-1 Do
        Begin
            Razm:=SD.Cells[7,i+1]+'*'+SD.Cells[8,i+1];
            if SD.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := SD.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := SD.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := SD.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := SD.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := SD.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := SD.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := SD.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := SD.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+5), 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
         XL.ActiveWorkBook.WorkSheets[1].Cells[2, 6] :=Label15.Caption;
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+6, 4] :='Комплектовщик (Заготовка)';
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+8, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+8, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+8, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+8].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+8].Font.Size := 12;
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+6].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+6].Font.Size := 12;
        //++++++++++++++++++++++++++++++++++++++++++++++++  Склад
        Q:=1;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[2].Range['B'+IntToStr(I+3)+':D'+IntToStr(I+3)].Merge;
               XL.ActiveWorkBook.WorkSheets[2].Rows[i+3].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[2].Rows[i+3].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[2].Cells[I+3, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[2].Cells[I+3, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[2].Cells[I+3, 5] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+2].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[2].Select;

        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 10] := 'Мат';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+2, 4] := 'Склад';
        For I:=0 To SD9.RowCount-1 Do
        Begin
            Razm:=SD9.Cells[7,i+1]+'*'+SD9.Cells[8,i+1];
            if SD9.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 2] := SD9.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 3] := SD9.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 4] := SD9.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 6] := SD9.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 7] := SD9.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 8] := SD9.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 9] := SD9.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 10] := SD9.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Range['B'+IntToStr(q+3), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
         XL.ActiveWorkBook.WorkSheets[2].Cells[2, 6] :=Label15.Caption;
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 3] :='Кладовщик';
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[2].Rows[q+8].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[2].Rows[q+8].Font.Size := 12;
        //++++++++++++++++++++++++++++++++++++++++++++++++  Склад Стенки
        Q:=1;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[3].Range['B'+IntToStr(I+3)+':D'+IntToStr(I+3)].Merge;
               XL.ActiveWorkBook.WorkSheets[3].Rows[i+3].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[3].Rows[i+3].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[3].Cells[I+3, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[3].Cells[I+3, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[3].Cells[I+3, 5] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+2].Font.Size := 14;
        //XL.ActiveWorkBook.WorkSheets[3].Select;

        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 10] := 'Мат';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+2, 4] := 'Склад Стенки';
        NC:=0;
        NC_S:='';
        For I:=0 To Stenki.RowCount-1 Do
        Begin
            Razm:=Stenki.Cells[7,i+1];
            if Stenki.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 2] := Stenki.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 3] := Stenki.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 4] := Stenki.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 6] := Stenki.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 7] := Stenki.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 8] := Stenki.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 9] := Stenki.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 10] := Stenki.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Range['B'+IntToStr(q+3), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
         XL.ActiveWorkBook.WorkSheets[3].Cells[2, 6] :=Label15.Caption;
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 3] :='Кладовщик';
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[3].Rows[q+8].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[3].Rows[q+8].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        H:=1;
        NC:=0;
        NC_S:='';
        For I:=0 To SD4.RowCount-1 Do //Ножницы
        Begin
             Elem:= SD4.Cells[3,h];
             Oboz:= SD4.Cells[6,h];
             O:=I+2;
             for Y:=0 To SD4.RowCount-1 Do
             Begin
                if Y>SD4.RowCount-1 Then
                Break;
                If (SD4.Cells[3,o]='') OR (SD4.Cells[3,h]='') Then
                        Continue;
                Res_Elem:=AnsiCompareStr(Elem,SD4.Cells[3,o]);
                Res_Oboz:=AnsiCompareStr(Oboz,SD4.Cells[6,o]);
                if (Res_Elem=0) AND (Res_Oboz=0) Then
                Begin
                        If (SD4.Cells[3,o]='') OR (SD4.Cells[3,h]='') Then
                        Continue;
                        Kol_Oboz:=StrToInt(SD4.Cells[4,o]);
                        NC:=StrToFloat(SD4.Cells[15,o]);
                        Kol_Gib := SD4.Cells[11,o];
                        SD4.Cells[4,h]:=FloatToStr(StrToInt(SD4.Cells[4,h])+Kol_Oboz);
                        SD4.Cells[15,h]:=FloatToStr(StrToFloat(SD4.Cells[15,h])+NC);
                        SD4.Cells[3,o]:='';
                        Form1.DeleteARow(SD4,o);
                        O:=O-1;
                end;
             Inc(O);
             end;
             Inc(H);
        End;
        NC:=0;
        NC_S:='';
        For I:=0 To SD4.RowCount-1 Do //Ножницы
        Begin
            if SD4.Cells[2,i+1]='' Then
            break;
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 2] := SD4.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 3] := SD4.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 4] := SD4.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 5] := SD4.Cells[9,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 6] := SD4.Cells[10,i+1]; //Shir
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 7] := SD4.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 8] := SD4.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 9] := SD4.Cells[15,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 10] := SD4.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 11] := SD4.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[4].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[4].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[4].Cells[4, 1] :=Label15.Caption;
            NC_S:= SD4.Cells[15,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[4].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[4].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[4].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[4].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[4].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[4].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[4].Cells[q+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[4].Cells[q+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 3] :='Резчик';
         XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[4].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[4].Rows[q+7].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To SD3.RowCount-1 Do  //Trumph
        Begin

            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 5] := SD3.Cells[0,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 6] := SD3.Cells[2,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 7] := SD3.Cells[3,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 9] := SD3.Cells[7,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 10] := SD3.Cells[8,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 11] := SD3.Cells[14,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 8] := SD3.Cells[4,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 12] := SD3.Cells[12,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 13] := SD3.Cells[6,i];
            XL.ActiveWorkBook.WorkSheets[5].Range['E21', 'M' +
                IntToStr(I+13)].Borders.LineStyle := 1;
            //XL.ActiveWorkBook.WorkSheets[5].Cells[18, 9] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[5].Cells[13, 2] :='СОПРОВОДИТЕЛЬНЫЙ ЛИСТ №  '+Label15.Caption;
            Q:=I+13;
        end;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[5].Range['F'+IntToStr(Q+4)+':G'+IntToStr(Q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[5].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[5].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[5].Cells[q+4, 5] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[5].Cells[q+4, 6] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[5].Cells[q+4, 8] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 6] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 7] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[5].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[5].Rows[q+7].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        //+++++++++++++++++++++++++++++++++++++++
        H:=1;
        NC:=0;
        NC_S:='';
        For I:=0 To SD5.RowCount-1 Do //Гибка
        Begin
             Elem:= SD5.Cells[3,h];
             Oboz:= SD5.Cells[6,h];
             NC:=0;
             O:=I+2;
             for Y:=0 To SD5.RowCount-1 Do
             Begin
                if Y>SD5.RowCount-1 Then
                Break;
                If (SD5.Cells[3,o]='') OR (SD5.Cells[3,h]='') Then
                        Continue;
                Res_Elem:=AnsiCompareStr(Elem,SD5.Cells[3,o]);
                Res_Oboz:=AnsiCompareStr(Oboz,SD5.Cells[6,o]);
                if (Res_Elem=0) AND (Res_Oboz=0) Then
                Begin
                        If (SD5.Cells[3,o]='') OR (SD5.Cells[3,h]='') Then
                        Continue;
                        Kol_Oboz:=StrToInt(SD5.Cells[4,o]);
                        NC:=StrToFloat(SD5.Cells[14,o]);
                        Kol_Gib := SD5.Cells[11,o];
                        SD5.Cells[4,h]:=FloatToStr(StrToInt(SD5.Cells[4,h])+Kol_Oboz);
                        SD5.Cells[14,h]:=FloatToStr(StrToFloat(SD5.Cells[14,h])+NC);
                        SD5.Cells[11,h]:=FloatToStr(StrToFloat(SD5.Cells[11,h])+StrToInt(Kol_Gib));
                        SD5.Cells[3,o]:='';
                        Form1.DeleteARow(SD5,o);
                        O:=O-1;
                end;
             Inc(O);
             end;
             Inc(H);
        End;
        NC:=0;
        NC_S:='';
        For I:=0 To SD5.RowCount-1 Do //Гибка
        Begin
            If (SD5.Cells[3,i+1]='')Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 2] := SD5.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 3] := SD5.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 4] := SD5.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 5] := SD5.Cells[4,i+1]; //  Kol
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 6] := SD5.Cells[11,i+1]; // KolGib
            Lenta:=Pos('Лента уплотнительная',SD5.Cells[3,i+1]);
            if Lenta<>0 Then
                XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 7] := FloatToStr(StrToFloat(SD5.Cells[14,i+1])+0.05) // N\C
            Else
                XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 7] := SD5.Cells[14,i+1];
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 8] := SD5.Cells[9,i+1]; //A
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 9] := SD5.Cells[10,i+1]; //B

            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 10] := SD5.Cells[15,i+1]; //B
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 11] := SD5.Cells[17,i+1]; //B
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 12] := SD5.Cells[18,i+1]; //B
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 13] := SD5.Cells[16,i+1]; //B

            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 14] := SD5.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 15] := SD5.Cells[6,i+1]; //Chert
            XL.ActiveWorkBook.WorkSheets[6].Range['B4', 'O' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[6].Cells[2, 5] :=Label15.Caption;
            NC_S:= SD5.Cells[14,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            Q:=I+5;
        end;
        XL.ActiveWorkBook.WorkSheets[6].Cells[(Q)+1, 6] :='Итого';
        NC_S:= FloatToStr(NC);
        XL.ActiveWorkBook.WorkSheets[6].Cells[(q)+1, 7] :=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[6].Range['C'+IntToStr(q+2)+':D'+IntToStr(q+2)].Merge;
               XL.ActiveWorkBook.WorkSheets[6].Rows[q+2].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[6].Rows[q+2].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[6].Cells[q+2, 2] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[6].Cells[q+2, 3] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[6].Cells[q+2, 6] := SG1.Cells[2,i+1];

               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[6].Rows[q+5].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[6].Rows[q+5].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        //+++++++++++++++++++++++++++++++++++++++
        H:=1;
        O:=1;
        For I:=0 To SD6.RowCount-1 Do //Pila
        Begin
             Elem:= SD6.Cells[3,h];
             Oboz:= SD6.Cells[6,h];
             O:=I+2;
             for Y:=0 To SD6.RowCount-1 Do
             Begin
                if Y>SD6.RowCount-1 Then
                Break;
                If (SD6.Cells[3,o]='') OR (SD6.Cells[3,h]='') Then
                        Continue;
                Res_Elem:=AnsiCompareStr(Elem,SD6.Cells[3,o]);
                Res_Oboz:=AnsiCompareStr(Oboz,SD6.Cells[6,o]);
                if (Res_Elem=0) AND (Res_Oboz=0) Then
                Begin
                        If (SD6.Cells[3,o]='') OR (SD6.Cells[3,h]='') Then
                        Continue;
                        Kol_Oboz:=StrToInt(SD6.Cells[4,o]);
                        NC:=StrToFloat(SD6.Cells[14,o]);
                        //Kol_Gib := SD6.Cells[11,o];
                        SD6.Cells[4,h]:=FloatToStr(StrToInt(SD6.Cells[4,h])+Kol_Oboz);
                        SD6.Cells[14,h]:=FloatToStr(StrToFloat(SD6.Cells[14,h])+NC);
                        //SD6.Cells[11,h]:=FloatToStr(StrToFloat(SD6.Cells[11,h])+StrToInt(Kol_Gib));
                        SD6.Cells[3,o]:='';
                        Form1.DeleteARow(SD6,o);
                        O:=O-1;
                end;
             Inc(O);
             end;
             Inc(H);
        End;
         NC:=0;
        For I:=0 To SD6.RowCount-1 Do //Pila
        Begin
            If (SD6.Cells[3,i+1]='')Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 2] := SD6.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 3] := SD6.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 4] := SD6.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 5] := SD6.Cells[9,i+1]; //Dlin


            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 7] := SD6.Cells[4,i+1]; //KolEd
            Res:=Pos( 'Профиль ФЭЗ.П.0043',SD6.Cells[12,i+1]);
            if Res<>0 Then
            Begin
                 XL.ActiveWorkBook.WorkSheets[7].Range['H'+IntToStr(I+5), 'H' +IntToStr(I+5)].Font.Italic := True;
                 XL.ActiveWorkBook.WorkSheets[7].Range['H'+IntToStr(I+5), 'H' +IntToStr(I+5)].Font.Bold := True;
            End;
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 8] := SD6.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 9] := SD6.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[7].Range['B4', 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[7].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[7].Cells[4, 1] :=Label15.Caption;
            if SD6.Cells[14,i+1]<>'' Then
            Begin
                NC2:=StrToFloat( SD6.Cells[14,i+1]); //*StrToInt(SD6.Cells[4,i+1]
                NC_S2:=FloatToStr(NC2);
                NC_S2:=Float(NC_S2);
                PosTrud := Pos(',', NC_S2);
                if PosTrud <> 0 then
                begin
                        Delete(NC_S2, PosTrud, 1);
                        Insert('.', NC_S2, PosTrud);
                end;
                XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 6] := NC_S2; //
                NC_S:= FloatToStr(NC2);
                if NC_S='' Then
                NC_S:='0';
                NC:=NC+StrToFloat(NC_S);
            End;
            Q:=I+5;
        end;
        XL.ActiveWorkBook.WorkSheets[7].Cells[(Q)+1, 5] :='Итого';
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        PosTrud := Pos(',', NC_S); //Трудоемкость FLOAT
        if PosTrud <> 0 then
        begin
                Delete(NC_S, PosTrud, 1);
                Insert('.', NC_S, PosTrud);
        end;
       // XL.ActiveWorkBook.WorkSheets[7].Range['F'+IntToStr(q+1)+':F'+IntToStr(q+1)].Select;
       // XL.ActiveWorkBook.WorkSheets[7].Selection.NumberFormat := '0.00';
        XL.ActiveWorkBook.WorkSheets[7].Cells[(q)+1, 6] :=NC_S;
                For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[7].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[7].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[7].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[7].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[7].Cells[q+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[7].Cells[q+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[7].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[7].Rows[q+7].Font.Size := 12;
                 //+++++++++++++++++++++++++++++++++++++++
                  NC:=0;
        For I:=0 To SD6.RowCount-1 Do //Abraziv
        Begin

            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 2] := SD6.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 3] := SD6.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 4] := SD6.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 5] := SD6.Cells[9,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 6] := SD6.Cells[14,i+1]; //Н\ч

            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 7] := SD6.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 8] := SD6.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 9] := SD6.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[8].Range['B4', 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[8].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[8].Cells[4, 1] :=Label15.Caption;
            Q:=I+5;
        end;
                For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[8].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[8].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[8].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[8].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[8].Cells[q+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[8].Cells[q+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[8].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[8].Rows[q+7].Font.Size := 12;
         //+++++++++++++++++++++++++++++++++++++++++++++++
         Q:=6;
        For I:=0 To SD3.RowCount-1 Do //Сопроводитр
        Begin
            //If (SD3.Cells[3,i+1]='')Then
            //Continue;
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 2] := SD3.Cells[0,i];
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 3] := SD3.Cells[2,i];//Тип
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 4] := SD3.Cells[3,i]; //Элемент
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 5] := SD3.Cells[7,i]; //Dlin
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 6] := SD3.Cells[8,i]; //Shir

            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 7] := SD3.Cells[4,i]; //KolEd
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 8] := SD3.Cells[12,i]; //Mat
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 9] := SD3.Cells[6,i]; //Obozn
            XL.ActiveWorkBook.WorkSheets[9].Range['B4', 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;


            Inc(Q);
        end;
        XL.ActiveWorkBook.WorkSheets[9].Cells[5, 4] :='Trumpf';
            XL.ActiveWorkBook.WorkSheets[9].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[9].Cells[4, 1] :=Label15.Caption;
        Q:=Q+4;
        //Nog
        XL.ActiveWorkBook.WorkSheets[9].Cells[q-2, 4] :='Ножницы';
        For I:=0 To SD4.RowCount-1 Do //Сопроводитр
        Begin
            If (SD4.Cells[3,i+1]='')Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 2] := SD4.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 3] := SD4.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[9].Cells[q, 4] := SD4.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 5] := SD4.Cells[7,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 6] := SD4.Cells[8,i+1]; //Shir

            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 7] := SD4.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 8] := SD4.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 9] := SD4.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[9].Range['B4', 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;

           Inc(Q);
        end;
        Q:=Q+4;
        //Nog
        XL.ActiveWorkBook.WorkSheets[9].Cells[q-2, 4] :='Гибка';
        For I:=0 To SD5.RowCount-1 Do //Сопроводитр
        Begin
            If (SD5.Cells[3,i+1]='')Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 2] := SD5.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 3] := SD5.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[9].Cells[q, 4] := SD5.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 5] := SD5.Cells[7,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 6] := SD5.Cells[8,i+1]; //Shir

            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 7] := SD5.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 8] := SD5.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 9] := SD5.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[9].Range['B4', 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;

           Inc(Q);
        end;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[9].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[9].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[9].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[9].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[9].Cells[q+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[9].Cells[q+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[9].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[9].Rows[q+7].Font.Size := 12;
        XL.Visible := True;
        XL := UnAssigned;
end;

procedure TFNAklVoz.SG2DblClick(Sender: TObject);
begin
        Label15.Caption := SG2.Cells[4, SG2.Row];
        Button3.Click;
end;

procedure TFNAklVoz.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFNAklVoz.SG2Click(Sender: TObject);
var I:Integer;
begin
{  for I:=0 to SG2.RowCount-1 do
  begin
    mmo1.Lines.Add(SG2.Cells[5,i]);
  end;
Form1.DeleteARow(SG2,2);
  for I:=0 to SG2.RowCount-1 do
  begin
    mmo2.Lines.Add(SG2.Cells[5,i]);
  end;  }
end;


procedure TFNAklVoz.btn1Click(Sender: TObject);
var I:Integer;
nog:TStringList;
begin
 Nog:=TStringList.Create;
 for I:=0 to 10 do
 begin
   Nog.Values['1']:=IntToStr(I);
   Nog.Values['2']:=IntToStr(I+2);
   Nog.Values['3']:=IntToStr(I+3);
 end;
 for I:=0 to 10 do
 begin
   mmo1.Lines.Add(Nog.Values['1']);
   mmo1.Lines.Add(Nog.Values['2']);
   mmo1.Lines.Add(Nog.Values['3']);
 end;
end;
procedure TFNAklVoz.btn2Click(Sender: TObject);

var
        Kol_ZAP, Kol, Kol_Zap_Ranee, Res, Nomer, Zak_Int, i, A, B, D, y, F_KPU,
        F_KPD,xy, Res1, e, Zag, Zap, j, g, k,Q,Res_Mat,Res_Proch,Res_Stand,qq,Res_Sbor,Res_Klap,Podstavka,Srav,Srav1:Integer;
        Res_VElem,Res_Elem,Res_Oboz,Flag,o,p,s,h,f,m,n,v,w,KolGib,AA,bb,Res_Tol,F2,F1,Res_KPU,Res_KPD,Res_list,Res_Ger:Integer;
        Res_Lop,Sten_V,Sten_G,hh,pp,Res_Kog_Priv,GG,Res_Kompl_priv,ww,www,Res_Detal,Os,R25,Ugol,Lenta,Kol_Otv,ID,Dl_Slova,Kol_Lop:Integer;
        Zak, Dat, Plan_Dat, Vn_Dat, Nom, Pos_Vst, Pos_Ml, R, Pereh, Privod,
                Zag_S,
                Zap_S, Dir_main, God, Mes, Fil,Dir,DlinRaz,ShirRaz,Mat,
                Dlina,Shir,Kol_Gib,Razm,Tol_S,Pos_Flan1,Pos_Flan,Pos_Dop,
                Pos_Sn,Pos_Privod,Pos_Isp,Pos_Ram,SvarkaS,Kol_Lops,BZ: string;
        Svar, Sbor, Izdel,VElem,Elem,Oboz, IDGP,IDKO,Tip,EI,NC_S,Pos1,Pos2,Pos3,Pos4,Pos5,OboznSh,NC_S2,N2,Diam,N_S: string;
       Res_Tol55, Kol_Ed,Kol_Oboz,Kol_Ed1,Kol_Oboz1,NC,NC2,ND22,C,DlinI,ShirI,Tol,Razmet_KPU,DlinD,A_Lenta,NC_OB,T,NCB,Diam_i: Double;
        Dat1, Dat2, Dat3: TDate;
        Kanban,Trumph,Nog,Ugl,Gib,Prok,Pila,Svarka,Tehnolog,VALCI,HAco,Kromka,Tokarn,Rotor:Boolean;
        XL, XL1, XL2, XL_Temp, V1: Variant;
        Res_Tol1,Res_Tol2,Res_Resh,Res_N2,Res_Svarka,PosTrud,DD,NERG,
        S1,S2,S3,S4:Integer;
        Nog_List:TStringList;
        RecNom,RecCol:Integer;
begin
        Nog_List :=TStringList.Create;
        F_KPU := 0;
        F_KPD := 0;
        y := 1;
        Xy := 1;
        g:=1;
        K:=1;
        O:=1;
        S:=1;
        P:=1;
        H:=1;
        N:=1;
        M:=1;
        v:=1;
        w:=1;
        ww:=1;
        www:=1;
        Q:=1;
        qq:=1;
        AA:=1;
        PP:=1;

        S1:=1;
        S2:=1;
        S3:=1;
        S4:=1;
        e := 30;

        //Vn_Dat := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);

        Clear_StringGrid(SD1);
        Clear_StringGrid(SD2);
        Clear_StringGrid(SD3);
        Clear_StringGrid(SD4);
        Clear_StringGrid(SD5);
        Clear_StringGrid(SD6);
        Clear_StringGrid(SD7);
        Clear_StringGrid(SD8);
        Clear_StringGrid(SD9);
        Clear_StringGrid(SD);
        Clear_StringGrid(Stenki);

        Clear_StringGrid(RSG1);
        Clear_StringGrid(RSG2);
        Clear_StringGrid(RSG3);
        Clear_StringGrid(RSG4);

        Memo2.Lines.Clear;
        SD1.Cells[2,0]:='Kanban True';
        SD2.Cells[2,0]:='Kanban False';
        SD3.Cells[2,0]:='Trumph';
        SD4.Cells[2,0]:='Nognicy';
        SD5.Cells[2,0]:='Gibka';
        SD6.Cells[2,0]:='Prokat';
        SD7.Cells[2,0]:='Pila';
        SD8.Cells[2,0]:='Uglorub';


        for i := 1 to SG1.RowCount do
        begin
                if SG1.Cells[2, i] <> '' then
                begin
                        Kol_Oboz:=0;
                        Kol_Zap:=0;
                        NC_OB:=0;
                        NC:=0;
                        NC_S:='';
                        NC_S2:='';
                        Tip:='';
                        BZ:=SG1.Cells[1, 0];
                        Nom := SG1.Cells[3, i];
                        Izdel := SG1.Cells[5, i];
                        if  SG1.Cells[14, i]='' then
                        Kol_Lop:=0
                        Else
                        Kol_Lop:= StrToInt(SG1.Cells[14, i]);
                        Kol_Zap := StrToInt(SG1.Cells[2, i]);
                        Tip:=Izdel;
                        Izdel := SG1.Cells[5, i];
                        IDGP:=SG1.Cells[4, i];
                        IDKO:=SG1.Cells[11, i];
                        //++++++++++++++++++++++
                        SQLConnector1.Connected:=True;
                        ADOQuery2.Close;
                        ADOQuery2.SQL.Clear;
                        ADOQuery2.SQL.Text :='Select * from СпецифВозд Where   (IdГП=' + #39 +
                                SG1.Cells[4, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[11, i] + #39 +') AND (Элемент=' + #39 +
                                'Лопатка' + #39 +') ' ;
                        ADOQuery2.Active := true;

                       { if not Form1.mkQuerySelect(Form1.ADOQuery2, //Вытащить кол лопаток
                                , ['СпецифВозд']) then
                                exit;   }
                        Kol_lopS:=ADOQuery2.FieldByName('Количество').AsString;
                        if Kol_Lops='' then
                        Kol_Lop:=0
                        else
                        Kol_Lop:= ADOQuery2.FieldByName('Количество').AsInteger;
                       { if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where   (IdГП=' + #39 +
                                SG1.Cells[4, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[11, i] + #39 +') ', ['СпецифВозд']) then
                                exit;  }
                       SQLConnector1.Connected:=True;
                        ADOQuery2.Close;
                        ADOQuery2.SQL.Clear;
                        ADOQuery2.SQL.Text :='Select * from СпецифВозд Where   (IdГП=' + #39 +
                                SG1.Cells[4, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[11, i] + #39 +') ' ;
                        ADOQuery2.Active := true;
                 for J:=0 To  ADOQuery2.RecordCount-1 Do
                 Begin
                      Flag:=0;
                      KolGib:=0;
                      VElem:= ADOQuery2.FieldByName('ВидЭлемента').AsString;
                      Elem:= ADOQuery2.FieldByName('Элемент').AsString;
                      Oboz:= ADOQuery2.FieldByName('Обозначение').AsString;
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                      //Trumph,Nog,Ugl,Gib,Prok,Pila
                      Dlina:=ADOQuery2.FieldByName('Длина').AsString;
                      Diam:=ADOQuery2.FieldByName('Диаметр').AsString;
                      Shir:= ADOQuery2.FieldByName('Ширина').AsString;
                      DlinRaz:= ADOQuery2.FieldByName('ДлинаРазв').AsString;
                      ShirRaz:=ADOQuery2.FieldByName('ШиринаРазв').AsString;
                      if DlinRaz ='' Then
                            DlinRaz:='0';
                      if ShirRaz ='' Then
                            ShirRaz:='0';
                      Kol_Gib:= ADOQuery2.FieldByName('КолГибов').AsString;
                      if Kol_Gib=''  Then
                      KolGib:=0
                      Else
                      KolGib:=StrToInt(Kol_Gib);

                      Mat:= ADOQuery2.FieldByName('Материал').AsString;
                      if Mat='' Then
                                        Mat:='0';
                      EI:= ADOQuery2.FieldByName('ЕИ').AsString;

                      Kanban:= ADOQuery2.FieldByName('Канбан').AsBoolean;
                      Trumph:= ADOQuery2.FieldByName('Trumph').AsBoolean;
                      Nog:= ADOQuery2.FieldByName('Ножницы').AsBoolean;
                      Ugl:= ADOQuery2.FieldByName('Углоруб').AsBoolean;
                      Gib:= ADOQuery2.FieldByName('Гибка').AsBoolean;
                      Prok:= ADOQuery2.FieldByName('Прокатка').AsBoolean;
                      Pila:= ADOQuery2.FieldByName('Пила').AsBoolean;

                      HACO:= ADOQuery2.FieldByName('HACO').AsBoolean;
                      VALCI:= ADOQuery2.FieldByName('Вальцы').AsBoolean;
                      KROMKA:=ADOQuery2.FieldByName('Кромкогиб').AsBoolean;
                      Res:=Pos('Обечайка',Elem); // (Res<>0) or
                      Res1:=Pos('Корпус выкатной',Elem);
                     if (Diam<>'') AND ((Res1<>0)) then
                     begin
                      Diam_i :=StrToFloat(Diam);
                      If Diam_I<400 then
                      begin
                         TOKARN:= True;
                         ROTOR:=  False;
                      end
                      else
                      begin
                         TOKARN:=False ;
                         ROTOR:=True  ;
                      end;
                      //TOKARN:= ADOQuery2.FieldByName('Токарный').AsBoolean;
                      //ROTOR:= ADOQuery2.FieldByName('Ротор').AsBoolean;
                     end;
                      //==================================================
                     Res_Mat:=Pos('Материалы',VELem);
                     Res_Proch:=Pos('Прочие изделия',VELem);
                     Res_Stand:=Pos('Стандартные изделия',VELem);
                     Res_Sbor:=Pos('Сборочные единицы',VELem);
                     Res_Klap:=Pos('Клапан',ELem);
                     Res_Kompl_priv:=Pos('Комплект привода',ELem);
                      //XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                      Res_Detal:=AnsiCompareStr('Детали',VElem);
                      OboznSh:=Oboz;
                      {Res:=Pos('-',OboznSh);
                      if Res<>0 Then                      //   AND ((Sten_G<>0) ANd (Sten_V<>0))
                        Delete(OboznSh,Res,30); }
                      if (Res_Detal=0) Then
                      Begin
                        if not Form1.mkQuerySelect66(Form1.ADOQuery4,
                                'Select * from %s Where   (Обозначение=' + #39 +
                                OboznSh + #39 +') AND (Технолог=1) ', ['СпецифОбщая']) then
                                exit;
                        Tehnolog:=Form1.ADOQuery4.FieldByName('Технолог').AsBoolean;
                        if Tehnolog=False then
                        begin
                                Kanban:=False ;
                                Trumph:=False ;
                                Nog:=False;
                                Ugl:= False;
                                Gib:= False;
                                Prok:= False;
                                Pila:= False;
                                Svarka:=False;
                        end
                        Else
                        begin
                        Kanban:= Form1.ADOQuery4.FieldByName('Канбан').AsBoolean;
                        Trumph:= Form1.ADOQuery4.FieldByName('Trumph').AsBoolean;
                        Nog:= Form1.ADOQuery4.FieldByName('Ножницы').AsBoolean;

                        //HACO:= Form1.ADOQuery4.FieldByName('HACO').AsBoolean;
                        VALCI:= Form1.ADOQuery4.FieldByName('Вальцы').AsBoolean;
                        KROMKA:= Form1.ADOQuery4.FieldByName('Кромкогиб').AsBoolean;
                        TOKARN:= Form1.ADOQuery4.FieldByName('Токарный').AsBoolean;
                        ROTOR:= Form1.ADOQuery4.FieldByName('Ротор').AsBoolean;
                        Res:=Pos('Обечайка',Elem);   //(Res<>0) or
                     Res1:=Pos('Корпус выкатной',Elem);
                     {if (Diam<>'') AND ((Res1<>0)) then
                     begin
                      Diam_i :=StrToFloat(Diam);
                      If Diam_I<400 then
                      begin
                         TOKARN:= True;
                         ROTOR:=  False;
                      end
                      else
                      begin
                         TOKARN:=False ;
                         ROTOR:=True  ;
                      end;  
                      //TOKARN:= ADOQuery2.FieldByName('Токарный').AsBoolean;
                      //ROTOR:= ADOQuery2.FieldByName('Ротор').AsBoolean;
                     end; 
                        if (chk1.Checked=True) then
                        Begin
                                if Nog=True then
                                begin
                                Trumph:= True;
                                Nog:=False;
                                end;
                        end;}
                        Ugl:= Form1.ADOQuery4.FieldByName('Углоруб').AsBoolean;
                        Gib:= Form1.ADOQuery4.FieldByName('Гибка').AsBoolean;
                        Prok:= Form1.ADOQuery4.FieldByName('Прокатка').AsBoolean;
                        Pila:= Form1.ADOQuery4.FieldByName('Пила').AsBoolean;
                        Svarka:=Form1.ADOQuery4.FieldByName('Сварка').AsBoolean;
                        end;
                        Ugol:= Form1.ADOQuery4.FieldByName('Угол').AsInteger;
                        Kol_Gib:= Form1.ADOQuery4.FieldByName('КолГибов').AsString;
                        if Kol_Gib=''  Then
                        KolGib:=0
                        Else
                        KolGib:=StrToInt(Kol_Gib);

                      //XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                      If (Svarka=True) Then
                      Begin
                            Dl_Slova:=Length(OboznSh);
                            Delete(OboznSh,Dl_Slova,1);
                      end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++++
                      NC:=0;
                      NC_S:='';
                      Tol:=0;
                      Tol_S:='';
                      //==============================================
                      Res:=Pos('Тяга',Elem);
                      NERG:=Pos('НЕРЖ',Oboz);
                      {if (Res<>0) and (Nerg<>0) then
                      begin
                        Trumph:=True;
                        Gib:=True;
                        Kanban:=False;
                        if not Form1.mkQueryUpdate2(Form1.ADOQuery1,
                                        'UPDATE %s SET [Trumph]=' +
                                        #39+ 'True' + #39 +
                                        ', Гибка='+#39+ 'True' + #39 +
                                        ', Канбан='+#39+ 'False' + #39 +
                                        ' WHERE (Обозначение=' + #39 + Oboz + #39 +') AND ([IdГП]=' + #39 + IDGP + #39 +
                                                ') ', ['СпецифВозд']) then
                                        Exit;
                        {DD:=1;
                        Continue;
                      end; }
                      //==============================
                      if VALCI =True then
                        begin
                                RSG1.Cells[0,s1]:=IntToStr(s1);
                                RSG1.Cells[1,s1]:=Tip;
                                RSG1.Cells[2,s1]:=Elem;
                                RSG1.Cells[3,s1]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG1.Cells[4,s1]:=Diam;
                                RSG1.Cells[5,s1]:=Oboz;
                                //RSG1.Cells[6,s1]:=Dlina;
                                RSG1.Cells[6,s1]:=Kol_Gib;
                                RSG1.Cells[7,s1]:=Mat;
                                RSG1.Cells[8,s1]:=N_S;//Н\ч
                                Inc(S1);
                                RSG1.RowCount:= RSG1.RowCount+1;
                        end;
                        //==============================
                        if KROMKA =True then
                        begin
                                RSG2.Cells[0,s2]:=IntToStr(s2);
                                RSG2.Cells[1,s2]:=Tip;
                                RSG2.Cells[2,s2]:=Elem;
                                RSG2.Cells[3,s2]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG2.Cells[4,s2]:=Diam;
                                RSG2.Cells[5,s2]:=Oboz;
                                //RSG2.Cells[6,s2]:=Dlina;
                                RSG2.Cells[6,s2]:=Kol_Gib;
                                RSG2.Cells[7,s2]:=Mat;
                                RSG2.Cells[8,s2]:=N_S;//Н\ч
                                Inc(S2);
                                RSG2.RowCount:= RSG2.RowCount+1;
                        end;
                        //==============================
                        if ROTOR =True then
                        begin
                                RSG3.Cells[0,s3]:=IntToStr(s3);
                                RSG3.Cells[1,s3]:=Tip;
                                RSG3.Cells[2,s3]:=Elem;
                                RSG3.Cells[3,s3]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG3.Cells[4,s3]:=Diam;
                                RSG3.Cells[5,s3]:=Oboz;
                                //RSG3.Cells[6,s3]:=Dlina;
                                RSG3.Cells[6,s3]:=Kol_Gib;
                                RSG3.Cells[7,s3]:=Mat;
                                RSG3.Cells[8,s3]:=N_S;//Н\ч
                                Inc(S3);
                                RSG3.RowCount:= RSG3.RowCount+1;
                        end;
                        //==============================
                        if TOKARN =True then
                        begin
                                RSG4.Cells[0,s4]:=IntToStr(s4);
                                RSG4.Cells[1,s4]:=Tip;
                                RSG4.Cells[2,s4]:=Elem;
                                RSG4.Cells[3,s4]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                RSG4.Cells[4,s4]:=Diam;
                                RSG4.Cells[5,s4]:=Oboz;
                                //RSG4.Cells[6,s]:=Dlina;
                                RSG4.Cells[6,s4]:=Kol_Gib;
                                RSG4.Cells[7,s4]:=Mat;
                                RSG4.Cells[8,s4]:=N_S;//Н\ч
                                Inc(S4);
                                RSG4.RowCount:= RSG4.RowCount+1;
                        end;

                      end;
                        //======================================================
                      Res_Stand:=Pos('Стандартные изделия',VELem);
                      if (Kanban=True) AND (REs_Stand=0) AND (Res_Sbor=0) AND (Res_Mat=0)   Then
                      Begin
                        For y:=1 To SD1.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD1.Cells[3,y]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD1.Cells[6,y]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin

                               Kol_Ed:=Kol_Ed+StrToFloat(SD1.Cells[4,y]);
                               SD1.Cells[4,y]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD1.Cells[5,y]);
                               SD1.Cells[5,y]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                        end;
                        If Flag=0 Then
                        Begin
                                SD1.Cells[2,k]:=Tip;
                                SD1.Cells[3,k]:=Elem;
                                SD1.Cells[4,k]:=FloatToStr(Kol_Ed);
                                SD1.Cells[5,k]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD1.Cells[6,k]:=Oboz;
                                SD1.Cells[0,k]:=IntToStr(k);
                                SD1.Cells[1,k]:=EI;
                                SD1.Cells[7,k]:=Dlina;
                                SD1.Cells[8,k]:=Shir;
                                SD1.Cells[9,k]:=DlinRaz;
                                SD1.Cells[10,k]:=ShirRaz;
                                SD1.Cells[11,k]:=Kol_Gib;
                                SD1.Cells[12,k]:=Mat;
                                SD1.Cells[13,k]:=VElem;
                                Inc(k);
                                SD1.RowCount:= SD1.RowCount+1;
                                //Break;
                        end;
                      end;

                     //Заготовка (Kanban=False) AND (Res_Mat=0) AND (Res_Proch=0) AND (Res_Stand=0) AND (Res_Sbor=0) AND (Res_Kompl_priv=0) AND (Res_Klap=0)
                      if (Kanban=False) AND (Res_Detal=0) AND (Tehnolog=True) Then
                      //For y:=1 To SD1.RowCount Do Заготовка
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                        For O:=1 To SD2.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD2.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD2.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(SD2.Cells[4,o]);
                               SD2.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD2.Cells[5,o]);
                               SD2.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                        end;
                        If Flag=0 Then
                        Begin
                                SD2.Cells[2,p]:=Tip;
                                SD2.Cells[3,p]:=Elem;
                                SD2.Cells[4,p]:=FloatToStr(Kol_Ed);
                                SD2.Cells[5,p]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD2.Cells[6,p]:=Oboz;
                                SD2.Cells[0,p]:=IntToStr(p);
                                SD2.Cells[1,p]:=EI;
                                SD2.Cells[7,p]:=Dlina;
                                SD2.Cells[8,p]:=Shir;
                                SD2.Cells[9,p]:=DlinRaz;
                                SD2.Cells[10,p]:=ShirRaz;
                                SD2.Cells[11,p]:=Kol_Gib;
                                SD2.Cells[12,p]:=Mat;
                                SD2.Cells[13,p]:=VElem;
                                Inc(p);
                                SD2.RowCount:= SD2.RowCount+1;
                                //Break;
                        end;
                      end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++ (Kanban=False)  AND
                      if (Res_Sbor<>0)   Then       //Сборка
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;

                        For O:=1 To SD.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(SD.Cells[4,o]);
                               SD.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD.Cells[5,o]);
                               SD.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                        end;
                        If Flag=0 Then
                        Begin
                                SD.Cells[2,aa]:=Tip;
                                SD.Cells[3,aa]:=Elem;
                                SD.Cells[4,aa]:=FloatToStr(Kol_Ed);
                                SD.Cells[5,aa]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD.Cells[6,aa]:=Oboz;
                                SD.Cells[0,aa]:=IntToStr(aa);
                                SD.Cells[1,aa]:=EI;
                                SD.Cells[7,aa]:=Dlina;
                                SD.Cells[8,aa]:=Shir;
                                SD.Cells[9,aa]:=DlinRaz;
                                SD.Cells[10,aa]:=ShirRaz;
                                SD.Cells[11,aa]:=Kol_Gib;
                                SD.Cells[12,aa]:=Mat;
                                SD.Cells[13,aa]:=VElem;
                                Inc(aa);
                                SD.RowCount:= SD.RowCount+1;
                                //Break;
                        end;
                      end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND ((Res_Stand<>0) OR (Res_Mat<>0) OR (Res_Proch<>0) OR  (Res_Kompl_priv<>0)) Then       //СCklad
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;

                        For O:=1 To SD9.RowCount Do
                        Begin
                           Res_Elem:=AnsiCompareStr(Elem,SD9.Cells[3,o]);
                           Res_Oboz:=AnsiCompareStr(Oboz,SD9.Cells[6,o]);
                           if (Res_Elem=0) AND (Res_Oboz=0) Then
                           Begin
                               Kol_Ed:=Kol_Ed+StrToFloat(SD9.Cells[4,o]);
                               SD9.Cells[4,o]:=FloatToStr(Kol_Ed);

                               Kol_Oboz:=(Kol_Oboz*Kol_Zap)+StrToFloat(SD9.Cells[5,o]);
                               SD9.Cells[5,o]:=FloatToStr(Kol_Oboz);
                               Flag:=1;
                               Break;
                           end ;
                        end;
                        If Flag=0 Then
                        Begin
                                SD9.Cells[2,qq]:=Tip;
                                SD9.Cells[3,qq]:=Elem;
                                SD9.Cells[4,qq]:=FloatToStr(Kol_Ed);
                                SD9.Cells[5,qq]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD9.Cells[6,qq]:=Oboz;
                                SD9.Cells[0,qq]:=IntToStr(qq);
                                SD9.Cells[1,qq]:=EI;
                                SD9.Cells[7,qq]:=Dlina;
                                SD9.Cells[8,qq]:=Shir;
                                SD9.Cells[9,qq]:=DlinRaz;
                                SD9.Cells[10,qq]:=ShirRaz;
                                SD9.Cells[11,qq]:=Kol_Gib;
                                SD9.Cells[12,qq]:=Mat;
                                SD9.Cells[13,qq]:=VElem;
                                Inc(qq);
                                SD9.RowCount:= SD9.RowCount+1;
                                //Break;
                        end;
                      end;
                     Res_N2:=AnsiCompareStr('2Н',N2);
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Trumph=True) AND (Res_Detal=0) Then
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                                SD3.Cells[2,s]:=Tip;
                                SD3.Cells[3,s]:=Elem;
                                SD3.Cells[4,s]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                //SD3.Cells[5,s]:=FloatToStr(Kol_Oboz);
                                SD3.Cells[6,s]:=Oboz;
                                SD3.Cells[0,s]:=IntToStr(s);
                                SD3.Cells[1,s]:=EI;
                                SD3.Cells[7,s]:=Dlina;
                                SD3.Cells[8,s]:=Shir;
                                SD3.Cells[9,s]:=DlinRaz;
                                SD3.Cells[10,s]:=ShirRaz;
                                SD3.Cells[11,s]:=Kol_Gib;
                                SD3.Cells[12,s]:=Mat;
                                if DlinRaz='' Then
                                DlinRaz:='0';
                                SD3.Cells[14,s]:=FloatToStr(C);
                                SD3.Cells[13,s]:=VElem;
                                Inc(s);
                                SD3.RowCount:= SD3.RowCount+1;
                     end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Nog=True) AND (Res_Detal=0) Then
                     Begin

                                Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                                Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                                SD4.Cells[2,h]:=Tip;
                                SD4.Cells[3,h]:=Elem;
                                SD4.Cells[4,h]:=FloatToStr(Kol_Oboz*Kol_Zap);
                                SD4.Cells[6,h]:=Oboz;
                                SD4.Cells[0,h]:=IntToStr(h);
                                SD4.Cells[1,h]:=EI;
                                SD4.Cells[7,h]:=Shir;
                                SD4.Cells[8,h]:=Dlina;
                                SD4.Cells[9,h]:=ShirRaz;
                                SD4.Cells[10,h]:=DlinRaz;
                                SD4.Cells[11,h]:=Kol_Gib;
                                SD4.Cells[12,h]:=Mat;
                                if DlinRaz='' Then
                                DlinRaz:='0';
                                //+ ++++++++++++++++++++++++++++++++++++++++++
                                //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                //Лист ДПРНМ 1,2х600х1500 Л63 ГОСТ 931-90
                                        Res_List:=Pos('ДПРНМ',Mat);
                                        if Res_List<>0 Then
                                        mat:='1,2';
                                 //Лента на шестьдесят НЕРЖ 0,3
                                 Res_List:=Pos('НЕРЖ',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                Res_List:=Pos('Л',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// Лента
                                                Delete(Mat,1,Res_Tol);
                                                //
                                                Res_Tol:=Pos(' ',Mat);//ОЦ
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('х',Mat);//x
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos('*',Mat);//*
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);
                                        end
                                        Else
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);
                                                Delete(Mat,1,Res_Tol);
                                        end;
                                        Res_List:=Pos('ОЦ',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);

                                //+++++++++++++++++++++++++++++++++++++++++++++
                                Tol:=StrToFloat(Mat);
                                Srav:=CompareValue(Tol,0.55);
                                if (Srav=0) OR (Srav<0) Then
                                Begin
                                        Tol_S:='0.55';
                                        Res_Tol55:=0.55;
                                End;
                                Srav:=CompareValue(Tol,0.7);
                                if ((Srav>0) OR (Srav=0)) AND (Tol<=1) Then
                                Begin
                                        Tol_S:='0.7';
                                        Res_Tol55:=7;
                                End;

                                Srav:=CompareValue(Tol,1.2);
                                //Srav1:= CompareValue(Tol,0.7);
                                if ((Srav=0) OR (Srav>0)) AND (Tol<=2) Then
                                Begin
                                        Tol_S:='1.2';
                                        Res_Tol55:=2;
                                End;
                                DlinI:=StrToFloat(Shir);
                                ShirI:=StrToFloat(Dlina);

                                if DlinI<=500 Then
                                        DlinI:=500;
                                //---------------
                                if (DlinI>500) AND (DlinI<=1000) Then
                                        DlinI:=1000;
                                //---------------
                                if (DlinI>1000) AND (DlinI<=1500) Then
                                        DlinI:=1500;
                                //---------------
                                if (DlinI>1500) AND (DlinI<=2000) Then
                                        DlinI:=2000;
                                //---------------
                                if (DlinI>2000) AND (DlinI<=2500) Then
                                        DlinI:=2500;
                                //+====================================
                                if Res_Tol55=0.55 Then
                                Begin
                                if ShirI<=50 Then
                                        ShirI:=50;
                                //---------------
                                if (ShirI>50) AND (ShirI<=100) Then
                                        ShirI:=100;
                                //---------------
                                if (ShirI>100) AND (ShirI<=200) Then
                                        ShirI:=200;
                                //---------------
                                if (ShirI>200) AND (ShirI<=500) Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) AND (ShirI<=600) Then
                                        ShirI:=600;
                                //---------------
                                if (ShirI>600) AND (ShirI<=1250) Then
                                        ShirI:=1250;
                                End;
                                //+====================================
                                if Res_Tol55=7 Then
                                Begin
                                if ShirI<=200 Then
                                        ShirI:=200;
                                //---------------
                                if (ShirI>200) AND (ShirI<=500) Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) Then
                                        ShirI:=1250;

                                End;
                                //+====================================
                                if Res_Tol55=2 Then
                                Begin
                                if ShirI<=200 Then
                                        ShirI:=200;
                                //---------------
                                if (ShirI>200) AND (ShirI<=500) Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) Then
                                        ShirI:=1250;

                                End;

                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'1'+ #39 +') And ([Толщина]='+Tol_S+') And ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([Высота]=' + #39 +FloatToStr(ShirI)+ #39 +')', ['[Резка Гибка]']) then
                                exit;
                                NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                if NC_S='' Then
                                        NC_S:='0';
                                //SD4.Cells[13,h]:=NC_S;
                                NC:=StrToFloat(NC_S)*Kol_Oboz*Kol_Zap;
//+++++++++++++++++++++++++++++++++++++++++++++

                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'11'+ #39 +') AND ([Высота]=' + #39 +IntToStr(Ugol)+ #39 +') ', ['[Резка Гибка]']) then
                                exit;
                                NC_S2:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                if NC_S2='' Then
                                        NC_S2:='0';

                                Razmet_KPU:=Form1.ADOQuery1.FieldByName('Норма').AsFloat;
                                SD4.Cells[16,h]:=FloatToStr(StrToFloat(NC_S)+Razmet_KPU);
                                NC:=NC+(Razmet_KPU*Kol_Oboz*Kol_Zap);
                                //end;
                                SD4.Cells[13,h]:=FloatToStr(C);
                                SD4.Cells[14,h]:=VElem;
                                SD4.Cells[15,h]:=FloatToStr(NC);
                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                'UPDATE %s SET [Н\ч Ножницы]=' + #39
                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                ') AND (IDКО=' + #39 + IDKO + #39 +
                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['СпецифВозд']) then
                                Exit;
                                //===========================================
                                Nog_List.Values['Typ']:=Tip;
                                Nog_List.Values['Elem']:=Elem;
                                Nog_List.Values['Kol']:=FloatToStr(Kol_Oboz*Kol_Zap);
                                Nog_List.Values['Oboz']:=Oboz;
                                Nog_List.Values['h']:=IntToStr(h);
                                Nog_List.Values['EI']:=EI;
                                Nog_List.Values['Shir']:=Shir;
                                Nog_List.Values['Dlina']:=Dlina;
                                Nog_List.Values['ShirRaz']:=ShirRaz;
                                Nog_List.Values['DlinRaz']:=DlinRaz;
                                Nog_List.Values['Kol_Gib']:=Kol_Gib;
                                Nog_List.Values['Mat']:=Mat;
                                Nog_List.Values['C']:=FloatToStr(C);
                                Nog_List.Values['VElem']:=VElem;
                                Nog_List.Values['NC']:=FloatToStr(NC);
                                //========================================
                                Inc(h);
                                SD4.RowCount:= SD4.RowCount+1;
                                //Break;
                       // End;
                     end;
                     //end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if  (Gib=True) AND (Res_Detal=0) Then

                     Begin
                                Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                                Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;
                                SD5.Cells[2,n]:=Tip;
                                SD5.Cells[3,n]:=Elem;
                                SD5.Cells[4,n]:=FloatToStr(Kol_Ed*Kol_Zap);
                                SD5.Cells[6,n]:=Oboz;
                                SD5.Cells[0,n]:=IntToStr(n);
                                SD5.Cells[1,n]:=EI;
                                SD5.Cells[7,n]:=Dlina;
                                SD5.Cells[8,n]:=Shir;
                                SD5.Cells[9,n]:=DlinRaz;
                                SD5.Cells[10,n]:=ShirRaz;
                                SD5.Cells[11,n]:=IntToStr(KolGib);
                                SD5.Cells[12,n]:=Mat;
                                SD5.Cells[13,n]:=VElem;
                                if  SG1.Cells[14, i]='' then
                                        Kol_Lop:=0
                                Else
                                        Kol_Lop:= StrToInt(SG1.Cells[14, i]);
                                Res_Lop:=AnsiCompareStr('Корпус лопатки',Elem);

                                if Res_Lop=0 Then
                                Begin
                                        SD5.Cells[14,n]:=FloatToStr(0.034*Kol_Ed*Kol_Zap);
                                        SD5.Cells[19,n]:=FloatToStr(0.034);
                                End
                                Else
                                Begin
                                        ShirI:=StrToFloat(ShirRaz);
                                        DlinI:=StrToFloat(DlinRaz);
                                        //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                        Res_List:=Pos('ДПРНМ',Mat);
                                        if Res_List<>0 Then
                                        mat:='1,2';
                                        //Лента Оц 1,5х264
                                        //Лист ОЦ 08пс 1,5х1250х2500
                                        //ОЦ 1,5
                                        //Лента на шестьдесят НЕРЖ 0,3

                                 Res_List:=Pos('НЕРЖ',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                        Res_List:=Pos('Л',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// Лента
                                                Delete(Mat,1,Res_Tol);
                                                //
                                                Res_Tol:=Pos(' ',Mat);//ОЦ
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('х',Mat);//x
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos('*',Mat);//*
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);
                                        end
                                        Else
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);
                                                Delete(Mat,1,Res_Tol);
                                        end;
                                        Res_List:=Pos('ОЦ',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);



                                        if Mat='' Then
                                        Mat:='0';
                                        Tol:=StrToFloat(Mat);
                                        if Tol<=1 Then
                                        Tol_S:='0.55';
                                        if (Tol>1) AND (Tol<=2) Then
                                        Tol_S:='2';
                                        if Tol>2 Then
                                        Tol_S:='3';
                                        if DlinI<=2000 Then
                                        DlinI:=2000
                                        Else
                                        DlinI:=3000;
                                        if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'3'+ #39 +') And ([Толщина]=' + #39 +(Tol_s)+ #39 +
                                ') And ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([Высота]=' + #39 +IntToStr(KolGib)+ #39 +')', ['[Резка Гибка]']) then
                                exit;
                                NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                If NC_S='' Then
                                NC:=0 Else
                                NC:=StrToFloat(NC_S);
                                SD5.Cells[14,n]:=FloatToStr(NC*Kol_Ed*Kol_Zap);
                                SD5.Cells[19,n]:=NC_S;
                                end;
                                //======================
                                Res_Resh:=Pos('Решетка жалюзийная',Elem);
                                if Res_Resh<>0 Then
                                Begin
                                       if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                        'Select * from %s Where   ([№]=' + #39 +'9'+ #39 +') And ([Толщина]=' + #39 +IntToStr(KolGib)+ #39 +
                                        ')', ['[Резка Гибка]']) then
                                        exit;
                                        NC_S:= (Form1.ADOQuery1.FieldByName('Высота').AsString);
                                        If NC_S='' Then
                                                NC:=0 Else
                                                NC:=StrToFloat(NC_S);
                                        //++++++++++++
                                        NC_S2:= (Form1.ADOQuery1.FieldByName('Длина').AsString);
                                        If NC_S2='' Then
                                                NC2:=0 Else
                                                NC2:=StrToFloat(NC_S2);
                                        SD5.Cells[19,n]:=FloatToStr(StrToFloat(NC_S)+StrToFloat(NC_S2));
                                        SD5.Cells[14,n]:=FloatToStr((NC*Kol_Zap)+(NC2/Kol_Zap));
                                end;
                                Res_Lop:=AnsiCompareStr('Стенка лопатки',Elem);

                                if Res_Lop=0 Then
                                Begin
                                     Kol_Lop:=ADOQuery2.FieldByName('Количество').AsInteger;;
                                end;
                                Lenta:=Pos('Лента уплотнительная',Elem);
                                if Lenta <>0 Then
                                Begin
                                      DlinD:=StrToInt(Dlina)-1.5;
                                      B:=StrToInt(Dlina);
                                      if B<335 Then
                                        DlinI:=170;
                                      if (B>=335) AND (B<500) Then
                                      Begin
                                        DlinI:=335;
                                      end;

                                      if (B>=500) AND (B<665) Then
                                      Begin
                                        DlinI:=500;
                                      end;

                                      if (B>=665) AND (B<830) Then
                                      Begin
                                        DlinI:=665;
                                      end;

                                      if (B>=830) AND (B<995) Then
                                      Begin
                                        DlinI:=830;
                                      end;

                                      if (B>=995) AND (B<1160) Then
                                      Begin
                                        DlinI:=995;
                                      end;

                                      if (B>=1160) AND (B<1325) Then
                                      Begin
                                        DlinI:=1160;
                                      end;

                                      if (B>=1325) AND (B<1490) Then
                                      Begin
                                        DlinI:=1325;
                                      end;

                                      if (B>=1490) AND (B<1655) Then
                                      Begin
                                        DlinI:=1490;
                                      end;

                                      if (B>=1655) AND (B<1820) Then
                                      Begin
                                        DlinI:=1655;
                                      end;

                                      if (B>=1820) AND (B<1985) Then
                                      Begin
                                        DlinI:=1820;
                                      end;

                                      if (B>=1985) AND (B<2150) Then
                                      Begin
                                        DlinI:=1985;
                                      end;

                                      if (B>=2150) AND (B<2315) Then
                                      Begin
                                        DlinI:=2150;
                                      end;

                                      if (B>=2315) AND (B<2440) Then
                                      Begin
                                        DlinI:=2315;
                                      end; 

                                      A_Lenta:=(StrToInt(Dlina)-1.5-(Kol_Lop-1)*150)/2;
                                      SD5.Cells[15,n]:=FloatToStr(DlinD);
                                      SD5.Cells[16,n]:= FloatToStr(A_Lenta);
                                      SD5.Cells[17,n]:=IntToStr(Kol_Lop);
                                      SD5.Cells[18,n]:='150';
                                      T:=0.008*Kol_Lop*Kol_Ed*Kol_Zap+0.008*Kol_Ed*Kol_Zap; //Пробивка отверстий в одной ленте
                                      SD5.Cells[20,n]:= FloatToStr(T);
                                      if SD5.Cells[20,n]='' Then
                                        NCB:=(StrToFloat(SD5.Cells[14,n]))
                                      Else
                                      Begin
                                        NCB:=(StrToFloat(SD5.Cells[14,n])+StrToFloat(SD5.Cells[20,n]));
                                        SD5.Cells[14,n]:=FloatToStr(NCB);
                                      End;

                                end;
                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                'UPDATE %s SET [Н\ч Гибка]=' + #39
                                + Form1.ConvertFloat(FloatToStr(NCB)) + #39 +
                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                ') AND (IDКО=' + #39 + IDKO + #39 +
                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['СпецифВозд']) then
                                Exit;
                                Inc(n);
                                SD5.RowCount:= SD5.RowCount+1;
                                //Break;
                     end;
                     //end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Pila=True) AND (Res_Detal=0) Then
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;

                                SD6.Cells[2,w]:=Tip;
                                SD6.Cells[3,w]:=Elem;
                                SD6.Cells[4,w]:=FloatToStr(Kol_Ed*Kol_Zap);
                               // SD7.Cells[5,w]:=FloatToStr(Kol_Oboz);
                                SD6.Cells[6,w]:=Oboz;
                                SD6.Cells[0,w]:=IntToStr(w);
                                SD6.Cells[1,w]:=EI;
                                SD6.Cells[7,w]:=Dlina;
                                SD6.Cells[8,w]:=Shir;
                                SD6.Cells[9,w]:=DlinRaz;
                                SD6.Cells[10,w]:=ShirRaz;
                                //SD6.Cells[11,w]:=Kol_Gib;
                                SD6.Cells[12,w]:=Mat;
                                SD6.Cells[13,w]:=VElem;
                                //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                Res_List:=Pos('ДПРНМ',Mat);
                                        if Res_List<>0 Then
                                        mat:='1,2';
                                        //Лента Оц 1,5х264
                                        //Лист ОЦ 08пс 1,5х1250х2500
                                        //ОЦ 1,5
                                        //Лента на шестьдесят НЕРЖ 0,3
                                 Res_List:=Pos('НЕРЖ',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                        Res_List:=Pos('Л',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// Лента
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//ОЦ
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('х',Mat);//x
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos('*',Mat);//*
                                                Delete(Mat,Res_Tol,50);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);
                                        end
                                        Else
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);
                                                Delete(Mat,1,Res_Tol);
                                        end;
                                        Res_List:=Pos('ОЦ',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);



                                        if Mat='' Then
                                        Mat:='0';
                                        //++++++++++++++++++++++++++++++++==
                                          DlinI:=StrToFloat(Dlina);
                                          if DlinI<=250 Then
                                          DlinI:=250;
                                          if (DlinI>250) AND (DlinI<=500) Then
                                          DlinI:=500;
                                          if (DlinI>500) AND (DlinI<=750) Then
                                          DlinI:=750;
                                          if (DlinI>750) AND (DlinI<=1000) Then
                                          DlinI:=1000;
                                          if (DlinI>1000) AND (DlinI<=1250) Then
                                          DlinI:=1250;
                                          if (DlinI>1250) AND (DlinI<=2500) Then
                                          DlinI:=2495;
                                          if (DlinI>2500)  Then
                                          DlinI:=2500;
                                          if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                          'Select * from %s Where   ([№]=' + #39 +'13'+ #39 +') And  ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[Резка Гибка]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S);            // 
                                          SD6.Cells[14,w]:=FloatToStr(NC*Kol_Ed*Kol_Zap);
                                       // End;
                                 if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                'UPDATE %s SET [Н\ч Пила]=' + #39
                                + Form1.ConvertFloat(SD6.Cells[14,w]) + #39 +
                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                ') AND (IDКО=' + #39 + IDKO + #39 +
                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['СпецифВозд']) then
                                Exit;
                                Inc(w);
                                SD6.RowCount:= SD6.RowCount+1;
                                //Break;
                     //end;
                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Pila=True) AND (Res_Detal=0) Then
                      Begin
                      Kol_Ed:= ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= ADOQuery2.FieldByName('Количество').AsFloat;

                                SD7.Cells[2,www]:=Tip;
                                SD7.Cells[3,www]:=Elem;
                                SD7.Cells[4,www]:=FloatToStr(Kol_Ed*Kol_Zap);
                               // SD7.Cells[5,w]:=FloatToStr(Kol_Oboz);
                                SD7.Cells[6,www]:=Oboz;
                                SD7.Cells[0,www]:=IntToStr(www);
                                SD7.Cells[1,www]:=EI;
                                SD7.Cells[7,www]:=Dlina;
                                SD7.Cells[8,www]:=Shir;
                                SD7.Cells[9,www]:=DlinRaz;
                                SD7.Cells[10,www]:=ShirRaz;
                                //SD7.Cells[11,w]:=Kol_Gib;
                                SD7.Cells[12,www]:=Mat;
                                SD7.Cells[13,www]:=VElem;
                                Inc(www);
                                SD7.RowCount:= SD7.RowCount+1;
                     end;
                       RecCol:=ADOQuery2.RecordCount;
                       RecNom:=ADOQuery2.RecNo ;
                       ADOQuery2.Next;
                   end;
                End;
        End;

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Dir_main := ExtractFileDir(ParamStr(0));
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);
        Dir :=Form1.Put_KTO+'\CKlapana\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\' + God + '\' + Mes+ '\';
        CreateDir(Dir);
        Puty:=(Dir +'\№ '+Label15.Caption+'.xlsx');
        XL := CreateOleObject('Excel.Application');
        CopyFile(PWideChar(Form1.Put_KTO+'\CKlapana\2013\Воздушные.xlsx'), PWideChar(Dir +'\№ '+Label15.Caption+'.xlsx'), False);
        XL.Workbooks.Open(Dir + '\№ '+Label15.Caption+'.xlsx');
        XL.Application.EnableEvents := false;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG1.RowCount-1 Do //Вальцы
        Begin

            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 2] := RSG1.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 3] := RSG1.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 4] := RSG1.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 5] := RSG1.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 6] := RSG1.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 7] := RSG1.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 8] := RSG1.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 9] := RSG1.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 10] := RSG1.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 11] := RSG1.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[10].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[10].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[10].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG1.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[10].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[10].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[10].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[10].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[10].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[10].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[10].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[10].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 3] :='Вальцовщик';
         XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[10].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[10].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[10].Rows[q+7].Font.Size := 12;
                //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG2.RowCount-1 Do //Кромкогиб
        Begin

            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 2] := RSG2.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 3] := RSG2.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 4] := RSG2.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 5] := RSG2.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(5), 6] := RSG2.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(5), 7] := RSG2.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 8] := RSG2.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 9] := RSG2.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 10] := RSG2.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 11] := RSG2.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[11].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[11].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[11].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG2.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[11].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[11].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[11].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[11].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[11].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[11].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[11].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[11].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 3] :='Кромкогибец';
         XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[11].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[11].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[11].Rows[q+7].Font.Size := 12;
                         //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG4.RowCount-1 Do //Токарный
        Begin

            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 2] := RSG4.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 3] := RSG4.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 4] := RSG4.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 5] := RSG4.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 6] := RSG4.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[10].Cells[(i)+(5), 7] := RSG4.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 8] := RSG4.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 9] := RSG4.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 10] := RSG4.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[12].Cells[(i)+(5), 11] := RSG4.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[12].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[12].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[12].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG4.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[12].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[12].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[12].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[12].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[12].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[12].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[12].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[12].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 3] :='Токарь';
         XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[12].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[12].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[12].Rows[q+7].Font.Size := 12;
                         //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To RSG3.RowCount-1 Do //Ротор
        Begin

            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 2] := RSG3.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 3] := RSG3.Cells[1,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 4] := RSG3.Cells[2,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 5] := RSG3.Cells[4,i+1]; //Dlin
            //XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 6] := RSG3.Cells[10,i+1]; //Shir
            //XL.ActiveWorkBook.WorkSheets[11].Cells[(i)+(5), 7] := RSG3.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 8] := RSG3.Cells[3,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 9] := RSG3.Cells[8,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 10] := RSG3.Cells[7,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[13].Cells[(i)+(5), 11] := RSG3.Cells[5,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[13].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[13].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[13].Cells[4, 1] :=Label15.Caption;
            NC_S:= RSG3.Cells[8,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[13].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[13].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[13].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[13].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[13].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[13].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[13].Cells[q+4, 2] := SG1.Cells[4,i+1];
               XL.ActiveWorkBook.WorkSheets[13].Cells[q+4, 6] := SG1.Cells[1,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 3] :='Роторист';
         XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[13].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[13].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[13].Rows[q+7].Font.Size := 12;

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  Kanban
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(I+4)+':E'+IntToStr(I+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[1].Rows[i+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[1].Rows[i+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[1].Cells[I+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        Q:=Q+2;   //Kanban
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+2].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Select;

        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+3, 10] := 'Мат';
        XL.ActiveWorkBook.WorkSheets[1].Cells[q+2, 4] := 'Kanban';
        For I:=0 To SD1.RowCount-1 Do
        Begin
            Razm:=SD1.Cells[7,i+1]+'*'+SD1.Cells[7,i+1];
            if SD1.Cells[0,i+1]='' Then
            Continue;


            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := SD1.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := SD1.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := SD1.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;//Razmer
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := SD1.Cells[4,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := SD1.Cells[5,i+1];

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := SD1.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := SD1.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := SD1.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+5), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[1].Cells[2, 6] :=Label15.Caption;
            Inc(Q);
        end;
        Q:=Q+2;

        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Cells[(Q)+(4), 4] :='Заготовка';
        For I:=0 To SD2.RowCount-1 Do
        Begin
            Razm:=SD2.Cells[7,i+1]+'*'+SD2.Cells[8,i+1];
            if SD2.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := SD2.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := SD2.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := SD2.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := SD2.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := SD2.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := SD2.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := SD2.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := SD2.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+3), 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Rows[q+4].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[1].Cells[(Q)+(4), 4] :='Сбор. Единицы';
        For I:=0 To SD.RowCount-1 Do
        Begin
            Razm:=SD.Cells[7,i+1]+'*'+SD.Cells[8,i+1];
            if SD.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 2] := SD.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 3] := SD.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 4] := SD.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 6] := SD.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 7] := SD.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 8] := SD.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 9] := SD.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Cells[(q)+(5), 10] := SD.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[1].Range['B'+IntToStr(Q+5), 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
         XL.ActiveWorkBook.WorkSheets[1].Cells[2, 6] :=Label15.Caption;
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+6, 4] :='Комплектовщик (Заготовка)';
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+8, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+8, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[1].Cells[Q+8, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+8].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+8].Font.Size := 12;
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+6].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[1].Rows[q+6].Font.Size := 12;
        //++++++++++++++++++++++++++++++++++++++++++++++++  Склад
        Q:=1;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[2].Range['B'+IntToStr(I+3)+':D'+IntToStr(I+3)].Merge;
               XL.ActiveWorkBook.WorkSheets[2].Rows[i+3].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[2].Rows[i+3].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[2].Cells[I+3, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[2].Cells[I+3, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[2].Cells[I+3, 5] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[2].Rows[q+2].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[2].Select;

        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+3, 10] := 'Мат';
        XL.ActiveWorkBook.WorkSheets[2].Cells[q+2, 4] := 'Склад';
        For I:=0 To SD9.RowCount-1 Do
        Begin
            Razm:=SD9.Cells[7,i+1]+'*'+SD9.Cells[8,i+1];
            if SD9.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 2] := SD9.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 3] := SD9.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 4] := SD9.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 6] := SD9.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 7] := SD9.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 8] := SD9.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 9] := SD9.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Cells[(q)+(5), 10] := SD9.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[2].Range['B'+IntToStr(q+3), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
         XL.ActiveWorkBook.WorkSheets[2].Cells[2, 6] :=Label15.Caption;
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 3] :='Кладовщик';
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[2].Cells[Q+8, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[2].Rows[q+8].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[2].Rows[q+8].Font.Size := 12;
        //++++++++++++++++++++++++++++++++++++++++++++++++  Склад Стенки
        Q:=1;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[3].Range['B'+IntToStr(I+3)+':D'+IntToStr(I+3)].Merge;
               XL.ActiveWorkBook.WorkSheets[3].Rows[i+3].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[3].Rows[i+3].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[3].Cells[I+3, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[3].Cells[I+3, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[3].Cells[I+3, 5] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        Q:=Q+2;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+2].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+3].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+3].Font.Size := 14;
        XL.ActiveWorkBook.WorkSheets[3].Rows[q+2].Font.Size := 14;
        //XL.ActiveWorkBook.WorkSheets[3].Select;

        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 2] := '№';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 3] := 'Тип клапана';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 4] := 'Изделие';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 5] := 'Размер';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 6] := 'Кол на 1 клапан';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 7] := 'Общее кол-во';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 8] := 'ЕдИзм';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 9] := 'Обознач';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+3, 10] := 'Мат';
        XL.ActiveWorkBook.WorkSheets[3].Cells[q+2, 4] := 'Склад Стенки';
        NC:=0;
        NC_S:='';
        For I:=0 To Stenki.RowCount-1 Do
        Begin
            Razm:=Stenki.Cells[7,i+1];
            if Stenki.Cells[0,i+1]='' Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 2] := Stenki.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 3] := Stenki.Cells[2,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 4] := Stenki.Cells[3,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 5] := Razm;
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 6] := Stenki.Cells[4,i+1];//KolEd
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 7] := Stenki.Cells[5,i+1];//KolOb

            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 8] := Stenki.Cells[1,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 9] := Stenki.Cells[6,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Cells[(q)+(5), 10] := Stenki.Cells[12,i+1];
            XL.ActiveWorkBook.WorkSheets[3].Range['B'+IntToStr(q+3), 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;
            Inc(Q);
        end;
         XL.ActiveWorkBook.WorkSheets[3].Cells[2, 6] :=Label15.Caption;
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 3] :='Кладовщик';
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[3].Cells[Q+8, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[3].Rows[q+8].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[3].Rows[q+8].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        H:=1;
        NC:=0;
        NC_S:='';
        For I:=0 To SD4.RowCount-1 Do //Ножницы
        Begin
             Elem:= SD4.Cells[3,h];
             Oboz:= SD4.Cells[6,h];
             O:=I+2;
             for Y:=0 To SD4.RowCount-1 Do
             Begin
                if Y>SD4.RowCount-1 Then
                Break;
                If (SD4.Cells[3,o]='') OR (SD4.Cells[3,h]='') Then
                        Continue;
                Res_Elem:=AnsiCompareStr(Elem,SD4.Cells[3,o]);
                Res_Oboz:=AnsiCompareStr(Oboz,SD4.Cells[6,o]);
                if (Res_Elem=0) AND (Res_Oboz=0) Then
                Begin
                        If (SD4.Cells[3,o]='') OR (SD4.Cells[3,h]='') Then
                        Continue;
                        Kol_Oboz:=StrToInt(SD4.Cells[4,o]);
                        NC:=StrToFloat(SD4.Cells[15,o]);
                        Kol_Gib := SD4.Cells[11,o];
                        SD4.Cells[4,h]:=FloatToStr(StrToInt(SD4.Cells[4,h])+Kol_Oboz);
                        SD4.Cells[15,h]:=FloatToStr(StrToFloat(SD4.Cells[15,h])+NC);
                        SD4.Cells[3,o]:='';
                        Form1.DeleteARow(SD4,o);
                        O:=O-1;
                end;
             Inc(O);
             end;
             Inc(H);
        End;
        NC:=0;
        NC_S:='';
        For I:=0 To SD4.RowCount-1 Do //Ножницы
        Begin
            if SD4.Cells[2,i+1]='' Then
            break;
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 2] := SD4.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 3] := SD4.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 4] := SD4.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 5] := SD4.Cells[9,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 6] := SD4.Cells[10,i+1]; //Shir
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 7] := SD4.Cells[13,i+1]; //Virub
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 8] := SD4.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 9] := SD4.Cells[15,i+1];//N\C
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 10] := SD4.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[4].Cells[(i)+(5), 11] := SD4.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[4].Range['B4', 'K' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[4].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[4].Cells[4, 1] :=Label15.Caption;
            NC_S:= SD4.Cells[15,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            q:=I+5;
        end;
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        XL.ActiveWorkBook.WorkSheets[4].Cells[q+1, 8]:='Итого';
        XL.ActiveWorkBook.WorkSheets[4].Cells[q+1, 9]:=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[4].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[4].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[4].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[4].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[4].Cells[q+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[4].Cells[q+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 3] :='Резчик';
         XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 9] :='Дата';
         XL.ActiveWorkBook.WorkSheets[4].Cells[Q+7, 11] :='____________________';
         XL.ActiveWorkBook.WorkSheets[4].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[4].Rows[q+7].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        For I:=0 To SD3.RowCount-1 Do  //Trumph
        Begin

            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 5] := SD3.Cells[0,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 6] := SD3.Cells[2,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 7] := SD3.Cells[3,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 9] := SD3.Cells[7,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 10] := SD3.Cells[8,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 11] := SD3.Cells[14,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 8] := SD3.Cells[4,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 12] := SD3.Cells[12,i];
            XL.ActiveWorkBook.WorkSheets[5].Cells[(i)+(13), 13] := SD3.Cells[6,i];
            XL.ActiveWorkBook.WorkSheets[5].Range['E21', 'M' +
                IntToStr(I+13)].Borders.LineStyle := 1;
            //XL.ActiveWorkBook.WorkSheets[5].Cells[18, 9] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[5].Cells[13, 2] :='СОПРОВОДИТЕЛЬНЫЙ ЛИСТ №  '+Label15.Caption;
            Q:=I+13;
        end;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[5].Range['F'+IntToStr(Q+4)+':G'+IntToStr(Q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[5].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[5].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[5].Cells[q+4, 5] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[5].Cells[q+4, 6] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[5].Cells[q+4, 8] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 6] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 7] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[5].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[5].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[5].Rows[q+7].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        //+++++++++++++++++++++++++++++++++++++++
        H:=1;
        NC:=0;
        NC_S:='';
        For I:=0 To SD5.RowCount-1 Do //Гибка
        Begin
             Elem:= SD5.Cells[3,h];
             Oboz:= SD5.Cells[6,h];
             NC:=0;
             O:=I+2;
             for Y:=0 To SD5.RowCount-1 Do
             Begin
                if Y>SD5.RowCount-1 Then
                Break;
                If (SD5.Cells[3,o]='') OR (SD5.Cells[3,h]='') Then
                        Continue;
                Res_Elem:=AnsiCompareStr(Elem,SD5.Cells[3,o]);
                Res_Oboz:=AnsiCompareStr(Oboz,SD5.Cells[6,o]);
                if (Res_Elem=0) AND (Res_Oboz=0) Then
                Begin
                        If (SD5.Cells[3,o]='') OR (SD5.Cells[3,h]='') Then
                        Continue;
                        Kol_Oboz:=StrToInt(SD5.Cells[4,o]);
                        NC:=StrToFloat(SD5.Cells[14,o]);
                        Kol_Gib := SD5.Cells[11,o];
                        SD5.Cells[4,h]:=FloatToStr(StrToInt(SD5.Cells[4,h])+Kol_Oboz);
                        SD5.Cells[14,h]:=FloatToStr(StrToFloat(SD5.Cells[14,h])+NC);
                        SD5.Cells[11,h]:=FloatToStr(StrToFloat(SD5.Cells[11,h])+StrToInt(Kol_Gib));
                        SD5.Cells[3,o]:='';
                        Form1.DeleteARow(SD5,o);
                        O:=O-1;
                end;
             Inc(O);
             end;
             Inc(H);
        End;
        NC:=0;
        NC_S:='';
        For I:=0 To SD5.RowCount-1 Do //Гибка
        Begin
            If (SD5.Cells[3,i+1]='')Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 2] := SD5.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 3] := SD5.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 4] := SD5.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 5] := SD5.Cells[4,i+1]; //  Kol
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 6] := SD5.Cells[11,i+1]; // KolGib
            Lenta:=Pos('Лента уплотнительная',SD5.Cells[3,i+1]);
            if Lenta<>0 Then
                XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 7] := FloatToStr(StrToFloat(SD5.Cells[14,i+1])+0.05) // N\C
            Else
                XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 7] := SD5.Cells[14,i+1];
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 8] := SD5.Cells[9,i+1]; //A
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 9] := SD5.Cells[10,i+1]; //B

            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 10] := SD5.Cells[15,i+1]; //B
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 11] := SD5.Cells[17,i+1]; //B
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 12] := SD5.Cells[18,i+1]; //B
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 13] := SD5.Cells[16,i+1]; //B

            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 14] := SD5.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[6].Cells[(i)+(5), 15] := SD5.Cells[6,i+1]; //Chert
            XL.ActiveWorkBook.WorkSheets[6].Range['B4', 'O' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[6].Cells[2, 5] :=Label15.Caption;
            NC_S:= SD5.Cells[14,i+1];
            if NC_S='' Then
            NC_S:='0';
            NC:=NC+StrToFloat(NC_S);
            Q:=I+5;
        end;
        XL.ActiveWorkBook.WorkSheets[6].Cells[(Q)+1, 6] :='Итого';
        NC_S:= FloatToStr(NC);
        XL.ActiveWorkBook.WorkSheets[6].Cells[(q)+1, 7] :=NC_S;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[6].Range['C'+IntToStr(q+2)+':D'+IntToStr(q+2)].Merge;
               XL.ActiveWorkBook.WorkSheets[6].Rows[q+2].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[6].Rows[q+2].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[6].Cells[q+2, 2] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[6].Cells[q+2, 3] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[6].Cells[q+2, 6] := SG1.Cells[2,i+1];

               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[6].Cells[Q+5, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[6].Rows[q+5].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[6].Rows[q+5].Font.Size := 12;
        //+++++++++++++++++++++++++++++++++++++++
        //+++++++++++++++++++++++++++++++++++++++
        H:=1;
        O:=1;
        For I:=0 To SD6.RowCount-1 Do //Pila
        Begin
             Elem:= SD6.Cells[3,h];
             Oboz:= SD6.Cells[6,h];
             O:=I+2;
             for Y:=0 To SD6.RowCount-1 Do
             Begin
                if Y>SD6.RowCount-1 Then
                Break;
                If (SD6.Cells[3,o]='') OR (SD6.Cells[3,h]='') Then
                        Continue;
                Res_Elem:=AnsiCompareStr(Elem,SD6.Cells[3,o]);
                Res_Oboz:=AnsiCompareStr(Oboz,SD6.Cells[6,o]);
                if (Res_Elem=0) AND (Res_Oboz=0) Then
                Begin
                        If (SD6.Cells[3,o]='') OR (SD6.Cells[3,h]='') Then
                        Continue;
                        Kol_Oboz:=StrToInt(SD6.Cells[4,o]);
                        NC:=StrToFloat(SD6.Cells[14,o]);
                        //Kol_Gib := SD6.Cells[11,o];
                        SD6.Cells[4,h]:=FloatToStr(StrToInt(SD6.Cells[4,h])+Kol_Oboz);
                        SD6.Cells[14,h]:=FloatToStr(StrToFloat(SD6.Cells[14,h])+NC);
                        //SD6.Cells[11,h]:=FloatToStr(StrToFloat(SD6.Cells[11,h])+StrToInt(Kol_Gib));
                        SD6.Cells[3,o]:='';
                        Form1.DeleteARow(SD6,o);
                        O:=O-1;
                end;
             Inc(O);
             end;
             Inc(H);
        End;
         NC:=0;
        For I:=0 To SD6.RowCount-1 Do //Pila
        Begin
            If (SD6.Cells[3,i+1]='')Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 2] := SD6.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 3] := SD6.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 4] := SD6.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 5] := SD6.Cells[9,i+1]; //Dlin


            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 7] := SD6.Cells[4,i+1]; //KolEd
            Res:=Pos( 'Профиль ФЭЗ.П.0043',SD6.Cells[12,i+1]);
            if Res<>0 Then
            Begin
                 XL.ActiveWorkBook.WorkSheets[7].Range['H'+IntToStr(I+5), 'H' +IntToStr(I+5)].Font.Italic := True;
                 XL.ActiveWorkBook.WorkSheets[7].Range['H'+IntToStr(I+5), 'H' +IntToStr(I+5)].Font.Bold := True;
            End;
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 8] := SD6.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 9] := SD6.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[7].Range['B4', 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[7].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[7].Cells[4, 1] :=Label15.Caption;
            if SD6.Cells[14,i+1]<>'' Then
            Begin
                NC2:=StrToFloat( SD6.Cells[14,i+1]); //*StrToInt(SD6.Cells[4,i+1]
                NC_S2:=FloatToStr(NC2);
                NC_S2:=Float(NC_S2);
                PosTrud := Pos(',', NC_S2);
                if PosTrud <> 0 then
                begin
                        Delete(NC_S2, PosTrud, 1);
                        Insert('.', NC_S2, PosTrud);
                end;
                XL.ActiveWorkBook.WorkSheets[7].Cells[(i)+(5), 6] := NC_S2; //
                NC_S:= FloatToStr(NC2);
                if NC_S='' Then
                NC_S:='0';
                NC:=NC+StrToFloat(NC_S);
            End;
            Q:=I+5;
        end;
        XL.ActiveWorkBook.WorkSheets[7].Cells[(Q)+1, 5] :='Итого';
        NC_S:= FloatToStr(NC);
        NC_S:=Float(NC_S);
        PosTrud := Pos(',', NC_S); //Трудоемкость FLOAT
        if PosTrud <> 0 then
        begin
                Delete(NC_S, PosTrud, 1);
                Insert('.', NC_S, PosTrud);
        end;
       // XL.ActiveWorkBook.WorkSheets[7].Range['F'+IntToStr(q+1)+':F'+IntToStr(q+1)].Select;
       // XL.ActiveWorkBook.WorkSheets[7].Selection.NumberFormat := '0.00';
        XL.ActiveWorkBook.WorkSheets[7].Cells[(q)+1, 6] :=NC_S;
                For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[7].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[7].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[7].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[7].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[7].Cells[q+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[7].Cells[q+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[7].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[7].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[7].Rows[q+7].Font.Size := 12;
                 //+++++++++++++++++++++++++++++++++++++++
                  NC:=0;
        For I:=0 To SD6.RowCount-1 Do //Abraziv
        Begin

            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 2] := SD6.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 3] := SD6.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 4] := SD6.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 5] := SD6.Cells[9,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 6] := SD6.Cells[14,i+1]; //Н\ч

            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 7] := SD6.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 8] := SD6.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[8].Cells[(i)+(5), 9] := SD6.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[8].Range['B4', 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;
            XL.ActiveWorkBook.WorkSheets[8].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[8].Cells[4, 1] :=Label15.Caption;
            Q:=I+5;
        end;
                For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[8].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[8].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[8].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[8].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[8].Cells[q+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[8].Cells[q+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[8].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[8].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[8].Rows[q+7].Font.Size := 12;
         //+++++++++++++++++++++++++++++++++++++++++++++++
         Q:=6;
        For I:=0 To SD3.RowCount-1 Do //Сопроводитр
        Begin
            //If (SD3.Cells[3,i+1]='')Then
            //Continue;
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 2] := SD3.Cells[0,i];
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 3] := SD3.Cells[2,i];//Тип
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 4] := SD3.Cells[3,i]; //Элемент
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 5] := SD3.Cells[7,i]; //Dlin
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 6] := SD3.Cells[8,i]; //Shir

            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 7] := SD3.Cells[4,i]; //KolEd
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 8] := SD3.Cells[12,i]; //Mat
            XL.ActiveWorkBook.WorkSheets[9].Cells[(i)+(7), 9] := SD3.Cells[6,i]; //Obozn
            XL.ActiveWorkBook.WorkSheets[9].Range['B4', 'J' +
                IntToStr(I+5)].Borders.LineStyle := 1;


            Inc(Q);
        end;
        XL.ActiveWorkBook.WorkSheets[9].Cells[5, 4] :='Trumpf';
            XL.ActiveWorkBook.WorkSheets[9].Cells[2, 5] :=Label15.Caption;
            XL.ActiveWorkBook.WorkSheets[9].Cells[4, 1] :=Label15.Caption;
        Q:=Q+4;
        //Nog
        XL.ActiveWorkBook.WorkSheets[9].Cells[q-2, 4] :='Ножницы';
        For I:=0 To SD4.RowCount-1 Do //Сопроводитр
        Begin
            If (SD4.Cells[3,i+1]='')Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 2] := SD4.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 3] := SD4.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[9].Cells[q, 4] := SD4.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 5] := SD4.Cells[7,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 6] := SD4.Cells[8,i+1]; //Shir

            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 7] := SD4.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 8] := SD4.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 9] := SD4.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[9].Range['B4', 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;

           Inc(Q);
        end;
        Q:=Q+4;
        //Nog
        XL.ActiveWorkBook.WorkSheets[9].Cells[q-2, 4] :='Гибка';
        For I:=0 To SD5.RowCount-1 Do //Сопроводитр
        Begin
            If (SD5.Cells[3,i+1]='')Then
            Continue;
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 2] := SD5.Cells[0,i+1];
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 3] := SD5.Cells[2,i+1];//Тип
            XL.ActiveWorkBook.WorkSheets[9].Cells[q, 4] := SD5.Cells[3,i+1]; //Элемент
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 5] := SD5.Cells[7,i+1]; //Dlin
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 6] := SD5.Cells[8,i+1]; //Shir

            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 7] := SD5.Cells[4,i+1]; //KolEd
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 8] := SD5.Cells[12,i+1]; //Mat
            XL.ActiveWorkBook.WorkSheets[9].Cells[(q), 9] := SD5.Cells[6,i+1]; //Obozn
            XL.ActiveWorkBook.WorkSheets[9].Range['B4', 'J' +
                IntToStr(q+5)].Borders.LineStyle := 1;

           Inc(Q);
        end;
        For I:=0 To SG1.RowCount-1 Do
        Begin
               if SG1.Cells[0,i+1]='' Then
               Break;
               XL.ActiveWorkBook.WorkSheets[9].Range['B'+IntToStr(q+4)+':E'+IntToStr(q+4)].Merge;
               XL.ActiveWorkBook.WorkSheets[9].Rows[q+4].Font.Bold := True;
               XL.ActiveWorkBook.WorkSheets[9].Rows[q+4].Font.Size := 12;
               XL.ActiveWorkBook.WorkSheets[9].Cells[q+4, 1] := SG1.Cells[0,i+1];
               XL.ActiveWorkBook.WorkSheets[9].Cells[q+4, 2] := SG1.Cells[5,i+1];
               XL.ActiveWorkBook.WorkSheets[9].Cells[q+4, 6] := SG1.Cells[2,i+1];
               Inc(q);
        end;
        XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 3] :='Оператор';
         XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 4] :='______________________________________';
         XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 8] :='Дата';
         XL.ActiveWorkBook.WorkSheets[9].Cells[Q+7, 9] :='____________________';
         XL.ActiveWorkBook.WorkSheets[9].Rows[q+7].Font.Bold := True;
         XL.ActiveWorkBook.WorkSheets[9].Rows[q+7].Font.Size := 12;
        XL.Visible := True;
        XL := UnAssigned;
end;
procedure TFNAklVoz.btn3Click(Sender: TObject);
Var S,Dir:String;
begin
        S := ExtractFileDir(ParamStr(0));
        //Puty:='\\192.168.7.1\Kto\CKlapana\Накладные(пожар)\2016\Декабрь\№ 6571.xlsx';
        Dir:=S +'\Воздушные\';
        CreateDir(Dir);
        CopyFile(PWideChar(Puty),PWideChar(Dir+'№ ' + Label15.Caption + '.xlsx'), False);
end;

end.




