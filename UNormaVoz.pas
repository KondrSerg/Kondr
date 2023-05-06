unit UNormaVoz;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Grids, ExtCtrls, ComObj,Math;

type
  //Второй
  TFNormaVoz = class(TForm)
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
    dtp1: TDateTimePicker;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure SG2DblClick(Sender: TObject);
    procedure Clear_StringGrid(SG:TStringGrid);
    function Float(var Str:String): String;
    procedure Button2Click(Sender: TObject);
    procedure Clear_StringGrid1(StringGrid: TStringGrid);
    procedure btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FNormaVoz: TFNormaVoz;

implementation

uses Main;

{$R *.dfm}
function TFNormaVoz.Float(var Str:String): String;
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
procedure TFNormaVoz.FormShow(Sender: TObject);
var
        I, Kol, Y, j, Pos, Pos1, x: Integer;
        Nom_Poz, Zak, IDGP, Briket, Briket1, Briket2, Br, Br1, Br2, Kol_Zap:
        string;
begin
        SG1.Enabled:=True;
        Clear_StringGrid(SG1);
        
        SG1.ColCount := 10;
        SG1.Cells[0, 0] := 'Заказ';
        SG1.Cells[1, 0] := 'БЗ';
        SG1.Cells[2, 0] := 'Кол во';
        SG1.Cells[3, 0] := 'Дата';
        SG1.Cells[4, 0] := 'IDКО';
        SG1.Cells[5, 0] := 'Клапан';
        SG1.Cells[6, 0] := 'Сборка';
        SG1.Cells[7, 0] := 'Сварка';
        SG1.Cells[8, 0] := 'Привод';
        SG1.Cells[9, 0] := 'IDGP';
        //  ([Заказ]= '+#39+Zak+#39+')AND
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where   ((СтатусНормы='
                + #39 + '0' + #39 + ') OR (СтатусНормы IS NULL)) AND (Статус='
                + #39 + '1' + #39 + ') AND (Отмена IS NULL) Order By Заказ ', ['KlapanaZap']) then
                exit;
        SG1.RowCount:=1;;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                SG1.RowCount := Form1.ADOQuery2.RecordCount + 1;

                for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        SG1.Cells[0, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SG1.Cells[1, i + 1] :=
                                Form1.ADOQuery2.FieldByName('БЗ').AsString;
                        SG1.Cells[2, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Кол во').AsString;
                        SG1.Cells[3, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Дата').AsString;
                        SG1.Cells[4, i + 1] :=
                                Form1.ADOQuery2.FieldByName('IDКО').AsString;
                        SG1.Cells[5, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Изделие').AsString;
                        SG1.Cells[6, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сборка Клапана').AsString;
                        SG1.Cells[7, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сварка').AsString;
                        SG1.Cells[8, i + 1] :=
                                Form1.ADOQuery2.FieldByName('МодПривода').AsString;
                        SG1.Cells[9, i + 1] :=
                                Form1.ADOQuery2.FieldByName('IdГП').AsString;
                        Form1.ADOQuery2.Next;
                end;
        end;
end;

procedure TFNormaVoz.FormClose(Sender: TObject; var Action: TCloseAction);
begin

        Label11.Caption := '0';
        Label9.Caption := '0';
        Label1.Caption := '0';
end;

//Очистка Грида

procedure TFNormaVoz.Clear_StringGrid1(StringGrid: TStringGrid);
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
procedure TFNormaVoz.Button3Click(Sender: TObject);
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
                + #39 + Label15.Caption + #39 + ') Order By Заказ ', ['ЗапускВозд']) then
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
procedure TFNormaVoz.Clear_StringGrid(SG:TStringGrid);
Var I:Integer;
Begin
        for i:=0 To SG.RowCount Do
        Begin
            SG.Rows[i].Clear;
        End;

end;
procedure TFNormaVoz.Button1Click(Sender: TObject);
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
                Pos_Sn,Pos_Privod,Pos_Isp,Pos_Ram,SvarkaS,Kol_Lops,DatStr: string;
        Svar, Sbor, Izdel,VElem,Elem,Oboz, IDGP,IDKO,Tip,EI,NC_S,Pos1,Pos2,Pos3,Pos4,Pos5,OboznSh,NC_S2,N2: string;
       Res_Tol55, Kol_Ed,Kol_Oboz,Kol_Ed1,Kol_Oboz1,NC,NC2,ND22,C,DlinI,ShirI,
       Tol,Razmet_KPU,DlinD,A_Lenta,NC_OB,T,NCB,
       NC_PILA,NC_NOG,NC_GIB: Double;
        Dat1, Dat2, Dat3: TDate;
        Kanban,Trumph,Nog,Ugl,Gib,Prok,Pila,Svarka,Tehnolog:Boolean;
        XL, XL1, XL2, XL_Temp, V1: Variant;
        Res_Tol1,Res_Tol2,Res_Resh,Res_N2,Res_Svarka,PosTrud:Integer;
        Nog_List:TStringList;
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
        e := 30;
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
        Memo2.Lines.Clear;
        DatStr := FormatDateTime('mm.dd.yyyy', dtp1.Date);
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
                NC_NOG:=0;
                NC_GIB:=0;
                NC_PILA:=0;
                if SG1.Cells[2, i] <> '' then
                begin
                        Kol_Oboz:=0;
                        Kol_Zap:=0;
                        NC_OB:=0;
                        NC:=0;
                        NC_S:='';
                        NC_S2:='';
                        Tip:='';
                        Nom := SG1.Cells[3, i];
                        Izdel := SG1.Cells[5, i];
                        if  SG1.Cells[14, i]='' then
                        Kol_Lop:=0
                        Else
                        Kol_Lop:= StrToInt(SG1.Cells[14, i]);
                        Kol_Zap := StrToInt(SG1.Cells[2, i]);
                        Tip:=Izdel;
                        Izdel := SG1.Cells[5, i];
                        IDGP:=SG1.Cells[9, i];
                        IDKO:=SG1.Cells[4, i];
                        if not Form1.mkQuerySelect(Form1.ADOQuery2, //Вытащить кол лопаток
                                'Select * from %s Where   (IdГП=' + #39 +
                                SG1.Cells[9, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[4, i] + #39 +') AND (Элемент=' + #39 +
                                'Лопатка' + #39 +') ', ['СпецифВозд']) then
                                exit;
                        Kol_lopS:=Form1.ADOQuery2.FieldByName('Количество').AsString;
                        if Kol_Lops='' then
                        Kol_Lop:=0
                        else
                        Kol_Lop:= Form1.ADOQuery2.FieldByName('Количество').AsInteger;
                        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where   (IdГП=' + #39 +
                                SG1.Cells[9, i] + #39 +') AND (IdКО=' + #39 +
                                SG1.Cells[4, i] + #39 +') ', ['СпецифВозд']) then
                                exit;
                 for J:=0 To  Form1.ADOQuery2.RecordCount-1 Do
                 Begin
                      Flag:=0;
                      KolGib:=0;
                      VElem:= Form1.ADOQuery2.FieldByName('ВидЭлемента').AsString;
                      Elem:= Form1.ADOQuery2.FieldByName('Элемент').AsString;
                      Oboz:= Form1.ADOQuery2.FieldByName('Обозначение').AsString;
                      //Kol_Ed:= Form1.ADOQuery2.FieldByName('КолНаЕд').AsFloat;
                      Kol_Ed:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                      //Trumph,Nog,Ugl,Gib,Prok,Pila
                      Dlina:=Form1.ADOQuery2.FieldByName('Длина').AsString;
                      Shir:= Form1.ADOQuery2.FieldByName('Ширина').AsString;
                      DlinRaz:= Form1.ADOQuery2.FieldByName('ДлинаРазв').AsString;
                      ShirRaz:=Form1.ADOQuery2.FieldByName('ШиринаРазв').AsString;
                      if DlinRaz ='' Then
                            DlinRaz:='0';
                      if ShirRaz ='' Then
                            ShirRaz:='0';
                      Kol_Gib:= Form1.ADOQuery2.FieldByName('КолГибов').AsString;
                      if Kol_Gib=''  Then
                      KolGib:=0
                      Else
                      KolGib:=StrToInt(Kol_Gib);
                      Ugol:= Form1.ADOQuery2.FieldByName('Угол').AsInteger;
                      Mat:= Form1.ADOQuery2.FieldByName('Материал').AsString;
                      if Mat='' Then
                                        Mat:='0';
                      EI:= Form1.ADOQuery2.FieldByName('ЕИ').AsString;

                      Kanban:= Form1.ADOQuery2.FieldByName('Канбан').AsBoolean;
                      Trumph:= Form1.ADOQuery2.FieldByName('Trumph').AsBoolean;
                      Nog:= Form1.ADOQuery2.FieldByName('Ножницы').AsBoolean;
                      Ugl:= Form1.ADOQuery2.FieldByName('Углоруб').AsBoolean;
                      Gib:= Form1.ADOQuery2.FieldByName('Гибка').AsBoolean;
                      Prok:= Form1.ADOQuery2.FieldByName('Прокатка').AsBoolean;
                      Pila:= Form1.ADOQuery2.FieldByName('Пила').AsBoolean;
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
                                Ugl:= Form1.ADOQuery4.FieldByName('Углоруб').AsBoolean;
                                Gib:= Form1.ADOQuery4.FieldByName('Гибка').AsBoolean;
                                Prok:= Form1.ADOQuery4.FieldByName('Прокатка').AsBoolean;
                                Pila:= Form1.ADOQuery4.FieldByName('Пила').AsBoolean;
                                Svarka:=Form1.ADOQuery4.FieldByName('Сварка').AsBoolean;
                        end;
                        {Ugol:= Form1.ADOQuery4.FieldByName('Угол').AsInteger;
                        Kol_Gib:= Form1.ADOQuery4.FieldByName('КолГибов').AsString;
                        if Kol_Gib=''  Then
                                KolGib:=0
                        Else
                                KolGib:=StrToInt(Kol_Gib);  }
                      end;
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
                     Res_N2:=AnsiCompareStr('2Н',N2);
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Trumph=True) Then
                      Begin
                        Kol_Ed:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                        Kol_Oboz:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                        SD3.Cells[2,s]:=Tip;
                        SD3.Cells[3,s]:=Elem;
                        SD3.Cells[4,s]:=FloatToStr(Kol_Oboz*Kol_Zap);
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
                      if (Kanban=False) AND (Nog=True) Then
                     Begin

                        Kol_Ed:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                        Kol_Oboz:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
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

                                res_Tol:=Pos(' ',Mat);//08pc
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
                        NC:=StrToFloat(NC_S)*Kol_Oboz;
                        //-----------------------------------------------------
                        NC_NOG:=NC_NOG+NC;
                        //-----------------------------------------------------
                        if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'11'+ #39 +') AND ([Высота]=' + #39 +IntToStr(Ugol)+ #39 +') ', ['[Резка Гибка]']) then
                                exit;
                                NC_S2:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                        if NC_S2='' Then
                                        NC_S2:='0';

                        Razmet_KPU:=Form1.ADOQuery1.FieldByName('Норма').AsFloat;
                        SD4.Cells[16,h]:=FloatToStr(StrToFloat(NC_S)+Razmet_KPU);
                        NC:=NC+(Razmet_KPU*Kol_Oboz);
                        //------------------------------------------------------
                        NC_NOG:=NC_NOG+(Razmet_KPU*Kol_Oboz);
                        //------------------------------------------------------
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
                        inc(h);
                        SD4.RowCount:= SD4.RowCount+1;
                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Gib=True) Then
                     Begin
                                Kol_Ed:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                                Kol_Oboz:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
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
                                 //----------------------------
                                 NC_GIB:=NC_GIB+(NC*Kol_Ed);
                                 //----------------------------
                                 SD5.Cells[14,n]:=FloatToStr(NC*Kol_Ed);
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
                                        //--------------------------------------
                                        NC_GIB:=NC_GIB+((NC*Kol_Ed)+(NC2*Kol_Ed));
                                        //--------------------------------------
                                        SD5.Cells[19,n]:=FloatToStr(StrToFloat(NC_S)+StrToFloat(NC_S2));
                                        SD5.Cells[14,n]:=FloatToStr((NC)+(NC2));
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
                                      T:=0.008*Kol_Otv*Kol_Ed+0.008*Kol_Ed; //Пробивка отверстий в одной ленте
                                      //-----------------------------------------
                                      //0.05 Нужно прибавлять прямо в Ексель один раз на на один тип ленты
                                        NC_GIB:=NC_GIB+T;
                                      //-----------------------------------------
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
                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Pila=True) Then
                      Begin
                      Kol_Ed:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                                SD6.Cells[2,w]:=Tip;
                                SD6.Cells[3,w]:=Elem;
                                SD6.Cells[4,w]:=FloatToStr(Kol_Ed*Kol_Zap);
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
                                 SD6.Cells[14,w]:=FloatToStr(NC*Kol_Ed);
                                 NC_Pila:=NC_PIla+(NC*Kol_Ed);
                                 if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                'UPDATE %s SET [Н\ч Пила]=' + #39
                                + Form1.ConvertFloat(SD6.Cells[14,w]) + #39 +
                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                ') AND (IDКО=' + #39 + IDKO + #39 +
                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['СпецифВозд']) then
                                Exit;
                                Inc(w);
                                SD6.RowCount:= SD6.RowCount+1;
                     end;
                      Form1.ADOQuery2.Next;
                  end;
                        if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                        'UPDATE %s SET [Ножницы]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_NOG)) + #39 +
                        ',[Пила]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_PILA)) + #39 +
                        ',[Гибка]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_GIB)) + #39 +
                        ',[СтатусНормы]=' + #39
                        + '1'+ #39 +',[Технолог Обраб]='+ #39+DatStr+#39+
                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                        ') ', ['KlapanaZap']) then
                        Exit;

                        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                        '[Ножницы]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_NOG)) + #39 +
                        ',[Пила]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_PILA)) + #39 +
                        ',[Гибка]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_GIB)) + #39 +
                        ',[СтатусНормы]=' + #39
                        + '1'+ #39 +',[Технолог Обраб]='+ #39+DatStr+#39+
                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +')', ['ЗапускВозд']) then
                        Exit;
                End;
         End;

        Close;
end;

procedure TFNormaVoz.SG2DblClick(Sender: TObject);
begin
        Label15.Caption := SG2.Cells[4, SG2.Row];
        Button3.Click;
end;

procedure TFNormaVoz.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFNormaVoz.btn1Click(Sender: TObject);
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

end.

