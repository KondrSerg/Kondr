unit UNorma1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, UNewNakl, ComCtrls, StdCtrls, Grids, Buttons, ExtCtrls,Math, DB,
  ADODB;

type
  TFNorma = class(TFNewNakl)
    dtp1: TDateTimePicker;
    procedure FormShow(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FNorma: TFNorma;

implementation

uses Main;

{$R *.dfm}

procedure TFNorma.FormShow(Sender: TObject);
var
        I, Kol, Y, j, Pos, Pos1, x: Integer;
        Nom_Poz, Zak, IDGP, Briket, Briket1, Briket2, Br, Br1, Br2, Kol_Zap:
        string;
begin
        SG1.Enabled:=True;
        Clear_StringGrid(SG1);
        
        SG1.ColCount := 9;
        SG1.Cells[0, 0] := 'Заказ';
        SG1.Cells[1, 0] := 'Кол во';
        SG1.Cells[2, 0] := 'Дата';
        SG1.Cells[3, 0] := 'ID';
        SG1.Cells[4, 0] := 'Клапан';
        SG1.Cells[5, 0] := 'Сборка';
        SG1.Cells[6, 0] := 'Сварка';
        SG1.Cells[7, 0] := 'Привод';
        SG1.Cells[8, 0] := 'IDGP';
        //  ([Заказ]= '+#39+Zak+#39+')AND
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where   ((СтатусНормы='
                + #39 + '0' + #39 + ') OR (СтатусНормы IS NULL)) AND (Статус='
                + #39 + '1' + #39 + ') AND (Отмена IS NULL) Order By Заказ ', ['Klapana']) then
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
                                Form1.ADOQuery2.FieldByName('Кол во').AsString;
                        SG1.Cells[2, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Дата').AsString;
                        SG1.Cells[3, i + 1] :=
                                Form1.ADOQuery2.FieldByName('ID').AsString;
                        SG1.Cells[4, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Изделие').AsString;
                        SG1.Cells[5, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сборка Клапана').AsString;
                        SG1.Cells[6, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сварка').AsString;
                        SG1.Cells[7, i + 1] :=
                                Form1.ADOQuery2.FieldByName('МодПривода').AsString;
                        SG1.Cells[8, i + 1] :=
                                Form1.ADOQuery2.FieldByName('IdГП').AsString;
                        Form1.ADOQuery2.Next;
                end;
        end;

end;

procedure TFNorma.Button4Click(Sender: TObject);
var     // Kol_ZAP,
         Kol, Kol_Zap_Ranee, Res, Nomer, Zak_Int, i, A, B, D, y, F_KPU,
        F_KPD,xy, Res1, e, Zag, Zap, j, g, k,Q,Res_Mat,Res_Proch,Res_Stand,qq,
        Res_Sbor,Res_Klap,Podstavka,Srav,Srav1:Integer;

        Res_VElem,Res_Elem,Res_Oboz,Flag,o,p,s,h,f,m,n,v,w,KolGib,AA,bb,Res_Tol,
        F2,F1,Res_KPU,Res_KPD,Res_list,Res_Ger:Integer;

        Res_Lop,Sten_V,Sten_G,hh,pp,Res_Kog_Priv,GG,Res_Kompl_priv,ww,www,
        Res_Detal,Os,R25,Ugol,Lenta,Kol_Otv,ID,Dl_Slova,
        DlinaI:Integer;

        Zak, Dat, Plan_Dat, Vn_Dat, Nom, Pos_Vst, Pos_Ml, R, Pereh, Privod,
        Zag_S,Zap_S, Dir_main, God, Mes, Fil,Dir,DlinRaz,ShirRaz,Mat,
        Dlina,Shir,Kol_Gib,Razm,Tol_S,Pos_Flan1,Pos_Flan,Pos_Dop,
        Pos_Sn,Pos_Privod,Pos_Isp,Pos_Ram,SvarkaS: string;

        Svar, Sbor, Izdel,VElem,Elem,Oboz, IDGP,Tip,EI,NC_S,Pos1,Pos2,Pos3,Pos4,
        Pos5,OboznSh,NC_S2,N2,DatStr,Kol_Ed_S,SH1,SH2: string;

       Res_Tol55, Kol_Ed,Kol_Oboz,Kol_Ed1,Kol_Oboz1,NC,NC2,ND22,C,DlinI,ShirI,Tol,Razmet_KPU,DlinD,A_Lenta,NC_OB
       ,L,Kol_Lop,T,NCB,
       NC_TRUMPF,nC_NOG,NC_PILA,NC_GIB: Double;

        Dat1, Dat2, Dat3: TDate;

        Kanban,Trumph,Nog,Ugl,Gib,Prok,Pila,PilaLent,Svarka:Boolean;
        Res_Tol1,Res_Tol2,Res_Resh,Res_N2,Res_Svarka,PosTrud,Res_N1,
        Flag_Razb_Klapana,// 1 Клапан состоит из двух половинок
        Klap
        :Integer;
begin
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
        SD1.RowCount:=2;
        SD2.RowCount:=2;
        SD3.RowCount:=2;
        SD4.RowCount:=2;
        SD5.RowCount:=2;
        SD6.RowCount:=2;
        SD7.RowCount:=2;
        SD8.RowCount:=2;
        SD9.RowCount:=2;
        SD.RowCount:=2;
        Stenki.RowCount:=2;
        DatStr := FormatDateTime('mm.dd.yyyy', dtp1.Date);
        Memo2.Lines.Clear;
        SD1.Cells[2,0]:='Kanban True';
        SD2.Cells[2,0]:='Kanban False';
        SD3.Cells[2,0]:='Trumph';
        SD4.Cells[2,0]:='Nognicy';
        SD5.Cells[2,0]:='Gibka';
        SD6.Cells[2,0]:='Prokat';
        SD7.Cells[2,0]:='Pila';
        SD8.Cells[2,0]:='Uglorub';

        ProgressBar1.Min:=0;
        ProgressBar1.Position:=0;
        ProgressBar1.Max:=SG1.RowCount-1;
        for i := 1 to SG1.RowCount do
        begin                                 //
                if SG1.Cells[2, i] <> '' then
                begin
                        Kol_Oboz:=0;
                       // Kol_Zap:=0;
                        NC_OB:=0;
                        NC:=0;
                        NC_S:='';
                        NC_S2:='';
                        Tip:='';
                        Nom := SG1.Cells[2, i];
                        Izdel := SG1.Cells[4, i];
                       // Kol_Zap := StrToInt(SG1.Cells[1, i]);
                        Res := Pos('КПД', Izdel);
                        if Res <> 0 then
                        Tip:='КПД';
                        Res := Pos('КПУ', Izdel);
                        if Res <> 0 then
                        Tip:='КПУ';
                        //Клапан ГЕРМИК-ДУ-710*710-2*ф-ЭМП220-сн-0-0
                        res:=Pos('ГЕРМИК-ДУ',IZdel);
                        if Res<>0 Then
                        Tip:='ГЕРМИК-ДУ';

                      Izdel := SG1.Cells[4, i];
                      Res_Kpd:=Pos('КПД',Izdel);
                      Res_Kpu:=Pos('КПУ',Izdel);
                      Res_Ger:=Pos('ГЕРМИК-ДУ',Izdel);
                      //++++++++++++++++++++++++++++++++
                      if Res_Kpu<>0 Then
                      Begin
                      Res := Pos(' ', Izdel);
                        Delete(Izdel, 1, Res);
                        //======================================== KPU
                        Res := Pos('-', Izdel);
                        Pos1 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== 1 H
                        Res := Pos('-', Izdel);
                        N2:=Copy(Izdel,1,2);
                        Pos2 := Copy(Izdel, 1, 1);
                        Pos3 := Copy(Izdel, 2, 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        //Клапан КПУ-1Н-О-Н-600*400-2*ф-MB220-сн-0-0-0-0-0-0
                        //Клапан КПД-4-01-500*500-1*ф-ЭМП220-вн-0-0
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos4 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos5 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        //Клапан КПУ-2Н-О-Н-200-2*ф-MB220-T-сн-0-0-0-2*100-0-0
                        //Клапан КПУ-1Н-О-Н-200*200-2*ф-MB220-сн-0-с-0-0-0-0
                        Res := Pos('-', Izdel);//Проверка на круглый если -
                        If Res >5 Then //200*200- Квадрат
                        Begin
                                Res := Pos('*', Izdel);
                                A := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                B := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                        end
                        else
                        begin
                                Res := Pos('-', Izdel);
                                A := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                B:=0;
                        end;

                        end;
                        //++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        //Клапан ГЕРМИК-ДУ-710*710-2*ф-ЭМП220-сн-0-0-
                        if Res_Ger<>0 Then
                        Begin
                        Res := Pos(' ', Izdel);
                        Delete(Izdel, 1, Res);
                        //======================================== KPU
                        Res := Pos('-', Izdel);
                        Pos1 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== 1 H
                        Res := Pos('-', Izdel);

                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('*', Izdel);
                        A := StrToInt(Copy(Izdel, 1, Res - 1));
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        B := StrToInt(Copy(Izdel, 1, Res - 1));
                        Delete(Izdel, 1, Res);
                        end;
                        //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                        Res_KPD := Pos('КПД', Izdel);
                        if Res_KPD <> 0 then
                        begin

                                Kol := StrToInt(SG1.Cells[1, i]);
                                Res := Pos(' ', Izdel);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos1 := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos2 := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Isp := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('*', Izdel);
                                A := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                B := StrToInt(Copy(Izdel, 1, Res - 1));
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('*', Izdel);
                                Pos_Flan := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Flan1 := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Privod := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Sn := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Res := Pos('-', Izdel);
                                Pos_Dop := Copy(Izdel, 1, Res - 1);
                                Delete(Izdel, 1, Res);
                                //========================================
                                Pos_Ram := Izdel;
                                //========================================

                        End;
                      Izdel := SG1.Cells[4, i];
                      Res_Kpd:=Pos('КПД',Izdel);
                      Res_Kpu:=Pos('КПУ',Izdel);
                      F2:=Pos('2*ф',Izdel);
                      F1:=Pos('1*ф',Izdel);
                      IDGP:=SG1.Cells[8, i];
                      if F2<>0 Then
                      Tip:=Tip+'-'+Pos2+Pos3+'-'+IntToStr(A)+'*'+IntToStr(B)+'-2*ф';
                      if F1<> 0 Then
                      Tip:=Tip+'-'+Pos2+Pos3+'-'+IntToStr(A)+'*'+IntToStr(B)+'-1*ф';

                        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where   (IdГП=' + #39 +
                                 IDGP+ #39 +') ', ['Специф']) then
                                exit;
                Flag_Razb_Klapana:=0;
                 for J:=0 To  Form1.ADOQuery2.RecordCount-1 Do  //Проверка на Клапан разбитый на 2
                 Begin
                        Elem:= Form1.ADOQuery2.FieldByName('Элемент').AsString;
                        Kol_Ed_S:= Form1.ADOQuery2.FieldByName('Количество').AsString;
                       { PosTrud:=Pos('.',Kol_Ed_S);
                        if PosTrud<>0 Then
                        Begin
                                Delete(Kol_Ed_S,PosTrud,1);
                                Insert(',',Kol_Ed_S,PosTrud);
                        End;  }
                        Kol_Ed:= StrToFloat(Kol_Ed_S);
                        Klap:=AnsiCompareStr('Клапан',Elem);
                        if (Klap=0) and (Kol_Ed>1) then
                        Begin
                                Flag_Razb_Klapana:=1;//Клапан разбит на 2
                                Break;
                        End;
                        Form1.ADOQuery2.Next;
                 end;
                 NC_TRUMPF:=0;
                 NC_NOG:=0;
                 NC_PILA:=0;
                 NC_GIB:=0;
                 Form1.ADOQuery2.First;
                 for J:=0 To  Form1.ADOQuery2.RecordCount-1 Do
                 Begin
                      Flag:=0;
                      KolGib:=0;
                      VElem:= Form1.ADOQuery2.FieldByName('ВидЭлемента').AsString;
                      Elem:= Form1.ADOQuery2.FieldByName('Элемент').AsString;
                      Oboz:= Form1.ADOQuery2.FieldByName('Обозначение').AsString;
                      //Kol_Ed:= Form1.ADOQuery2.FieldByName('КолНаЕд').AsFloat;
                      Kol_Ed_S:= Form1.ADOQuery2.FieldByName('Количество').AsString;
                       { PosTrud:=Pos('.',Kol_Ed_S);
                        if PosTrud<>0 Then
                        Begin
                                Delete(Kol_Ed_S,PosTrud,1);
                                Insert(',',Kol_Ed_S,PosTrud);
                        End;}
                        Kol_Ed:= StrToFloat(Kol_Ed_S);
                      Kol_Oboz:= StrToFloat(Kol_Ed_S);
                      //Trumph,Nog,Ugl,Gib,Prok,Pila
                      Dlina:= Form1.ADOQuery2.FieldByName('Ширина').AsString;
                      Shir:= Form1.ADOQuery2.FieldByName('Длина').AsString;
                      DlinRaz:= Form1.ADOQuery2.FieldByName('ШиринаРазв').AsString;
                      ShirRaz:= Form1.ADOQuery2.FieldByName('ДлинаРазв').AsString;
                      if DlinRaz ='' Then
                            DlinRaz:='0';
                      if ShirRaz ='' Then
                            ShirRaz:='0';
                      Kol_Gib:= Form1.ADOQuery2.FieldByName('КолГибов').AsString;
                      if Kol_Gib=''  Then
                      KolGib:=0
                      Else
                      KolGib:=StrToInt(Kol_Gib);

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
                      PilaLent:= Form1.ADOQuery2.FieldByName('Пила ленточная').AsBoolean;

                      SH1:= Form1.ADOQuery2.FieldByName('ШиринаПолки1').AsString;
                      SH2:= Form1.ADOQuery2.FieldByName('ШиринаПолки2').AsString;
                      //==================================================


                      Res_Mat:=Pos('Материалы',VELem);
                      Res_Proch:=Pos('Прочие изделия',VELem);
                      Res_Stand:=Pos('Стандартные изделия',VELem);
                      Res_Sbor:=Pos('Сборочные единицы',VELem);
                      Res_Klap:=Pos('Клапан',ELem);
                      Res_Kompl_priv:=Pos('Комплект привода',ELem);
                      Os:= Pos('Ось',ELem);
                      hh:=A MOD 50;
                      gg:=B MOD 50;

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
                                OboznSh + #39 +') ', ['Шаблон']) then
                                exit;
                        Kanban:= Form1.ADOQuery4.FieldByName('Канбан').AsBoolean;
                        Trumph:= Form1.ADOQuery4.FieldByName('Trumph').AsBoolean;
                        Nog:= Form1.ADOQuery4.FieldByName('Ножницы').AsBoolean;
                        Ugl:= Form1.ADOQuery4.FieldByName('Углоруб').AsBoolean;
                        Gib:= Form1.ADOQuery4.FieldByName('Гибка').AsBoolean;
                        Prok:= Form1.ADOQuery4.FieldByName('Прокатка').AsBoolean;
                        Pila:= Form1.ADOQuery4.FieldByName('Пила').AsBoolean;
                        PilaLent:= Form1.ADOQuery4.FieldByName('Пила ленточная').AsBoolean;
                        Svarka:=Form1.ADOQuery4.FieldByName('Сварка').AsBoolean;
                        Ugol:= Form1.ADOQuery4.FieldByName('Угол').AsInteger;
                        Kol_Gib:= Form1.ADOQuery4.FieldByName('КолГибов').AsString;
                        if Kol_Gib=''  Then
                                KolGib:=0
                        Else
                                KolGib:=StrToInt(Kol_Gib);
                      end;
                      //XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                      If (Svarka=True) Then
                      Begin
                            Dl_Slova:=Length(OboznSh);
                            Delete(OboznSh,Dl_Slova,1);
                      end;
        //SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS    STENKI
                      Sten_V:=Pos('Стенка вертикальная',Elem);
                      Sten_G:=Pos('Стенка горизонтальная',Elem);
                      
                      //gg:=B MOD 50;
                      F2:=Pos('2*ф',Izdel);
                      if (Res_KPU<>0) AND (F2<>0) AND (Res_Detal=0)  Then
                      Begin
                         if  ((Sten_V<>0) OR (Sten_G<>0)) then
                         begin
                                Dlina:= Form1.ADOQuery2.FieldByName('Длина').AsString;
                                DlinaI:=StrToInt(Dlina);
                                if (Sten_G<>0) then
                                        DlinaI:=StrToInt(Dlina)-60;
                                hh:=DlinaI MOD 50;
                          If ((DlinaI>=150) AND (DlinaI<=1000))  AND (hh=0)   Then
                          Begin
                          Kanban:=True ;
                          Trumph:=False;
                          Gib:=False;

                          end;
                         End;
                      end;
                      if (Res_KPU<>0) AND (F2<>0) AND (Res_Detal=0)  Then
                      Begin

                       if  ((Sten_G<>0) OR (Sten_V<>0)) Then
                       Begin
                                Dlina:= Form1.ADOQuery2.FieldByName('Длина').AsString;
                                DlinaI:=StrToInt(Dlina);
                                if (Sten_G<>0) then
                                        DlinaI:=StrToInt(Dlina)-60;
                                hh:=DlinaI MOD 50;
                          if  ((DlinaI<150)  OR (DlinaI>1000)) OR (hh<>0) Then
                          Begin
                          Kanban:=False ;
                          Trumph:=True;
                          Gib:=True;
                          if not Form1.mkQueryUpdate2(Form1.ADOQuery1,
                                        'UPDATE %s SET [Trumph]=' +
                                        #39+ 'True' + #39 +
                                        ', Гибка='+#39+ 'True' + #39 +
                                        ', Канбан='+#39+ 'False' + #39 +
                                        ' WHERE (Обозначение=' + #39 + Oboz + #39 +') AND ([IdГП]=' + #39 + IDGP + #39 +
                                                ') ', ['Специф']) then
                                        Exit;

                          end;

                       end;
        //SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS    STENKI
                      end;
                     Res_N2:=AnsiCompareStr('2Н',N2);
                     Res_N1:=AnsiCompareStr('1Н',N2);
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Trumph=True) Then
                      Begin
                      Kol_Ed:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                        if Res_KPD<>0 Then
                        Begin
                                SD3.Cells[2,s]:=Tip;
                                SD3.Cells[3,s]:=Elem;              //*Kol_Zap
                                SD3.Cells[4,s]:=FloatToStr(Kol_Oboz);
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
                               { if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0) Then
                                C := StrToFloat(DlinRaz)-(((A - 24) /(Kol_Ed/2) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND ((A=100) OR (B=100))  Then
                                C := StrToFloat(DlinRaz)-((A /(Kol_Ed/2)- 14)/2-13) - 24.3;
                                if (Res_KPU<>0) AND (F1<>0) AND (Res_N2=0)  Then
                                C := StrToFloat(DlinRaz)-(StrToFloat(DlinRaz)/2-34.3) - 24.3; }
                                C := StrToFloat(SH1);
                                SD3.Cells[14,s]:=FloatToStr(C);
                                SD3.Cells[13,s]:=VElem;
                                Inc(s);
                                SD3.RowCount:= SD3.RowCount+1;
                        end;
                        if Res_KPU<>0 Then
                        Begin
                                SD3.Cells[2,s]:=Tip;
                                SD3.Cells[3,s]:=Elem;            //  *Kol_Zap
                                SD3.Cells[4,s]:=FloatToStr(Kol_Oboz);
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
                                {if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0) Then
                                        C := StrToFloat(DlinRaz)-(((A - 24) /(Kol_Ed/2) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND ((A=100) OR (B=100))  Then
                                        C := StrToFloat(DlinRaz)-((A /(Kol_Ed/2)- 14)/2-13) - 24.3;
                                if (Res_KPU<>0) AND (F1<>0) AND (Res_N2=0)  Then
                                        C := StrToFloat(DlinRaz)-(StrToFloat(DlinRaz)/2-34.3) - 24.3;  }
                                C := StrToFloat(SH1);
                                SD3.Cells[14,s]:=FloatToStr(C);
                                SD3.Cells[13,s]:=VElem;
                                Inc(s);
                                SD3.RowCount:= SD3.RowCount+1;
                        end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        if (Res_KPU=0) AND (Res_KPD=0) Then  //Гермик
                        Begin
                                SD3.Cells[2,s]:=Tip;
                                SD3.Cells[3,s]:=Elem;            //  *Kol_Zap
                                SD3.Cells[4,s]:=FloatToStr(Kol_Oboz);
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
                                {if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0) Then
                                        C := StrToFloat(DlinRaz)-(((A - 24) /(Kol_Ed/2) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND ((A=100) OR (B=100))  Then
                                        C := StrToFloat(DlinRaz)-((A /(Kol_Ed/2)- 14)/2-13) - 24.3;
                                if (Res_KPU<>0) AND (F1<>0)  Then
                                        C := StrToFloat(DlinRaz)-((A /(Kol_Ed/2)- 14)/2-13) - 24.3; }
                                C := StrToFloat(SH1);
                                SD3.Cells[14,s]:=FloatToStr(C);
                                SD3.Cells[13,s]:=VElem;
                                Inc(s);
                                SD3.RowCount:= SD3.RowCount+1;

                        end;

                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++ Ножницы
                      if (Kanban=False) AND (Nog=True) Then
                      Begin
                      Kol_Ed:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                                SD4.Cells[2,h]:=Tip;
                                SD4.Cells[3,h]:=Elem;                // *Kol_Zap
                                SD4.Cells[4,h]:=FloatToStr(Kol_Oboz);
                                SD4.Cells[6,h]:=Oboz;
                                SD4.Cells[0,h]:=IntToStr(h);
                                SD4.Cells[1,h]:=EI;
                                SD4.Cells[7,h]:=Shir;
                                SD4.Cells[8,h]:=Dlina;
                                SD4.Cells[9,h]:=ShirRaz;
                                SD4.Cells[10,h]:=DlinRaz;
                                if Res_KPD<>0 Then
                                Begin
                                        SD4.Cells[7,h]:=Dlina;
                                        SD4.Cells[8,h]:=Shir;
                                        SD4.Cells[9,h]:=DlinRaz;
                                        SD4.Cells[10,h]:=ShirRaz;
                                end;

                                SD4.Cells[11,h]:=Kol_Gib;
                                SD4.Cells[12,h]:=Mat;
                                if DlinRaz='' Then
                                DlinRaz:='0';
                                //+ ++++++++++++++++++++++++++++++++++++++++++
                                //Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90
                                 Res_List:=Pos('Лист ДПРНМ 1,0х600х1500 Л63 ГОСТ 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
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
                                if ((Srav=0) OR (Srav>0)) AND (Tol<=2) Then
                                Begin
                                        Tol_S:='1.2';
                                        Res_Tol55:=2;
                                End;
                                DlinI:=StrToFloat(ShirRAz);
                                ShirI:=StrToFloat(DlinRaz);

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
                                NC_NOG:=NC_NOG+NC;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=
                               { if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0) Then
                                        C := StrToFloat(DlinRaz)-(((A - 24) /(Kol_Ed/2) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND (A>100) AND (B>100) AND (F2<>0)AND (Flag_Razb_Klapana =1) Then
                                        C := StrToFloat(DlinRaz)-(((A - 24) /(Kol_Ed/4) - 28)/2) - 34.14;
                                if (Res_KPU<>0) AND ((A=100) OR   (B=100))  Then
                                        C := StrToFloat(DlinRaz)-((A /(Kol_Ed/2) - 14)/2-13) - 24.3;
                                if (Res_KPU<>0) AND (F1<>0)  AND (Res_N2=0)  Then
                                        C := StrToFloat(DlinRaz)-(StrToFloat(DlinRaz) /(Kol_Ed/2)- 34.3) - 24.3;}
                                C := StrToFloat(SH1);
                                if Flag_Razb_Klapana =1 Then
                                        Kol_Lop:=Kol_Ed/4
                                else
                                        Kol_Lop:=Kol_Ed/2;
                               { L :=(A/Kol_LOp-14)/2 -13;
                                if (Res_KPU<>0) AND (F1<>0)  AND (Res_N1=0)  Then
                                        C := StrToFloat(DlinRaz)-(L) - 24.3;  }
                                //+ ++++++++++++++++++++++++++++++++++++++++++
                                Res_Lop:=AnsiCompareStr('Корпус лопатки',SD4.Cells[3,h]);
                                 if (Res_KPU<>0) AND (Res_Lop=0) AND (C<220)  Then //4 Ugla
                                 Begin
                                       Ugol:=4;
                                 end;
                                 if (Res_KPU<>0) AND (Res_Lop=0) AND ((C>=220) AND (C<440))  Then //6 Uglov
                                 Begin
                                       Ugol:=6;
                                 end;
                                 if (Res_KPU<>0) AND (Res_Lop=0) AND (C>=400)   Then //8 Uglov
                                 Begin
                                       Ugol:=8;
                                 end;
                                 //+====================================
                                if Res_Tol55=0.55 Then
                                Begin
                                if ShirI<=600 Then
                                        ShirI:=600;
                                //---------------
                                if (ShirI>600)  Then
                                        ShirI:=1250;
                                End;
                                //+====================================
                                if Res_Tol55=7 Then
                                Begin
                                if ShirI<=500 Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500)  Then
                                        ShirI:=1250;

                                End;
                                //+====================================
                                if Res_Tol55=2 Then
                                Begin
                                if ShirI<=500 Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500)  Then
                                        ShirI:=500;
                                //---------------
                                if (ShirI>500) Then
                                        ShirI:=1250;

                                End;
                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'11'+ #39 +') And ([Толщина]='+Tol_S+') And ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([Высота]=' + #39 +FloatToStr(ShirI)+ #39 +')', ['[Резка Гибка]']) then
                                exit;
                                NC_S2:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                if NC_S2='' Then
                                        NC_S2:='0';
                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([№]=' + #39 +'12'+ #39 +') ', ['[Резка Гибка]']) then
                                exit;
                                if C>200 Then
                                Razmet_KPU:=Form1.ADOQuery1.FieldByName('Норма').AsFloat
                                Else
                                Razmet_KPU:=0;
                                SD4.Cells[16,h]:=FloatToStr(StrToFloat(NC_S)+(StrToFloat(NC_S2)+Razmet_KPU)*Ugol);
                                NC:=NC+(((StrToFloat(NC_S2)*Ugol+Razmet_KPU))*Kol_Oboz);
                                NC_NOG:=NC_NOG+(((StrToFloat(NC_S2)*Ugol+Razmet_KPU))*Kol_Oboz);
                                //end;
                                SD4.Cells[13,h]:=FloatToStr(C);
                                SD4.Cells[14,h]:=VElem;
                                SD4.Cells[15,h]:=FloatToStr(NC);
                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [Н\ч Ножницы]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                        Exit;
                                Inc(h);
                                SD4.RowCount:= SD4.RowCount+1;
                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Gib=True) Then
                      Begin
                      Kol_Ed:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                      Kol_Oboz:= Form1.ADOQuery2.FieldByName('Количество').AsFloat;
                                SD5.Cells[2,n]:=Tip;
                                SD5.Cells[3,n]:=Elem;           // *Kol_Zap
                                SD5.Cells[4,n]:=FloatToStr(Kol_Ed);
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
                                Res_Lop:=AnsiCompareStr('Корпус лопатки',Elem);
                                if Res_Lop=0 Then
                                Begin                                          //*Kol_Zap
                                        SD5.Cells[14,n]:=FloatToStr(0.034*Kol_Ed);
                                        SD5.Cells[19,n]:=FloatToStr(0.034);
                                        NC_GIB:=NC_GIB+(0.034*Kol_Ed);
                                End
                                Else
                                Begin
                                        ShirI:=StrToFloat(ShirRaz);
                                        DlinI:=StrToFloat(DlinRaz);
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
                                        NC:=StrToFloat(NC_S)*Kol_Ed;                //*Kol_Zap
                                        NC_GIB:=NC_GIB+NC;
                                        SD5.Cells[14,n]:=FloatToStr(NC);
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
                                                NC:=0
                                        Else
                                                NC:=StrToFloat(NC_S)*Kol_Ed;
                                        //++++++++++++
                                        NC_S2:= (Form1.ADOQuery1.FieldByName('Длина').AsString);
                                        If NC_S2='' Then
                                                NC2:=0
                                        Else
                                                NC2:=StrToFloat(NC_S2)*Kol_Ed;    //*Kol_Zap /Kol_Zap
                                        NC_GIB:=NC_GIB+NC+NC2;
                                        SD5.Cells[19,n]:=FloatToStr(StrToFloat(NC_S)+StrToFloat(NC_S2));
                                        SD5.Cells[14,n]:=FloatToStr((NC)+(NC2));
                                end;
                                Lenta:=Pos('Лента уплотнительная',Elem);
                                if Lenta <>0 Then
                                Begin
                                      DlinD:=B-1.5;
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


                                      if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                        'Select * from %s Where   ([№]=' + #39 +'12'+ #39 +') And ([Высота]=' + #39 +FloatToStr(DlinI)+ #39 +
                                        ')', ['[Резка Гибка]']) then
                                        exit;
                                      Kol_otv:= (Form1.ADOQuery1.FieldByName('Норма').AsInteger);
                                      A_Lenta:=(B-1.5-(Kol_Otv-1)*150)/2;
                                      SD5.Cells[15,n]:=FloatToStr(DlinD);//В Клапана-1,5
                                      SD5.Cells[16,n]:= FloatToStr(A_Lenta);
                                      SD5.Cells[17,n]:=IntToStr(Kol_Otv);
                                      SD5.Cells[18,n]:='150';      //*Kol_Zap   *Kol_Zap
                                      T:=0.008*Kol_Otv*Kol_Ed+0.008*Kol_Ed; //Пробивка отверстий в одной ленте
                                      //0.05 Нужно прибавлять прямо в Ексель один раз на на один тип ленты
                                      NC_GIB:=NC_GIB+T;
                                      SD5.Cells[20,n]:= FloatToStr(T);
                                end;
                                if SD5.Cells[20,n]='' Then
                                NCB:=(StrToFloat(SD5.Cells[14,n])+0)
                                Else NCB:=(StrToFloat(SD5.Cells[14,n])+StrToFloat(SD5.Cells[20,n]));

                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [Н\ч Гибка]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NCB)) + #39 +
                                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
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
                                SD6.Cells[3,w]:=Elem;            //*Kol_Zap
                                SD6.Cells[4,w]:=FloatToStr(Kol_Ed);
                                SD6.Cells[6,w]:=Oboz;
                                SD6.Cells[0,w]:=IntToStr(w);
                                SD6.Cells[1,w]:=EI;
                                SD6.Cells[7,w]:=Shir;
                                SD6.Cells[8,w]:=Dlina;
                                SD6.Cells[9,w]:=ShirRaz;
                                SD6.Cells[10,w]:=DlinRaz ;
                                SD6.Cells[12,w]:=Mat;
                                SD6.Cells[13,w]:=VElem;
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
                                        Res_Detal:=AnsiCompareStr('Детали',VElem);
                                        Res_Lop:=Pos('Лопатка',Elem);
                                        Res_Ger:=Pos('ГЕРМИК',Izdel);
                                        R25:=Pos('р25',Izdel);
                                        Res_List:=Pos('20001',Mat);//Лопатка Р25
                                        if Res_Ger=0 Then
                                          SD6.Cells[9,w]:=ShirRaz;                         //  AND (Res_List=0)
                                        if (Res_Detal=0) AND (Res_Lop<>0) AND (Res_Ger<>0)  Then
                                        Begin

                                          SD6.Cells[9,w]:=ShirRaz;
                                          DlinI:=StrToFloat(ShirRaz);
                                          if DlinI<=500 Then
                                          DlinI:=500;
                                          if DlinI>500 Then
                                          DlinI:=1250;
                                          if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                          'Select * from %s Where   ([№]=' + #39 +'6'+ #39 +') And  ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[Резка Гибка]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S)+0.02*Kol_Ed;      //  *Kol_Zap  *Kol_Zap
                                          NC_PILA:=NC_PILA+NC;
                                          SD6.Cells[14,w]:=FloatToStr(NC);
                                          if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                                'UPDATE %s SET [Н\ч Пила]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                                Exit;
                                          Inc(w);
                                                SD6.RowCount:= SD6.RowCount+1;
                                        End;
                                        Ugol:=Pos('Стенка',Elem);
                                        //++++++++++++++++++++++++++++++++++++++++++++++++
                                        if (Res_Detal=0) AND (Ugol<>0) AND (R25<>0) Then
                                        Begin
                                          DlinI:=StrToFloat(Shir)+78;
                                          SD6.Cells[9,w]:=FloatToStr(DlinI);
                                          if DlinI<=350 Then
                                          DlinI:=350;
                                          if (DlinI>350) AND (DlinI<=500) Then
                                          DlinI:=500;
                                          if (DlinI>500) AND (DlinI<=750) Then
                                          DlinI:=750;
                                          if (DlinI>750) Then
                                          DlinI:=1000;
                                          if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                          'Select * from %s Where   ([№]=' + #39 +'7'+ #39 +') And  ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[Резка Гибка]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S)+0.007*Kol_Ed;   //  *Kol_Zap *Kol_Zap
                                          SD6.Cells[14,w]:=FloatToStr(NC);
                                          NC_PILA:=NC_PILA+NC;
                                          if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                                'UPDATE %s SET [Н\ч Пила]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                                ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                                Exit;
                                          Inc(w);
                                                SD6.RowCount:= SD6.RowCount+1;
                                        End;
                                        //++++++++++++++++++++++++++++++++++++++++++++++++
                                        if (Res_Detal=0) AND (Res_Lop<>0) AND (R25<>0) AND (Res_List<>0) Then
                                        Begin
                                          DlinI:=StrToFloat(Shir);
                                          SD6.Cells[9,w]:=FloatToStr(DlinI);
                                          if DlinI<=250 Then
                                          DlinI:=250;
                                          if (DlinI>250) AND (DlinI<=500) Then
                                          DlinI:=500;
                                          if (DlinI>500) AND (DlinI<=750) Then
                                          DlinI:=750;
                                          if (DlinI>750) Then
                                          DlinI:=1000;
                                          if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                          'Select * from %s Where   ([№]=' + #39 +'8'+ #39 +') And  ([Длина]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[Резка Гибка]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('Норма').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S)*Kol_Ed;             // *Kol_Zap *Kol_Zap
                                          SD6.Cells[14,w]:=FloatToStr(NC);
                                          NC_PILA:=NC_PILA+NC;
                                          if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [Н\ч Пила]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
                                                ') AND (Элемент=' + #39 + Elem + #39 +') AND (Обозначение=' + #39 + Oboz + #39 +')  ', ['Специф']) then
                                        Exit;
                                          Inc(w);
                                                SD6.RowCount:= SD6.RowCount+1;

                                        End;
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
                        ') ', ['Klapana']) then
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
                        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +')', ['Запуск']) then
                        Exit;
                End;
                ProgressBar1.Position:=I;
        End;
        ProgressBar1.Position:=0;
        Close;

end;

end.
