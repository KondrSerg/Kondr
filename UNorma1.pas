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
        SG1.Cells[0, 0] := '�����';
        SG1.Cells[1, 0] := '��� ��';
        SG1.Cells[2, 0] := '����';
        SG1.Cells[3, 0] := 'ID';
        SG1.Cells[4, 0] := '������';
        SG1.Cells[5, 0] := '������';
        SG1.Cells[6, 0] := '������';
        SG1.Cells[7, 0] := '������';
        SG1.Cells[8, 0] := 'IDGP';
        //  ([�����]= '+#39+Zak+#39+')AND
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where   ((�����������='
                + #39 + '0' + #39 + ') OR (����������� IS NULL)) AND (������='
                + #39 + '1' + #39 + ') AND (������ IS NULL) Order By ����� ', ['Klapana']) then
                exit;
        SG1.RowCount:=1;;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                SG1.RowCount := Form1.ADOQuery2.RecordCount + 1;

                for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        SG1.Cells[0, i + 1] :=
                                Form1.ADOQuery2.FieldByName('�����').AsString;
                        SG1.Cells[1, i + 1] :=
                                Form1.ADOQuery2.FieldByName('��� ��').AsString;
                        SG1.Cells[2, i + 1] :=
                                Form1.ADOQuery2.FieldByName('����').AsString;
                        SG1.Cells[3, i + 1] :=
                                Form1.ADOQuery2.FieldByName('ID').AsString;
                        SG1.Cells[4, i + 1] :=
                                Form1.ADOQuery2.FieldByName('�������').AsString;
                        SG1.Cells[5, i + 1] :=
                                Form1.ADOQuery2.FieldByName('�\� ������ �������').AsString;
                        SG1.Cells[6, i + 1] :=
                                Form1.ADOQuery2.FieldByName('�\� ������').AsString;
                        SG1.Cells[7, i + 1] :=
                                Form1.ADOQuery2.FieldByName('����������').AsString;
                        SG1.Cells[8, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Id��').AsString;
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
        Flag_Razb_Klapana,// 1 ������ ������� �� ���� ���������
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
                        Res := Pos('���', Izdel);
                        if Res <> 0 then
                        Tip:='���';
                        Res := Pos('���', Izdel);
                        if Res <> 0 then
                        Tip:='���';
                        //������ ������-��-710*710-2*�-���220-��-0-0
                        res:=Pos('������-��',IZdel);
                        if Res<>0 Then
                        Tip:='������-��';

                      Izdel := SG1.Cells[4, i];
                      Res_Kpd:=Pos('���',Izdel);
                      Res_Kpu:=Pos('���',Izdel);
                      Res_Ger:=Pos('������-��',Izdel);
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
                        //������ ���-1�-�-�-600*400-2*�-MB220-��-0-0-0-0-0-0
                        //������ ���-4-01-500*500-1*�-���220-��-0-0
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos4 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos5 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        //������ ���-2�-�-�-200-2*�-MB220-T-��-0-0-0-2*100-0-0
                        //������ ���-1�-�-�-200*200-2*�-MB220-��-0-�-0-0-0-0
                        Res := Pos('-', Izdel);//�������� �� ������� ���� -
                        If Res >5 Then //200*200- �������
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
                        //������ ������-��-710*710-2*�-���220-��-0-0-
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
                        Res_KPD := Pos('���', Izdel);
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
                      Res_Kpd:=Pos('���',Izdel);
                      Res_Kpu:=Pos('���',Izdel);
                      F2:=Pos('2*�',Izdel);
                      F1:=Pos('1*�',Izdel);
                      IDGP:=SG1.Cells[8, i];
                      if F2<>0 Then
                      Tip:=Tip+'-'+Pos2+Pos3+'-'+IntToStr(A)+'*'+IntToStr(B)+'-2*�';
                      if F1<> 0 Then
                      Tip:=Tip+'-'+Pos2+Pos3+'-'+IntToStr(A)+'*'+IntToStr(B)+'-1*�';

                        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where   (Id��=' + #39 +
                                 IDGP+ #39 +') ', ['������']) then
                                exit;
                Flag_Razb_Klapana:=0;
                 for J:=0 To  Form1.ADOQuery2.RecordCount-1 Do  //�������� �� ������ �������� �� 2
                 Begin
                        Elem:= Form1.ADOQuery2.FieldByName('�������').AsString;
                        Kol_Ed_S:= Form1.ADOQuery2.FieldByName('����������').AsString;
                       { PosTrud:=Pos('.',Kol_Ed_S);
                        if PosTrud<>0 Then
                        Begin
                                Delete(Kol_Ed_S,PosTrud,1);
                                Insert(',',Kol_Ed_S,PosTrud);
                        End;  }
                        Kol_Ed:= StrToFloat(Kol_Ed_S);
                        Klap:=AnsiCompareStr('������',Elem);
                        if (Klap=0) and (Kol_Ed>1) then
                        Begin
                                Flag_Razb_Klapana:=1;//������ ������ �� 2
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
                      VElem:= Form1.ADOQuery2.FieldByName('�����������').AsString;
                      Elem:= Form1.ADOQuery2.FieldByName('�������').AsString;
                      Oboz:= Form1.ADOQuery2.FieldByName('�����������').AsString;
                      //Kol_Ed:= Form1.ADOQuery2.FieldByName('�������').AsFloat;
                      Kol_Ed_S:= Form1.ADOQuery2.FieldByName('����������').AsString;
                       { PosTrud:=Pos('.',Kol_Ed_S);
                        if PosTrud<>0 Then
                        Begin
                                Delete(Kol_Ed_S,PosTrud,1);
                                Insert(',',Kol_Ed_S,PosTrud);
                        End;}
                        Kol_Ed:= StrToFloat(Kol_Ed_S);
                      Kol_Oboz:= StrToFloat(Kol_Ed_S);
                      //Trumph,Nog,Ugl,Gib,Prok,Pila
                      Dlina:= Form1.ADOQuery2.FieldByName('������').AsString;
                      Shir:= Form1.ADOQuery2.FieldByName('�����').AsString;
                      DlinRaz:= Form1.ADOQuery2.FieldByName('����������').AsString;
                      ShirRaz:= Form1.ADOQuery2.FieldByName('���������').AsString;
                      if DlinRaz ='' Then
                            DlinRaz:='0';
                      if ShirRaz ='' Then
                            ShirRaz:='0';
                      Kol_Gib:= Form1.ADOQuery2.FieldByName('��������').AsString;
                      if Kol_Gib=''  Then
                      KolGib:=0
                      Else
                      KolGib:=StrToInt(Kol_Gib);

                      Mat:= Form1.ADOQuery2.FieldByName('��������').AsString;
                      if Mat='' Then
                                        Mat:='0';
                      EI:= Form1.ADOQuery2.FieldByName('��').AsString;

                      Kanban:= Form1.ADOQuery2.FieldByName('������').AsBoolean;
                      Trumph:= Form1.ADOQuery2.FieldByName('Trumph').AsBoolean;
                      Nog:= Form1.ADOQuery2.FieldByName('�������').AsBoolean;
                      Ugl:= Form1.ADOQuery2.FieldByName('�������').AsBoolean;
                      Gib:= Form1.ADOQuery2.FieldByName('�����').AsBoolean;
                      Prok:= Form1.ADOQuery2.FieldByName('��������').AsBoolean;
                      Pila:= Form1.ADOQuery2.FieldByName('����').AsBoolean;
                      PilaLent:= Form1.ADOQuery2.FieldByName('���� ���������').AsBoolean;

                      SH1:= Form1.ADOQuery2.FieldByName('�����������1').AsString;
                      SH2:= Form1.ADOQuery2.FieldByName('�����������2').AsString;
                      //==================================================


                      Res_Mat:=Pos('���������',VELem);
                      Res_Proch:=Pos('������ �������',VELem);
                      Res_Stand:=Pos('����������� �������',VELem);
                      Res_Sbor:=Pos('��������� �������',VELem);
                      Res_Klap:=Pos('������',ELem);
                      Res_Kompl_priv:=Pos('�������� �������',ELem);
                      Os:= Pos('���',ELem);
                      hh:=A MOD 50;
                      gg:=B MOD 50;

                      //XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                      Res_Detal:=AnsiCompareStr('������',VElem);
                      OboznSh:=Oboz;
                      Res:=Pos('-',OboznSh);
                      if Res<>0 Then                      //   AND ((Sten_G<>0) ANd (Sten_V<>0))
                        Delete(OboznSh,Res,30);
                      if (Res_Detal=0) Then
                      Begin
                        if not Form1.mkQuerySelect66(Form1.ADOQuery4,
                                'Select * from %s Where   (������������=' + #39 +
                                OboznSh + #39 +') ', ['������']) then
                                exit;
                        Kanban:= Form1.ADOQuery4.FieldByName('������').AsBoolean;
                        Trumph:= Form1.ADOQuery4.FieldByName('Trumph').AsBoolean;
                        Nog:= Form1.ADOQuery4.FieldByName('�������').AsBoolean;
                        Ugl:= Form1.ADOQuery4.FieldByName('�������').AsBoolean;
                        Gib:= Form1.ADOQuery4.FieldByName('�����').AsBoolean;
                        Prok:= Form1.ADOQuery4.FieldByName('��������').AsBoolean;
                        Pila:= Form1.ADOQuery4.FieldByName('����').AsBoolean;
                        PilaLent:= Form1.ADOQuery4.FieldByName('���� ���������').AsBoolean;
                        Svarka:=Form1.ADOQuery4.FieldByName('������').AsBoolean;
                        Ugol:= Form1.ADOQuery4.FieldByName('����').AsInteger;
                        Kol_Gib:= Form1.ADOQuery4.FieldByName('��������').AsString;
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
                      Sten_V:=Pos('������ ������������',Elem);
                      Sten_G:=Pos('������ ��������������',Elem);
                      
                      //gg:=B MOD 50;
                      F2:=Pos('2*�',Izdel);
                      if (Res_KPU<>0) AND (F2<>0) AND (Res_Detal=0)  Then
                      Begin
                         if  ((Sten_V<>0) OR (Sten_G<>0)) then
                         begin
                                Dlina:= Form1.ADOQuery2.FieldByName('�����').AsString;
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
                                Dlina:= Form1.ADOQuery2.FieldByName('�����').AsString;
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
                                        ', �����='+#39+ 'True' + #39 +
                                        ', ������='+#39+ 'False' + #39 +
                                        ' WHERE (�����������=' + #39 + Oboz + #39 +') AND ([Id��]=' + #39 + IDGP + #39 +
                                                ') ', ['������']) then
                                        Exit;

                          end;

                       end;
        //SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS    STENKI
                      end;
                     Res_N2:=AnsiCompareStr('2�',N2);
                     Res_N1:=AnsiCompareStr('1�',N2);
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Trumph=True) Then
                      Begin
                      Kol_Ed:= Form1.ADOQuery2.FieldByName('����������').AsFloat;
                      Kol_Oboz:= Form1.ADOQuery2.FieldByName('����������').AsFloat;
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
                        if (Res_KPU=0) AND (Res_KPD=0) Then  //������
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
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++ �������
                      if (Kanban=False) AND (Nog=True) Then
                      Begin
                      Kol_Ed:= Form1.ADOQuery2.FieldByName('����������').AsFloat;
                      Kol_Oboz:= Form1.ADOQuery2.FieldByName('����������').AsFloat;
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
                                //���� ����� 1,0�600�1500 �63 ���� 931-90
                                 Res_List:=Pos('���� ����� 1,0�600�1500 �63 ���� 931-90',Mat);
                                 if Res_List<>0 Then
                                 Mat:='1';
                                 //����� �� ���������� ���� 0,3
                                 Res_List:=Pos('����',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                Res_List:=Pos('�',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// �����
                                                Delete(Mat,1,Res_Tol);
                                                //
                                                Res_Tol:=Pos(' ',Mat);//��
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('�',Mat);//x
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
                                        Res_List:=Pos('��',Mat);
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
                                'Select * from %s Where   ([�]=' + #39 +'1'+ #39 +') And ([�������]='+Tol_S+') And ([�����]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([������]=' + #39 +FloatToStr(ShirI)+ #39 +')', ['[����� �����]']) then
                                exit;
                                NC_S:= (Form1.ADOQuery1.FieldByName('�����').AsString);
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
                                Res_Lop:=AnsiCompareStr('������ �������',SD4.Cells[3,h]);
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
                                'Select * from %s Where   ([�]=' + #39 +'11'+ #39 +') And ([�������]='+Tol_S+') And ([�����]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([������]=' + #39 +FloatToStr(ShirI)+ #39 +')', ['[����� �����]']) then
                                exit;
                                NC_S2:= (Form1.ADOQuery1.FieldByName('�����').AsString);
                                if NC_S2='' Then
                                        NC_S2:='0';
                                if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                'Select * from %s Where   ([�]=' + #39 +'12'+ #39 +') ', ['[����� �����]']) then
                                exit;
                                if C>200 Then
                                Razmet_KPU:=Form1.ADOQuery1.FieldByName('�����').AsFloat
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
                                        'UPDATE %s SET [�\� �������]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                        ' WHERE ([Id��]=' + #39 + IDGP + #39 +
                                                ') AND (�������=' + #39 + Elem + #39 +') AND (�����������=' + #39 + Oboz + #39 +')  ', ['������']) then
                                        Exit;
                                Inc(h);
                                SD4.RowCount:= SD4.RowCount+1;
                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Gib=True) Then
                      Begin
                      Kol_Ed:= Form1.ADOQuery2.FieldByName('����������').AsFloat;
                      Kol_Oboz:= Form1.ADOQuery2.FieldByName('����������').AsFloat;
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
                                Res_Lop:=AnsiCompareStr('������ �������',Elem);
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
                                 Res_List:=Pos('����',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                        Res_List:=Pos('�',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// �����
                                                Delete(Mat,1,Res_Tol);
                                                //
                                                Res_Tol:=Pos(' ',Mat);//��
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('�',Mat);//x
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
                                        Res_List:=Pos('��',Mat);
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
                                        'Select * from %s Where   ([�]=' + #39 +'3'+ #39 +') And ([�������]=' + #39 +(Tol_s)+ #39 +
                                        ') And ([�����]=' + #39 +FloatToStr(DlinI)+ #39 +') And ([������]=' + #39 +IntToStr(KolGib)+ #39 +')', ['[����� �����]']) then
                                        exit;
                                        NC_S:= (Form1.ADOQuery1.FieldByName('�����').AsString);
                                        If NC_S='' Then
                                        NC:=0 Else
                                        NC:=StrToFloat(NC_S)*Kol_Ed;                //*Kol_Zap
                                        NC_GIB:=NC_GIB+NC;
                                        SD5.Cells[14,n]:=FloatToStr(NC);
                                        SD5.Cells[19,n]:=NC_S;
                                end;
                                //======================
                                Res_Resh:=Pos('������� ����������',Elem);
                                if Res_Resh<>0 Then
                                Begin
                                       if not Form1.mkQuerySelect1(Form1.ADOQuery1,
                                        'Select * from %s Where   ([�]=' + #39 +'9'+ #39 +') And ([�������]=' + #39 +IntToStr(KolGib)+ #39 +
                                        ')', ['[����� �����]']) then
                                        exit;
                                        NC_S:= (Form1.ADOQuery1.FieldByName('������').AsString);
                                        If NC_S='' Then
                                                NC:=0
                                        Else
                                                NC:=StrToFloat(NC_S)*Kol_Ed;
                                        //++++++++++++
                                        NC_S2:= (Form1.ADOQuery1.FieldByName('�����').AsString);
                                        If NC_S2='' Then
                                                NC2:=0
                                        Else
                                                NC2:=StrToFloat(NC_S2)*Kol_Ed;    //*Kol_Zap /Kol_Zap
                                        NC_GIB:=NC_GIB+NC+NC2;
                                        SD5.Cells[19,n]:=FloatToStr(StrToFloat(NC_S)+StrToFloat(NC_S2));
                                        SD5.Cells[14,n]:=FloatToStr((NC)+(NC2));
                                end;
                                Lenta:=Pos('����� ��������������',Elem);
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
                                        'Select * from %s Where   ([�]=' + #39 +'12'+ #39 +') And ([������]=' + #39 +FloatToStr(DlinI)+ #39 +
                                        ')', ['[����� �����]']) then
                                        exit;
                                      Kol_otv:= (Form1.ADOQuery1.FieldByName('�����').AsInteger);
                                      A_Lenta:=(B-1.5-(Kol_Otv-1)*150)/2;
                                      SD5.Cells[15,n]:=FloatToStr(DlinD);//� �������-1,5
                                      SD5.Cells[16,n]:= FloatToStr(A_Lenta);
                                      SD5.Cells[17,n]:=IntToStr(Kol_Otv);
                                      SD5.Cells[18,n]:='150';      //*Kol_Zap   *Kol_Zap
                                      T:=0.008*Kol_Otv*Kol_Ed+0.008*Kol_Ed; //�������� ��������� � ����� �����
                                      //0.05 ����� ���������� ����� � ������ ���� ��� �� �� ���� ��� �����
                                      NC_GIB:=NC_GIB+T;
                                      SD5.Cells[20,n]:= FloatToStr(T);
                                end;
                                if SD5.Cells[20,n]='' Then
                                NCB:=(StrToFloat(SD5.Cells[14,n])+0)
                                Else NCB:=(StrToFloat(SD5.Cells[14,n])+StrToFloat(SD5.Cells[20,n]));

                                if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [�\� �����]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NCB)) + #39 +
                                        ' WHERE ([Id��]=' + #39 + IDGP + #39 +
                                                ') AND (�������=' + #39 + Elem + #39 +') AND (�����������=' + #39 + Oboz + #39 +')  ', ['������']) then
                                        Exit;
                                Inc(n);
                                SD5.RowCount:= SD5.RowCount+1;
                     end;
                      //+++++++++++++++++++++++++++++++++++++++++++++++++++++
                      if (Kanban=False) AND (Pila=True) Then
                      Begin
                        Kol_Ed:= Form1.ADOQuery2.FieldByName('����������').AsFloat;
                        Kol_Oboz:= Form1.ADOQuery2.FieldByName('����������').AsFloat;
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
                                 Res_List:=Pos('����',Mat);
                                 if Res_List<>0 Then
                                 Delete(Mat,1,Res_List+4);
                                        Res_List:=Pos('�',Mat);
                                        if Res_List<>0 Then
                                        Begin
                                                Res_Tol:=Pos(' ',Mat);// �����
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//��
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos(' ',Mat);//08pc
                                                Delete(Mat,1,Res_Tol);

                                                Res_Tol:=Pos('�',Mat);//x
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
                                        Res_List:=Pos('��',Mat);
                                        if Res_List<>0 Then
                                        Delete(Mat,1,2);
                                        if Mat='' Then
                                        Mat:='0';
                                        //++++++++++++++++++++++++++++++++==
                                        Res_Detal:=AnsiCompareStr('������',VElem);
                                        Res_Lop:=Pos('�������',Elem);
                                        Res_Ger:=Pos('������',Izdel);
                                        R25:=Pos('�25',Izdel);
                                        Res_List:=Pos('20001',Mat);//������� �25
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
                                          'Select * from %s Where   ([�]=' + #39 +'6'+ #39 +') And  ([�����]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[����� �����]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('�����').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S)+0.02*Kol_Ed;      //  *Kol_Zap  *Kol_Zap
                                          NC_PILA:=NC_PILA+NC;
                                          SD6.Cells[14,w]:=FloatToStr(NC);
                                          if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                                'UPDATE %s SET [�\� ����]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                                ' WHERE ([Id��]=' + #39 + IDGP + #39 +
                                                ') AND (�������=' + #39 + Elem + #39 +') AND (�����������=' + #39 + Oboz + #39 +')  ', ['������']) then
                                                Exit;
                                          Inc(w);
                                                SD6.RowCount:= SD6.RowCount+1;
                                        End;
                                        Ugol:=Pos('������',Elem);
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
                                          'Select * from %s Where   ([�]=' + #39 +'7'+ #39 +') And  ([�����]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[����� �����]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('�����').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S)+0.007*Kol_Ed;   //  *Kol_Zap *Kol_Zap
                                          SD6.Cells[14,w]:=FloatToStr(NC);
                                          NC_PILA:=NC_PILA+NC;
                                          if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                                'UPDATE %s SET [�\� ����]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                                ' WHERE ([Id��]=' + #39 + IDGP + #39 +
                                                ') AND (�������=' + #39 + Elem + #39 +') AND (�����������=' + #39 + Oboz + #39 +')  ', ['������']) then
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
                                          'Select * from %s Where   ([�]=' + #39 +'8'+ #39 +') And  ([�����]=' + #39 +FloatToStr(DlinI)+ #39 +')', ['[����� �����]']) then
                                          exit;
                                          NC_S:= (Form1.ADOQuery1.FieldByName('�����').AsString);
                                          If NC_S='' Then
                                          NC:=0 Else
                                          NC:=StrToFloat(NC_S)*Kol_Ed;             // *Kol_Zap *Kol_Zap
                                          SD6.Cells[14,w]:=FloatToStr(NC);
                                          NC_PILA:=NC_PILA+NC;
                                          if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                                        'UPDATE %s SET [�\� ����]=' + #39
                                                + Form1.ConvertFloat(FloatToStr(NC)) + #39 +
                                        ' WHERE ([Id��]=' + #39 + IDGP + #39 +
                                                ') AND (�������=' + #39 + Elem + #39 +') AND (�����������=' + #39 + Oboz + #39 +')  ', ['������']) then
                                        Exit;
                                          Inc(w);
                                                SD6.RowCount:= SD6.RowCount+1;

                                        End;
                      end;

                      Form1.ADOQuery2.Next;

                 end;
                        if not Form1.mkQueryUpdate2(Form1.ADOQuery4,
                        'UPDATE %s SET [�������]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_NOG)) + #39 +
                        ',[����]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_PILA)) + #39 +
                        ',[�����]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_GIB)) + #39 +
                        ',[�����������]=' + #39
                        + '1'+ #39 +',[�������� �����]='+ #39+DatStr+#39+
                        ' WHERE ([Id��]=' + #39 + IDGP + #39 +
                        ') ', ['Klapana']) then
                        Exit;

                        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET '+
                        '[�������]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_NOG)) + #39 +
                        ',[����]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_PILA)) + #39 +
                        ',[�����]=' + #39
                        + Form1.ConvertFloat(FloatToStr(NC_GIB)) + #39 +
                        ',[�����������]=' + #39
                        + '1'+ #39 +',[�������� �����]='+ #39+DatStr+#39+
                        ' WHERE ([Id��]=' + #39 + IDGP + #39 +')', ['������']) then
                        Exit;
                End;
                ProgressBar1.Position:=I;
        End;
        ProgressBar1.Position:=0;
        Close;

end;

end.
