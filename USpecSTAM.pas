unit USpecSTAM;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, Grids, Menus,Clipbrd;

type
  TFSTAM = class(TForm)
    SpecGrid: TStringGrid;
    cbb1: TComboBox;
    mmo1: TMemo;
    mmo2: TMemo;
    lbl1: TLabel;
    pnl1: TPanel;
    lbl2: TLabel;
    lbl3: TLabel;
    btn1: TSpeedButton;
    lbl4: TLabel;
    lbl5: TLabel;
    edt1: TEdit;
    btn2: TButton;
    pm1: TPopupMenu;
    N1: TMenuItem;
    Label3: TLabel;
    Memo2: TMemo;
    Memo1: TMemo;
    procedure FormShow(Sender: TObject);
    procedure SpecGridDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure cbb1Click(Sender: TObject);
    procedure SpecGridMouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure SpecGridMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure SpecGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure N1Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    AColG,ARowG,TmpCol,TmpRow :Integer;
    Shablon:Integer;
  end;

var
  FSTAM: TFSTAM;

implementation

uses Main, UNewNakl;

{$R *.dfm}

procedure TFSTAM.FormShow(Sender: TObject);
var I:Integer;
S:string;
begin
FNewNakl.Clear_StringGrid(SpecGrid);
S := ExtractFileDir(ParamStr(0));
        Mmo1.Lines.Clear;
        Mmo1.Lines.LoadFromFile(S + '\SpecGrid.ini');
        SpecGrid.ColCount := Mmo1.Lines.Count;
        for I := 0 to Mmo1.Lines.Count - 1 do
                SpecGrid.ColWidths[i] := StrToInt(Mmo1.Lines.Strings[i]);
if Form1.Luk=2 then
  begin
       { if Form1.Spec=1 Then
        begin
           if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s ', ['����������']) then
                exit;
        end; }
        if Form1.Spec=0 Then
        begin
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where [Id��]=' +
                #39+Lbl1.Caption+#39+' Order By �����������,�������', ['������750']) then
                exit;
        end;
  end;
  if Form1.Luk=1 then
  begin
       { if Form1.Spec=1 Then
        begin
           if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s ', ['����������']) then
                exit;
        end; }
        if Form1.Spec=0 Then
        begin
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where [Id��]=' +
                #39+Lbl1.Caption+#39+' Order By �����������,�������', ['���������']) then
                exit;
        end;
  end;
  if Form1.Luk=0 then
  begin
        if Form1.Spec=0 Then
        begin
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where [Id��]=' +
                #39+Lbl1.Caption+#39+' Order By �����������,�������', ['����������']) then
                exit;
        end;
  end;
        SpecGrid.ColCount := I_FN_NOM1 + 16;
        SpecGrid.Cells[0, 0] := 'Id';

        SpecGrid.Cells[I_FN_ZAK1, 0] := '��������' ;

       // SpecGrid.Cells[I_FN_ZAK1, 0] := FN_ZAK1;
        SpecGrid.Cells[I_FN_POS_KOMP, 0] := '�����������';
        SpecGrid.Cells[I_FN_POSS, 0] := '�������';
        SpecGrid.Cells[I_FN_OBOZN, 0] := '�������';

        SpecGrid.Cells[I_FN_TYP_S, 0] := '����������';
        SpecGrid.Cells[I_FN_NOM_NOM, 0] := '��';

        SpecGrid.Cells[I_FN_ZBOR_ED, 0] := '�����������';
        SpecGrid.Cells[I_FN_MODEL, 0] := '�������';
        SpecGrid.Cells[I_FN_KATEG, 0] := '�����';
        SpecGrid.Cells[I_FN_KOL_S, 0] := '�����';
        SpecGrid.Cells[I_FN_SPRAM, 0] := '������';
        SpecGrid.Cells[I_FN_UKRUG, 0] := '���������';
        SpecGrid.Cells[I_FN_KPD, 0] := '����������';
        SpecGrid.Cells[I_FN_KOMPL_W, 0] := '�����';
        SpecGrid.Cells[I_FN_NOM_GUR, 0] := '��������';
        SpecGrid.Cells[I_FN_NOM1, 0] := '������';
        SpecGrid.Cells[I_FN_NOM1+1, 0] := '��������';
        SpecGrid.Cells[I_FN_NOM1+2, 0] := 'Trumph';
        SpecGrid.Cells[I_FN_NOM1+3, 0] := '�������';
        SpecGrid.Cells[I_FN_NOM1+4, 0] := '�������';
        SpecGrid.Cells[I_FN_NOM1+5, 0] := '����';
        SpecGrid.Cells[I_FN_NOM1+6, 0] := '�����';
        SpecGrid.Cells[I_FN_NOM1+7, 0] := '��������';
        SpecGrid.Cells[I_FN_NOM1+8, 0] := '����';
        SpecGrid.Cells[I_FN_NOM1+9, 0] := '���� ���������';
        SpecGrid.Cells[I_FN_NOM1 + 10, 0] := '������';
        SpecGrid.Cells[I_FN_NOM1+11, 0] := 'Id��';
        SpecGrid.Cells[I_FN_NOM1 + 12, 0] := '�\� �����';
        SpecGrid.Cells[I_FN_NOM1 + 13, 0] := '�\� �����';
        SpecGrid.Cells[I_FN_NOM1 + 14, 0] := '����';
        SpecGrid.Cells[I_FN_NOM1 + 15, 0] := 'id��';
        //SpecGrid.ColCount := I_FN_NOM1 + 3;
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
                SpecGrid.Cells[0, I + 1] := Form1.ADOQuery1.FieldByName('Id').AsString;
               if Form1.Luk=2 then
                SpecGrid.Cells[I_FN_ZAK1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Parent').AsString
                Else
                SpecGrid.Cells[I_FN_ZAK1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��������').AsString;

                //SpecGrid.Cells[I_FN_ZAK1, I + 1] :=
               //         Form1.ADOQuery1.FieldByName(FN_ZAK1).AsString;
                SpecGrid.Cells[I_FN_POS_KOMP, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����������').AsString;
                SpecGrid.Cells[I_FN_POSS, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SpecGrid.Cells[I_FN_OBOZN, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;

                SpecGrid.Cells[I_FN_TYP_S, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����������').AsString;
                SpecGrid.Cells[I_FN_NOM_NOM, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��').AsString;

               SpecGrid.Cells[I_FN_ZBOR_ED, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����������').AsString;
                SpecGrid.Cells[I_FN_MODEL, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SpecGrid.Cells[I_FN_KATEG, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����').AsString;
                SpecGrid.Cells[I_FN_KOL_S, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����').AsString;
                SpecGrid.Cells[I_FN_SPRAM, I + 1] :=
                        Form1.ADOQuery1.FieldByName('������').AsString;
                SpecGrid.Cells[I_FN_UKRUG, I + 1] :=
                        Form1.ADOQuery1.FieldByName('���������').AsString;
                SpecGrid.Cells[I_FN_KPD, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����������').AsString;
                SpecGrid.Cells[I_FN_KOMPL_W, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����').AsString;
                SpecGrid.Cells[I_FN_NOM_GUR, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��������').AsString;
                SpecGrid.Cells[I_FN_NOM1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 2, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Trumph').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 3, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 4, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 5, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����').AsString;

                SpecGrid.Cells[I_FN_NOM1 + 6, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 7, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 8, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 9, I + 1] :=
                        Form1.ADOQuery1.FieldByName('���� ���������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 10, I + 1] :=
                        Form1.ADOQuery1.FieldByName('������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 11, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Id��').AsString;

                SpecGrid.Cells[I_FN_NOM1 + 12, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�\� �����').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 13, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�\� �������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 14, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 15, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Id��').AsString;
                Form1.ADOQuery1.Next;
        end;
        if Form1.Luk=2 then
  begin
       { if Form1.Spec=1 Then
        begin
           if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s ', ['����������']) then
                exit;
        end; }
        Memo1.Lines.Clear;
        if Form1.Spec=0 Then
        begin
                S:='������� LIKE '+#39+'%����������� �������������%'+#39;
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where [Id��]=' +
                #39+Lbl1.Caption+#39+' AND (%s) Order By �����������,�������', ['������750',S]) then
                exit;
                Memo1.Lines.Add( Form1.ADOQuery1.FieldByName('�������').AsString);
            Memo1.Lines.Add(Form1.ADOQuery1.FieldByName('����������').AsString);
            Memo1.Lines.Add(Form1.ADOQuery1.FieldByName('�����������').AsString);
            Form1.ADOQuery1.Next;
        end;
  end;
end;

procedure TFSTAM.SpecGridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
  Var Str:String;
  res:Integer;
begin
        Str:=SpecGrid.Cells[I_FN_ZAK1, ARow];
        Res:=AnsiCompareStr('TRUE',Str);
        if  Res=0 then
        begin
                                SpecGrid.canvas.brush.Color := RGB(100, 149,
                                        237); //������
                                SpecGrid.Canvas.Font.Color := clBlack;
                                SpecGrid.Canvas.FillRect(Rect);
                                SpecGrid.Canvas.TextOut(Rect.Left, Rect.Top,
                                SpecGrid.Cells[ACol, ARow]);
        end;
end;

procedure TFSTAM.cbb1Click(Sender: TObject);
begin
        if Form1.Spec=1 Then
        Begin

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' + Label3.Caption + ']=' +
                #39 + Cbb1.Text+ #39 +
                ' WHERE ([������������]=' + #39 + SpecGrid.Cells[I_FN_ZBOR_ED,ARowG] + #39+')',
                ['����������']) then
                Exit;
                SpecGrid.Cells[AColG,SpecGrid.Row]:=Cbb1.Items[Cbb1.ItemIndex];
        end;
end;

procedure TFSTAM.SpecGridMouseWheelDown(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
        Handled := true;
        SpecGrid.Perform(WM_VSCROLL, SB_LINEDOWN, 0);
end;

procedure TFSTAM.SpecGridMouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
        Handled := true;
        SpecGrid.Perform(WM_VSCROLL, SB_LINEUP, 0);
end;

procedure TFSTAM.SpecGridSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
  Var R,Rect1:TRect;
begin
                AColG:=ACol;
        ARowG:=ARow;
        //StringGrid2.LeftCol:=0 ;
        SpecGrid.Selection:=TGridRect(Rect(AColG, ARow, ARowG, ARow));
        Rect1:=SpecGrid.CellRect(ACol,ARow);
        TmpCol:=ACol;
        TmpRow:=ARow;
        if (Form1.FlagDolg=1) OR (Form1.FlagDolg=2) Then
        Begin
        AColG:=ACol;
        ARowG:=ARow;
        //StringGrid2.LeftCol:=0 ;
        SpecGrid.Selection:=TGridRect(Rect(AColG, ARow, ARowG, ARow));
        Rect1:=SpecGrid.CellRect(ACol,ARow);
        TmpCol:=ACol;
        TmpRow:=ARow;

Cbb1.Visible:=False;
  if ((ACol = I_FN_NOM1) OR (ACol = I_FN_NOM1+2) OR (ACol = I_FN_NOM1+3) OR (ACol = I_FN_NOM1+4) OR (ACol = I_FN_NOM1+6) OR (ACol = I_FN_NOM1+7) OR (ACol = I_FN_NOM1+8) OR (ACol = I_FN_NOM1+9) OR (ACol = I_FN_NOM1+10) OR ((ACol=I_FN_ZAK1) AND (Form1.Spec=1))  ) and (ARow <> 0) then   //����
  begin
    {������ � ��������� ComboBox ������ ��������������� ������ StringGrid}
    R := SpecGrid.CellRect(ACol, ARow);
    R.Left := R.Left + SpecGrid.Left;
    R.Right := R.Right + SpecGrid.Left;
    R.Top := R.Top + SpecGrid.Top;
    R.Bottom := R.Bottom + SpecGrid.Top;
    Cbb1.Left := R.Left + 1;
    Cbb1.Top := R.Top + 1; Cbb1.Width := (R.Right + 1) - R.Left;
    Cbb1.Height := (R.Bottom + 1) - R.Top; {������� combobox}
    Cbb1.Visible := True;
    Label3.caption:=SpecGrid.Cells[ACol,0];
    //ComboBox10.SetFocus;
  end;

  end;
end;

procedure TFSTAM.N1Click(Sender: TObject);
Var Str1:String;
begin
        Str1 := SpecGrid.Cells[TmpCol, tmpRow];
        Memo2.Lines.Clear;
        Memo2.Lines.Add(Str1);
        Clipboard.AsText := Memo2.Lines.Strings[0];
end;

procedure TFSTAM.btn1Click(Sender: TObject);
begin
Form1.ExportGridtoExcel1(SpecGrid);
end;

procedure TFSTAM.btn2Click(Sender: TObject);
var I:Integer;
begin
   if Form1.Luk=1 then
  begin

        {if Form1.Spec=1 Then
        begin
           if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where ������������ = '+#39+Edt1.Text+#39+' Order By ��������', ['����������']) then
                exit;
        end; }
        if Form1.Spec=0 Then
        begin
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where ����������� = '+#39+Edt1.Text+#39+' Order By ��������', ['���������']) then
                exit;
        end;
  end;
  if Form1.Luk=0 then
  begin

        if Form1.Spec=1 Then
        begin
           if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where ������������ = '+#39+Edt1.Text+#39+' Order By ��������', ['����������']) then
                exit;
        end;
        if Form1.Spec=0 Then
        begin
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where ����������� = '+#39+Edt1.Text+#39+' Order By ��������', ['����������']) then
                exit;
        end;
  end;
        SpecGrid.ColCount := I_FN_NOM1 + 16;
        SpecGrid.Cells[0, 0] := '�';
        if Form1.Spec=1 Then
        SpecGrid.Cells[I_FN_ZAK1, 0] := '��������'
        Else
        SpecGrid.Cells[I_FN_ZAK1, 0] := FN_ZAK1;
        SpecGrid.Cells[I_FN_POS_KOMP, 0] := '�����������';
        SpecGrid.Cells[I_FN_POSS, 0] := '�������';
        SpecGrid.Cells[I_FN_OBOZN, 0] := '�������';

        SpecGrid.Cells[I_FN_TYP_S, 0] := '����������';
        SpecGrid.Cells[I_FN_NOM_NOM, 0] := '��';

        SpecGrid.Cells[I_FN_ZBOR_ED, 0] := '�����������';
        SpecGrid.Cells[I_FN_MODEL, 0] := '�������';
        SpecGrid.Cells[I_FN_KATEG, 0] := '�����';
        SpecGrid.Cells[I_FN_KOL_S, 0] := '�����';
        SpecGrid.Cells[I_FN_SPRAM, 0] := '������';
        SpecGrid.Cells[I_FN_UKRUG, 0] := '���������';
        SpecGrid.Cells[I_FN_KPD, 0] := '����������';
        SpecGrid.Cells[I_FN_KOMPL_W, 0] := '�����';
        SpecGrid.Cells[I_FN_NOM_GUR, 0] := '��������';
        SpecGrid.Cells[I_FN_NOM1, 0] := '������';
        SpecGrid.Cells[I_FN_NOM1+1, 0] := '��������';
        SpecGrid.Cells[I_FN_NOM1+2, 0] := 'Trumph';
        SpecGrid.Cells[I_FN_NOM1+3, 0] := '�������';
        SpecGrid.Cells[I_FN_NOM1+4, 0] := '�������';
        SpecGrid.Cells[I_FN_NOM1+5, 0] := '����';
        SpecGrid.Cells[I_FN_NOM1+6, 0] := '�����';
        SpecGrid.Cells[I_FN_NOM1+7, 0] := '��������';
        SpecGrid.Cells[I_FN_NOM1+8, 0] := '����';
        SpecGrid.Cells[I_FN_NOM1+9, 0] := '���� ���������';
        SpecGrid.Cells[I_FN_NOM1 + 10, 0] := '������';
        SpecGrid.Cells[I_FN_NOM1+11, 0] := 'Id��';
        SpecGrid.Cells[I_FN_NOM1 + 12, 0] := '�\� �����';
        SpecGrid.Cells[I_FN_NOM1 + 13, 0] := '�\� �����';
        SpecGrid.Cells[I_FN_NOM1 + 14, 0] := '����';
        SpecGrid.Cells[I_FN_NOM1 + 15, 0] := 'id��';
        //SpecGrid.ColCount := I_FN_NOM1 + 3;
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
                SpecGrid.Cells[0, I + 1] := IntToStr(I + 1);
                if Form1.Spec=1 Then

                SpecGrid.Cells[I_FN_ZAK1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��������').AsString
                Else
                SpecGrid.Cells[I_FN_ZAK1, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZAK1).AsString;
                SpecGrid.Cells[I_FN_POS_KOMP, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����������').AsString;
                SpecGrid.Cells[I_FN_POSS, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SpecGrid.Cells[I_FN_OBOZN, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;

                SpecGrid.Cells[I_FN_TYP_S, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����������').AsString;
                SpecGrid.Cells[I_FN_NOM_NOM, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��').AsString;
{[�����������], [�������], [�������], [����������], [��], [�����������],'+
            '[���], [���������], [����������], [�������], [�����], [�����],   [������], [���������],'+
            '[����������], [�����], [��������], [������], [��������], [��������], [Trumph], [�������],'+
            '[�������], [�����], [��������], [����]
            }
                if Form1.Spec=1 Then
                SpecGrid.Cells[I_FN_ZBOR_ED, I + 1] :=
                        Form1.ADOQuery1.FieldByName('������������').AsString
               Else
               SpecGrid.Cells[I_FN_ZBOR_ED, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����������').AsString;
                SpecGrid.Cells[I_FN_MODEL, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SpecGrid.Cells[I_FN_KATEG, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����').AsString;
                SpecGrid.Cells[I_FN_KOL_S, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����').AsString;
                SpecGrid.Cells[I_FN_SPRAM, I + 1] :=
                        Form1.ADOQuery1.FieldByName('������').AsString;
                SpecGrid.Cells[I_FN_UKRUG, I + 1] :=
                        Form1.ADOQuery1.FieldByName('���������').AsString;
                SpecGrid.Cells[I_FN_KPD, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����������').AsString;
                SpecGrid.Cells[I_FN_KOMPL_W, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����').AsString;
                SpecGrid.Cells[I_FN_NOM_GUR, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��������').AsString;
                SpecGrid.Cells[I_FN_NOM1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 2, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Trumph').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 3, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 4, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 5, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����').AsString;

                SpecGrid.Cells[I_FN_NOM1 + 6, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�����').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 7, I + 1] :=
                        Form1.ADOQuery1.FieldByName('��������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 8, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 9, I + 1] :=
                        Form1.ADOQuery1.FieldByName('���� ���������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 10, I + 1] :=
                        Form1.ADOQuery1.FieldByName('������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 11, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Id��').AsString;

                SpecGrid.Cells[I_FN_NOM1 + 12, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�\� �����').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 13, I + 1] :=
                        Form1.ADOQuery1.FieldByName('�\� �������').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 14, I + 1] :=
                        Form1.ADOQuery1.FieldByName('����').AsString;
                SpecGrid.Cells[I_FN_NOM1 + 15, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Id��').AsString;
                Form1.ADOQuery1.Next;
        end;
end;

end.
