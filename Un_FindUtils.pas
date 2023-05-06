unit Un_FindUtils;

interface
uses Windows, SysUtils, classes, NiceGrid, StdCtrls, Graphics;
const
    FIND_PRIORITY_COLUMN = 0;
    FIND_PRIORITY_ROW = 1;
    FIND_COMPARE_ANYPART = 0;
    FIND_COMPARE_START = 1;
    FIND_COMPARE_FULL = 2;
    //������� ������
    FIND_IN_COLUMN = 0;
    FIND_IN_TABLE = 1;    
    
    CAPTION_ALL_COLUMN = '��� ������';
    CAPTION_COLUMN = '������� �';
type
    TFindGrid = class( TComponent )
    private
        FCheckRegister : Boolean;
        FCurMove, FMovePriority, FMethodCompare, FCurCol  : Integer;
        //������� ������
        FLocation : Integer;
        FGrid : TNiceGrid;
        FColumnNames : TStrings;
        //����������� ������ ������
        FFindStr : string;
        //������� ������
        FFindCount : integer;

        FtmpRowCount,               
        FtmpColCount,
        FtmpSoonGridRowCount,
        FtmpSoonGridColCount, 
        FPosRow,
        FPosCol, FSoonPosCol  : Integer;
        
        
        //������� ��������� � ������ ������
        procedure FindEditChange( Sender : TObject );
        //������� �� ������ "����"        
        procedure FindDownClick( Sender : TObject );
        //��� ������� �� ����� ����(�������� �������)
        procedure FindEditOnClick( Sender : TObject );
        //������� �� ������ "�����"
        procedure FindUpClick( Sender : TObject );
        //������� �� ������ "��������� ������"
        procedure FindParamClick( Sender : TObject );

        //����� ���� � ����� �������
        procedure _FindNext( const AStr : string );overload;
        //����� ���� �� ������ �������
        procedure _FindNext();overload;
        //����� �����
        procedure _FindPrev();
        //����� ����� ���������� ������
        procedure _ShowParamForm();
        //�������� �����
        procedure _NewFind( Const ANewStr : string );
        //��������� �������� �� �������
        function _ChangeSizeGrid : boolean;
        //��������� ������� �� ��������� ������
        function _ChangePos : boolean;
        
        //�������� ������� ������
        procedure __FindNext( const AStr : string );    
        //�������� ��� (��)���������� ��������
        function _BadResult : boolean;         
        function _ToResult( AidX, AidY : Integer ) : boolean;
    public
        FindEdit, FindDown, FindUp, FindParam : TObject;       
        constructor Create( AOwner : TComponent );override;
        Destructor Destroy;override;
        procedure ColStringstoList( SGrid : TNiceGrid; SCol : integer; SList : TStrings );
        procedure AssignGrid( AGrid : TNiceGrid );
        //������ �������� ����������
        procedure SetControl( AFindEdit,
                              AFindDown, 
                              AFindUp, 
                              AFindParam : TObject );
        function ColumnNames : TStrings;
//        procedure Find
        procedure FromParamForm( const
                                {find location}
                                InColumn : boolean;
                                const
                                {��������� ������� ��� �������}
                                AnValuePrior,
                                {����� ������ ����������}
                                AnValueMethodCompare : integer;
                                {��������� ��������}
                                AbCheckRegister : boolean );
        published
        property Location : integer read FLocation write FLocation ;//default FIND_IN_COLUMN;
        property MovePriority : integer read FMovePriority write FMovePriority ;//default FIND_PRIORITY_COLUMN;
        property MethodCompare : integer read FMethodCompare write FMethodCompare;// default FIND_COMPARE_ANYPART;
        property CheckRegister : boolean read FCheckRegister write FCheckRegister;// default false;
    end;
implementation

uses Frm_SetFindGrid;

//���� ��������� �� �����
constructor TFindGrid.Create( AOwner : TComponent );
begin
    inherited Create( AOwner );
    FColumnNames := nil;

    FindEdit  := nil;
    FindDown  := nil;
    FindUp    := nil;
    FindParam := nil;
    
end;

Destructor TFindGrid.Destroy;
begin
    if FColumnNames <> nil then
        FColumnNames.Destroy;
    inherited Destroy;
end;
procedure TFindGrid.ColStringstoList( SGrid : TNiceGrid; SCol : integer; SList : TStrings );
var
    Y : Integer;
begin
    SList.Clear;
    for Y := 0 to SGrid.RowCount - 1 do
        SList.AddObject( SGrid.Cells[ SCol, Y  ], SGrid.Objects[ SCol, Y  ] );
end;

procedure TFindGrid.AssignGrid( AGrid : TNiceGrid );
begin
    if FGrid = AGrid then exit;
    FGrid := AGrid;
    if ( FGrid <> nil ) and ( FGrid.ColCount > 0 ) then
        FCurCol := 0;            
end;

{$REGION '���������� �������'}

procedure TFindGrid.SetControl( AFindEdit,
                      AFindDown, 
                      AFindUp, 
                      AFindParam : TObject );
begin
    FindEdit  := AFindEdit;
    TEdit( FindEdit ).OnClick := FindEditOnClick;
    with TEdit( FindEdit ) do
    begin 
        Color := clRed;
        OnChange := FindEditChange;
    end;

    FindDown  := AFindDown;
        TButton( FindDown ).OnClick := FindDownClick;
    FindUp    := AFindUp;
        TButton( FindUp ).OnClick := FindUpClick;
    FindParam := AFindParam;
        TButton( FindParam ).OnClick := FindParamClick;
end;

procedure TFindGrid.FindEditChange( Sender : TObject );
begin
    _FindNext( TEdit( Sender ).Text );
end;

procedure TFindGrid.FindDownClick( Sender : TObject );
begin
    _FindNext();
    FGrid.SetFocus;
end;

procedure TFindGrid.FindEditOnClick( Sender : TObject );
begin
    TEdit( Sender ).SelectAll;
end;


procedure TFindGrid.FindUpClick( Sender : TObject );
begin
    _FindPrev();
end; 

procedure TFindGrid.FindParamClick( Sender : TObject );
begin
    _ShowParamForm();
end; 
                     
{$ENDREGION}

{$REGION '������� ������'}

procedure TFindGrid._FindNext( const AStr : string );
begin
    __FindNext( AStr );        
end;

procedure TFindGrid._FindNext();
begin
    _FindNext( FFindStr );
end; 

procedure TFindGrid._ShowParamForm();
var
   SetForm : TFrmSetFindGrid;
begin
    SetForm := TFrmSetFindGrid.Create( nil );
    SetForm.SetFinder( Self );
    SetForm.ShowModal;    
end; 
//�������� �����
procedure TFindGrid._NewFind( Const ANewStr : string );
begin
    FFindCount := -1;
    FFindStr := ANewStr;
    FtmpRowCount := FGrid.RowCount;
    FtmpColCount := FGrid.ColCount;
    FPosRow := FGrid.Row;
    FPosCol := FGrid.Col;           
    FCurCol := FGrid.Col;
end;

//��������� �������� �� �������
function TFindGrid._ChangeSizeGrid : boolean;
begin
    Result := ( FtmpRowCount <> FGrid.RowCount ) or 
              ( FtmpColCount <> FGrid.ColCount );
end;
//��������� ������� �� ��������� ������
function TFindGrid._ChangePos : boolean;
begin
    Result := ( FGrid.Row <> FPosRow ) or ( FGrid.Col <> FPosCol );
    if Result then
    begin
        FPosRow := FGrid.Row;
        FPosCol := FGrid.Col;
    end;
end;
//�� �������
function TFindGrid._BadResult : boolean;
begin
    Result := true;
    if FindEdit <> nil then 
    with TEdit( FindEdit ) do
        Color := clRed;
end;
//�������
function TFindGrid._ToResult( AidX, AidY : Integer ) : boolean;
begin
    Result := true;
    if FindEdit <> nil then 
    with TEdit( FindEdit ) do
        Color := clWhite;

    FGrid.Row := AidY;
    FGrid.Col := AidX;
    FGrid.EnsureVisible( FGrid.Col, FGrid.Row );
    _ChangePos;
end;

procedure TFindGrid.__FindNext( const AStr : string );
var
    sl : TStrings;
    GetStr : string;
    I : Integer;
    CountResult, IdX, idY, X, Y : Integer;
    procedure InitProc();
    begin
        idY := -1;
        sl := TStringList.Create;
        CountResult := -1;
    end;
    function CloseProc : boolean;
    begin
        Result := true;
        sl.Destroy;
    end;
begin
    if ( AStr = '' ) and _BadResult then exit;
    if ( AStr <> FFindStr ) or _ChangeSizeGrid or _ChangePos then
        _NewFind( AStr );
    InitProc;
    //����� ������ �� ������� �������
    if FLocation = FIND_IN_COLUMN then
    begin
        ColStringstoList( FGrid, FCurCol, sl );
//        sl.Assign( FGrid.Columns[ FCurCol ].Strings );
        //����� �������� � ������ �����������
        I := 0;
        while I < sl.count do
        begin
            GetStr := sl[ I ];
            if not CheckRegister then
                GetStr := LowerCase( GetStr );
            if ( ( FMethodCompare = FIND_COMPARE_FULL ) and ( GetStr = FFindStr ) ) or
            ( ( FMethodCompare = FIND_COMPARE_START ) and ( Pos( FFindStr, GetStr ) = 1 ) ) or
            ( ( FMethodCompare = FIND_COMPARE_ANYPART ) and ( Pos( FFindStr, GetStr ) > 0 ) )                    
            then
            begin
                inc( CountResult );
                IdY := I;
                if CountResult >= FFindCount + 1  then 
                begin
                    FFindCount := CountResult;
                    break;
                end;
            end;
            inc( I );                            
        end;
    end;
    if ( IdY < 0 ) and _BadResult and CloseProc then
        Exit
    else
    if ( IdY >= 0 ) and _ToResult( FCurCol, IdY ) and CloseProc then exit;
end;

procedure TFindGrid._FindPrev();
var
    sl : TStrings;
    I : Integer;
    GetStr : string;
    CountResult, IdY, IDX : Integer;
    procedure InitProc();
    begin
        idY := -1;
        sl := TStringList.Create;
        CountResult := -1;
    end;
    function CloseProc : boolean;
    begin
        Result := true;
        sl.Destroy;
    end;
begin
    if ( ( FFindStr = '' ) or ( FFindCount < 1 ) ) and _BadResult then exit;
    
    if _ChangeSizeGrid or _ChangePos then
    begin
        _NewFind( FFindStr );
        exit;
    end;

    
    InitProc;
    //����� ������ �� ������� �������
    if FLocation = FIND_IN_COLUMN then
    begin
        ColStringstoList( FGrid, FCurCol, sl );
        //sl.Assign( FGrid.Columns[ FCurCol ].Strings );
        //����� �������� � ������ �����������
        I := 0;
        while I < sl.count do
        begin
            GetStr := sl[ I ];
            if not CheckRegister then
                GetStr := LowerCase( GetStr );
       
            if ( ( FMethodCompare = FIND_COMPARE_FULL ) and ( GetStr = FFindStr ) ) or
            ( ( FMethodCompare = FIND_COMPARE_START ) and ( Pos( FFindStr, GetStr ) = 1 ) ) or
            ( ( FMethodCompare = FIND_COMPARE_ANYPART ) and ( Pos( FFindStr, GetStr ) > 0 ) )                    
            then
            begin
                inc( CountResult );
                IdY := I;
                if CountResult >= FFindCount - 1 then 
                begin
                    FFindCount := CountResult;
                    break;
                end;
            end;
            inc( I );                            
        end;
    end;
    if ( IdY < 0 ) and _BadResult and CloseProc then Exit
    else
    if ( IdY >= 0 ) and _ToResult( FCurCol, IdY ) and CloseProc then exit;
end;

{**********************************************}

{$ENDREGION}

{$REGION '�������������� �����������'}
function TFindGrid.ColumnNames : TStrings;
var
    I : integer;
    Str : string;
begin
    Result := FColumnNames; 
    if FColumnNames = nil then
        FColumnNames := TStringList.Create;
    FColumnNames.Clear;
    if ( FGrid = nil ) or ( FGrid.ColCount < 1 )  then exit;

    FColumnNames.Add( CAPTION_ALL_COLUMN );

    with FGrid do
    for I := 0 to ColCount - 1 do
    begin
        Str := Columns[ I ].Title;
        if Str = '' then
            Str := CAPTION_COLUMN + inttostr( I );
        FColumnNames.Add( Str );
    end;

end;

procedure TFindGrid.FromParamForm( const
                                {find location}
                                InColumn : boolean;
                                const
                                {��������� ������� ��� �������}
                                AnValuePrior,
                                {����� ������ ����������}
                                AnValueMethodCompare : integer;
                                {��������� ��������}
                                AbCheckRegister : boolean );
begin
    FLocation := integer( not InColumn );
    MovePriority := AnValuePrior;
    MethodCompare := AnValueMethodCompare;
    CheckRegister := AbCheckRegister;

end;
{$ENDREGION}
end.
