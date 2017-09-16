unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, ComObj,
  System.Variants, System.Classes, Vcl.Graphics, Vcl.Controls,
  Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.IdAntiFreeze, IdHTTP, Parsing, Vcl.Grids, AdvObj, BaseGrid,
  AdvGrid, IdGlobal, AdvPageControl, System.DateUtils, Vcl.Menus, frmctrllink,
  WordXP, AdvEdit, ShLwApi, AdvSplitter, ShellCtrls, AdvEdBtn,
  AdvDateTimePicker, Vcl.CheckLst, Vcl.AppEvnts, JvPanel, JvExControls,
  JvGIFCtrl, Vcl.OleCtrls, SqlExpr, Data.DB, acShellCtrls, AdvAppStyler,
  Data.DbxSqlite, Data.FMTBcd, IdAntiFreezeBase, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, sTreeView, sListView, JvExExtCtrls,
  JvExtComponent, JvAnimatedImage;

type
  TMainForm = class(TForm)
    idhtp1: TIdHTTP;
    idntfrz1: TIdAntiFreeze;
    advpgcntrl1: TAdvPageControl;
    WordSheet: TAdvTabSheet;
    SSWebSheet: TAdvTabSheet;
    pmGridContract: TPopupMenu;
    GetReqmniC1: TMenuItem;
    CreateContractmniN1: TMenuItem;
    pm2grid1: TPopupMenu;
    mniMove: TMenuItem;
    pmCreateOrderForm: TPopupMenu;
    mnia1CreateOrderForm: TMenuItem;
    pmGridListOfServices: TPopupMenu;
    mniAddRow: TMenuItem;
    mniDelRow: TMenuItem;
    pnl2: TPanel;
    lbl3: TLabel;
    GridOrderForm: TAdvStringGrid;
    lblListOfServices: TLabel;
    GridListOfServices: TAdvStringGrid;
    GridContract: TAdvStringGrid;
    lbl2: TLabel;
    edtlnkPrice: TFormControlEditLink;
    mniN1CreateOrderForm: TMenuItem;
    N1: TMenuItem;
    mnis1Search: TMenuItem;
    mniN2Search: TMenuItem;
    pm4PriceEdit: TPopupMenu;
    mnid1SavePrice: TMenuItem;
    lbl5: TLabel;
    GridSSWebReq: TAdvStringGrid;
    DocsSheet: TAdvTabSheet;
    FileGrid4: TAdvStringGrid;
    advspltrContractSheet: TAdvSplitter;
    Gridedtlnk: TFormControlEditLink;
    pnlStatus: TPanel;
    lblStatus: TLabel;
    pmChklstPrice: TPopupMenu;
    mnic1: TMenuItem;
    pnlContractsList: TPanel;
    lbl6: TLabel;
    chklstPrice: TCheckListBox;
    lblPrice: TLabel;
    aplctnvnts1: TApplicationEvents;
    lblNotFound: TLabel;
    GridAllServices: TAdvStringGrid;
    GridFiltredPrice: TAdvStringGrid;
    edtServiceName: TAdvEdit;
    jvpnlPrice: TJvPanel;
    jvpnlContract: TJvPanel;
    pnl3Price: TJvPanel;
    pnl4date: TJvPanel;
    rbDefaultDate: TRadioButton;
    dtmpckrActionDate: TAdvDateTimePicker;
    rb2: TRadioButton;
    jvpnlRequisites: TJvPanel;
    jvpnl4: TJvPanel;
    edtInsertService: TAdvEditBtn;
    chkAddId: TCheckBox;
    edtContract: TAdvEdit;
    Gif: TJvGIFAnimator;
    pnlTreeView: TPanel;
    edt3: TAdvEdit;
    conSQL: TSQLConnection;
    sqlQry: TSQLQuery;
    tv1: TsShellTreeView;
    lv1: TsShellListView;
    advfrmstylr1: TAdvFormStyler;
    GridPosts: TAdvStringGrid;
    jvpnlPriceEditor: TJvPanel;
    lblPriceEditor: TLabel;
    jvpnlPosts: TJvPanel;
    lbl4: TLabel;
    Grid4PriceEditor: TAdvStringGrid;
    edtFilterPrice: TAdvEdit;
    pmPosts: TPopupMenu;
    mnia1: TMenuItem;
    mnid1: TMenuItem;
    mniS1: TMenuItem;
    mniN2: TMenuItem;
    advspltrWord: TAdvSplitter;
    advspltrEditorSheet: TAdvSplitter;
    procedure FormCreate(Sender: TObject);
    procedure btn7Click(Sender: TObject);
    procedure edtContractKeyPress(Sender: TObject; var Key: Char);
    procedure edtServiceNameChange(Sender: TObject);
    procedure edtContractClick(Sender: TObject);
    procedure GridSSWebReqGetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure GetReqmniC1Click(Sender: TObject);
    procedure CreateContractmniN1Click(Sender: TObject);
    procedure GridContractGetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure GridListOfServicesGetEditorType(Sender: TObject;
      ACol, ARow: Integer; var AEditor: TEditorType);
    procedure mniMoveClick(Sender: TObject);
    procedure GridOrderFormGetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure GridContractDateTimeChange(Sender: TObject; ACol, ARow: Integer;
      ADateTime: TDateTime);
    procedure mnia1CreateOrderFormClick(Sender: TObject);
    procedure mniAddRowClick(Sender: TObject);
    procedure mniDelRowClick(Sender: TObject);
    procedure GridContractEditCellDone(Sender: TObject; ACol, ARow: Integer);
    procedure edtlnkPriceSetEditorProperties(Sender: TObject;
      Grid: TAdvStringGrid; AControl: TWinControl);
    procedure edtlnkPriceSetEditorFocus(Sender: TObject; Grid: TAdvStringGrid;
      AControl: TWinControl);
    procedure edtServiceNameKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure GridListOfServicesCanEditCell(Sender: TObject;
      ARow, ACol: Integer; var CanEdit: Boolean);
    procedure edtlnkPriceGetEditorValue(Sender: TObject; Grid: TAdvStringGrid;
      var AValue: string);
    procedure GridListOfServicesEditingDone(Sender: TObject);
    procedure mniN1CreateOrderFormClick(Sender: TObject);
    procedure GridAllServicesClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure mnis1SearchClick(Sender: TObject);
    procedure mniN2SearchClick(Sender: TObject);
    procedure mnid1SavePriceClick(Sender: TObject);
    procedure chkAddIdClick(Sender: TObject);
    procedure edtServiceNameEnter(Sender: TObject);
    procedure GridedtlnkGetEditorValue(Sender: TObject; Grid: TAdvStringGrid;
      var AValue: string);
    procedure rbDefaultDateMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dtmpckrActionDateClick(Sender: TObject);
    procedure rb2Click(Sender: TObject);
    procedure GridedtlnkSetEditorValue(Sender: TObject; Grid: TAdvStringGrid;
      AValue: string);
    procedure FormResize(Sender: TObject);
    procedure mnic1Click(Sender: TObject);
    procedure edt3LookupSelect(Sender: TObject; var Value: string);
    procedure tv1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GridAllServicesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure aplctnvnts1Message(var Msg: tagMSG; var Handled: Boolean);
    procedure btn4Click(Sender: TObject);
    procedure dtmpckrActionDateKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edtInsertServiceClickBtn(Sender: TObject);
    procedure edtInsertServiceKeyPress(Sender: TObject; var Key: Char);
    procedure btn5Click(Sender: TObject);
    procedure GridAllServicesExpandNode(Sender: TObject;
      ARow, ARowreal: Integer);
    procedure GridAllServicesBeforeExpandNode(Sender: TObject;
      ARow, ARowreal: Integer; var Allow: Boolean);
    procedure tv1Editing(Sender: TObject; Node: TTreeNode;
      var AllowEdit: Boolean);
    procedure tv1Change(Sender: TObject; Node: TTreeNode);
    procedure DocsSheetShow(Sender: TObject);
    procedure GridListOfServicesCanDeleteRow(Sender: TObject; ARow: Integer;
      var CanDelete: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure mniS1Click(Sender: TObject);
    procedure mniA1Click(Sender: TObject);
    procedure mniD1Click(Sender: TObject);
    procedure edtFilterPriceEnter(Sender: TObject);
    procedure edtFilterPriceChange(Sender: TObject);
  private
    function FindAndReplace(const FindText, ReplaceText: string): Boolean;
    procedure InsertReqsContract;
    procedure GetReqsWord;
    procedure InsertReqsSSWEB;
    function GetNextMoonth(d: TDate): TDate;
    procedure SetContractGroup();
    procedure CreateContract;
    procedure InsertReqsOrderForm;
    procedure GetReqsSSWeb;
    procedure GetReqMap;
    procedure GetReqsOrderForm;
    procedure MoveContract;
    function DataProp(d: TDateTime): string;
    procedure CreateOrderForm;
    procedure AddPriceRow;
    procedure GetPriceList(CID: string);
    procedure GetPermanentServices(s: string);
    procedure GetPeriodicServices(s: string);
    procedure ShowPreloader;
    procedure HidePreloader;
    procedure TreeSearch(Tree: TsShellTreeView; SearchTarget: string);
    procedure EnableCtrBackspace(Handle: HWND);
    procedure AuthSSWeb;
    function DeclineBoss(s: string): string;
    function GetWordInString(s: string): TStringList;
    procedure ShowWordDocument();
    procedure ExtractAndOpenWordDocument(ResName, FileName: string);
    procedure GetAllData;
    procedure FilterPrice;
    procedure FilterActivate;
    procedure PriceListActivate;
    procedure ShowNotFound;
    procedure PrepareTreewiev;
    procedure GetGDrivePath;
    function GetDbPatch: string;
    procedure GridsInit;
    procedure AskAboutTransition(redirectURL: string);
    procedure AskAboutOpeningUrl(redirectURL: string);
  end;

var
  MainForm: TMainForm;
  WordDoc: variant;
  list: TStringList;
  b: Boolean;
  ReqMap: string;
  OldRow: string;
  AppPatch: string;
  GDrivePatch: string;
  FilePath: string;

implementation

uses
  Winapi.ShellAPI, Login;
{$R *.dfm}

procedure TMainForm.FormCreate(Sender: TObject);
var
  s: string;
begin
  GetGDrivePath;
  AppPatch := ExtractFilePath(application.ExeName);
  list := TStringList.Create;
  PrepareTreewiev;
  GridsInit;
  dtmpckrActionDate.Date := strtodate('01.' + formatdatetime('mm.yyyy', now));
  tv1.AutoRefresh := True;
  lv1.AutoRefresh := True;
  AuthSSWeb;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(list);
end;

function TMainForm.DataProp(d: TDateTime): string;
const
  Genitive: array [1 .. 12] of string = ('января', 'февраля', 'марта', 'апреля',
    'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября',
    'декабря');
var
  Year, Month, Day: Word;
begin
  DecodeDate(d, Year, Month, Day);
  Result := '«' + Day.ToString + '»' + ' ' + Genitive[Month] + ' ' +
    Year.ToString + ' г.';
end;

procedure TMainForm.edtServiceNameChange(Sender: TObject);
begin
  FilterPrice;
end;

procedure TMainForm.edtServiceNameEnter(Sender: TObject);
begin
  EnableCtrBackspace(edtServiceName.Handle);
  if Length(edtServiceName.Text) = 0 then
    GridAllServices.Show
end;

procedure TMainForm.edtServiceNameKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  service: string;
  ActiveGrid: TAdvStringGrid;
begin
  if GridAllServices.Visible = false then
    ActiveGrid := GridFiltredPrice
  else
    ActiveGrid := GridAllServices;
  if Key in [VK_DOWN, VK_UP] then
  begin
    ActiveGrid.SetFocus;
    keybd_event(Key, 0, 0, 0);
    keybd_event(Key, 0, KEYEVENTF_KEYUP, 0);
    Key := 0;
  end;
  service := GridListOfServices.Cells[0, GridListOfServices.Row];
  case Key of
    VK_ESCAPE:
      begin
        if ActiveGrid.Visible = false then
        begin
          edtServiceName.Text := service;
          GridListOfServices.Cells[1, GridListOfServices.Row] := OldRow;
          GridListOfServices.HideCellEdit;
        end
        else
        begin
          ActiveGrid.Hide;
          pnl3Price.Height := 25;
        end;
      end;
    VK_RETURN:
      if ActiveGrid.IsNode(ActiveGrid.Row) = false then
        with ActiveGrid do
        begin
          if Visible = false then
          begin
            if RowCount > 1 then
              GridListOfServices.Cells[1, GridListOfServices.Row] :=
                Cells[2, Row];
            GridListOfServices.HideCellEdit;
            GridListOfServices.Cells[0, GridListOfServices.Row] :=
              edtServiceName.Text;
            edtServiceName.Clear;
          end
          else
          begin
            if RowCount > 1 then
            begin
              edtServiceName.Text := Cells[1, Row];
              GridListOfServices.Cells[1, GridListOfServices.Row] :=
                Cells[2, Row];
            end;
            edtServiceName.SelLength := Length(edtServiceName.Text);
            Hide;
            pnl3Price.Height := 24;
          end;
        end
      else if ActiveGrid.NodeState[ActiveGrid.Row] = True then
        ActiveGrid.ExpandNode(ActiveGrid.RealRow)
  end;
end;

procedure TMainForm.edt3LookupSelect(Sender: TObject; var Value: string);
begin
  TreeSearch(tv1, Value);
  lbl6.Caption := Value;
  tv1.SetFocus;
end;

procedure TMainForm.edtInsertServiceClickBtn(Sender: TObject);
begin
  if Length(edtInsertService.Text) > 6 then
    GetPriceList(edtInsertService.Text);
end;

procedure TMainForm.edtInsertServiceKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    if Length(edtInsertService.Text) > 6 then
      GetPriceList(edtInsertService.Text);
end;

procedure TMainForm.TreeSearch(Tree: TsShellTreeView; SearchTarget: string);
var
  TrText: string;
  i, j: Integer;
begin
  i := 0;
  Tree.Items.Item[0].Selected := True;
  for i := Tree.Selected.AbsoluteIndex + 1 to Tree.Items.Count - 1 do
  begin
    TrText := Tree.Items.Item[i].Text;
    if Pos(ansiUppercase(SearchTarget), ansiUppercase(TrText)) > 0 then
    begin
      Tree.Items[i].MakeVisible;
      Tree.Items.Item[i].Selected := True;
      break;
    end;
  end;
end;

procedure TMainForm.tv1Change(Sender: TObject; Node: TTreeNode);
begin
  lbl6.Caption := tv1.Selected.Text;
end;

procedure TMainForm.tv1Editing(Sender: TObject; Node: TTreeNode;
  var AllowEdit: Boolean);
begin
  AllowEdit := false;
end;

procedure TMainForm.tv1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key in [VK_UP, VK_DOWN, VK_RETURN] = false then
  begin
    edt3.SetFocus;
    edt3.SelStart := Length(edt3.Text);
    keybd_event(Key, 0, 0, 0);
    keybd_event(Key, 0, KEYEVENTF_KEYUP, 0);
    Key := 0;
  end;
  if Key = VK_RETURN then
    if tv1.Selected.Expanded then
      tv1.Selected.Collapse(false)
    else
      tv1.Selected.Expand(false);
end;

function TMainForm.FindAndReplace(const FindText, ReplaceText: string): Boolean;
const
  wdReplaceAll = 2;
begin
  WordDoc.Selection.Find.MatchSoundsLike := false;
  WordDoc.Selection.Find.MatchAllWordForms := false;
  WordDoc.Selection.Find.Format := false;
  WordDoc.Selection.Find.Forward := True;
  WordDoc.Selection.Find.ClearFormatting;
  WordDoc.Selection.Find.Text := FindText;
  WordDoc.Selection.Find.Replacement.Text := ReplaceText;
  FindAndReplace := WordDoc.Selection.Find.Execute(EmptyParam, EmptyParam,
    EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
    EmptyParam, EmptyParam, wdReplaceAll, EmptyParam, EmptyParam, EmptyParam,
    EmptyParam);
end;

procedure TMainForm.FormResize(Sender: TObject);
begin
  if pnlStatus.Visible = True then
  begin
    pnlStatus.Left := (MainForm.ClientWidth - pnlStatus.Width) div 2;
    pnlStatus.Top := (MainForm.ClientHeight - pnlStatus.Height -
      (pnlStatus.Height div 3)) div 2;
  end;
  edtlnkPriceSetEditorProperties(nil, GridListOfServices, pnl3Price);
end;

procedure TMainForm.edtlnkPriceGetEditorValue(Sender: TObject;
  Grid: TAdvStringGrid; var AValue: string);
begin
  AValue := edtServiceName.Text;
end;

procedure TMainForm.edtlnkPriceSetEditorFocus(Sender: TObject;
  Grid: TAdvStringGrid; AControl: TWinControl);
begin
  edtServiceName.SetFocus;

end;

procedure TMainForm.edtlnkPriceSetEditorProperties(Sender: TObject;
  Grid: TAdvStringGrid; AControl: TWinControl);
begin
  if GridFiltredPrice.RowCount > 1 then
  begin
    AControl.Height := 224;
  end
  else
    AControl.Height := 50;
end;

procedure TMainForm.InsertReqsContract;
begin
  with GridContract do
  begin
    Cells[2, 0] := 'dog';
    Cells[2, 1] := 'date';
    Cells[2, 2] := 'C org';
    Cells[2, 3] := 'C bosst';
    Cells[2, 4] := 'C bossa';
    Cells[2, 5] := 'under';
    Cells[2, 6] := 'start';
    Cells[2, 7] := 'end';
    Cells[2, 8] := 'C adr';
    Cells[2, 9] := 'C padr';
    Cells[2, 10] := 'C phone';
    Cells[2, 11] := 'C master';
    Cells[2, 12] := 'D masterPhone';
    Cells[2, 13] := 'C fax';
    Cells[2, 14] := 'C bank';
    Cells[2, 15] := 'C inn';
    Cells[2, 16] := 'C kpp';
    Cells[2, 17] := 'C mfo';
    Cells[2, 18] := 'C pm';
    Cells[2, 19] := 'D boss';
  end;
end;

procedure TMainForm.InsertReqsOrderForm;
begin
  with GridOrderForm do
  begin
    Cells[2, 0] := 'C boss';
    Cells[2, 1] := 'C padr';
    Cells[2, 2] := 'date start';
  end;
end;

procedure TMainForm.GetReqsSSWeb;
var
  s: string;
  ReplaceString: string;
  i: Integer;
  AreqMap: TStringList;
begin
  GridSSWebReq.ClearNormalCols(1, 1);
  s := ReqMap;
  GridSSWebReq.StartUpdate;
  for i := 0 to GridSSWebReq.RowCount - 1 do
  begin
    if s.Contains(GridSSWebReq.Cells[2, i]) then
      if GridSSWebReq.Cells[2, i].Contains('C ') then
      begin
        GridSSWebReq.Cells[1, i] := Pars(s, GridSSWebReq.Cells[2, i] + ' ',
          ''#10'').Trim;
        ReplaceString := GridSSWebReq.Cells[2, i] + ' ' +
          Pars(s, GridSSWebReq.Cells[2, i] + ' ', #10) + #10;
        s := s.Replace(ReplaceString, '', []);
      end;
  end;
  GridSSWebReq.Dates[1, 2] := GetNextMoonth(Date);
  if GridSSWebReq.Cells[1, 21] = 'ok' then
    GridSSWebReq.AddCheckBox(1, 21, True, True)
  else
    GridSSWebReq.AddCheckBox(1, 21, false, True);
  SetContractGroup;
  GridSSWebReq.Enabled := True;
  GridSSWebReq.EndUpdate;
end;

procedure TMainForm.GetReqMap;
begin
  ReqMap := idhtp1.Get
    ('http://dyatel.antar.bryansk.me/cgi-bin/ss-web/contract.cgi?cid=' +
    Trim(edtContract.Text));
  if ReqMap.Contains('ошибка') then
  begin
    HidePreloader;
    ShowMessage('Договор не найден');
    Abort;
  end;
  ReqMap := Pars(ReqMap, 'con_req">', '#C cred');
  ReqMap := ClearTeg(ReqMap, '');
  ReqMap := ReqMap.Replace('&quot;', '"');
  ReqMap := ReqMap.Replace('  БИК', '');
  ReqMap := ReqMap.Replace('XXX', '');
end;

procedure TMainForm.InsertReqsSSWEB;
begin
  with GridSSWebReq do
  begin
    Cells[2, 0] := 'con_cat';
    Cells[2, 1] := 'new_cid';;
    Cells[2, 2] := 'ctime';
    Cells[2, 3] := 'C org';
    Cells[2, 4] := 'C orgs';
    Cells[2, 5] := 'C orgpok';
    Cells[2, 6] := 'C adr';
    Cells[2, 7] := 'C padr';
    Cells[2, 8] := 'C adrpok';
    Cells[2, 9] := 'C phone';
    Cells[2, 10] := 'C master';
    Cells[2, 11] := 'C fax';
    Cells[2, 12] := 'C bank';
    Cells[2, 13] := 'C inn';
    Cells[2, 14] := 'C kpp';
    Cells[2, 15] := 'C mfo';
    Cells[2, 16] := 'C boss';
    Cells[2, 17] := 'C bossa';
    Cells[2, 18] := 'C bosst';
    Cells[2, 20] := 'C pm';
    Cells[2, 19] := 'C to';
    Cells[2, 21] := 'C send';
  end;
end;

procedure TMainForm.mniA1Click(Sender: TObject);
begin
  GridPosts.AddRow;
end;

procedure TMainForm.mnia1CreateOrderFormClick(Sender: TObject);
begin
  CreateOrderForm;
end;

procedure TMainForm.CreateOrderForm;
var
  i: Integer;
  tbl: variant;
  tag: string;
  repstr: string;
  s: string;
  s1: string;
begin
  lblStatus.Caption := 'Создание Бланка заказа';
  ShowPreloader;
  ExtractAndOpenWordDocument('OrderForm', 'Бланк заказ');
  FindAndReplace(GridOrderForm.Cells[2, 0], GridOrderForm.Cells[1, 0]);
  FindAndReplace(GridOrderForm.Cells[2, 1], GridOrderForm.Cells[1, 1]);
  FindAndReplace(GridOrderForm.Cells[2, 2],
    DataProp(GridOrderForm.Dates[1, 2]));
  FindAndReplace(GridContract.Cells[2, 1], DataProp(GridContract.Dates[1, 1]));
  s := WordDoc.Selection.Text;
  WordDoc.Selection.SetRange(0, 0);
  for i := 0 to GridContract.RowCount do
  begin
    tag := GridContract.Cells[2, i];
    repstr := GridContract.Cells[1, i];
    FindAndReplace(tag, repstr);
  end;
  tbl := WordDoc.ActiveDocument.Tables.Item(3);
  WordDoc.Selection.Font.Bold := 0;
  WordDoc.Selection.Font.Italic := 0;
  for i := 1 to GridListOfServices.RowCount - 1 do
    if (GridListOfServices.Cells[0, i] <> '') and
      (GridListOfServices.Cells[1, i] <> '') then
    begin
      tbl.Rows.Add(EmptyParam);
      tbl.Cell(i + 1, 1).range.Paragraphs.Alignment := wdAlignParagraphLeft;
      tbl.Cell(i + 1, 1).range.Font.Bold := false;
      tbl.Cell(i + 1, 1).range.Font.Italic := false;
      tbl.Cell(i + 1, 1).range.InsertBefore(GridListOfServices.Cells[0, i]);
      tbl.Cell(i + 1, 2).range.Font.Bold := false;
      tbl.Cell(i + 1, 2).range.Font.Italic := false;
      tbl.Cell(i + 1, 2).range.InsertBefore(GridListOfServices.Cells[1, i]);
    end;
  ShowWordDocument();
  HidePreloader;
end;

procedure TMainForm.AddPriceRow;
begin
  if (edtServiceName.Text <> ' ') and (edtServiceName.Text <> '') then
    if GridListOfServices.Row = GridListOfServices.RowCount - 1 then
      GridListOfServices.AddRow;
end;

procedure TMainForm.mniAddRowClick(Sender: TObject);
begin
  GridListOfServices.AddRow;
end;

procedure TMainForm.mnic1Click(Sender: TObject);
var
  i: Integer;
begin
  lblStatus.Caption := 'Создание прайса';
  ShowPreloader;
  if (chklstPrice.Checked[0] or chklstPrice.Checked[1]) = True then
  begin
    for i := 0 to chklstPrice.Items.Count - 1 do
      if chklstPrice.Checked[i] then
      begin
        ExtractAndOpenWordDocument(chklstPrice.Items[i],
          'Прайс-лист(' + chklstPrice.Items[i] + ')');
        FindAndReplace(GridContract.Cells[2, 0], GridContract.Cells[1, 0]);
        WordDoc.ActiveDocument.Close;
        WordDoc.Quit;
      end
  end
  else
    ShowMessage('Не выбран прайс лист');
  HidePreloader;
end;

procedure TMainForm.mniD1Click(Sender: TObject);
begin
  GridPosts.RemoveRows(GridPosts.Row, 1);
end;

procedure TMainForm.mnid1SavePriceClick(Sender: TObject);
begin
  Grid4PriceEditor.SaveToFile(ExtractFilePath(application.ExeName) +
    'price.dat');
  GridAllServices.UnGroup;
  GridAllServices.ClearAll;
  GridAllServices.LoadFromFile(ExtractFilePath(application.ExeName) +
    'price.dat');
end;

procedure TMainForm.mniDelRowClick(Sender: TObject);
begin
  with GridListOfServices do
  begin
    if RowCount > 2 then
    begin
      RemoveRows(Row, 1);
    end
    else
      ClearNormalCells;
  end;
end;

procedure TMainForm.mniMoveClick(Sender: TObject);
begin
  case application.MessageBox('Перенести договор?', 'Подтверждение',
    MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) of
    IDYES:
      MoveContract;
  end;
end;

procedure TMainForm.mniN1CreateOrderFormClick(Sender: TObject);
begin
  case application.MessageBox('Создать Бланк Заказ?', 'Подтверждение',
    MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) of
    IDYES:
      CreateOrderForm;
  end;
end;

procedure TMainForm.mniN2SearchClick(Sender: TObject);
begin
  if GridSSWebReq.SearchFooter.Visible = false then
  begin
    GridSSWebReq.SearchFooter.Visible := True;
  end
  else
    GridSSWebReq.SearchFooter.Visible := false;
end;

procedure TMainForm.mnis1SearchClick(Sender: TObject);
begin
  if GridContract.SearchFooter.Visible = false then
  begin
    GridContract.SearchFooter.Visible := True;
    GridContract.SearchPanel.EditControl.SetFocus;
  end
  else
  begin
    GridContract.SearchFooter.Visible := false;
    GridContract.SetFocus;
  end;
end;

procedure TMainForm.mniS1Click(Sender: TObject);
begin
  GridPosts.SaveToFile(AppPatch + 'list of posts.list');
end;

function TMainForm.GetNextMoonth(d: TDate): TDate;
begin
  d := EncodeDate(YearOf(d), MonthOf(d), 1);
  d := IncMonth(d, 1);
  Result := d;
end;

procedure TMainForm.SetContractGroup();
var
  i: Integer;
  s: string;
begin
  list.LoadFromFile(AppPatch + 'Groups.dat');
  if Pos('C', edtContract.Text) <> 0 then
    for i := 0 to list.Count - 1 do
    begin
      s := Pars(Trim(edtContract.Text), '', '-');
      if list.Strings[i].Contains(s) then
        GridSSWebReq.Cells[1, 0] := list.Strings[i];
    end
  else
    GridSSWebReq.Cells[1, 0] := list.Strings[0];
  list.Clear;
end;

procedure TMainForm.GetReqmniC1Click(Sender: TObject);
begin
  GetReqsWord;
end;

procedure TMainForm.GetReqsWord;
var
  s: string;
  i: Integer;
  masterPhone: Ansistring;
begin
  GridContract.ClearNormalCols(1, 1);
  s := ReqMap;
  GridContract.StartUpdate;
  for i := 0 to GridContract.RowCount - 1 do
  begin
    if s.Contains(GridContract.Cells[2, i]) then
      if GridContract.Cells[2, i].Contains('C ') then
        GridContract.Cells[1, i] := Pars(s, GridContract.Cells[2, i],
          ''#10'').Trim;
  end;
  if GridContract.Cells[1, 18] = '' then
    GridContract.Cells[1, 18] := Pars(s, 'C to', ''#10'').Trim;
  GridContract.Cells[1, 0] := edtContract.Text;
  GridContract.Dates[1, 1] := Date;
  GridContract.Cells[1, 6] := 'момента подписания ';
  GridContract.Cells[1, 7] := 'момента расторжения';
  masterPhone := GetOnlyNumber(GridContract.Cells[1, 11]);
  if Length(masterPhone) > 0 then
    while masterPhone[1] in [',', ';', '-', ' ', '(', ')'] <> false do
      Delete(masterPhone, 1, 1);
  GridContract.Cells[1, 12] := masterPhone;
  GridContract.Cells[1, 11] := GetOnlyWords(GridContract.Cells[1, 11]);
  GridContract.Cells[1, 19] := DeclineBoss(GridContract.Cells[1, 3]);
  GridContract.EndUpdate;
end;

function TMainForm.DeclineBoss(s: string): string;
var
  WList: TStringList;
  i, Row: Integer;
  s1, s2: string;
begin
  WList := GetWordInString(s);
  for i := 0 to WList.Count - 1 do
  begin
    s1 := WList.Strings[i];
    for Row := 0 to GridPosts.RowCount - 1 do
      if Pos(AnsiLowerCase(s1), AnsiLowerCase(GridPosts.Cells[0, Row])) <> 0
      then
      begin
        s2 := GridPosts.Cells[1, Row];
        if s1[1] <> s2[1] then
          s2[1] := ansiUppercase(s2[1])[1];
        WList.Strings[i] := s2;
      end;
  end;
  Result := WList.Text.Replace(#13#10, ' ').Trim;
  FreeAndNil(WList);
end;

function TMainForm.GetWordInString(s: string): TStringList;
var
  i: Integer;
  Word: string;
begin
  Result := TStringList.Create;
  For i := 1 to Length(s) do
  begin
    if s[i] <> ' ' then
      Word := Word + s[i]
    else if Word <> '' then
    begin
      Result.Add(Word);
      Word := '';
    end;
  end;
  if Word <> '' then
    Result.Add(Word);
end;

procedure TMainForm.ExtractAndOpenWordDocument(ResName, FileName: string);
var
  res: TResourceStream;
begin
  FileName := '\' + FileName + '.docx';
  res := TResourceStream.Create(HInstance, ResName, RT_RCDATA);
  FilePath := GDrivePatch + '\' + formatdatetime('yyyy', Date) + '\' +
    GridContract.Cells[1, 0] + ' ' + Pars(ReqMap, 'C orgs', ''#10'').Trim;
  FilePath := FilePath.Replace('"', '');
  if DirectoryExists(FilePath) = false then
    ForceDirectories(FilePath);
  if GridContract.Cells[1, 2] <> '' then
    try
      res.SaveToFile(FilePath + FileName);
      FreeAndNil(res);
    Except
      on E: Exception do
      begin
        HidePreloader;
        ShowMessage(E.Message);
        Abort;
      end;
    end
  else
  begin
    ShowMessage('Не указан новый номер договора');
    HidePreloader;
    Abort;
  end;
  WordDoc := CreateOleObject('Word.Application');
  WordDoc.Visible := false;
  WordDoc.Documents.Open(FilePath + FileName);
  WordDoc.ActiveWindow.Visible := false;
end;

procedure TMainForm.ShowWordDocument();
begin
  ShellExecute(MainForm.Handle, 'open', PWideChar(FilePath), nil, nil,
    SW_SHOWNORMAL);
  WordDoc.Documents.Save;
  WordDoc.Visible := Visible;
  WordDoc.WindowState := 2;
  WordDoc.WindowState := wdWindowStateMaximize;
end;

procedure TMainForm.GetAllData;
begin
  lblStatus.Caption := 'Получение данных';
  chklstPrice.CheckAll(cbUnchecked, True, false);
  edtContract.Text := edtContract.Text.Trim;
  GetReqMap;
  GetReqsWord;
  GetReqsSSWeb;
  GetReqsOrderForm;
  GetPriceList(edtContract.Text);
  GridContract.Enabled := True;
end;

procedure TMainForm.FilterPrice;
begin
  if Length(edtServiceName.Text) > 0 then
    FilterActivate
  else
    PriceListActivate;
  if GridFiltredPrice.RowCount = 1 then
    ShowNotFound;
end;

procedure TMainForm.FilterActivate;
begin
  GridAllServices.Hide;
  with GridFiltredPrice do
  begin
    BeginUpdate;
    Hide;
    lblNotFound.Visible := false;
    Filter.Clear;
    filteractive := false;
    with GridFiltredPrice.Filter.Add do
    begin
      condition := edtServiceName.Text;
      Data := fcRow;
      CaseSensitive := false;
    end;
    filteractive := True;
    if RowCount > 1 then
    begin
      pnl3Price.Height := 224;
      Show;
    end;
    EndUpdate;
  end;
end;

procedure TMainForm.PriceListActivate;
begin
  lblNotFound.Visible := false;
  GridFiltredPrice.Filter.Clear;
  GridFiltredPrice.filteractive := false;
  lblNotFound.Hide;
  GridFiltredPrice.Hide;
  GridAllServices.ContractAll;
  GridAllServices.Row := 1;
  GridAllServices.Show;
  pnl3Price.Height := 224;
end;

procedure TMainForm.ShowNotFound;
begin
  GridFiltredPrice.Hide;
  GridAllServices.Hide;
  pnl3Price.Height := 50;
  lblNotFound.Visible := True;
end;

procedure TMainForm.GetGDrivePath;
var
  Query: string;
  str: string;
  DbPatch: string;
begin
  conSQL.Params.Add('Database=' + GetDbPatch);
  conSQL.Connected := True;
  sqlQry.Active := True;
  sqlQry.Open;
  GDrivePatch := sqlQry.FieldByName('data_value').AsString;
  GDrivePatch := GDrivePatch.Replace('\\?\', '') + '\Договоры';
end;

function TMainForm.GetDbPatch: string;
begin
  Result := GetEnvironmentVariable('LOCALAPPDATA');
  if Result <> '' then
    Result := IncludeTrailingPathDelimiter(Result) +
      'Google\Drive\user_default\sync_config.db';
end;

procedure TMainForm.GridsInit;
begin
  GridContract.HideColumn(2);
  GridSSWebReq.HideColumn(2);
  GridOrderForm.HideColumn(2);
  InsertReqsContract;
  InsertReqsSSWEB;
  InsertReqsOrderForm;
  GridAllServices.LoadFromFile(AppPatch + 'price.dat');
  GridFiltredPrice.LoadFromFile(AppPatch + 'price.dat');
  GridFiltredPrice.RemoveCols(3, 1);
  GridFiltredPrice.RemoveCols(3, 1);
  Grid4PriceEditor.LoadFromFile(AppPatch + 'price.dat');
  GridPosts.LoadFromFile(AppPatch + 'list of posts.list');
  GridAllServices.HideColumn(4);
  GridAllServices.Group(3);
  GridAllServices.ContractAll;
  GridListOfServices.HideColumn(2);
end;

procedure TMainForm.AskAboutTransition(redirectURL: string);
var
  OldCid: string;
begin
  if application.MessageBox('Перейти на вкладку создания Word договора?',
    'Вопрос', MB_YESNO + MB_ICONQUESTION + MB_TOPMOST) = IDYES then
  begin
    OldCid := edtContract.Text;
    edtContract.Text := Pars(redirectURL, 'cid=', '');
    GetAllData;
    GetPriceList(OldCid);
    advpgcntrl1.ActivePage := WordSheet;
  end;
end;

procedure TMainForm.AskAboutOpeningUrl(redirectURL: string);
begin
  if application.MessageBox('Договор успешно перенесён. ' + ''#13''#10'' +
    'Открыть ссылку на новый договор в браузере?',
    'Завершение переноса договора', MB_OKCANCEL + MB_ICONINFORMATION +
    MB_TOPMOST) = IDOK then
  begin
    ShellExecute(Handle, 'open', PWideChar(redirectURL), nil, nil, SW_NORMAL);
  end;
end;

procedure TMainForm.PrepareTreewiev;
var
  i, j: Integer;
  MainNode, YearNode, ContractNode: TTreeNode;
begin
  tv1.Root := GDrivePatch;
  SendMessage(pnlTreeView.Handle, WM_SETREDRAW, Ord(false), 0);
  tv1.FullExpand;
  MainNode := tv1.Items.Item[0];
  for i := 0 to MainNode.Count - 1 do
  begin
    YearNode := MainNode.Item[i];
    for j := 0 to YearNode.Count - 1 do
    begin
      ContractNode := YearNode.Item[j];
      edt3.Lookup.DisplayList.Add(ContractNode.Text);
    end;
    YearNode.Collapse(false);
  end;
  SendMessage(pnlTreeView.Handle, WM_SETREDRAW, Ord(True), 0);
  InvalidateRect(edt3.Handle, nil, True);
end;

procedure TMainForm.GetReqsOrderForm;
begin
  GridListOfServices.ClearNormalCells;
  GridListOfServices.RowCount := 2;
  GridOrderForm.StartUpdate;
  GridOrderForm.ClearNormalCols(1, 1);
  GridOrderForm.Cells[1, 0] := Pars(ReqMap, 'C boss ', ''#10'').Trim;
  GridOrderForm.Cells[1, 1] := GridContract.Cells[1, 9];
  GridOrderForm.Cells[1, 2] := GridContract.Cells[1, 1];
  GridOrderForm.EndUpdate;
end;

procedure TMainForm.MoveContract;
var
  i: Integer;
  url, redirectURL: string;
begin
  ShowPreloader;
  lblStatus.Caption := 'Перенос договора';
  list.Clear;
  list.Add(GridSSWebReq.Cells[2, 0] + '=' + Pars(GridSSWebReq.Cells[1, 0],
    '(', ')'));
  for i := 1 to GridSSWebReq.RowCount - 2 do
  begin
    if GridSSWebReq.Cells[2, i].Contains('C ') then
      list.Add(GridSSWebReq.Cells[2, i].Replace('C ', 'C_') + '=' +
        GridSSWebReq.Cells[1, i])
    else
      list.Add(GridSSWebReq.Cells[2, i] + '=' + GridSSWebReq.Cells[1, i]);
  end;
  if GridSSWebReq.IsChecked(1, i) then
    list.Add('C send=yes')
  else
    list.Add('C send=no');
  list.Add('juro=1');
  list.Add('confirm=1');
  url := SSWeb + 'add-contract-wizard.cgi';
  idhtp1.Post(url, list, IndyTextEncoding(20866));
  list.Clear;
  redirectURL := SSWeb + idhtp1.Response.Location;
  AskAboutOpeningUrl(redirectURL);
  AskAboutTransition(redirectURL);
  HidePreloader;
end;

procedure TMainForm.rbDefaultDateMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  rbDefaultDate.Checked := True;
  GridContract.Cells[1, 6] := 'момента подписания ';
  GridContract.Cells[1, 7] := 'момента расторжения';
  pnl4date.Hide;
  GridContract.HideCellEdit;
end;

procedure TMainForm.rb2Click(Sender: TObject);
begin
  dtmpckrActionDate.SetFocus;
end;

procedure TMainForm.GridSSWebReqGetEditorType(Sender: TObject;
  ACol, ARow: Integer; var AEditor: TEditorType);
begin
  with GridSSWebReq do
  begin
    if ACol = 1 then
      case ARow of
        2:
          begin
            AEditor := edDateEdit;
            GridSSWebReq.DateTimePicker.Format := 'dd.MM.yyyy';
          end;
        0:
          begin
            AEditor := edComboList;
            ClearComboString;
            Combobox.Items.LoadFromFile('Groups.dat');
          end;
        13, 14, 15:
          begin
            AEditor := edNumeric;
          end;
      end;
  end;
end;

procedure TMainForm.GridListOfServicesCanDeleteRow(Sender: TObject;
  ARow: Integer; var CanDelete: Boolean);
begin
  if ARow = GridListOfServices.LastRow then
    CanDelete := false;
end;

procedure TMainForm.GridListOfServicesCanEditCell(Sender: TObject;
  ARow, ACol: Integer; var CanEdit: Boolean);
begin
  if GridListOfServices.Cells[0, GridListOfServices.Row].Length > 0 then
  begin
    if GridListOfServices.EditMode = false then
    begin
      edtServiceName.Text := GridListOfServices.Cells[0, ARow];
    end;
  end;
  OldRow := GridListOfServices.Cells[1, GridListOfServices.Row];
end;

procedure TMainForm.GridListOfServicesEditingDone(Sender: TObject);
var
  CurentService: string;
begin
  CurentService := GridListOfServices.Cells[0, GridListOfServices.Row];
  if (CurentService = '') or (CurentService = ' ') = True then
    GridListOfServices.ClearRows(GridListOfServices.Row, 1);
  AddPriceRow;
end;

procedure TMainForm.GridListOfServicesGetEditorType(Sender: TObject;
  ACol, ARow: Integer; var AEditor: TEditorType);
begin
  if ACol = 0 then
    AEditor := edCustom;
  GridListOfServices.EditLink := edtlnkPrice;
end;

procedure TMainForm.GridOrderFormGetEditorType(Sender: TObject;
  ACol, ARow: Integer; var AEditor: TEditorType);
begin
  if ARow = 2 then
    AEditor := edDateEdit;
end;

procedure TMainForm.GridContractDateTimeChange(Sender: TObject;
  ACol, ARow: Integer; ADateTime: TDateTime);
begin
  if ARow = 1 then
    GridOrderForm.Cells[1, 2] := DateToStr(GridContract.DateTimePicker.Date);
end;

procedure TMainForm.GridContractEditCellDone(Sender: TObject;
  ACol, ARow: Integer);
begin
  if ARow = 1 then
    GridOrderForm.Cells[1, 2] := DateToStr(GridContract.DateTimePicker.Date);
end;

procedure TMainForm.GridedtlnkGetEditorValue(Sender: TObject;
  Grid: TAdvStringGrid; var AValue: string);
begin
  if rbDefaultDate.Checked then
    case Grid.Row of
      6:
        AValue := 'момента подписания';
      7:
        AValue := 'момента расторжения';
    end;
  if rb2.Checked then
  begin
    if (Grid.Cells[1, 6] = 'момента подписания') or
      (Grid.Cells[1, 7] = 'момента расторжения') then
    begin
      Grid.Cells[1, 6] := '';
      Grid.Cells[1, 7] := '';
    end;
    AValue := DateToStr(dtmpckrActionDate.Date);
  end;
end;

procedure TMainForm.GridedtlnkSetEditorValue(Sender: TObject;
  Grid: TAdvStringGrid; AValue: string);
begin
  if Grid.Row = 7 then
    dtmpckrActionDate.MinDate := dtmpckrActionDate.Date
  else
    dtmpckrActionDate.MinDate := strtodate('01.01.1800');
  try
    dtmpckrActionDate.Date := strtodate(AValue);
  except
  end;
end;

procedure TMainForm.GridContractGetEditorType(Sender: TObject;
  ACol, ARow: Integer; var AEditor: TEditorType);
begin
  with GridContract do
    if ACol = 1 then
      case ARow of
        1:
          begin
            AEditor := edDateEdit;
          end;
        5:
          begin
            AEditor := edComboEdit;
            ClearComboString;
            AddComboString('Устава');
            AddComboString('Положения');
            AddComboString('Приказа');
            AddComboString('Доверенности');
          end;
        6, 7:
          begin
            AEditor := edCustom;
            GridContract.EditLink := Gridedtlnk;
          end;
      end;
end;

procedure TMainForm.dtmpckrActionDateClick(Sender: TObject);
begin
  rb2.Checked := True;
end;

procedure TMainForm.dtmpckrActionDateKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
    GridContract.HideCellEdit;
end;

procedure TMainForm.aplctnvnts1Message(var Msg: tagMSG; var Handled: Boolean);
var
  ActiveGrid: TAdvStringGrid;
  Pos: Integer;
begin
  if GridAllServices.Visible = false then
    ActiveGrid := GridFiltredPrice
  else
    ActiveGrid := GridAllServices;
  with Msg do
    if message = WM_MOUSEWHEEL Then
      if not ActiveGrid.Focused then
      begin
        Pos := GetScrollPos(ActiveGrid.Handle, SB_VERT);
        if wParam = 4287102976 then
          ActiveGrid.Perform(WM_VSCROLL, MAKEWPARAM(SB_THUMBPOSITION,
            Pos + 40), 0)
        else
          ActiveGrid.Perform(WM_VSCROLL, MAKEWPARAM(SB_THUMBPOSITION,
            Pos - 40), 0)
      end;
end;

procedure TMainForm.btn4Click(Sender: TObject);
begin
  ShowPreloader;
end;

procedure TMainForm.btn5Click(Sender: TObject);
begin
  HidePreloader
end;

procedure TMainForm.btn7Click(Sender: TObject);
begin
  if GridListOfServices.RowCount > 2 then
    GridListOfServices.RemoveRows(2, 1);
end;

procedure TMainForm.chkAddIdClick(Sender: TObject);
var
  i: Integer;
var
  id: string;
begin
  with GridListOfServices do
  begin
    BeginUpdate;
    for i := 1 to LastRow do
    begin
      id := Cells[2, i];
      if chkAddId.Checked then
        Cells[0, i] := Cells[0, i] + ' ' + id
      else
        Cells[0, i] := Cells[0, i].Replace(id, '').Trim;
    end;
    EndUpdate;
  end;
end;

procedure TMainForm.CreateContract;
var
  s: string;
  tag: string;
  repstr: string;
  i: Integer;
begin
  lblStatus.Caption := 'Создание договора';
  ShowPreloader;
  ExtractAndOpenWordDocument('Contract', 'Договор');
  WordDoc.Selection.WholeStory;
  s := WordDoc.Selection.Text;
  for i := 0 to GridContract.RowCount do
  begin
    tag := GridContract.Cells[2, i];
    repstr := GridContract.Cells[1, i];
    if repstr.Length > 3 then
      if Pars(s, tag, ' ').Contains('__') then
        FindAndReplace(Pars(s, tag, ' ') + ' ', '');
    FindAndReplace(tag, repstr);
  end;
  FindAndReplace(tag, DataProp(GridContract.Dates[1, 1]));
  FindAndReplace('_ ', '_');
  FindAndReplace('&', '');
  WordDoc.Selection.SetRange(0, 0);
  ShowWordDocument;
  HidePreloader;
  TreeSearch(tv1, GridContract.Cells[1, 0] + ' ' + Pars(ReqMap, 'C orgs',
    ''#10'').Trim);
end;

procedure TMainForm.EnableCtrBackspace(Handle: HWND);
begin
  SHAutoComplete(Handle, SHACF_AUTOAPPEND_FORCE_OFF or
    SHACF_AUTOSUGGEST_FORCE_OFF);
end;

procedure TMainForm.AuthSSWeb;
begin
  AuthForm := TAuthForm.Create(self);
  AuthForm.ShowModal;
  if AuthForm.ModalResult = mrcancel then
    application.Terminate;
  if AuthForm.ModalResult = mrOk then
  begin
    AuthForm.Close;
    MainForm.Show;
  end;
end;

procedure TMainForm.DocsSheetShow(Sender: TObject);
begin
  edt3.SetFocus;
end;

procedure TMainForm.CreateContractmniN1Click(Sender: TObject);
begin
  case application.MessageBox('Создать договор? ', 'Подтверждение',
    MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) of
    IDYES:
      CreateContract;
  end;
end;

procedure TMainForm.edtContractClick(Sender: TObject);
begin
  if edtContract.Text = 'Номер исходного договора' then
  begin
    edtContract.Clear;
    edtContract.Font.Color := clBlack;
    edtContract.Font.Style := [];
  end;
end;

procedure TMainForm.edtContractKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) and (Length(edtContract.Text) > 4) then
  begin
    ShowPreloader;
    GetAllData;
    HidePreloader;
  end;
end;

procedure TMainForm.edtFilterPriceChange(Sender: TObject);
begin
  Grid4PriceEditor.NarrowDown(edtFilterPrice.Text);
end;

procedure TMainForm.edtFilterPriceEnter(Sender: TObject);
begin
  LoadKeyboardLayout('00000419', KLF_ACTIVATE);
  EnableCtrBackspace(edt3.Handle);
end;

procedure TMainForm.GetPriceList(CID: string);
var
  s, url: string;
begin
  url := ServicesCgi + 'cid=' + CID + ';active=1';
  s := idhtp1.Get(url);
  GetPermanentServices(s);
  GetPeriodicServices(s);
end;

procedure TMainForm.GetPermanentServices(s: string);
var
  service: string;
  sid: string;
  speed: Integer;
  i: Integer;
begin
  s := Pars(s, 'Аккаунты', '');
  s := Pars(s, '<tbody>', '</tbody>');
  if s <> '' then
    repeat
      s := Pars(s, '<tr>', '');
      s := Pars(s, '</span></td>', '');
      service := Pars(s, '">', '<').Trim;
      sid := Pars(s, 'sid=', ';').Replace('%40', '@');
      if service <> '&empty;' then
        if Grid4PriceEditor.Cols[4].Text.Contains(service) = false then
        begin
          speed := Pars(s, '|', 'Kbit').Trim.ToInteger;
          if speed >= 1000 then
            speed := speed div 1000;
          GridListOfServices.Cells[0, GridListOfServices.RowCount - 1] :=
            'Оплата интернет канала ' + speed.ToString + ' Мбит/с';
          GridListOfServices.Cells[1, GridListOfServices.RowCount - 1] :=
            rpars(Pars(s, '', '|'), 'р', '"') + ' руб./мес.';
          GridListOfServices.Cells[2, GridListOfServices.LastRow] := sid;
          GridListOfServices.AddRow;
        end
        else
          for i := 1 to Grid4PriceEditor.RowCount - 1 do
          begin
            if service = Grid4PriceEditor.Cells[4, i] then
            begin
              if chkAddId.Checked = True then
                GridListOfServices.Cells[0, GridListOfServices.RowCount - 1] :=
                  Grid4PriceEditor.Cells[1, i] + ' ' + sid
              else
                GridListOfServices.Cells[0, GridListOfServices.RowCount - 1] :=
                  Grid4PriceEditor.Cells[1, i];
              GridListOfServices.Cells[1, GridListOfServices.RowCount - 1] :=
                Grid4PriceEditor.Cells[2, i];
              GridListOfServices.Cells[2, GridListOfServices.LastRow] := sid;
              GridListOfServices.AddRow;
              break;
            end;
          end;
    until s.Contains('<tr>') = false;
end;

procedure TMainForm.GetPeriodicServices(s: string);
var
  sid: string;
  service: string;
  i: Integer;
begin
  s := Pars(s, 'Периодические:', '');
  s := Pars(s, '<tbody>', '</tbody>');
  if s <> '' then
    repeat
      s := Pars(s, '<tr>', '');
      sid := Pars(s, 'srid=', '');
      sid := Pars(sid, '>', '<');
      s := Pars(s, '</a></td>', '');
      service := Pars(s, '">', '<').Trim;
      for i := 1 to Grid4PriceEditor.RowCount - 1 do
        if service = Grid4PriceEditor.Cells[4, i] then
        begin
          if chkAddId.Checked then
            GridListOfServices.Cells[0, GridListOfServices.RowCount - 1] :=
              Grid4PriceEditor.Cells[1, i] + ' ' + sid
          else
            GridListOfServices.Cells[0, GridListOfServices.RowCount - 1] :=
              Grid4PriceEditor.Cells[1, i];
          GridListOfServices.Cells[1, GridListOfServices.RowCount - 1] :=
            Grid4PriceEditor.Cells[2, i];
          if i <> Grid4PriceEditor.RowCount - 1 then
          begin
            GridListOfServices.Cells[2, GridListOfServices.LastRow] := sid;
            GridListOfServices.AddRow;
          end;
          break;
        end;
    until s.Contains('<tr>') = false;
end;

procedure TMainForm.ShowPreloader;
begin
  advpgcntrl1.Hide;
  edtContract.Hide;
  pnlStatus.Left := (MainForm.ClientWidth - pnlStatus.Width) div 2;
  pnlStatus.Top := (MainForm.ClientHeight - pnlStatus.Height -
    (pnlStatus.Height div 3)) div 2;
  pnlStatus.Show;
  Gif.Animate := True;
end;

procedure TMainForm.HidePreloader;
begin
  Gif.Animate := false;
  advpgcntrl1.Show;
  edtContract.Show;
  pnlStatus.Hide;
end;

procedure TMainForm.GridAllServicesBeforeExpandNode(Sender: TObject;
  ARow, ARowreal: Integer; var Allow: Boolean);
begin
  GridAllServices.BeginUpdate;
end;

procedure TMainForm.GridAllServicesClickCell(Sender: TObject;
  ARow, ACol: Integer);
var
  Grid: TAdvStringGrid;
begin
  Grid := Sender as TAdvStringGrid;
  if Grid.IsNode(ARow) = false then
  begin
    if ARow > 0 then
    begin
      edtServiceName.SetFocus;
      edtServiceName.Text := Grid.Cells[1, Grid.Row];
      GridListOfServices.Cells[1, GridListOfServices.Row] :=
        Grid.Cells[2, Grid.Row];
      edtServiceName.SelLength := Length(edtServiceName.Text);
      pnl3Price.Height := 23;
      pnl3Price.BorderStyle := bsNone;
      Grid.Hide;
    end
  end
  else if ACol > 0 then
    if Grid.IsNode(ARow) = True then
    begin
      if Grid.NodeState[Grid.Row] = True then
        Grid.ExpandNode(Grid.RealRow)
      else
        Grid.ContractNode(Grid.RealRow);
    end;
end;

procedure TMainForm.GridAllServicesExpandNode(Sender: TObject;
  ARow, ARowreal: Integer);
begin
  GridAllServices.EndUpdate;
end;

procedure TMainForm.GridAllServicesKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key in [VK_DOWN, VK_UP] = false then
  begin
    edtServiceName.SetFocus;
    edtServiceName.SelStart := Length(edtServiceName.Text);
    keybd_event(Key, 0, 0, 0);
    keybd_event(Key, 0, KEYEVENTF_KEYUP, 0);
    Key := 0;
  end;
end;

end.
