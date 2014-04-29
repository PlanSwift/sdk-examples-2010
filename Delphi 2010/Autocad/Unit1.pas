unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, planswift9_tlb,AutoCAD_TLB,Unit2;

type
  TForm1 = class(TForm)
    PageCBX: TComboBox;
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Processtxt: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  ps:IPlanswift;
   pglst: TStringList;
  slst: TStringList;
  ACad: IACadApplication;
  AMS: IAcadModelSpace;
  Form2:TForm2;
implementation

{$R *.dfm}
uses
comobj;


Procedure populateitemList(itm:IItem; lst:TStringlist; itmType:String);
var
  idx: Integer;
  citm: IItem;
  isItm: Boolean;
begin
  for idx := 0 to itm.ChildCount  - 1 do begin
    citm := itm[idx];
    if citm.GetProperty('Type').ResultAsString = 'Folder' then begin
      //(recursive) if item is a folder load all its childPages
      PopulateItemList(citm,lst,itmType);
    end else begin
       isItm := citm.GetPropertyResultAsBoolean(itmType,false);
       if isItm then
        lst.Add(citm.GUID);

       //(recursive)if this item has child items we will want to include those as well
       if citm.ChildCount > 0 then
        populateItemList(citm, lst, itmType);
    end;

  end;
end;

procedure GetpageSections(pg:String; lst:TStringList);
var
  idx: Integer;
  itm: IItem;
begin
  for idx := 0 to slst.Count - 1 do begin
    itm := ps.GetItem(slst.Strings[idx]);
    if itm.GetPropertyResultAsString('PageGUID','') = pg then
      lst.Add(itm.GUID);
  end;
end;
procedure itemToAcad(itm:IItem);
var
  idx: Integer;

  SX: Double;
  SY: Double;
  aptx: Double;
  apty: Double;
  isclosed: Boolean;

  thisline: IAcadLWPolyline;
  pt: IPoint;
  pline: IAcadPolyline;
  ary: olevariant;
  ipt: IPoint;
begin
//Exclude Count Sections
  if itm.GetProperty('Type').ResultAsString = 'Count Section' then
    Exit;
// Exclude all Items with less the 2 Points
  if itm.PointCount < 2 then
    Exit;
  form1.Processtxt.Text := form1.Processtxt.Text + 'Exporting Item: ' + itm.Name + #13#10;
  SX := ps.GetItem(itm.GetProperty('PageGUID').ResultAsString).GetProperty('ScaleX').ResultAsFloat;
  SY := ps.GetItem(itm.GetProperty('PageGUID').ResultAsString).GetProperty('ScaleY').ResultAsFloat;

  ary := varArrayCreate([0,(itm.PointCount * 2) -1],varDouble);
  for idx := 0 to itm.PointCount - 1 do begin
      pt := itm.GetPoint(idx);
      Ary[idx * 2] := (pt.X / sx) ;
      Ary[idx * 2 + 1] := 0 - (pt.Y / sy) ;
  end;

  thisline := AMS.AddLightWeightPolyline(Ary);
  if ((ps.GetPropertyResultAsBoolean(itm.ParentItem.GUID,'IsArea',false) = true) or (itm.GetProperty('Type').ResultAsString = 'Area Subtract Section')) then
    thisline.Closed := true;
  thisline.color := acGreen;

end;
procedure TForm1.Button1Click(Sender: TObject);
var
  pg: string;
  pgSlst: TStringList;
  idx: Integer;
  Adoc: IAcadDocument;
begin
Processtxt.Text := '';
pgSlst := TStringList.Create;
pg := pglst.Strings[pagecbx.itemIndex];

//Populate Page Sections
GetpageSections(pg,pgSLst);
if pgSlst.Count = 0 then begin
  ShowMessage('Could Not Find Any Sections on Page: Please Select a Different Page');
  PageCBX.SetFocus;
  pgSlst.Free;
  Exit;
end;
 try
  //Load Autoocad if Not already open


try
  Acad := GetActiveOleObject('AutoCad.Application') as IACadApplication;
except
    Form2 := TForm2.Create(self);
    Form2.ProgressBar1.Min := 0;
    Form2.ProgressBar1.Max := 10;
    Form2.Show;
    ACad := CreateOleObject('AutoCad.Application') as IACadApplication;
    for idx := 0 to 10 do begin
      Sleep(500);         //Wait for AutoCad To Respond;
      Form2.ProgressBar1.Position := idx;
      Form2.Update;
    End;
    Form2.Close;
    Form2.Free;
end;

  Application.ProcessMessages;

  if Acad = nil then begin
    ShowMessage('No Object');
    Exit;
  end;
    Acad.Visible := true;
  //Get Active Document and ModelSpace
  Adoc := Acad.ActiveDocument;
  AMS := Adoc.ModelSpace;

  //Cycle Through All Page Items and add them to the model Space
  for idx  := 0 to pgslst.count - 1 do
    itemToAcad(ps.GetItem(pgslst.Strings[idx]));
  Acad.ZoomExtents;
 finally
   pgSlst.Free;
   AMS := nil;
   ADoc := nil;
 end;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
self.Close;
Application.Terminate;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//Close and Free All Applications and Lists
try
if Acad <> nil then
  Acad.Quit;
except
  Acad := nil;
end;
 if PS <> nil then
  PS := nil;
  pglst.Free;
  slst.Free;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  pgPath: WideString;
  tkPath: WideString;
  pgsItem: IItem;
  sItem: IItem;
  idx: Integer;


begin
  //Load Planswift if not already open
  ps := coPlanswift.Create;
  pgPath := ps.Root.FullPath + '\Job\Pages';
  tkPath := ps.Root.FullPath + '\Job\TakeOff';
  //get Items
  pgsItem := ps.GetItem(pgPath);
  sItem := ps.GetItem(tkpath);
  //Create Page and takeoff list
  pglst := TstringList.Create;
  slst := TStringList.Create;
  // load Pages into List
  PopulateItemList(pgsitem,pgLst,'IsPage');
  // load Sections into list
  PopulateItemList(SItem,slst,'IsSection');
  // Load Pages into combobox
  for idx := 0 to pglst.count - 1 do
    pagecbx.Items.Add(ps.GetItem(pglst.Strings[idx]).Name);
end;

end.
