unit VirtualStringTreeHelper;
{Note these Helper Routines work in conjuction to help link the Dataset to the
VirtualTree

YOU MUST First setup you Columns with AddColumnRef-This links your ColumnName(HeaderName)
of your Tree to your FieldName, if you plan on usings GetColumnAsString;

GetColumnAsString assumes that you Load your Dataset with LoadDataset or you
manually call SetupTree and BuildNode;

}

interface

uses
  VirtualTrees, mtCheck, DB, classes;

type

  PDataRec = ^TDataRec;
  TDataRec = record
    PData: Pointer;
  end;

  TData = class(TObject)
    BookMark: string;
    DataSet: TDataset;
    Node: PVirtualNode;
    Tree: TBaseVirtualTree;
    Columns: TStrings;
    Other: Pointer;
  end;


  TVirtualStringTreeHelper = class helper for TBaseVirtualTree
  public
    procedure SetupTree(aTree: TBaseVirtualTree);
    procedure AddColumnRef(aTree: TBaseVirtualTree;aColumns: TStrings; aColumnName: string; aField: TField; aColumnWidth: integer);
    function BuildNode(aTree: TBaseVirtualTree; aParentNode: PVirtualNode;aData: TData): PVirtualNode;
    function GetColumnAsString(aNode: PVirtualNode; aColumnName: string; var aColumnIndex: integer): string; overload;
    function GetColumnAsString(aNode: PVirtualNode; aColumnIndex: integer): string; overload;
    procedure LoadDataSet(aTree: TBaseVirtualTree; aParentNode: PVirtualNode;
      aDataSet: TDataSet; aColumns: TStrings; aDataList: TList);
  end;


implementation

{ TVirtualStringTreeHelper }

procedure TVirtualStringTreeHelper.AddColumnRef(aTree: TBaseVirtualTree;aColumns: TStrings;
  aColumnName: string; aField: TField; aColumnWidth: integer);
var
  a_Column: TVirtualTreeColumn;
begin
 {Note the ColumnName should match Tree.Header.Column[x].Text...this way we can
 find the Column that matches the Field}
  aColumns.AddObject(aColumnName, aField);
  if aTree is TVirtualStringTree then
  begin
    a_Column := TVirtualTreeColumn.Create(TVirtualStringTree(aTree).Header.Columns);
    a_Column.Text := aColumnName;
    a_Column.Width := aColumnWidth;
    TVirtualStringTree(aTree).Header.InValidate(a_Column);
  end;
end;

function TVirtualStringTreeHelper.BuildNode(aTree: TBaseVirtualTree;
  aParentNode: PVirtualNode;aData: TData): PVirtualNode;
var
  a_DataRec: PDataRec;
begin
  Result := aTree.AddChild(aParentNode);
  a_DataRec := aTree.GetNodeData(Result);
  if Assigned(a_DataRec) then
  begin
    a_DataRec.PData := nil;
    aData.Tree := aTree;
    aData.Node := Result;
    a_DataRec.PData := aData;
  end;
end;

function TVirtualStringTreeHelper.GetColumnAsString(aNode: PVirtualNode;
  aColumnIndex: integer): string;
var
  a_Data: TData;
  a_DataRec: PDataRec;
  a_Tree: TBaseVirtualTree;
begin
  a_Tree := TreeFromNode(aNode);
  if (aColumnIndex <> - 1) and Assigned(a_Tree) then
  begin
    a_DataRec := a_Tree.GetNodeData(aNode);
    if Assigned(a_DataRec) then
    begin
      a_Data := a_DataRec.PData;
      if Assigned(a_Data) and Assigned(a_Data.Columns) and (aColumnIndex <= (a_Data.Columns.Count -1)) then
      begin
        if Assigned(TField(a_Data.Columns.Objects[aColumnIndex])) then
          Result := TField(a_Data.Columns.Objects[aColumnIndex]).AsString;
      end;
    end;
  end;
  Result := '';
end;

function TVirtualStringTreeHelper.GetColumnAsString(aNode: PVirtualNode;
  aColumnName: string; var aColumnIndex: integer): string;
var
  a_Data: TData;
  a_DataRec: PDataRec;
  a_Tree: TBaseVirtualTree;
  a_Index: integer;
begin
  a_Tree := TreeFromNode(aNode);
  if Assigned(a_Tree) then
  begin
    a_DataRec := a_Tree.GetNodeData(aNode);
    if Assigned(a_DataRec) then
    begin
      a_Data := TData(a_DataRec.PData);
      if Assigned(a_Data) and Assigned(a_Data.Columns) then
        for a_Index := 0 to a_Data.Columns.Count - 1 do
          if a_Data.Columns[a_Index] = aColumnName then
          begin
            aColumnIndex := a_Index;
            Result := TField(a_Data.Columns.Objects[a_Index]).AsString;
            exit;
          end;
    end;
  end;
  Result := '';
end;


procedure TVirtualStringTreeHelper.LoadDataSet(aTree: TBaseVirtualTree; aParentNode: PVirtualNode;
  aDataSet: TDataSet; aColumns: TStrings; aDataList: TList);
var
  a_Node, a_Parent: PVirtualNode;
  a_Data: TData;
begin
  SetupTree(aTree);
  a_Parent := aParentNode;

  if not Assigned(a_Parent) then
    a_Parent := aTree.RootNode;
  aDataset.First;
  while not aDataset.Eof do
  begin
    a_Data := TData.Create;
    a_Data.Columns := aColumns;
    a_Data.DataSet := aDataset;
    a_Data.BookMark := aDataset.Bookmark;

    a_Node := BuildNode(aTree, a_Parent, a_Data);
    aDataList.Add(a_Node);
    aDataset.Next;
  end;
end;

procedure TVirtualStringTreeHelper.SetupTree(aTree: TBaseVirtualTree);
begin
  if aTree is TVirtualStringTree then
    TVirtualStringTree(aTree).NodeDataSize := SizeOf(TDataRec);
end;

end.
