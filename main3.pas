unit main3;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, ExtCtrls, ImgList, ShellApi;

type
  TForm1 = class(TForm)
    Button1: TButton;
    rgScope: TRadioGroup;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    cbTypeAny: TCheckBox;
    cbTypeDisk: TCheckBox;
    cbTypePrint: TCheckBox;
    cbUsageAll: TCheckBox;
    cbUsageConnectable: TCheckBox;
    cbUsageContainer: TCheckBox;
    NetTree: TTreeView;
    ImageList1: TImageList;
    procedure Button1Click(Sender: TObject);
    procedure NetTreeGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure NetTreeDblClick(Sender: TObject);
    procedure NetTreeCustomDrawItem(Sender: TCustomTreeView;
      Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
  private
    { Private declarations }
  public
    procedure Open_Do_Close_Enum(const ParentNode: TTreeNode;
         ResScope, ResType, ResUsage: DWORD; const NetContainerToOpen: PNetResource);
    function OpenEnum(const NetContainerToOpen: PNetResource;
         ResScope, ResType, ResUsage: DWORD): THandle;
    function EnumResources(const ParentNode: TTreeNode;
         ResScope, ResType, ResUsage: DWORD; hNetEnum: THandle): UINT;
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.Button1Click(Sender: TObject);
var
 ResScope, ResType, ResUsage: dword;
begin
 Button1.Caption:='Поиск сетевых ресурсов. Ждите...';
 Button1.Enabled:=false;
 //
 NetTree.Items.Clear;
 case rgScope.ItemIndex of
  1: ResScope:=RESOURCE_GLOBALNET;
  2: ResScope:=RESOURCE_REMEMBERED;
  else ResScope:=RESOURCE_CONNECTED;
 end;
 ResType:=0;
 if cbTypeAny.Checked
 then ResType:=ResType or RESOURCETYPE_ANY;
 if cbTypeDisk.Checked
 then ResType:=ResType or RESOURCETYPE_DISK;
 if cbTypePrint.Checked
 then ResType:=ResType or RESOURCETYPE_PRINT;
 ResUsage:=0;
 if cbUsageConnectable.Checked
 then ResUsage:=ResUsage or RESOURCEUSAGE_CONNECTABLE;
 if cbUsageContainer.Checked
 then ResUsage:=ResUsage or RESOURCEUSAGE_CONTAINER;
 Open_Do_Close_Enum(NetTree.Items.Add(nil, 'Network Resources'),
                               ResScope, ResType, ResUsage, nil);
 //
 Button1.Caption:='Обновить список ресурсов';
 Button1.Enabled:=true;
end;


procedure TForm1.Open_Do_Close_Enum(const ParentNode: TTreeNode;
ResScope, ResType, ResUsage: DWORD; const NetContainerToOpen: PNetResource);
var
 hNetEnum: THandle;
begin
 hNetEnum:=OpenEnum(NetContainerToOpen, ResScope, ResType, ResUsage);
 if (hNetEnum=0)
 then Exit;
 EnumResources(ParentNode, ResScope, ResType, ResUsage, hNetEnum);
 if (NO_ERROR<>WNetCloseEnum(hNetEnum))
 then ShowMessage('WNetCloseEnum Error');
end;

function TForm1.OpenEnum(const NetContainerToOpen: PNetResource;
                         ResScope, ResType, ResUsage: DWORD): THandle;
var
 hNetEnum: THandle;
begin
 Result:=0;
 if (NO_ERROR<>WNetOpenEnum(ResScope, ResType, ResUsage,
                                 NetContainerToOpen, hNetEnum))
 then ShowMessage('Error!')
 else Result:=hNetEnum;
end;


function TForm1.EnumResources(const ParentNode: TTreeNode;
ResScope, ResType, ResUsage: DWORD; hNetEnum: THandle): UINT;
 function ShowResource(const ParentNode: TTreeNode; Res: TNetResource): TTreeNode;
  begin
   Result:=NetTree.Items.AddChild(ParentNode, string(Res.lpRemoteName));
  end;

const
 RESOURCE_BUF_ENTRIES = 2000;

var
 ResourceBuffer: array[1..RESOURCE_BUF_ENTRIES] of TNetResource;
 i, ResourceBuf, EntriesToGet: dword;
 NewNode: TTreeNode;
begin
 Result:=0;
 while true do
  begin
   ResourceBuf:=sizeof(ResourceBuffer);
   EntriesToGet:=RESOURCE_BUF_ENTRIES;
   if (NO_ERROR<>WNetEnumResource(hNetEnum, EntriesToGet,
                              @ResourceBuffer, ResourceBuf))
   then
    begin
     case GetLastError() of
      NO_ERROR:  // Drop out of the switch, walk the buffer
      Break;
      ERROR_NO_MORE_ITEMS:
          // Return with 0 code because this only happens when we got
          // RESOURCE_BUF_ENTRIES entries on the previous call to
          // WNetEnumResource, and there were coincidentally exactly
          // RESOURCE_BUF_ENTRIES entries total in the enum at the time of
          // that previous call
      Exit;
      else ShowMessage('Error!');
      Result:=1;
      Exit;
     end;
    end;
   for i:=1 to EntriesToGet do
    begin
     NewNode:=ShowResource(ParentNode, ResourceBuffer[i]);
     if (ResourceBuffer[i].dwUsage and RESOURCEUSAGE_CONTAINER)<>0
     then Open_Do_Close_Enum(NewNode, ResScope, ResType, ResUsage, @ResourceBuffer[i]);
     Application.ProcessMessages;
    end;
  end;
end;

(*
function TForm1.EnumNetResources(const ParentNode: TTreeNode;
ResScope, ResType, ResUsage: DWORD; const NetResource: PNetResource): boolean;
 function AddNetResource(const Res: TNetResource): TTreeNode;
  begin
   Result:=NetTree.Items.AddChild(ParentNode, Res.lpRemoteName);
  end;
var
 EnumHandle: THandle;
 Res, ResEnum: DWORD;
 LocalNetResource: array[1..2000] of TNetResource;
 i, NumEntries, BufSize: DWORD;
 NewNode: TTreeNode;
begin
 Result:=false;
 Res:=WNetOpenEnum(ResScope, ResType, ResUsage, NetResource, EnumHandle);
 if Res=NO_ERROR then begin
  repeat
   BufSize:=SizeOf(LocalNetResource);
   NumEntries:=2000; // as many as possible.
   ResEnum:=WNetEnumResource(EnumHandle, NumEntries, @LocalNetResource, BufSize);
   if ResEnum=NO_ERROR
   then
    begin
   for i:=1 to NumEntries do
    begin
   // Range checking *has* to be off for this to work.
   {$IFOPT R+} {$DEFINE RANGECHECKWASON} {$R-} {$ENDIF}
   NewNode:=AddNetResource(LocalNetResource[i]);
   if ((LocalNetResource[i].dwUsage and RESOURCEUSAGE_CONTAINER)<>0) then
    begin
     // Danger Will Robinson, recursion in progress in recursion in progress in ...
     EnumNetResources(NewNode, ResScope, ResType, ResUsage, @LocalNetResource[i]);
    end;
   {$IFDEF RANGECHECKWASON} {$R+} {$UNDEF RANGECHECKWASON} {$ENDIF}
   end;
  end
 else if ResEnum<>ERROR_NO_MORE_ITEMS then
  begin
   ShowMessage(Format('WNetEnumResource Error: %d', [Res]));
   break;
  end;
  until ResEnum = ERROR_NO_MORE_ITEMS;
  end
  else
   begin
    ShowMessage(Format('WNetOpenEnum Error: %d', [Res]));
   end;
  Res:=WNetCloseEnum(EnumHandle);
  if(Res<>NO_ERROR)
  then ShowMessage(Format('WNetCloseEnum Error: %d', [Res]))
  else Result:=true;
end;
*)
procedure TForm1.NetTreeGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
 if Node.HasChildren
 then Node.ImageIndex:=1
 else Node.ImageIndex:=0;
end;

procedure TForm1.NetTreeDblClick(Sender: TObject);
begin
 try
  ShellExecute(0,'open',PChar(NetTree.Selected.Text),'','',SW_SHOW);
 except

 end;
end;

procedure TForm1.NetTreeCustomDrawItem(Sender: TCustomTreeView;
  Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
 if cdsSelected in State
 then Sender.Canvas.Font.Style:=Sender.Canvas.Font.Style+[fsUnderline];
end;

end.



