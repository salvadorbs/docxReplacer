unit docxreplacerapp;

{$mode delphi}

interface

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  Classes, SysUtils, CustApp, zipper, XMLRead, DOM, fpjson, jsonparser, FileUtil,
  StrUtils, LazFileUtils, XMLWrite;

type

  { TDocxReplacerApp }

  TDocxReplacerApp = class(TCustomApplication)
  private
    FDocxPath: string;
    FPlaceholdersPath: string;
    FOutputPath: string;
    FXMLDocument: TXMLDocument;

    function GetDocxPath: string;
    function GetOutputPath: string;
    function GetInternalDocXmlPath: string;
    function GetTempPath: string;
    procedure UncompressAndReadDocXml;
    procedure UncompressDocx;
    function GetXMLDocument: TXMLDocument;
    procedure IteratePlaceholders;
    procedure IterateXMLDocument(APlaceholder, AReplacement: string);
    procedure ReplacePlaceholder(Node: TDOMNode;
      APlaceholder, AReplacement: string);
    procedure SaveDocx;
    procedure CleanTemp;
  protected
    procedure DoRun; override;
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;
    procedure WriteHelp; virtual;

    property XMLDocument: TXMLDocument read GetXMLDocument;
    property DocxPath: string read GetDocxPath;
    property TempPath: string read GetTempPath;
    property OutputPath: string read GetOutputPath;
    property InternalDocXmlPath: string read GetInternalDocXmlPath;
  end;

implementation

{ TDocxReplacerApp }

procedure TDocxReplacerApp.UncompressDocx;
var
  UnZip: TUnZipper;
begin
  UnZip := TUnZipper.Create;
  try
    UnZip.FileName := DocxPath;
    UnZip.OutputPath := TempPath;
    UnZip.Examine;
    UnZip.UnZipAllFiles;
  finally
    UnZip.Free;
  end;
end;

function TDocxReplacerApp.GetDocxPath: string;
begin
  Result := ConcatPaths([Location, FDocxPath]);
end;

function TDocxReplacerApp.GetOutputPath: string;
begin
  Result := ExpandFileName(FOutputPath);
end;

function TDocxReplacerApp.GetInternalDocXmlPath: string;
begin
  Result := ConcatPaths([Location, 'temp', 'word/document.xml']);
end;

function TDocxReplacerApp.GetTempPath: string;
begin
  Result := ExpandFileName('temp');

  ForceDirectories(Result);
end;

procedure TDocxReplacerApp.UncompressAndReadDocXml;
begin
  if not (Assigned(FXMLDocument)) then
  begin
    UncompressDocx;

    // Read the contents of the XML file into the TXMLDocument object
    ReadXMLFile(FXMLDocument, InternalDocXmlPath);
  end;
end;

function TDocxReplacerApp.GetXMLDocument: TXMLDocument;
begin
  Result := FXMLDocument;
end;

procedure TDocxReplacerApp.IteratePlaceholders;
var
  Placeholder, Replacement: string;
  i: integer;
  json: TJSONData;
  jsonArray: TJSONArray;
  jsonFile: TFileStream;
begin
  // Load the JSON file containing the placeholder data
  jsonFile := TFileStream.Create(FPlaceholdersPath, fmOpenRead);
  json := GetJSON(jsonFile);
  try
    // Get the number of placeholders in the JSON file
    jsonArray := json.FindPath('placeholder') as TJSONArray;

    // Iterate through all the placeholders in the JSON file
    for i := 0 to jsonArray.Count - 1 do
    begin
      // Get the placeholder and replacement text from the JSON file
      Placeholder := jsonArray.Items[i].FindPath('name').AsString;
      Replacement := jsonArray.Items[i].FindPath('value').AsString;

      // Replace all occurrences of the placeholder with the replacement text
      IterateXMLDocument(Placeholder, Replacement);
    end;
  finally
    json.Free;
    jsonFile.Free;
  end;
end;

procedure TDocxReplacerApp.IterateXMLDocument(APlaceholder, AReplacement: string);
var
  iNode: TDOMNode;

  procedure ProcessNode(Node: TDOMNode);
  var
    cNode: TDOMNode;
    s: string;
  begin
    if Node = nil then Exit; // Stops if reached a leaf

    ReplacePlaceholder(Node, APlaceholder, AReplacement);

    // Goes to the child node
    cNode := Node.FirstChild;

    // Processes all child nodes
    while cNode <> nil do
    begin
      ProcessNode(cNode);
      cNode := cNode.NextSibling;
    end;
  end;

begin
  Assert(Assigned(Self.XMLDocument));

  // Iterate through all the nodes in the XML document
  iNode := Self.XMLDocument.DocumentElement.FirstChild;
  while iNode <> nil do
  begin
    ProcessNode(iNode); // Recursive
    iNode := iNode.NextSibling;
  end;
end;

procedure TDocxReplacerApp.ReplacePlaceholder(Node: TDOMNode;
  APlaceholder, AReplacement: string);
begin
  // Check if the node is a text node
  if Node.NodeType = TEXT_NODE then
  begin
    // Replace the placeholder text with the replacement text
    if Pos(APlaceholder, Node.TextContent) > 0 then
      Node.TextContent := StringReplace(Node.TextContent, APlaceholder,
        AReplacement, [rfReplaceAll]);
  end;
end;

procedure TDocxReplacerApp.SaveDocx;
var
  AZipper: TZipper;
  szPathEntry: string;
  i: integer;
  ZEntries: TZipFileEntries;
  TheFileList: TStringList;
  RelativeDirectory: string;
begin
  WriteXMLFile(FXMLDocument, InternalDocXmlPath);

  AZipper := TZipper.Create;
  try
  try
    AZipper.Filename := Self.OutputPath;
    RelativeDirectory := TempPath;
    AZipper.Clear;
    ZEntries := TZipFileEntries.Create(TZipFileEntry);
    // Verify valid directory
    if DirPathExists(RelativeDirectory) then
    begin
      // Construct the path to the directory BELOW RelativeDirectory
      szPathEntry := IncludeTrailingPathDelimiter(RelativeDirectory);

      // Use the FileUtils.FindAllFiles function to get everything (files and folders) recursively
      TheFileList := TStringList.Create;
      try
        FindAllFiles(TheFileList, RelativeDirectory);
        for i := 0 to TheFileList.Count - 1 do
        begin
          // Make sure the RelativeDirectory files are not in the root of the ZipFile
          ZEntries.AddFileEntry(TheFileList[i], CreateRelativePath(
            TheFileList[i], szPathEntry));
        end;
      finally
        TheFileList.Free;
      end;
    end;
    if (ZEntries.Count > 0) then
      AZipper.ZipFiles(ZEntries);
  except
    On E: EZipError do
      E.CreateFmt('Zipfile could not be created%sReason: %s',
        [LineEnding, E.Message])
  end;
  finally
    FreeAndNil(ZEntries);
    AZipper.Free;
  end;
end;

procedure TDocxReplacerApp.CleanTemp;
begin
  DeleteDirectory(Self.TempPath, True);
end;

procedure TDocxReplacerApp.DoRun;
var
  ErrorMsg: string;
begin
  // quick check parameters
  ErrorMsg := CheckOptions('h', 'help');
  if ErrorMsg <> '' then
  begin
    ShowException(Exception.Create(ErrorMsg));
    Terminate;
    Exit;
  end;

  // parse parameters
  if HasOption('h', 'help') then
  begin
    WriteHelp;
    Terminate;
    Exit;
  end;

  // parse parameters
  if HasOption('d', 'docx') then
  begin
    FDocxPath := GetOptionValue('i', 'inputdocx');
  end;

  // parse parameters
  if HasOption('p', 'placeholders') then
  begin
    FPlaceholdersPath := GetOptionValue('p', 'placeholders');
  end;

  // parse parameters
  if HasOption('o', 'outputfile') then
  begin
    FOutputPath := GetOptionValue('o', 'outputdocx');
  end;

  FDocxPath := 'test.docx';
  FPlaceholdersPath := 'test.json';
  FOutputPath := 'newtest.docx';

  UncompressAndReadDocXml;
  IteratePlaceholders;
  SaveDocx;
  CleanTemp;

  // stop program loop
  Terminate;
end;

constructor TDocxReplacerApp.Create(TheOwner: TComponent);
begin
  inherited Create(TheOwner);
  StopOnException := True;
end;

destructor TDocxReplacerApp.Destroy;
begin
  inherited Destroy;
  FXMLDocument.Free;
end;

procedure TDocxReplacerApp.WriteHelp;
begin
  { add your help code here }
  writeln('Usage: ', ExeName, ' -h');
end;

end.
