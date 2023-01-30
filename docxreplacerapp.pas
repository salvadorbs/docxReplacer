unit docxreplacerapp;

{$mode delphi}

interface

uses
  Classes, SysUtils, CustApp, Zipper, XMLRead, DOM, fpjson, FileUtil, LazFileUtils,
  XMLWrite, rcmdline, jsonparser;

type

  { TDocxReplacerApp }

  TDocxReplacerApp = class(TCustomApplication)
  private
    FCommandLineReader: TCommandLineReader;
    FInputFilePath: string;
    FPlaceholdersPath: string;
    FOutputPath: string;
    FXMLDocument: TXMLDocument;

    function GetParamFile(AParam: String): String;
    function GetInputFilePath: string;
    function GetOutputPath: string;
    function GetInternalDocXmlPath: string;
    function GetPlaceholdersPath: string;
    function GetTempPath: string;
    procedure UncompressDocx;
    procedure IteratePlaceholders;
    procedure IterateXMLDocument(APlaceholder, AReplacement: string);
    procedure ReplacePlaceholder(Node: TDOMNode;
      APlaceholder, AReplacement: string);
    procedure SaveDocx;
    procedure CleanTemp;

    procedure DeclareParams;
    function ParseParams: boolean;
    procedure WriteHelp;
  protected
    procedure DoRun; override;
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;

    property InputFilePath: string read GetInputFilePath;
    property PlaceholdersPath: string read GetPlaceholdersPath;
    property TempPath: string read GetTempPath;
    property OutputPath: string read GetOutputPath;
    property InternalDocXmlPath: string read GetInternalDocXmlPath;
  end;

const
  PARAM_INPUT_DOC = 'inputDoc';
  PARAM_OUTPUT_DOC = 'outputDoc';
  PARAM_TOKENS_JSON = 'placeholdersJson';

implementation

{ TDocxReplacerApp }

procedure TDocxReplacerApp.UncompressDocx;
var
  UnZip: TUnZipper;
begin
  UnZip := TUnZipper.Create;
  try
    UnZip.FileName := InputFilePath;
    UnZip.OutputPath := TempPath;
    UnZip.Examine;
    UnZip.UnZipAllFiles;
  finally
    UnZip.Free;
  end;
end;

function TDocxReplacerApp.GetInputFilePath: string;
begin
  Result := CleanAndExpandFilename(FInputFilePath);
end;

function TDocxReplacerApp.GetParamFile(AParam: String): String;
var
  flagError: Boolean;
begin
  Result := FCommandLineReader.readString(AParam);
  flagError := (not FCommandLineReader.existsProperty(AParam))
    or (Result = '');

  if flagError then
    WriteLn('Error: Parameter ' + AParam + ' is mandatory!' + LineEnding)
  else
    if not(FileExists(CleanAndExpandFilename(Result))) then
    begin
      Result := '';
      WriteLn('Error: Filename ' + AParam + ' is not found!' + LineEnding);
    end;
end;

function TDocxReplacerApp.GetOutputPath: string;
begin
  Result := ExpandFileName(FOutputPath);
end;

function TDocxReplacerApp.GetInternalDocXmlPath: string;
begin
  if (LowerCase(ExtractFileExt(FInputFilePath)) = '.docx') then
    Result := ConcatPaths([Location, 'temp', 'word', 'document.xml'])
  else
    Result := ConcatPaths([Location, 'temp', 'content.xml']);
end;

function TDocxReplacerApp.GetPlaceholdersPath: string;
begin
  Result := CleanAndExpandFilename(FPlaceholdersPath);
end;

function TDocxReplacerApp.GetTempPath: string;
begin
  Result := ExpandFileName('temp');

  ForceDirectories(Result);
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
  jsonFile := TFileStream.Create(PlaceholdersPath, fmOpenRead);
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
  Assert(Assigned(Self.FXMLDocument));

  // Iterate through all the nodes in the XML document
  iNode := Self.FXMLDocument.DocumentElement.FirstChild;
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

procedure TDocxReplacerApp.DeclareParams;
begin
  FCommandLineReader.declareFile(PARAM_INPUT_DOC, 'Filepath to input docx/odt');
  FCommandLineReader.addAbbreviation('i');

  FCommandLineReader.declareFile(PARAM_TOKENS_JSON, 'Filepath to placeholders json file');
  FCommandLineReader.addAbbreviation('p');

  FCommandLineReader.declareFile(PARAM_OUTPUT_DOC, 'Filepath to output docx/odt', 'newDocx.docx');
  FCommandLineReader.addAbbreviation('o');
end;

function TDocxReplacerApp.ParseParams: boolean;
begin
  Result := False;
  try
    try
      FCommandLineReader.parse();

      FInputFilePath := GetParamFile(PARAM_INPUT_DOC);
      FPlaceholdersPath := GetParamFile(PARAM_TOKENS_JSON);
      FOutputPath := FCommandLineReader.readString(PARAM_OUTPUT_DOC);
    except
      on E: Exception do
        WriteLn(E.Message + LineEnding);
    end;

  finally
    Result := (FInputFilePath <> '') and (FPlaceholdersPath <> '');
  end;
end;

procedure TDocxReplacerApp.WriteHelp;
begin
  WriteLn('The following command line options are valid: ' + LineEnding +
    LineEnding + FCommandLineReader.availableOptions);
end;

procedure TDocxReplacerApp.DoRun;
var
  FileDocPath: String;
begin
  DeclareParams;

  if not ParseParams then
  begin
    WriteHelp;
    Terminate;
    Exit;
  end;

  UncompressDocx;
  try
    //Read xml in FXMLDocument
    FileDocPath := InternalDocXmlPath;
    if (FileExists(FileDocPath)) then
    begin
      ReadXMLFile(FXMLDocument, InternalDocXmlPath);

      //Iterate placeholders.json tokens to replace them in xml
      IteratePlaceholders;

      //Compress again in a docx file
      SaveDocx;

      CleanTemp;
    end
    else begin
      WriteLn('Error: File document xml [' + InternalDocXmlPath + '] not found!');
      Terminate;
      Exit;
    end;
  finally
    FXMLDocument.Free;
  end;

  // stop program loop
  Terminate;
end;

constructor TDocxReplacerApp.Create(TheOwner: TComponent);
begin
  inherited Create(TheOwner);
  StopOnException := True;

  FCommandLineReader := TCommandLineReader.Create;
  //showErrorAutomatically = false because we won't halt when find wrong param
  FCommandLineReader.showErrorAutomatically := False;
end;

destructor TDocxReplacerApp.Destroy;
begin
  inherited Destroy;
  FCommandLineReader.Free;
end;

end.
