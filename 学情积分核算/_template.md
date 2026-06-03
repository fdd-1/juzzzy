Usage:
  lark-cli docs [flags]
  lark-cli docs [command]

Available Commands:
  +create            Create a Lark document
  +fetch             Fetch Lark document content
  +media-download    Download document media or whiteboard thumbnail (auto-detects extension)
  +media-insert      Insert a local image or file into a Lark document (4-step orchestration + auto-rollback); appends to end by default, or inserts relative to a text selection with --selection-with-ellipsis
  +media-preview     Preview document media file (auto-detects extension)
  +media-upload      Upload media file (image/attachment) to a document block
  +search            Search Lark docs, Wiki, and spreadsheet files (Search v2: doc_wiki/search)
  +update            Update a Lark document
  +whiteboard-update Update an existing whiteboard in lark document with mermaid, plantuml or whiteboard dsl. refer to lark-whiteboard skill for more details.

Flags:
      --api-version string   show docs help for API version (v1|v2)
  -h, --help                 help for docs

Use "lark-cli docs [command] --help" for more information about a command.

node.exe : Error: unknown flag: --doc
At C:\Users\fengjianyi\AppData\Roaming\npm\lark-cli.ps1:24 char:5
+     & "node$exe"  "$basedir/node_modules/@larksuite/cli/scripts/run.j ...
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : NotSpecified: (Error: unknown flag: --doc:String) [], RemoteException
    + FullyQualifiedErrorId : NativeCommandError
 
