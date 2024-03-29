'! A simple interactive VBScript shell.
'!
'! @see http://www.kryogenix.org/days/2004/04/01/interactiveVbscript
'!
'! For my own convenience I added line-continuation, some helper functions (cd,
'! pwd, ls, import), and a help message.

Option Explicit

Private Const InitScript    = "%USERPROFILE%\init.vbs"
Private Const Documentation = "script56.chm"

Private helpKeywords, helpfile

Set helpKeywords = CreateObject("Scripting.Dictionary")
	helpKeywords.CompareMode = vbTextCompare
	helpKeywords.Add "-", "html/bf104d71-8802-4335-8098-243ed0f33d59.htm"
	helpKeywords.Add "&", "html/5fe0cfd1-130b-4977-8700-fe599864017d.htm"
	helpKeywords.Add "*", "html/24dd4043-d87e-4611-9d5d-49fef5e8764c.htm"
	helpKeywords.Add "/", "html/f5ac59e8-5a7c-432b-b75b-0d0f6f99f78e.htm"
	helpKeywords.Add "\", "html/af6bebf3-387b-46c9-959c-e32f992d90be.htm"
	helpKeywords.Add "^", "html/26663910-58de-4cde-93a1-cb1ddc214b37.htm"
	helpKeywords.Add "+", "html/e18860b9-78f5-47b0-b124-ee316aed1a9c.htm"
	helpKeywords.Add "=", "html/daf7070e-6669-4d35-ab24-ea5aa9012142.htm"
	helpKeywords.Add "Abs", "html/d26a4c53-b8b4-41f0-a8f9-cb5c2dfa99f7.htm"
	helpKeywords.Add "Add Dictionary", "html/45c14ad6-fa79-4ec7-a567-9ec5ffbbcf52.htm"
	helpKeywords.Add "Add Folder", "html/cdf82cf0-f5c5-4ab7-87de-e436b3d537df.htm"
	helpKeywords.Add "Add Folders", "html/cdf82cf0-f5c5-4ab7-87de-e436b3d537df.htm"
	helpKeywords.Add "Add", "html/45c14ad6-fa79-4ec7-a567-9ec5ffbbcf52.htm"
	helpKeywords.Add "Addition Operator", "html/e18860b9-78f5-47b0-b124-ee316aed1a9c.htm"
	helpKeywords.Add "Addition", "html/e18860b9-78f5-47b0-b124-ee316aed1a9c.htm"
	helpKeywords.Add "AddPrinterConnection", "html/672f0518-a698-44b9-9f94-2a0bb201e16a.htm"
	helpKeywords.Add "AddWindowsPrinterConnection", "html/999614b9-1d1e-4d7a-af43-e22ab22fb782.htm"
	helpKeywords.Add "And Operator", "html/631fdafa-38a8-482b-84ca-479230737220.htm"
	helpKeywords.Add "AppActivate", "html/2b9476ce-54a7-4a00-b761-25bf9f36e83f.htm"
	helpKeywords.Add "Argument Count", "html/d621189c-d562-4ae1-970a-c15536e5b78d.htm"
	helpKeywords.Add "Argument Exists", "html/515c00ed-f035-4744-8f9f-7456feef7f3e.htm"
	helpKeywords.Add "Argument length", "html/f73ceff9-4651-41af-a953-1bd432c9a899.htm"
	helpKeywords.Add "Argument Shortcut", "html/15b932ee-9d9d-41f0-81ed-f6575e0d067b.htm"
	helpKeywords.Add "Argument", "html/dfa1ed1c-d88e-415a-b2cc-958eb6a12c38.htm"
	helpKeywords.Add "Arguments Count", "html/d621189c-d562-4ae1-970a-c15536e5b78d.htm"
	helpKeywords.Add "Arguments Exists", "html/515c00ed-f035-4744-8f9f-7456feef7f3e.htm"
	helpKeywords.Add "Arguments length", "html/f73ceff9-4651-41af-a953-1bd432c9a899.htm"
	helpKeywords.Add "Arguments Shortcut", "html/15b932ee-9d9d-41f0-81ed-f6575e0d067b.htm"
	helpKeywords.Add "Arguments WScript", "html/dfa1ed1c-d88e-415a-b2cc-958eb6a12c38.htm"
	helpKeywords.Add "Arguments", "html/dfa1ed1c-d88e-415a-b2cc-958eb6a12c38.htm"
	helpKeywords.Add "Arithmetic Operators", "html/877e8d67-610f-4b68-8d39-1b9a0ecf1b2c.htm"
	helpKeywords.Add "Array", "html/fc7449c3-fec1-4f0d-9166-a7b08bc4bc10.htm"
	helpKeywords.Add "Asc", "html/c847a40b-9a73-4434-8f65-c52c0085b059.htm"
	helpKeywords.Add "ASCII", "html/c60e2712-20e6-40f2-8fe2-cfb74ca6bca1.htm"
	helpKeywords.Add "Assignment Operator", "html/daf7070e-6669-4d35-ab24-ea5aa9012142.htm"
	helpKeywords.Add "Assignment", "html/daf7070e-6669-4d35-ab24-ea5aa9012142.htm"
	helpKeywords.Add "AtEndOfLine File", "html/f5603906-6958-4a64-9eaa-bd99aecb62b7.htm"
	helpKeywords.Add "AtEndOfLine Files", "html/f5603906-6958-4a64-9eaa-bd99aecb62b7.htm"
	helpKeywords.Add "AtEndOfLine Script Host", "html/fc673f11-e815-440e-ad58-74bd42722e30.htm"
	helpKeywords.Add "AtEndOfLine StdIn", "html/fc673f11-e815-440e-ad58-74bd42722e30.htm"
	helpKeywords.Add "AtEndOfLine Stream", "html/f5603906-6958-4a64-9eaa-bd99aecb62b7.htm"
	helpKeywords.Add "AtEndOfLine TextStream", "html/f5603906-6958-4a64-9eaa-bd99aecb62b7.htm"
	helpKeywords.Add "AtEndOfLine Windows Script Host", "html/fc673f11-e815-440e-ad58-74bd42722e30.htm"
	helpKeywords.Add "AtEndOfLine WSH", "html/fc673f11-e815-440e-ad58-74bd42722e30.htm"
	helpKeywords.Add "AtEndOfLine", "html/f5603906-6958-4a64-9eaa-bd99aecb62b7.htm"
	helpKeywords.Add "AtEndOfStream File", "html/0bb47056-1e5b-4d51-9fb3-9fa12d4ec90c.htm"
	helpKeywords.Add "AtEndOfStream Files", "html/0bb47056-1e5b-4d51-9fb3-9fa12d4ec90c.htm"
	helpKeywords.Add "AtEndOfStream Script Host", "html/af97a6d0-1924-420f-af93-848c5b137856.htm"
	helpKeywords.Add "AtEndOfStream StdIn", "html/af97a6d0-1924-420f-af93-848c5b137856.htm"
	helpKeywords.Add "AtEndOfStream Stream", "html/0bb47056-1e5b-4d51-9fb3-9fa12d4ec90c.htm"
	helpKeywords.Add "AtEndOfStream TextStream", "html/0bb47056-1e5b-4d51-9fb3-9fa12d4ec90c.htm"
	helpKeywords.Add "AtEndOfStream Windows Script Host", "html/af97a6d0-1924-420f-af93-848c5b137856.htm"
	helpKeywords.Add "AtEndOfStream WSH", "html/af97a6d0-1924-420f-af93-848c5b137856.htm"
	helpKeywords.Add "AtEndOfStream", "html/0bb47056-1e5b-4d51-9fb3-9fa12d4ec90c.htm"
	helpKeywords.Add "Atn", "html/c9def5d0-c712-475e-81c2-fca5bdfa5b16.htm"
	helpKeywords.Add "Attributes", "html/423ca96b-6877-4268-a6cc-3139e034f88c.htm"
	helpKeywords.Add "AvailableSpace", "html/6c72bec6-7f83-492f-b7f3-6c896d3f9cb9.htm"
	helpKeywords.Add "BuildPath", "html/e8dc898e-95d3-4d9a-b724-52108ad64bd8.htm"
	helpKeywords.Add "BuildVersion", "html/2211622b-f2cf-45a1-8a08-a8c448b0e250.htm"
	helpKeywords.Add "Call", "html/b8097176-5f29-419c-9146-2faf52dba613.htm"
	helpKeywords.Add "Case", "html/91c340af-8ceb-4f46-86fa-7871eefb3b01.htm"
	helpKeywords.Add "CBool", "html/ff782ecb-f15f-4e3c-a28e-43670557de40.htm"
	helpKeywords.Add "CByte", "html/88e74e98-0801-497d-9d7b-6ee0e3e815c3.htm"
	helpKeywords.Add "CCur", "html/50e319b9-ba46-4919-b2a7-a0db3e2e7c37.htm"
	helpKeywords.Add "CDate", "html/3b133369-8b81-4b7a-a88d-597a1c4f7f1a.htm"
	helpKeywords.Add "CDbl", "html/c576a7a2-0a88-4577-abb0-ca3413fa3b1c.htm"
	helpKeywords.Add "Character", "html/ce4afafe-0bc3-485c-b201-6f213a6bf8b1.htm"
	helpKeywords.Add "Chr", "html/a1240025-063b-4a37-9684-f647c2210998.htm"
	helpKeywords.Add "CInt", "html/38eee725-da5b-469b-b3f7-c818a74037c1.htm"
	helpKeywords.Add "Class", "html/31910d85-2ee9-4234-9444-2ef61f669ec5.htm"
	helpKeywords.Add "Class_Initialize", "html/c47f6152-bffc-4d34-bb63-b59b1146d97f.htm"
	helpKeywords.Add "Class_Terminate", "html/e91761c6-cfea-4e42-8955-128869e68275.htm"
	helpKeywords.Add "Clear Method", "html/7653b031-3ff0-477b-a63d-8e1fd774619b.htm"
	helpKeywords.Add "CLng", "html/5a06dda1-fbf2-4126-9a7b-dca4143a50e7.htm"
	helpKeywords.Add "Close File", "html/ad39aab0-8af6-48ac-87e0-f6fb972a33f8.htm"
	helpKeywords.Add "Close Files", "html/ad39aab0-8af6-48ac-87e0-f6fb972a33f8.htm"
	helpKeywords.Add "Close Script Host", "html/14b767c3-1c6f-4d1f-be05-331b11aad04f.htm"
	helpKeywords.Add "Close StdErr", "html/14b767c3-1c6f-4d1f-be05-331b11aad04f.htm"
	helpKeywords.Add "Close StdIn", "html/14b767c3-1c6f-4d1f-be05-331b11aad04f.htm"
	helpKeywords.Add "Close StdOut", "html/14b767c3-1c6f-4d1f-be05-331b11aad04f.htm"
	helpKeywords.Add "Close Stream", "html/ad39aab0-8af6-48ac-87e0-f6fb972a33f8.htm"
	helpKeywords.Add "Close TextStream", "html/ad39aab0-8af6-48ac-87e0-f6fb972a33f8.htm"
	helpKeywords.Add "Close Windows Script Host", "html/14b767c3-1c6f-4d1f-be05-331b11aad04f.htm"
	helpKeywords.Add "Close WSH", "html/14b767c3-1c6f-4d1f-be05-331b11aad04f.htm"
	helpKeywords.Add "Close", "html/ad39aab0-8af6-48ac-87e0-f6fb972a33f8.htm"
	helpKeywords.Add "Color Constants", "html/96025e1a-85cd-43d5-8be0-99971544b612.htm"
	helpKeywords.Add "Column File", "html/22a9298b-7697-4320-88eb-1b7aa0bb8c94.htm"
	helpKeywords.Add "Column Files", "html/22a9298b-7697-4320-88eb-1b7aa0bb8c94.htm"
	helpKeywords.Add "Column Script Host", "html/8271d0a2-e9e2-4fa4-aa66-48374b89ac23.htm"
	helpKeywords.Add "Column StdIn", "html/8271d0a2-e9e2-4fa4-aa66-48374b89ac23.htm"
	helpKeywords.Add "Column Stream", "html/22a9298b-7697-4320-88eb-1b7aa0bb8c94.htm"
	helpKeywords.Add "Column TextStream", "html/22a9298b-7697-4320-88eb-1b7aa0bb8c94.htm"
	helpKeywords.Add "Column Windows Script Host", "html/8271d0a2-e9e2-4fa4-aa66-48374b89ac23.htm"
	helpKeywords.Add "Column WSH", "html/8271d0a2-e9e2-4fa4-aa66-48374b89ac23.htm"
	helpKeywords.Add "Column", "html/22a9298b-7697-4320-88eb-1b7aa0bb8c94.htm"
	helpKeywords.Add "CompareMode", "html/65e3818b-c68f-4fe0-9f74-4738cb7ca931.htm"
	helpKeywords.Add "Comparison Constants", "html/04723862-22a7-4556-b261-efc929620b96.htm"
	helpKeywords.Add "Comparison Operators", "html/adb6e3ac-d925-4987-9d43-f6486d5f1e30.htm"
	helpKeywords.Add "ComputerName", "html/82049413-c8ce-4079-b05e-606849b1c476.htm"
	helpKeywords.Add "Concatenation Operator", "html/5fe0cfd1-130b-4977-8700-fe599864017d.htm"
	helpKeywords.Add "Concatenation Operators", "html/7eab59dc-15c1-411f-9a5f-813bf6f74bec.htm"
	helpKeywords.Add "Concatenation", "html/5fe0cfd1-130b-4977-8700-fe599864017d.htm"
	helpKeywords.Add "ConnectObject", "html/fffd2ade-9fa7-42e8-a0c9-a6a7e631b378.htm"
	helpKeywords.Add "Const", "html/61ad308e-6c17-4cd5-a9a6-8a47dc628ba1.htm"
	helpKeywords.Add "Constants", "html/0d289784-5655-4d16-91dc-732b319d0aea.htm"
	helpKeywords.Add "Conversion Functions", "html/51b7aeaf-e58d-4f7b-b992-ca6cfd0a3fe1.htm"
	helpKeywords.Add "Copy", "html/c98c7d8c-ba9e-46b7-bc29-14ab1b7f06f3.htm"
	helpKeywords.Add "CopyFile", "html/94e39f9a-c7bf-42be-ae71-5768f034e070.htm"
	helpKeywords.Add "CopyFolder", "html/d3695dd5-45ce-410b-ac96-42a92c68344b.htm"
	helpKeywords.Add "Cos", "html/d3e52048-aba6-430b-a512-4cd705e62e3b.htm"
	helpKeywords.Add "Count Argument", "html/d621189c-d562-4ae1-970a-c15536e5b78d.htm"
	helpKeywords.Add "Count Arguments", "html/d621189c-d562-4ae1-970a-c15536e5b78d.htm"
	helpKeywords.Add "Count Dictionary", "html/08130c63-ea3a-4beb-ae84-6e1b4d9f1c82.htm"
	helpKeywords.Add "Count", "html/08130c63-ea3a-4beb-ae84-6e1b4d9f1c82.htm"
	helpKeywords.Add "CreateFolder", "html/481d39c7-dbd3-4cfb-b416-8187d73eb9d6.htm"
	helpKeywords.Add "CreateObject Method", "html/a0d71def-47b9-4797-837c-bef653d958aa.htm"
	helpKeywords.Add "CreateObject", "html/b1545c3f-5cab-4c66-af6f-be9b7daa1131.htm"
	helpKeywords.Add "CreateScript", "html/e06559a3-ac57-41c4-89d8-212519fb64ef.htm"
	helpKeywords.Add "CreateShortcut", "html/d91b9d23-a7e5-4ec2-8b55-ef6ffe9c777d.htm"
	helpKeywords.Add "CreateTextFile", "html/62045015-70d8-4308-a74a-71de0166e6ec.htm"
	helpKeywords.Add "CSng", "html/84d2f785-ba4a-47c3-8989-b5c9d6db036a.htm"
	helpKeywords.Add "CStr", "html/25333070-f879-4904-9fa7-2f2d4e4762ce.htm"
	helpKeywords.Add "CurrentDirectory", "html/a36f684c-efef-4069-9102-21b3d1d55e9e.htm"
	helpKeywords.Add "Date and Time Constants", "html/8c1bfc98-c011-4825-adb1-8d5eb52928d1.htm"
	helpKeywords.Add "Date Constants", "html/8c1bfc98-c011-4825-adb1-8d5eb52928d1.htm"
	helpKeywords.Add "Date Format Constants", "html/bcdaaad6-a693-4f59-8158-59a8554800c6.htm"
	helpKeywords.Add "Date Time Constants", "html/8c1bfc98-c011-4825-adb1-8d5eb52928d1.htm"
	helpKeywords.Add "Date", "html/ab4fe211-df30-4702-9317-5869c5668c6a.htm"
	helpKeywords.Add "DateAdd", "html/f0ff19c4-0a74-44ca-867f-4739a814a88d.htm"
	helpKeywords.Add "DateCreated", "html/b4743178-f759-48de-bc69-9efc4b9f49c7.htm"
	helpKeywords.Add "DateDiff", "html/fe4e5302-9602-4e0f-8496-559c7f58c8fa.htm"
	helpKeywords.Add "DateLastAccessed", "html/685c15d5-ea6b-4b0e-9452-560a8296d586.htm"
	helpKeywords.Add "DateLastModified", "html/a42dfb9f-7355-4fb0-b238-a11a278f20e2.htm"
	helpKeywords.Add "DatePart", "html/4e45cd84-b22b-437f-8410-6f4b8ff1c769.htm"
	helpKeywords.Add "DateSerial", "html/47ae21c7-193d-4216-a571-7deac1027662.htm"
	helpKeywords.Add "DateValue", "html/d859df49-ab8d-4964-ac7e-774e51cf53ca.htm"
	helpKeywords.Add "Day", "html/0cf777f2-635c-4e88-95e7-fcdae164f065.htm"
	helpKeywords.Add "Debug Write", "html/0530c5f1-c079-4d1a-aa42-b3f9bbf74e41.htm"
	helpKeywords.Add "Debug WriteLine", "html/8f5593a4-3abe-49ca-9a81-e96e1607d725.htm"
	helpKeywords.Add "Debug", "html/948adf05-b2d4-4f3a-bdbf-cf281b912eaa.htm"
	helpKeywords.Add "Debug.Write", "html/0530c5f1-c079-4d1a-aa42-b3f9bbf74e41.htm"
	helpKeywords.Add "Debug.WriteLine", "html/8f5593a4-3abe-49ca-9a81-e96e1607d725.htm"
	helpKeywords.Add "Debugger Write", "html/0530c5f1-c079-4d1a-aa42-b3f9bbf74e41.htm"
	helpKeywords.Add "Debugger WriteLine", "html/8f5593a4-3abe-49ca-9a81-e96e1607d725.htm"
	helpKeywords.Add "Delete", "html/bc78e712-a749-4037-a4c8-4aa445c38f09.htm"
	helpKeywords.Add "DeleteFile", "html/502f0ddc-1bbc-4723-a65d-f355fef6555e.htm"
	helpKeywords.Add "DeleteFolder", "html/7991e577-f2ac-4b33-8231-761971d6c0d9.htm"
	helpKeywords.Add "Description Err", "html/44212188-70d3-4356-a7f6-c9902e419e08.htm"
	helpKeywords.Add "Description Error", "html/44212188-70d3-4356-a7f6-c9902e419e08.htm"
	helpKeywords.Add "Description Shortcut", "html/31fd3558-705c-4c37-9a27-1399baf0cf41.htm"
	helpKeywords.Add "Description WshRemoteError", "html/a769de1b-b007-492f-9d72-66053ade2874.htm"
	helpKeywords.Add "Description", "html/44212188-70d3-4356-a7f6-c9902e419e08.htm"
	helpKeywords.Add "Dictionary Exists", "html/157fb2fe-7453-480c-bb8b-3aa5c0b21cfd.htm"
	helpKeywords.Add "Dictionary Remove", "html/a5a8b684-65ab-4995-83fa-c4f22c482321.htm"
	helpKeywords.Add "Dictionary", "html/b4a7ddb3-2474-49ef-8540-8d67a747c8db.htm"
	helpKeywords.Add "Dim", "html/87113e7d-b3d4-41c7-85c3-4b9925586707.htm"
	helpKeywords.Add "DisconnectObject", "html/5c3b18d5-f137-4a93-b9d2-3aa005fa3b2f.htm"
	helpKeywords.Add "Division Operator", "html/f5ac59e8-5a7c-432b-b75b-0d0f6f99f78e.htm"
	helpKeywords.Add "Division", "html/f5ac59e8-5a7c-432b-b75b-0d0f6f99f78e.htm"
	helpKeywords.Add "Do Loop", "html/4d86d9ea-6b84-4f59-a179-296b5087cdd1.htm"
	helpKeywords.Add "Do", "html/4d86d9ea-6b84-4f59-a179-296b5087cdd1.htm"
	helpKeywords.Add "Do...Loop", "html/4d86d9ea-6b84-4f59-a179-296b5087cdd1.htm"
	helpKeywords.Add "Do..Loop", "html/4d86d9ea-6b84-4f59-a179-296b5087cdd1.htm"
	helpKeywords.Add "Drive Path", "html/c988522c-3cb1-462e-ad84-b44f0266bd00.htm"
	helpKeywords.Add "Drive Object", "html/4162395a-e7b9-4e25-87b5-f2f528b7fe4b.htm"
	helpKeywords.Add "Drive", "html/be580244-3425-4fd9-88d0-83b0f2a34d74.htm"
	helpKeywords.Add "DriveExists", "html/28fab0b2-8a2c-424f-b8e0-1efc519f4b7f.htm"
	helpKeywords.Add "DriveLetter", "html/269b9ef4-f08b-4ef6-82db-ee1d7341afd4.htm"
	helpKeywords.Add "Drives Collection", "html/6c1124d7-1bab-477a-9ba6-229e082121d9.htm"
	helpKeywords.Add "Drives", "html/9ea90237-0888-4813-8aeb-7fa8597100e7.htm"
	helpKeywords.Add "DriveType", "html/4ba94046-87fc-41e9-aaed-0d321c55ecb0.htm"
	helpKeywords.Add "Echo", "html/8dac8dcb-d237-4921-b2ba-0c6a34b60265.htm"
	helpKeywords.Add "Else", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "ElseIf", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "EnumNetworkDrive Item", "html/61580f5c-e7f9-49b3-8d53-09cdeae822f7.htm"
	helpKeywords.Add "EnumNetworkDrives", "html/3bf85ca6-9647-448f-a5aa-bdf47a50500b.htm"
	helpKeywords.Add "EnumPrinterConnections Item", "html/61580f5c-e7f9-49b3-8d53-09cdeae822f7.htm"
	helpKeywords.Add "EnumPrinterConnections", "html/a0c66893-4599-405f-911b-5eedb091753a.htm"
	helpKeywords.Add "Environment Item", "html/61580f5c-e7f9-49b3-8d53-09cdeae822f7.htm"
	helpKeywords.Add "Environment length", "html/b8172567-22b9-4496-9cf2-8cd732fbf492.htm"
	helpKeywords.Add "Environment Remove", "html/05426477-f6e4-4dcb-8101-3f7d979c8ed5.htm"
	helpKeywords.Add "Environment", "html/7544483a-b1b3-4b00-bb0e-0d260f1b099a.htm"
	helpKeywords.Add "Eqv", "html/5285cd35-e650-495a-97bb-05499130ab50.htm"
	helpKeywords.Add "Erase", "html/cfd6f862-fcd9-4625-b6f9-ffd31e8e3e86.htm"
	helpKeywords.Add "Err Description", "html/44212188-70d3-4356-a7f6-c9902e419e08.htm"
	helpKeywords.Add "Err Number", "html/35cc596b-e2fa-431b-b6de-31167187b941.htm"
	helpKeywords.Add "Err Source", "html/8e3cd971-4c3c-4f22-9559-775864be71c6.htm"
	helpKeywords.Add "Err", "html/70223b47-3bb6-4b15-b967-f3f8082fdbfe.htm"
	helpKeywords.Add "Err.Description", "html/44212188-70d3-4356-a7f6-c9902e419e08.htm"
	helpKeywords.Add "Err.Number", "html/35cc596b-e2fa-431b-b6de-31167187b941.htm"
	helpKeywords.Add "Err.Source", "html/8e3cd971-4c3c-4f22-9559-775864be71c6.htm"
	helpKeywords.Add "Error Description", "html/44212188-70d3-4356-a7f6-c9902e419e08.htm"
	helpKeywords.Add "Error Number", "html/35cc596b-e2fa-431b-b6de-31167187b941.htm"
	helpKeywords.Add "Error Source", "html/8e3cd971-4c3c-4f22-9559-775864be71c6.htm"
	helpKeywords.Add "Error", "html/6848ab32-569d-49b4-9b5b-343146ceb1de.htm"
	helpKeywords.Add "Errors", "html/58ee8744-f5fe-42cd-a612-e3158682d891.htm"
	helpKeywords.Add "Escape", "html/4200f159-3e9e-4b39-8b02-0048917f29f5.htm"
	helpKeywords.Add "Eval", "html/a4ede5a0-5d7a-4a3e-adea-dad5105d0a2f.htm"
	helpKeywords.Add "Exec Status", "html/6e874fcd-4efd-4891-b098-242509cbc1f9.htm"
	helpKeywords.Add "Exec StdErr", "html/fed33bf2-907f-4043-9900-2cb0da528992.htm"
	helpKeywords.Add "Exec StdIn", "html/b26de977-dcae-4bde-8871-21fa4cf698c2.htm"
	helpKeywords.Add "Exec StdOut", "html/85684a76-6d66-4a1a-a3c4-cf3f48baa595.htm"
	helpKeywords.Add "Exec Terminate", "html/a1ca9cc6-4b46-4190-8333-a03f1cad05cd.htm"
	helpKeywords.Add "Exec", "html/5593b353-ef4b-4c99-8ae1-f963bac48929.htm"
	helpKeywords.Add "Execute RegEx", "html/711116fb-9c47-47cb-b664-db8141b8cc69.htm"
	helpKeywords.Add "Execute RegExp", "html/711116fb-9c47-47cb-b664-db8141b8cc69.htm"
	helpKeywords.Add "Execute Regular Expression", "html/711116fb-9c47-47cb-b664-db8141b8cc69.htm"
	helpKeywords.Add "Execute WshRemote", "html/742524c8-990a-4f07-8130-336803d68f67.htm"
	helpKeywords.Add "Execute", "html/c9cddcbd-2d2b-4139-bb21-c21136f2df81.htm"
	helpKeywords.Add "ExecuteGlobal", "html/25ebfa26-d3b9-4f82-b3c9-a8568a389dbc.htm"
	helpKeywords.Add "Exists Argument", "html/515c00ed-f035-4744-8f9f-7456feef7f3e.htm"
	helpKeywords.Add "Exists Arguments", "html/515c00ed-f035-4744-8f9f-7456feef7f3e.htm"
	helpKeywords.Add "Exists Dictionary", "html/157fb2fe-7453-480c-bb8b-3aa5c0b21cfd.htm"
	helpKeywords.Add "Exists", "html/157fb2fe-7453-480c-bb8b-3aa5c0b21cfd.htm"
	helpKeywords.Add "Exit", "html/20ec6708-0725-488c-8110-2dceaf9755c4.htm"
	helpKeywords.Add "ExitCode", "html/4c5b06ac-dc45-4ec2-aca1-f168bab75483.htm"
	helpKeywords.Add "Exp", "html/7d2561e8-e099-466d-871c-000a440ddf1c.htm"
	helpKeywords.Add "ExpandEnvironmentStrings", "html/c1c1f29c-2a30-46a1-bae7-c17f8af55a19.htm"
	helpKeywords.Add "Exponent", "html/26663910-58de-4cde-93a1-cb1ddc214b37.htm"
	helpKeywords.Add "Exponentiation Operator", "html/26663910-58de-4cde-93a1-cb1ddc214b37.htm"
	helpKeywords.Add "Exponentiation", "html/26663910-58de-4cde-93a1-cb1ddc214b37.htm"
	helpKeywords.Add "File Name", "html/c750e811-79cc-4cc1-a5b6-f6fac0e7043d.htm"
	helpKeywords.Add "File Path", "html/c988522c-3cb1-462e-ad84-b44f0266bd00.htm"
	helpKeywords.Add "File", "html/f88bc0b2-346e-4ed8-8577-e026b4fc976e.htm"
	helpKeywords.Add "FileExists", "html/0f2bdb53-5821-45d7-9044-ef1abc4ddc89.htm"
	helpKeywords.Add "Files Collection", "html/ad23fd3a-5492-4a97-b83d-0ad0804d1e8f.htm"
	helpKeywords.Add "Files", "html/f9356e12-b08e-4771-bbd1-d00d1c143563.htm"
	helpKeywords.Add "FileSystem", "html/105828a4-5b45-4b80-98ae-d3b50c91488b.htm"
	helpKeywords.Add "FileSystemObject", "html/af4423b2-4ee8-41d6-a704-49926cd4d2e8.htm"
	helpKeywords.Add "Filter", "html/9f4d308c-6c2f-46a6-823c-29b9b6a54c56.htm"
	helpKeywords.Add "FirstIndex", "html/6ddce8a8-5fc6-4403-b5cb-5e389715e4e0.htm"
	helpKeywords.Add "Fix", "html/69196bd0-f222-4056-aac5-1989ef3696dd.htm"
	helpKeywords.Add "Folder Name", "html/c750e811-79cc-4cc1-a5b6-f6fac0e7043d.htm"
	helpKeywords.Add "Folder Path", "html/c988522c-3cb1-462e-ad84-b44f0266bd00.htm"
	helpKeywords.Add "Folder", "html/b3e21591-b52e-4dfd-926d-992843eff3ce.htm"
	helpKeywords.Add "FolderExists", "html/db4df1be-1e13-46bd-8a2c-b8ba776cac03.htm"
	helpKeywords.Add "Folders", "html/d562a09d-c530-4f97-a62f-08e0f15b3d66.htm"
	helpKeywords.Add "For Each Next", "html/5920e7da-3ce3-4009-975b-d7aa1b0ea826.htm"
	helpKeywords.Add "For Each", "html/5920e7da-3ce3-4009-975b-d7aa1b0ea826.htm"
	helpKeywords.Add "For Each...Next", "html/5920e7da-3ce3-4009-975b-d7aa1b0ea826.htm"
	helpKeywords.Add "For Each..Next", "html/5920e7da-3ce3-4009-975b-d7aa1b0ea826.htm"
	helpKeywords.Add "For Next", "html/4a7b5b81-bba7-49b1-b176-f6b221df8a1c.htm"
	helpKeywords.Add "For", "html/4a7b5b81-bba7-49b1-b176-f6b221df8a1c.htm"
	helpKeywords.Add "For...Next", "html/4a7b5b81-bba7-49b1-b176-f6b221df8a1c.htm"
	helpKeywords.Add "For..Next", "html/4a7b5b81-bba7-49b1-b176-f6b221df8a1c.htm"
	helpKeywords.Add "FormatCurrency", "html/3fa17fdd-c24d-45b3-bf63-c905c4b9f6d8.htm"
	helpKeywords.Add "FormatDateTime", "html/1d3db34d-159e-45fc-b242-ae5d87d75725.htm"
	helpKeywords.Add "FormatNumber", "html/0f8d2abb-085b-447b-9899-f85521255e4b.htm"
	helpKeywords.Add "FormatPercent", "html/69a085c0-6b18-497b-be40-be4dd16d5f9e.htm"
	helpKeywords.Add "FreeSpace", "html/81ad2760-97db-4945-8db9-9ead7c2ffd96.htm"
	helpKeywords.Add "FullName Script", "html/bb249bc2-e80d-44a3-ac66-16ce03db5d61.htm"
	helpKeywords.Add "FullName Shortcut", "html/64d008bc-53c2-4138-9a3e-ce8f141fc371.htm"
	helpKeywords.Add "FullName URL", "html/416b8d2b-2219-4879-a94a-92d003cbe34d.htm"
	helpKeywords.Add "FullName UrlShortcut", "html/416b8d2b-2219-4879-a94a-92d003cbe34d.htm"
	helpKeywords.Add "FullName WScript", "html/bb249bc2-e80d-44a3-ac66-16ce03db5d61.htm"
	helpKeywords.Add "FullName", "html/bb249bc2-e80d-44a3-ac66-16ce03db5d61.htm"
	helpKeywords.Add "Function", "html/41d6c677-d975-4f1c-954f-d55cc0bfebce.htm"
	helpKeywords.Add "Get Property", "html/05167024-a817-405f-a0ce-2057d01f804a.htm"
	helpKeywords.Add "Get", "html/05167024-a817-405f-a0ce-2057d01f804a.htm"
	helpKeywords.Add "GetAbsolutePathName", "html/8c2c1519-1ca1-49d4-9998-44d721eeb868.htm"
	helpKeywords.Add "GetBaseName", "html/02dc148f-2cc7-480d-abb8-adc6c5faa3a2.htm"
	helpKeywords.Add "GetDrive", "html/552d86b1-91c1-4d6f-9ef1-369876ec11b0.htm"
	helpKeywords.Add "GetDriveName", "html/f298573a-2b91-4116-aa94-a2b07c2f0ee3.htm"
	helpKeywords.Add "GetExtensionName", "html/92bbb0d5-3f9c-4f82-a105-e3c47a2a151d.htm"
	helpKeywords.Add "GetFile", "html/04b19915-6d68-4f3c-ac97-6a221daa8f83.htm"
	helpKeywords.Add "GetFileName", "html/e9878c39-88ae-4091-a392-7f70591982a0.htm"
	helpKeywords.Add "GetFileVersion", "html/45588625-2d2b-4efa-95e1-6cf1feb8f3e7.htm"
	helpKeywords.Add "GetFolder", "html/867057a0-50cf-4613-84be-28bf908b23a7.htm"
	helpKeywords.Add "GetLocale", "html/a7255149-98a6-4559-be7e-a5547e216bb0.htm"
	helpKeywords.Add "GetObject Method", "html/2bf7937a-5386-4a86-9b5f-0a7740bcb0e6.htm"
	helpKeywords.Add "GetObject", "html/5f3671b6-4729-4898-825b-9223765130fd.htm"
	helpKeywords.Add "GetParentFolderName", "html/315b5106-d147-47e2-af13-7605f969d37a.htm"
	helpKeywords.Add "GetRef", "html/f9643b79-b4a1-4f70-9017-2b2e2f01ef3a.htm"
	helpKeywords.Add "getResource", "html/3a03277b-1e24-403c-a0d7-066aa963b112.htm"
	helpKeywords.Add "GetSpecialFolder", "html/328b505e-6dfd-4f4a-b819-250ca46689a1.htm"
	helpKeywords.Add "GetStandardStream", "html/6ae9a1dc-35ae-4e06-94b2-1578ba153fce.htm"
	helpKeywords.Add "GetTempName", "html/f540ca8a-ddb4-4b6a-87dc-1940cc59dffd.htm"
	helpKeywords.Add "Global", "html/34b002c0-91dd-4c49-8742-f0d62fb272cd.htm"
	helpKeywords.Add "HelpContext", "html/0263378d-f517-490a-b68f-79f6740c8e55.htm"
	helpKeywords.Add "HelpFile", "html/245f7dd4-b2f5-4e8e-9af1-a46edf4f0181.htm"
	helpKeywords.Add "Hex", "html/9d1c8e64-74de-45e1-8da2-ac0cdf69ee5a.htm"
	helpKeywords.Add "Hotkey", "html/f732fcf1-2333-4d5a-8c6b-37e5f53da09d.htm"
	helpKeywords.Add "Hour", "html/8890ced0-3578-49b2-ba44-f8b9d58cf56d.htm"
	helpKeywords.Add "IconLocation", "html/68b5a4bb-b235-4123-bb17-580f044da45a.htm"
	helpKeywords.Add "If Then Else", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "If Then", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "If", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "If...Then", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "If...Then...Else", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "If..Then", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "If..Then..Else", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "IgnoreCase", "html/efb59f88-8977-4aad-bc17-134ec0b27224.htm"
	helpKeywords.Add "Imp", "html/9b2eb295-77e3-49f2-b77c-2faa91e0fce8.htm"
	helpKeywords.Add "Initialize", "html/c47f6152-bffc-4d34-bb63-b59b1146d97f.htm"
	helpKeywords.Add "InputBox", "html/c911e626-555a-41c5-8343-5e8b941469c4.htm"
	helpKeywords.Add "InStr", "html/cf0c9b24-ea84-420b-9393-2dbe0c58d7d1.htm"
	helpKeywords.Add "InStrRev", "html/014390e0-5e51-4238-a131-2aa30eed493a.htm"
	helpKeywords.Add "Int", "html/69196bd0-f222-4056-aac5-1989ef3696dd.htm"
	helpKeywords.Add "Integer Division Operator", "html/af6bebf3-387b-46c9-959c-e32f992d90be.htm"
	helpKeywords.Add "Integer Division", "html/af6bebf3-387b-46c9-959c-e32f992d90be.htm"
	helpKeywords.Add "Interactive", "html/6ddd3b66-22d9-4674-9aad-8f3ca9581783.htm"
	helpKeywords.Add "Is", "html/49dafc9c-d1de-4193-8685-1d3fad510189.htm"
	helpKeywords.Add "IsArray", "html/7c36354c-0aa8-4ee6-9f2f-1f3543c0d8ac.htm"
	helpKeywords.Add "IsDate", "html/00eb08e0-457e-4ca4-85b6-167d35c5d7ba.htm"
	helpKeywords.Add "IsEmpty", "html/feb0bd25-8377-42aa-a3c0-3d2c62f83c39.htm"
	helpKeywords.Add "IsNull", "html/c6a16689-8b58-48ec-af0b-41faf888afc1.htm"
	helpKeywords.Add "IsNumeric", "html/467a0ae1-b9b6-4fa0-b2e4-984704033cbb.htm"
	helpKeywords.Add "IsObject", "html/364b2da5-9d01-424f-8b15-838216d3dcad.htm"
	helpKeywords.Add "IsReady", "html/bc05f16d-4465-4b14-b382-9e9bb2f25716.htm"
	helpKeywords.Add "IsRootFolder", "html/2163a5a8-9059-4eef-a44b-6a9efd03ede1.htm"
	helpKeywords.Add "Item Dictionary", "html/8c9581aa-cead-48be-9951-9bc9866c0c55.htm"
	helpKeywords.Add "Item EnumNetworkDrive", "html/61580f5c-e7f9-49b3-8d53-09cdeae822f7.htm"
	helpKeywords.Add "Item EnumPrinterConnections", "html/61580f5c-e7f9-49b3-8d53-09cdeae822f7.htm"
	helpKeywords.Add "Item Environment", "html/61580f5c-e7f9-49b3-8d53-09cdeae822f7.htm"
	helpKeywords.Add "Item Named Argument", "html/f4b72d03-1714-4d0d-81c9-7b5dd6327e50.htm"
	helpKeywords.Add "Item Named Arguments", "html/f4b72d03-1714-4d0d-81c9-7b5dd6327e50.htm"
	helpKeywords.Add "Item SpecialFolders", "html/61580f5c-e7f9-49b3-8d53-09cdeae822f7.htm"
	helpKeywords.Add "Item Unnamed Argument", "html/fa6fae44-a5d6-4eff-b79f-430c6b5c7ee8.htm"
	helpKeywords.Add "Item Unnamed Arguments", "html/fa6fae44-a5d6-4eff-b79f-430c6b5c7ee8.htm"
	helpKeywords.Add "Item WshNamed", "html/f4b72d03-1714-4d0d-81c9-7b5dd6327e50.htm"
	helpKeywords.Add "Item WshUnnamed", "html/fa6fae44-a5d6-4eff-b79f-430c6b5c7ee8.htm"
	helpKeywords.Add "Items", "html/0c15c9e5-8f05-4f1f-a3b7-6bb2c7da3380.htm"
	helpKeywords.Add "Join", "html/cc221524-8f4c-4463-8fa6-578c39a939d4.htm"
	helpKeywords.Add "Key", "html/e01c4c98-21fb-44cc-9a35-f192c47d0d3d.htm"
	helpKeywords.Add "Keys", "html/6ca8d51a-a235-4a68-becb-c08c51ae12df.htm"
	helpKeywords.Add "LBound", "html/6efa4ad0-0941-4349-a680-236c3200794e.htm"
	helpKeywords.Add "LCase", "html/1a3ee93f-f175-4358-b13d-b8f601092174.htm"
	helpKeywords.Add "LCID", "html/882ca1eb-81b6-4a73-839d-154c6440bf70.htm"
	helpKeywords.Add "Left", "html/6ae9b6ff-c2ee-46a9-8cf6-da0d0469a5d8.htm"
	helpKeywords.Add "Len", "html/517758dd-b3ad-47d5-86c0-5889d4ecac2a.htm"
	helpKeywords.Add "length Argument", "html/f73ceff9-4651-41af-a953-1bd432c9a899.htm"
	helpKeywords.Add "length Arguments", "html/f73ceff9-4651-41af-a953-1bd432c9a899.htm"
	helpKeywords.Add "length Environment", "html/b8172567-22b9-4496-9cf2-8cd732fbf492.htm"
	helpKeywords.Add "Length Match", "html/5efd34f5-cb70-42cd-b4b1-86fdf6c6e811.htm"
	helpKeywords.Add "Length RegEx", "html/5efd34f5-cb70-42cd-b4b1-86fdf6c6e811.htm"
	helpKeywords.Add "Length RegExp", "html/5efd34f5-cb70-42cd-b4b1-86fdf6c6e811.htm"
	helpKeywords.Add "Length Regular Expression", "html/5efd34f5-cb70-42cd-b4b1-86fdf6c6e811.htm"
	helpKeywords.Add "length SpecialFolders", "html/12beb9da-a714-4de4-b6cc-b5b9c6fdc5d5.htm"
	helpKeywords.Add "Length", "html/5efd34f5-cb70-42cd-b4b1-86fdf6c6e811.htm"
	helpKeywords.Add "Let Property", "html/23a9ca52-91c1-4257-a090-fdb397380ece.htm"
	helpKeywords.Add "Let", "html/23a9ca52-91c1-4257-a090-fdb397380ece.htm"
	helpKeywords.Add "Line StdErr", "html/1684fcd7-7762-4fb2-acf6-dacfa44cdd6d.htm"
	helpKeywords.Add "Line StdIn", "html/1684fcd7-7762-4fb2-acf6-dacfa44cdd6d.htm"
	helpKeywords.Add "Line StdOut", "html/1684fcd7-7762-4fb2-acf6-dacfa44cdd6d.htm"
	helpKeywords.Add "Line TextStream", "html/a9d774cb-1e17-457e-9f1d-01b77bdec36b.htm"
	helpKeywords.Add "Line WshRemoteError", "html/e60dcf7b-3507-4cf4-b652-cc2f21cf88ca.htm"
	helpKeywords.Add "Line", "html/a9d774cb-1e17-457e-9f1d-01b77bdec36b.htm"
	helpKeywords.Add "LoadPicture", "html/cc8ec37c-34d7-468c-87b9-40f0d111b12e.htm"
	helpKeywords.Add "Locale ID", "html/882ca1eb-81b6-4a73-839d-154c6440bf70.htm"
	helpKeywords.Add "Locale", "html/882ca1eb-81b6-4a73-839d-154c6440bf70.htm"
	helpKeywords.Add "Log", "html/b08b8347-12b8-44b4-9bbf-84484485ff0f.htm"
	helpKeywords.Add "LogEvent", "html/03f770db-b59f-4523-ad9d-5f2b34f986ac.htm"
	helpKeywords.Add "Logical Operators", "html/1ae45658-5228-4b77-9f26-cc42bd710c89.htm"
	helpKeywords.Add "Loop", "html/4d86d9ea-6b84-4f59-a179-296b5087cdd1.htm"
	helpKeywords.Add "LTrim", "html/511e2176-d9d0-4911-9742-a455f681d10f.htm"
	helpKeywords.Add "MapNetworkDrive", "html/fef5a591-c633-4c18-91a3-848e72a36ca7.htm"
	helpKeywords.Add "Match Length", "html/5efd34f5-cb70-42cd-b4b1-86fdf6c6e811.htm"
	helpKeywords.Add "Match", "html/6a4d98b7-5b77-4c63-971c-48075af7ba65.htm"
	helpKeywords.Add "Matches", "html/4993bfad-7392-4536-9703-62f545863129.htm"
	helpKeywords.Add "Math Functions", "html/ed045a5e-d5fb-4227-a370-89d6c6023e74.htm"
	helpKeywords.Add "Methods", "html/ba2c4d21-b177-4966-a551-caee7a209e85.htm"
	helpKeywords.Add "Mid", "html/3021b949-3c89-475a-bb38-e87b1b1d3854.htm"
	helpKeywords.Add "Minute", "html/318a938f-34bb-45c5-b9f7-16f87c8d62d6.htm"
	helpKeywords.Add "Minux", "html/bf104d71-8802-4335-8098-243ed0f33d59.htm"
	helpKeywords.Add "Miscellaneous Constants", "html/132e7a70-0bb5-4e56-be3f-d2da086200bb.htm"
	helpKeywords.Add "Mod", "html/42bf3cd0-f15f-4bd9-9f31-270fc4ee4d59.htm"
	helpKeywords.Add "Modulo Operator", "html/42bf3cd0-f15f-4bd9-9f31-270fc4ee4d59.htm"
	helpKeywords.Add "Modulo", "html/42bf3cd0-f15f-4bd9-9f31-270fc4ee4d59.htm"
	helpKeywords.Add "Month", "html/0cf0416f-b251-46f2-8490-21b4e834b4ac.htm"
	helpKeywords.Add "MonthName", "html/99221560-8a0a-4f65-a2c2-f263a9bdb241.htm"
	helpKeywords.Add "Move", "html/da6b8343-2db6-4a50-af04-ce47cae3e066.htm"
	helpKeywords.Add "MoveFile", "html/277621e8-9f65-4500-bc35-89610894a3cf.htm"
	helpKeywords.Add "MoveFolder", "html/2b6287d7-a6e3-40b5-97e4-899932059b28.htm"
	helpKeywords.Add "MsgBox Constants", "html/517be4ea-7c55-4780-a48d-0b0224b29315.htm"
	helpKeywords.Add "MsgBox", "html/ae073d50-e4a4-4e23-8e46-0cb1369965e7.htm"
	helpKeywords.Add "Multiplication Operator", "html/24dd4043-d87e-4611-9d5d-49fef5e8764c.htm"
	helpKeywords.Add "Multiplication", "html/24dd4043-d87e-4611-9d5d-49fef5e8764c.htm"
	helpKeywords.Add "Name File", "html/c750e811-79cc-4cc1-a5b6-f6fac0e7043d.htm"
	helpKeywords.Add "Name Folder", "html/c750e811-79cc-4cc1-a5b6-f6fac0e7043d.htm"
	helpKeywords.Add "Name Script", "html/d511bdf9-ec04-4557-b4fd-f51c123bc835.htm"
	helpKeywords.Add "Name WScript", "html/d511bdf9-ec04-4557-b4fd-f51c123bc835.htm"
	helpKeywords.Add "Name", "html/c750e811-79cc-4cc1-a5b6-f6fac0e7043d.htm"
	helpKeywords.Add "Named Argument Item", "html/f4b72d03-1714-4d0d-81c9-7b5dd6327e50.htm"
	helpKeywords.Add "Named Arguments Item", "html/f4b72d03-1714-4d0d-81c9-7b5dd6327e50.htm"
	helpKeywords.Add "Named", "html/73b6fde7-5414-44ea-ab6d-6b27bcabb07c.htm"
	helpKeywords.Add "Next", "html/4a7b5b81-bba7-49b1-b176-f6b221df8a1c.htm"
	helpKeywords.Add "Not", "html/3aa44a43-2c6c-424f-8801-6ef2675227d3.htm"
	helpKeywords.Add "Now", "html/74f3ca98-792a-43e6-9d7e-c9a2bab0be7b.htm"
	helpKeywords.Add "Number Err", "html/35cc596b-e2fa-431b-b6de-31167187b941.htm"
	helpKeywords.Add "Number Error", "html/35cc596b-e2fa-431b-b6de-31167187b941.htm"
	helpKeywords.Add "Number WshRemoteError", "html/6fed1428-5ce4-4cd5-8ea3-599848ac4b6a.htm"
	helpKeywords.Add "Number", "html/35cc596b-e2fa-431b-b6de-31167187b941.htm"
	helpKeywords.Add "Oct", "html/e7628835-bb23-4d62-a58c-3114338844c6.htm"
	helpKeywords.Add "On Error Goto 0", "html/0675d0b2-5c1a-4f20-94f3-6749c74984a9.htm"
	helpKeywords.Add "On Error Resume Next", "html/0675d0b2-5c1a-4f20-94f3-6749c74984a9.htm"
	helpKeywords.Add "On Error", "html/0675d0b2-5c1a-4f20-94f3-6749c74984a9.htm"
	helpKeywords.Add "OpenAsTextStream", "html/37611221-cbff-4a11-965d-f7b05c6a8fe3.htm"
	helpKeywords.Add "OpenTextFile", "html/8575e5c4-dec5-48e7-92a2-790cac708c7f.htm"
	helpKeywords.Add "Operator Precedence", "html/1c1ec63f-f6d1-4ead-9660-90858ddcf023.htm"
	helpKeywords.Add "Option Explicit", "html/29da309d-81ab-4bb4-ba4b-8c7e17ef0e05.htm"
	helpKeywords.Add "Or", "html/f6eeaf84-9613-427a-a3d8-e14aefbee57e.htm"
	helpKeywords.Add "ParentFolder", "html/2d66e4df-6ec6-434e-a089-789bf94b75fd.htm"
	helpKeywords.Add "Path Drive", "html/c988522c-3cb1-462e-ad84-b44f0266bd00.htm"
	helpKeywords.Add "Path File", "html/c988522c-3cb1-462e-ad84-b44f0266bd00.htm"
	helpKeywords.Add "Path Folder", "html/c988522c-3cb1-462e-ad84-b44f0266bd00.htm"
	helpKeywords.Add "Path Script", "html/b5158c13-dd38-4052-b904-c33d993247c4.htm"
	helpKeywords.Add "Path WScript", "html/b5158c13-dd38-4052-b904-c33d993247c4.htm"
	helpKeywords.Add "Path", "html/c988522c-3cb1-462e-ad84-b44f0266bd00.htm"
	helpKeywords.Add "Pattern", "html/648fb4cf-2968-491c-b9de-51a7dad965f1.htm"
	helpKeywords.Add "Popup", "html/f482c739-3cf9-4139-a6af-3bde299b8009.htm"
	helpKeywords.Add "Private", "html/dab3dadf-6424-4464-88e1-5c0151c1a08e.htm"
	helpKeywords.Add "ProcessID Property (Windows Script Host)", "html/f3862a9f-df48-452a-a97c-0ee102a73b97.htm"
	helpKeywords.Add "Property Get", "html/05167024-a817-405f-a0ce-2057d01f804a.htm"
	helpKeywords.Add "Property Let", "html/23a9ca52-91c1-4257-a090-fdb397380ece.htm"
	helpKeywords.Add "Property Set", "html/8a5ba1a2-2a66-4d7b-bb36-2062fea3595b.htm"
	helpKeywords.Add "Public", "html/5cfb9cec-2bba-483a-b857-c0540de43fd3.htm"
	helpKeywords.Add "Quit", "html/277933cd-478e-4a45-86c8-9d96547bb515.htm"
	helpKeywords.Add "Raise", "html/feb4d98d-11ff-457f-9e73-3417f117281f.htm"
	helpKeywords.Add "Randomize", "html/ac1ef1bb-f1d8-4369-af7f-ddd89c926250.htm"
	helpKeywords.Add "Read StdIn", "html/c17eb6b3-5656-4e7a-a825-b1261a55ee42.htm"
	helpKeywords.Add "Read TextStream", "html/456d1491-dbc5-4315-876e-d181feef2884.htm"
	helpKeywords.Add "Read", "html/456d1491-dbc5-4315-876e-d181feef2884.htm"
	helpKeywords.Add "ReadAll StdIn", "html/7c92dbc9-57b0-4193-927c-fe0739a7bf1b.htm"
	helpKeywords.Add "ReadAll TextStream", "html/cc47419a-259b-4c22-a454-bac2ed150866.htm"
	helpKeywords.Add "ReadAll", "html/cc47419a-259b-4c22-a454-bac2ed150866.htm"
	helpKeywords.Add "ReadLine StdIn", "html/a53db1d2-add8-461d-8c13-a7bd60525dca.htm"
	helpKeywords.Add "ReadLine TextStream", "html/74df9812-6f1e-4a67-b76f-e05a19b3dad3.htm"
	helpKeywords.Add "ReadLine", "html/74df9812-6f1e-4a67-b76f-e05a19b3dad3.htm"
	helpKeywords.Add "ReDim", "html/5c12ce79-6616-4144-b3b6-4cffe3884dfd.htm"
	helpKeywords.Add "RegDelete", "html/161db13b-c4ca-4aec-8899-697a4183e82c.htm"
	helpKeywords.Add "RegEx Execute", "html/711116fb-9c47-47cb-b664-db8141b8cc69.htm"
	helpKeywords.Add "RegEx Length", "html/5efd34f5-cb70-42cd-b4b1-86fdf6c6e811.htm"
	helpKeywords.Add "RegEx Replace", "html/810607c5-5926-43d9-b7e8-4126e97000d2.htm"
	helpKeywords.Add "RegExp Execute", "html/711116fb-9c47-47cb-b664-db8141b8cc69.htm"
	helpKeywords.Add "RegExp Length", "html/5efd34f5-cb70-42cd-b4b1-86fdf6c6e811.htm"
	helpKeywords.Add "RegExp Replace", "html/810607c5-5926-43d9-b7e8-4126e97000d2.htm"
	helpKeywords.Add "RegExp", "html/05f9ee2e-982f-4727-839e-b1b8ed696d0a.htm"
	helpKeywords.Add "RegRead", "html/1b567504-59f4-40a9-b586-0be49ab3a015.htm"
	helpKeywords.Add "Regular Expression Execute", "html/711116fb-9c47-47cb-b664-db8141b8cc69.htm"
	helpKeywords.Add "Regular Expression Length", "html/5efd34f5-cb70-42cd-b4b1-86fdf6c6e811.htm"
	helpKeywords.Add "Regular Expression Replace", "html/810607c5-5926-43d9-b7e8-4126e97000d2.htm"
	helpKeywords.Add "Regular Expression", "html/05f9ee2e-982f-4727-839e-b1b8ed696d0a.htm"
	helpKeywords.Add "Regular Expressions", "html/05f9ee2e-982f-4727-839e-b1b8ed696d0a.htm"
	helpKeywords.Add "RegWrite", "html/678e6992-ddc4-4333-a78c-6415c9ebcc77.htm"
	helpKeywords.Add "RelativePath", "html/66f4e2bb-770d-4f48-800d-caaaab2254b6.htm"
	helpKeywords.Add "Rem", "html/674f7cdb-24f3-447c-aece-a4572c7f594f.htm"
	helpKeywords.Add "Remove Dictionary", "html/a5a8b684-65ab-4995-83fa-c4f22c482321.htm"
	helpKeywords.Add "Remove Environment", "html/05426477-f6e4-4dcb-8101-3f7d979c8ed5.htm"
	helpKeywords.Add "Remove", "html/a5a8b684-65ab-4995-83fa-c4f22c482321.htm"
	helpKeywords.Add "RemoveAll", "html/1553a16a-a7ef-42ac-9e45-ea7b918809f7.htm"
	helpKeywords.Add "RemoveNetworkDrive", "html/9ad65378-e0c9-4975-a0f3-dfecafaa7e6e.htm"
	helpKeywords.Add "RemovePrinterConnection", "html/a90ad91e-3d1f-4f2d-99d3-338a95c110d4.htm"
	helpKeywords.Add "Replace RegEx", "html/810607c5-5926-43d9-b7e8-4126e97000d2.htm"
	helpKeywords.Add "Replace RegExp", "html/810607c5-5926-43d9-b7e8-4126e97000d2.htm"
	helpKeywords.Add "Replace Regular Expression", "html/810607c5-5926-43d9-b7e8-4126e97000d2.htm"
	helpKeywords.Add "Replace", "html/65e15b2c-99b6-4f82-88e7-8c657489dd34.htm"
	helpKeywords.Add "RGB", "html/7f6c2fb0-ef0b-4774-9e3e-d7cc594996b1.htm"
	helpKeywords.Add "Right", "html/ca808da5-85e2-49c3-8abe-d66f2dc92f9d.htm"
	helpKeywords.Add "Rnd", "html/618fcba9-cbf2-409b-9963-566e28f67fd1.htm"
	helpKeywords.Add "RootFolder", "html/dba2b2d5-0a95-4199-b23e-c1c89a59e701.htm"
	helpKeywords.Add "Round", "html/c25665d8-cea5-4dbc-a320-c2c9722184bc.htm"
	helpKeywords.Add "RTrim", "html/511e2176-d9d0-4911-9742-a455f681d10f.htm"
	helpKeywords.Add "Run-time Errors", "html/48394ca0-beec-4051-9354-b47849725218.htm"
	helpKeywords.Add "Run", "html/6f28899c-d653-4555-8a59-49640b0e32ea.htm"
	helpKeywords.Add "Runtime Errors", "html/48394ca0-beec-4051-9354-b47849725218.htm"
	helpKeywords.Add "Save", "html/a6e33ae0-25ba-4a11-80ee-94764565be54.htm"
	helpKeywords.Add "Script FullName", "html/bb249bc2-e80d-44a3-ac66-16ce03db5d61.htm"
	helpKeywords.Add "Script Name", "html/d511bdf9-ec04-4557-b4fd-f51c123bc835.htm"
	helpKeywords.Add "Script Path", "html/b5158c13-dd38-4052-b904-c33d993247c4.htm"
	helpKeywords.Add "Script StdErr", "html/ad8b57d3-8ef2-4603-afe7-5807a03cb0d0.htm"
	helpKeywords.Add "Script StdIn", "html/330e1184-04e3-4314-8051-f7f24be00223.htm"
	helpKeywords.Add "Script StdOut", "html/cb7c65bf-2dce-40ab-b769-3fd59941f74b.htm"
	helpKeywords.Add "ScriptEngine", "html/dc307168-ad11-48bd-bbe8-f9df9212c085.htm"
	helpKeywords.Add "ScriptEngineBuildVersion", "html/73b1ef45-fbb1-4317-b0c1-65ff27872718.htm"
	helpKeywords.Add "ScriptEngineMajorVersion", "html/200755fe-d862-4bce-9223-b3326e659e31.htm"
	helpKeywords.Add "ScriptEngineMinorVersion", "html/2c942ec2-ea3c-45fd-b17f-9fc1c1c82f83.htm"
	helpKeywords.Add "ScriptFullName", "html/00f4fa13-345d-492f-b879-4b34160f8cc1.htm"
	helpKeywords.Add "Scripting Signer", "html/d514daef-c309-4802-abb7-cc6f731f152d.htm"
	helpKeywords.Add "Scripting.Signer", "html/d514daef-c309-4802-abb7-cc6f731f152d.htm"
	helpKeywords.Add "ScriptName", "html/b102af67-7b22-4f24-86f6-efef06402d92.htm"
	helpKeywords.Add "Second", "html/928cbc7b-5b99-4f11-96fd-83d10b3e7520.htm"
	helpKeywords.Add "Select Case", "html/91c340af-8ceb-4f46-86fa-7871eefb3b01.htm"
	helpKeywords.Add "Select", "html/91c340af-8ceb-4f46-86fa-7871eefb3b01.htm"
	helpKeywords.Add "SendKeys", "html/4b032417-ebda-4d30-88a4-2b56c24affdd.htm"
	helpKeywords.Add "SerialNumber", "html/0976513c-8be2-4af8-9a99-01f6f2d795be.htm"
	helpKeywords.Add "Set Property", "html/8a5ba1a2-2a66-4d7b-bb36-2062fea3595b.htm"
	helpKeywords.Add "Set", "html/0693dea5-c491-4d02-8eb5-f1b8a1d6c011.htm"
	helpKeywords.Add "SetDefaultPrinter", "html/66b55665-0ba7-4216-b3b3-06c9fb837f68.htm"
	helpKeywords.Add "SetLocale", "html/3c3d7685-afc6-4933-a668-21709030c021.htm"
	helpKeywords.Add "Sgn", "html/9dcee1e9-2203-4b38-b080-df305d0860e7.htm"
	helpKeywords.Add "ShareName", "html/dc2c468a-e8e1-4f8e-9dbb-ccb4d5ee9c3a.htm"
	helpKeywords.Add "Shortcut Argument", "html/15b932ee-9d9d-41f0-81ed-f6575e0d067b.htm"
	helpKeywords.Add "Shortcut Arguments", "html/15b932ee-9d9d-41f0-81ed-f6575e0d067b.htm"
	helpKeywords.Add "Shortcut Description", "html/31fd3558-705c-4c37-9a27-1399baf0cf41.htm"
	helpKeywords.Add "Shortcut FullName", "html/64d008bc-53c2-4138-9a3e-ce8f141fc371.htm"
	helpKeywords.Add "ShortName", "html/6c25d56e-466e-44d9-9221-6bf4dfa17c88.htm"
	helpKeywords.Add "ShortPath", "html/78a2c3b6-9274-4411-a234-99f457e8bfd4.htm"
	helpKeywords.Add "ShowUsage", "html/a115485e-3e68-4781-b32c-2a74c895a720.htm"
	helpKeywords.Add "Sign", "html/a7fde1db-f33c-43f7-b9bd-17bb47c895ba.htm"
	helpKeywords.Add "SignFile", "html/aff04355-a7fe-4670-9ddb-226c712a262e.htm"
	helpKeywords.Add "Sin", "html/9843cf30-c25b-44fe-b4b0-24573723eebe.htm"
	helpKeywords.Add "Size", "html/ff05d604-2bcb-4d81-853d-2332f3063814.htm"
	helpKeywords.Add "Skip StdIn", "html/4030698a-7467-4ffd-b6ab-fa81e26f82ed.htm"
	helpKeywords.Add "Skip TextStream", "html/ff7483e5-39ad-47ab-844c-816465f7c7be.htm"
	helpKeywords.Add "Skip", "html/ff7483e5-39ad-47ab-844c-816465f7c7be.htm"
	helpKeywords.Add "SkipLine StdIn", "html/b42d1b08-4273-4f56-b196-817be4246d2d.htm"
	helpKeywords.Add "SkipLine TextStream", "html/fbb27f13-701d-44f3-992f-f3cdc4e838a7.htm"
	helpKeywords.Add "SkipLine", "html/fbb27f13-701d-44f3-992f-f3cdc4e838a7.htm"
	helpKeywords.Add "Sleep", "html/af46bac4-3d73-4340-91f8-5bff61e8c996.htm"
	helpKeywords.Add "Source Err", "html/8e3cd971-4c3c-4f22-9559-775864be71c6.htm"
	helpKeywords.Add "Source Error", "html/8e3cd971-4c3c-4f22-9559-775864be71c6.htm"
	helpKeywords.Add "Source WshRemoteError", "html/255f6548-eb7c-4c5f-a158-fc56c4497665.htm"
	helpKeywords.Add "Source", "html/8e3cd971-4c3c-4f22-9559-775864be71c6.htm"
	helpKeywords.Add "SourceText", "html/e0a706c3-8408-4e4f-9886-eaaf992b92b8.htm"
	helpKeywords.Add "Space", "html/059a7c2b-daaf-40a7-8c50-6dcbe3824f0f.htm"
	helpKeywords.Add "SpecialFolders Item", "html/61580f5c-e7f9-49b3-8d53-09cdeae822f7.htm"
	helpKeywords.Add "SpecialFolders length", "html/12beb9da-a714-4de4-b6cc-b5b9c6fdc5d5.htm"
	helpKeywords.Add "SpecialFolders", "html/14761fa3-19be-4742-9f91-23b48cd9228f.htm"
	helpKeywords.Add "Split", "html/fb2bbb28-85bc-42fc-85fb-ccc7da8abe8c.htm"
	helpKeywords.Add "Sqr", "html/47bd7101-9cac-4274-b048-295608557178.htm"
	helpKeywords.Add "Status Exec", "html/6e874fcd-4efd-4891-b098-242509cbc1f9.htm"
	helpKeywords.Add "Status WshRemote", "html/7e47f7ce-10cc-4934-87ab-3675fda3fe5e.htm"
	helpKeywords.Add "Status WshScriptExec", "html/6e874fcd-4efd-4891-b098-242509cbc1f9.htm"
	helpKeywords.Add "Status", "html/6e874fcd-4efd-4891-b098-242509cbc1f9.htm"
	helpKeywords.Add "StdErr Exec", "html/fed33bf2-907f-4043-9900-2cb0da528992.htm"
	helpKeywords.Add "StdErr Line", "html/1684fcd7-7762-4fb2-acf6-dacfa44cdd6d.htm"
	helpKeywords.Add "StdErr Script", "html/ad8b57d3-8ef2-4603-afe7-5807a03cb0d0.htm"
	helpKeywords.Add "StdErr Write", "html/d73e4b78-1827-4864-945b-731373c36655.htm"
	helpKeywords.Add "StdErr WriteBlankLines", "html/2018c4c8-9d42-4ebc-8e22-2fb5c39ed053.htm"
	helpKeywords.Add "StdErr WriteLine", "html/0b1a80c3-7115-4643-a83f-7679659885b5.htm"
	helpKeywords.Add "StdErr WScript", "html/ad8b57d3-8ef2-4603-afe7-5807a03cb0d0.htm"
	helpKeywords.Add "StdErr WshScriptExec", "html/fed33bf2-907f-4043-9900-2cb0da528992.htm"
	helpKeywords.Add "StdErr", "html/ad8b57d3-8ef2-4603-afe7-5807a03cb0d0.htm"
	helpKeywords.Add "StdErr.Write", "html/d73e4b78-1827-4864-945b-731373c36655.htm"
	helpKeywords.Add "StdErr.WriteBlankLines", "html/2018c4c8-9d42-4ebc-8e22-2fb5c39ed053.htm"
	helpKeywords.Add "StdErr.WriteLine", "html/0b1a80c3-7115-4643-a83f-7679659885b5.htm"
	helpKeywords.Add "StdIn Exec", "html/b26de977-dcae-4bde-8871-21fa4cf698c2.htm"
	helpKeywords.Add "StdIn Line", "html/1684fcd7-7762-4fb2-acf6-dacfa44cdd6d.htm"
	helpKeywords.Add "StdIn Read", "html/c17eb6b3-5656-4e7a-a825-b1261a55ee42.htm"
	helpKeywords.Add "StdIn ReadAll", "html/7c92dbc9-57b0-4193-927c-fe0739a7bf1b.htm"
	helpKeywords.Add "StdIn ReadLine", "html/a53db1d2-add8-461d-8c13-a7bd60525dca.htm"
	helpKeywords.Add "StdIn Script", "html/330e1184-04e3-4314-8051-f7f24be00223.htm"
	helpKeywords.Add "StdIn Skip", "html/4030698a-7467-4ffd-b6ab-fa81e26f82ed.htm"
	helpKeywords.Add "StdIn SkipLine", "html/b42d1b08-4273-4f56-b196-817be4246d2d.htm"
	helpKeywords.Add "StdIn WScript", "html/330e1184-04e3-4314-8051-f7f24be00223.htm"
	helpKeywords.Add "StdIn WshScriptExec", "html/b26de977-dcae-4bde-8871-21fa4cf698c2.htm"
	helpKeywords.Add "StdIn", "html/330e1184-04e3-4314-8051-f7f24be00223.htm"
	helpKeywords.Add "StdOut Exec", "html/85684a76-6d66-4a1a-a3c4-cf3f48baa595.htm"
	helpKeywords.Add "StdOut Line", "html/1684fcd7-7762-4fb2-acf6-dacfa44cdd6d.htm"
	helpKeywords.Add "StdOut Script", "html/cb7c65bf-2dce-40ab-b769-3fd59941f74b.htm"
	helpKeywords.Add "StdOut Write", "html/d73e4b78-1827-4864-945b-731373c36655.htm"
	helpKeywords.Add "StdOut WriteBlankLines", "html/2018c4c8-9d42-4ebc-8e22-2fb5c39ed053.htm"
	helpKeywords.Add "StdOut WriteLine", "html/0b1a80c3-7115-4643-a83f-7679659885b5.htm"
	helpKeywords.Add "StdOut WScript", "html/cb7c65bf-2dce-40ab-b769-3fd59941f74b.htm"
	helpKeywords.Add "StdOut WshScriptExec", "html/85684a76-6d66-4a1a-a3c4-cf3f48baa595.htm"
	helpKeywords.Add "StdOut", "html/cb7c65bf-2dce-40ab-b769-3fd59941f74b.htm"
	helpKeywords.Add "StdOut.Write", "html/d73e4b78-1827-4864-945b-731373c36655.htm"
	helpKeywords.Add "StdOut.WriteBlankLines", "html/2018c4c8-9d42-4ebc-8e22-2fb5c39ed053.htm"
	helpKeywords.Add "StdOut.WriteLine", "html/0b1a80c3-7115-4643-a83f-7679659885b5.htm"
	helpKeywords.Add "Stop", "html/3ff21ea0-54e5-4f95-9c77-7f2d02977463.htm"
	helpKeywords.Add "StrComp", "html/0e7fae39-e07b-4888-8d86-f1ba18d352f3.htm"
	helpKeywords.Add "String Constants", "html/dada61a7-ec5f-4686-8ccf-f3a293b281c8.htm"
	helpKeywords.Add "String", "html/ac6f31ac-b1ff-4170-aa2e-9931c18f3dd4.htm"
	helpKeywords.Add "StrReverse", "html/44ce0dc7-3179-4654-9354-a48518bb5ecf.htm"
	helpKeywords.Add "Sub", "html/bb367eeb-58d4-4a9c-9dad-ce6b6c21ad11.htm"
	helpKeywords.Add "SubFolders", "html/1fddd555-caa0-4f77-851d-0a2d3082e13d.htm"
	helpKeywords.Add "SubMatches", "html/e84ef1f4-dc6f-4d30-8b5d-dd452efec2d5.htm"
	helpKeywords.Add "Subtraction ", "html/bf104d71-8802-4335-8098-243ed0f33d59.htm"
	helpKeywords.Add "Subtraction Operator", "html/bf104d71-8802-4335-8098-243ed0f33d59.htm"
	helpKeywords.Add "Syntax Errors", "html/5ce74817-31a0-4248-a7d7-2d30b728d415.htm"
	helpKeywords.Add "Tan", "html/5f5e44c6-aa49-45e6-a48f-3b9817a047f7.htm"
	helpKeywords.Add "TargetPath", "html/05736c92-ee7c-4614-bd28-fb66e188e7c7.htm"
	helpKeywords.Add "Terminate Exec", "html/a1ca9cc6-4b46-4190-8333-a03f1cad05cd.htm"
	helpKeywords.Add "Terminate WshScriptExec", "html/a1ca9cc6-4b46-4190-8333-a03f1cad05cd.htm"
	helpKeywords.Add "Terminate", "html/a1ca9cc6-4b46-4190-8333-a03f1cad05cd.htm"
	helpKeywords.Add "Test", "html/d898bc7a-2169-4171-9d1b-941c72dac8eb.htm"
	helpKeywords.Add "TextStream Line", "html/a9d774cb-1e17-457e-9f1d-01b77bdec36b.htm"
	helpKeywords.Add "TextStream Read", "html/456d1491-dbc5-4315-876e-d181feef2884.htm"
	helpKeywords.Add "TextStream ReadAll", "html/cc47419a-259b-4c22-a454-bac2ed150866.htm"
	helpKeywords.Add "TextStream ReadLine", "html/74df9812-6f1e-4a67-b76f-e05a19b3dad3.htm"
	helpKeywords.Add "TextStream Skip", "html/ff7483e5-39ad-47ab-844c-816465f7c7be.htm"
	helpKeywords.Add "TextStream SkipLine", "html/fbb27f13-701d-44f3-992f-f3cdc4e838a7.htm"
	helpKeywords.Add "TextStream Write", "html/7223daff-c95d-490e-8933-213982b42bb6.htm"
	helpKeywords.Add "TextStream WriteBlankLines", "html/9617e99f-958f-4f52-9731-00a41b87a0f3.htm"
	helpKeywords.Add "TextStream WriteLine", "html/b220a016-5799-4c40-b483-bc09f86a2d0a.htm"
	helpKeywords.Add "TextStream", "html/0fd59100-adcf-441d-8fa5-5e0800ad34d6.htm"
	helpKeywords.Add "Then", "html/b56cad2e-4f7c-40cf-9f08-b1934c2a868a.htm"
	helpKeywords.Add "Time Constants", "html/8c1bfc98-c011-4825-adb1-8d5eb52928d1.htm"
	helpKeywords.Add "Time", "html/3eb7a72b-f6bd-4d2f-8e5b-feca09d8a8ef.htm"
	helpKeywords.Add "Timer", "html/700b6bc7-b482-4e3f-a20a-894fb5f0e970.htm"
	helpKeywords.Add "TimeSerial", "html/1e879b67-2986-4e31-9bf1-0c3294f39c73.htm"
	helpKeywords.Add "TimeValue", "html/f97ecda8-816a-480a-ae88-3e30d96a98f6.htm"
	helpKeywords.Add "TotalSize", "html/e9d3046f-1fda-458a-9be5-011c3ade1730.htm"
	helpKeywords.Add "Trim", "html/511e2176-d9d0-4911-9742-a455f681d10f.htm"
	helpKeywords.Add "Tristate Constants", "html/b2113505-c8f4-4fdf-b306-2336cf01b477.htm"
	helpKeywords.Add "Type", "html/6491b833-7189-469b-a6ac-4f15d31d648d.htm"
	helpKeywords.Add "TypeName", "html/49a1baa0-24a3-427b-ad6f-4ccc0e94dae0.htm"
	helpKeywords.Add "UBound", "html/e6e62bd2-7b22-4848-b04c-97a3ddde109c.htm"
	helpKeywords.Add "UCase", "html/b62fee94-99cc-46ba-ab19-6a4fe278cc59.htm"
	helpKeywords.Add "Unescape", "html/8379cf57-a742-4fa2-8bee-ca25ef1a484b.htm"
	helpKeywords.Add "Unnamed Argument Item", "html/fa6fae44-a5d6-4eff-b79f-430c6b5c7ee8.htm"
	helpKeywords.Add "Unnamed Arguments Item", "html/fa6fae44-a5d6-4eff-b79f-430c6b5c7ee8.htm"
	helpKeywords.Add "Unnamed", "html/cbd8efd3-6de1-4140-8f73-d199ca5c9e3b.htm"
	helpKeywords.Add "URL FullName", "html/416b8d2b-2219-4879-a94a-92d003cbe34d.htm"
	helpKeywords.Add "UrlShortcut FullName", "html/416b8d2b-2219-4879-a94a-92d003cbe34d.htm"
	helpKeywords.Add "UserDomain", "html/d54e73f3-2932-4c33-b1ab-6c5b6c4afa6c.htm"
	helpKeywords.Add "UserName", "html/a96adec2-4898-46c2-8154-9b02a49e52c3.htm"
	helpKeywords.Add "Value", "html/cb884066-362c-4aa0-9bda-c4d4a18a940c.htm"
	helpKeywords.Add "VarType Constants", "html/e60cd976-2708-4d1b-8722-57ba65139800.htm"
	helpKeywords.Add "VarType", "html/3c8b0c10-a185-4006-bb88-79406968634b.htm"
	helpKeywords.Add "Verify", "html/96b63a3f-6024-4185-807a-f882a6f4d24e.htm"
	helpKeywords.Add "VerifyFile", "html/dd43ad9b-83a6-4cc6-9dcc-4a9f437c0e51.htm"
	helpKeywords.Add "Version", "html/d1f1c300-d8f5-4fe8-8a51-4cb40aa81190.htm"
	helpKeywords.Add "VolumeName", "html/88a0b5b9-943a-4893-8974-95fdeadc8d75.htm"
	helpKeywords.Add "Weekday", "html/c6c8d3d6-cf6b-4545-8844-2b09a72491bd.htm"
	helpKeywords.Add "WeekdayName", "html/91ad6dd2-fd34-4777-a1d4-cca4421d18bf.htm"
	helpKeywords.Add "Wend", "html/86059917-4a88-41d9-bb7a-35142c222725.htm"
	helpKeywords.Add "While Wend", "html/86059917-4a88-41d9-bb7a-35142c222725.htm"
	helpKeywords.Add "While", "html/86059917-4a88-41d9-bb7a-35142c222725.htm"
	helpKeywords.Add "While...Wend", "html/86059917-4a88-41d9-bb7a-35142c222725.htm"
	helpKeywords.Add "While..Wend", "html/86059917-4a88-41d9-bb7a-35142c222725.htm"
	helpKeywords.Add "WindowStyle", "html/a239a3ac-e51c-4e70-859e-d2d8c2eb3135.htm"
	helpKeywords.Add "With", "html/888c3bb3-3e52-480e-95e1-89f294b35df8.htm"
	helpKeywords.Add "WorkingDirectory", "html/cf9f2343-3fde-42aa-acd4-4037eb33bbf0.htm"
	helpKeywords.Add "Write Debug", "html/0530c5f1-c079-4d1a-aa42-b3f9bbf74e41.htm"
	helpKeywords.Add "Write Debugger", "html/0530c5f1-c079-4d1a-aa42-b3f9bbf74e41.htm"
	helpKeywords.Add "Write StdErr", "html/d73e4b78-1827-4864-945b-731373c36655.htm"
	helpKeywords.Add "Write StdOut", "html/d73e4b78-1827-4864-945b-731373c36655.htm"
	helpKeywords.Add "Write TextStream", "html/7223daff-c95d-490e-8933-213982b42bb6.htm"
	helpKeywords.Add "Write", "html/7223daff-c95d-490e-8933-213982b42bb6.htm"
	helpKeywords.Add "WriteBlankLines StdErr", "html/2018c4c8-9d42-4ebc-8e22-2fb5c39ed053.htm"
	helpKeywords.Add "WriteBlankLines StdOut", "html/2018c4c8-9d42-4ebc-8e22-2fb5c39ed053.htm"
	helpKeywords.Add "WriteBlankLines TextStream", "html/9617e99f-958f-4f52-9731-00a41b87a0f3.htm"
	helpKeywords.Add "WriteBlankLines", "html/9617e99f-958f-4f52-9731-00a41b87a0f3.htm"
	helpKeywords.Add "WriteLine Debug", "html/8f5593a4-3abe-49ca-9a81-e96e1607d725.htm"
	helpKeywords.Add "WriteLine Debugger", "html/8f5593a4-3abe-49ca-9a81-e96e1607d725.htm"
	helpKeywords.Add "WriteLine StdErr", "html/0b1a80c3-7115-4643-a83f-7679659885b5.htm"
	helpKeywords.Add "WriteLine StdOut", "html/0b1a80c3-7115-4643-a83f-7679659885b5.htm"
	helpKeywords.Add "WriteLine TextStream", "html/b220a016-5799-4c40-b483-bc09f86a2d0a.htm"
	helpKeywords.Add "WriteLine", "html/b220a016-5799-4c40-b483-bc09f86a2d0a.htm"
	helpKeywords.Add "WScript FullName", "html/bb249bc2-e80d-44a3-ac66-16ce03db5d61.htm"
	helpKeywords.Add "WScript Name", "html/d511bdf9-ec04-4557-b4fd-f51c123bc835.htm"
	helpKeywords.Add "WScript Path", "html/b5158c13-dd38-4052-b904-c33d993247c4.htm"
	helpKeywords.Add "WScript StdErr", "html/ad8b57d3-8ef2-4603-afe7-5807a03cb0d0.htm"
	helpKeywords.Add "WScript StdIn", "html/330e1184-04e3-4314-8051-f7f24be00223.htm"
	helpKeywords.Add "WScript StdOut", "html/cb7c65bf-2dce-40ab-b769-3fd59941f74b.htm"
	helpKeywords.Add "WScript", "html/4dc0e2db-3234-4383-afba-147154041dfd.htm"
	helpKeywords.Add "WScript.FullName", "html/bb249bc2-e80d-44a3-ac66-16ce03db5d61.htm"
	helpKeywords.Add "WScript.Name", "html/d511bdf9-ec04-4557-b4fd-f51c123bc835.htm"
	helpKeywords.Add "WScript.Path", "html/b5158c13-dd38-4052-b904-c33d993247c4.htm"
	helpKeywords.Add "WScript.StdErr", "html/ad8b57d3-8ef2-4603-afe7-5807a03cb0d0.htm"
	helpKeywords.Add "WScript.StdIn", "html/330e1184-04e3-4314-8051-f7f24be00223.htm"
	helpKeywords.Add "WScript.StdOut", "html/cb7c65bf-2dce-40ab-b769-3fd59941f74b.htm"
	helpKeywords.Add "WshArguments", "html/d1754ef6-9181-419f-8280-293443a41778.htm"
	helpKeywords.Add "WshController", "html/dd269b51-cf9e-41e9-ac28-3c01e32ee59b.htm"
	helpKeywords.Add "WshEnvironment", "html/af2cc3bc-f468-4d76-be12-1472fe8abef2.htm"
	helpKeywords.Add "WshNamed Item", "html/f4b72d03-1714-4d0d-81c9-7b5dd6327e50.htm"
	helpKeywords.Add "WshNamed", "html/23f429b3-72c4-453c-bb91-d9e15ad21cbf.htm"
	helpKeywords.Add "WshNetwork", "html/438d1705-fb98-4a63-86e5-a8cc01f4ce16.htm"
	helpKeywords.Add "WshRemote Execute", "html/742524c8-990a-4f07-8130-336803d68f67.htm"
	helpKeywords.Add "WshRemote Status", "html/7e47f7ce-10cc-4934-87ab-3675fda3fe5e.htm"
	helpKeywords.Add "WshRemote", "html/f9f0e5da-824d-4cde-a67b-02108178ee45.htm"
	helpKeywords.Add "WshRemoteError Description", "html/a769de1b-b007-492f-9d72-66053ade2874.htm"
	helpKeywords.Add "WshRemoteError Line", "html/e60dcf7b-3507-4cf4-b652-cc2f21cf88ca.htm"
	helpKeywords.Add "WshRemoteError", "html/1718aa52-0485-4b76-a683-5d41fd689c95.htm"
	helpKeywords.Add "WshScriptExec Status", "html/6e874fcd-4efd-4891-b098-242509cbc1f9.htm"
	helpKeywords.Add "WshScriptExec StdErr", "html/fed33bf2-907f-4043-9900-2cb0da528992.htm"
	helpKeywords.Add "WshScriptExec StdIn", "html/b26de977-dcae-4bde-8871-21fa4cf698c2.htm"
	helpKeywords.Add "WshScriptExec StdOut", "html/85684a76-6d66-4a1a-a3c4-cf3f48baa595.htm"
	helpKeywords.Add "WshScriptExec Terminate", "html/a1ca9cc6-4b46-4190-8333-a03f1cad05cd.htm"
	helpKeywords.Add "WshScriptExec", "html/f3358e96-3d5a-46c2-b43b-3107e586736e.htm"
	helpKeywords.Add "WshShell", "html/7b956233-c1aa-4b59-b36d-f3e97a9b02f0.htm"
	helpKeywords.Add "WshShortcut", "html/5ce04e4b-871a-4378-a192-caa644bd9c55.htm"
	helpKeywords.Add "WshSpecialFolders", "html/7682257e-4042-4f7d-b266-03382021d0aa.htm"
	helpKeywords.Add "WshUnnamed Item", "html/fa6fae44-a5d6-4eff-b79f-430c6b5c7ee8.htm"
	helpKeywords.Add "WshUnnamed", "html/7bb7a47b-7e57-4071-987c-9ee77c5e2b18.htm"
	helpKeywords.Add "WshUrlShortcut", "html/e1b7f981-e8b9-4f41-bb4c-5462d664f184.htm"
	helpKeywords.Add "Xor", "html/1ad4dd65-f8b1-4d76-996f-306d19106941.htm"
	helpKeywords.Add "Year", "html/9569b7bb-8533-40d9-bc13-372d922244c3.htm"

Main


Sub Main()
	Dim line

	ImportInitScript
	Usage

	Do While True
		WScript.StdOut.Write(">>> ")

		line = Trim(WScript.StdIn.ReadLine)
		Do While Right(line, 2) = " _" Or line = "_"
			line = RTrim(Left(line, Len(line)-1)) & " " & Trim(WScript.StdIn.ReadLine)
		Loop

		If LCase(line) = "exit" Then Exit Do

		If line = "help" Then
			Usage
		ElseIf line = "?" Then
			Help Null
		ElseIf Left(line, 2) = "? " Then
			Help Trim(Replace(Mid(line, 3), """", ""))
		ElseIf Left(line, 2) = "! " Then
			WScript.Echo Trim(Replace(Mid(line, 3), """", ""))
		Else
			If Left(line, 2) = "! " Then line = "WScript.Echo " & Mid(line, 3)
			On Error Resume Next
			Err.Clear
			Execute line
			If Err.Number <> 0 Then WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
			On Error Goto 0
		End If
	Loop
End Sub

'! Import initialization script if present.
Private Sub ImportInitScript()
	Dim sh, fso, path, initScriptExists

	Set sh  = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")

	path = sh.ExpandEnvironmentStrings(InitScript)
	initScriptExists = fso.FileExists(path)

	Set sh  = Nothing
	Set fso = Nothing

	If initScriptExists Then Import path
End Sub

'! Look up the given keyword in the VBScript documentation (if the help file is
'! located in the current working directory, the Windows help directory, or
'! somewhere in the %PATH%).
'!
'! @param  keyword  Keyword to look up.
Private Sub Help(ByVal keyword)
	Dim sh, fso, chm, dir

	Set sh  = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")

	If isEmpty(helpfile) Then
		' check current directory
		chm = fso.BuildPath(sh.CurrentDirectory, Documentation)
		If Not fso.FileExists(chm) Then
			' check help directory
			chm = fso.BuildPath(sh.ExpandEnvironmentStrings("%SystemRoot%\Help"), Documentation)
		End If
		If Not fso.FileExists(chm) Then
			' check %PATH%
			For Each dir In Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
				chm = fso.BuildPath(fso.GetAbsolutePathName(dir), Documentation)
				If fso.FileExists(chm) Then Exit For
			Next
		End If
		If fso.FileExists(chm) Then helpfile = chm
	End If

	If Not fso.FileExists(helpfile) Then
		WScript.StdErr.WriteLine "'" & Documentation & "' not found. Make sure the file is present in the current working working" & vbNewLine _
			& "directory or the Windows help directory, or add its location to the PATH environment" & vbNewLine _
			& "variable."
		Exit Sub
	End If

	If Not IsNull(keyword) And Not helpKeywords.Exists(keyword) Then
		WScript.StdOut.WriteLine "'" & keyword & "' not found in the help index."
		keyword = Null
	End If

	If IsNull(keyword) Then
		sh.Run "hh.exe """ & helpfile & """", 1, False
	Else
		sh.Run "hh.exe ""mk:@MSITStore:" & helpfile & "::/" & helpKeywords(keyword) & """", 1, False
	End If
End Sub

'! Print usage information.
Private Sub Usage()
	WScript.StdOut.Write "A simple interactive VBScript Shell." & vbNewLine & vbNewLine _
		& vbTab & "help                      Print this help." & vbNewLine _
		& vbTab & "! EXPRESSION              Shortcut for 'WScript.Echo'." & vbNewLine _
		& vbTab & "?                         Open the VBScript documentation." & vbNewLine _
		& vbTab & "? ""keyword""               Look up ""keyword"" in the documentation." & vbNewLine _
		& vbTab & "                          The helpfile (" & Documentation & ") must be installed" & vbNewLine _
		& vbTab & "                          in either the Windows help directory, %PATH%" & vbNewLine _
		& vbTab & "                          or the current working directory." & vbNewLine _
		& vbTab & "import ""\PATH\TO\my.vbs""  Load and execute the contents of the script." & vbNewLine _
		& vbTab & "exit                      Exit the shell." & vbNewLine _
		& vbNewLine & "Customize with an (optional) init script '" & InitScript & "'." & vbNewLine _
		& vbNewLine
End Sub

'! Import the first occurrence of the given filename from the working directory
'! or any directory in the %PATH%.
'!
'! @param  filename   Name of the file to import.
'!
'! @see http://gazeek.com/coding/importing-vbs-files-in-your-vbscript-project/
Private Sub Import(ByVal filename)
	Dim fso, sh, file, code, dir

	' Create my own objects, so the function is self-contained and can be called
	' before anything else in the script.
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh = CreateObject("WScript.Shell")

	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	If Not (Left(filename, 2) = "\\" Or Mid(filename, 2, 2) = ":\") Then
		' filename is not absolute
		If Not fso.FileExists(fso.GetAbsolutePathName(filename)) Then
			' file doesn't exist in the working directory => iterate over the
			' directories in the %PATH% and take the first occurrence
			' if no occurrence is found => use filename as-is, which will result
			' in an error when trying to open the file
			For Each dir In Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
				If fso.FileExists(fso.BuildPath(dir, filename)) Then
					filename = fso.BuildPath(dir, filename)
					Exit For
				End If
			Next
		End If
		filename = fso.GetAbsolutePathName(filename)
	End If

	Set file = fso.OpenTextFile(filename, 1, False)
	code = file.ReadAll
	file.Close

	ExecuteGlobal(code)

	Set fso = Nothing
	Set sh = Nothing
End Sub
