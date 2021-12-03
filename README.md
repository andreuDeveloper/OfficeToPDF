# OfficeToPDF
This library allows to Transform an Office document or other documents to a final PDF.

IMPORTANT: This library needs a LibreOffice versión installed. (It can be a Portable Version)


Example:

```
const OfficePath = @"userPath\LibreOfficePortable_7.0\App\libreoffice\program\soffice.exe"; //or soffice.bin
const inputFilePath = @"";
const outputDirPath = @""; //OPTIONAL

OfficeToPDF transformer = new OfficeToPDF(OfficePath);
transformer.ConvertToPDF(inputFilePath, outputDirPath);

```

**Example: Specifying max timeout**

*Sometimes.. huge files or documents that will generate a lot of pages (like Excel) can take a long time.
The default timeout is: **30 MIN***
```
const OfficePath = @"userPath\LibreOfficePortable_7.0\App\libreoffice\program\soffice.exe"; //or soffice.bin
const inputFilePath = @"";
const outputDirPath = @""; //OPTIONAL
const maxTimeout = 5; //Min

OfficeToPDF transformer = new OfficeToPDF(OfficePath);
transformer.ConvertToPDF(inputFilePath, outputDirPath, maxTimeout);

```


Common extensions:
  - ".doc"
  - ".docx"
  - ".txt"
  - ".rtf"
  - ".html"
  - ".htm"
  - ".xml"
  - ".odt"
  - ".wps"
  - ".wpd"
  - ".css"
  - ".json"
  - ".xls"
  - ".xlsb"
  - ".xlsx"
  - ".ods"
  - ".csv"
  - ".ppt"
  - ".pptx"
  - ".odp"
    
Can transform other formats like: .csv, .json..
