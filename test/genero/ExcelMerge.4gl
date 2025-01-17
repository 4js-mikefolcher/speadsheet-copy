IMPORT os

IMPORT JAVA java.io.FileInputStream
IMPORT JAVA java.io.FileOutputStream
IMPORT JAVA java.io.InputStream
IMPORT JAVA com.fourjs.spreadsheet.CopySheets
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFWorkbook
IMPORT JAVA org.apache.poi.ss.usermodel.Workbook

PRIVATE TYPE CopyExcel CopySheets
PRIVATE TYPE InputStreamArray ARRAY[] OF InputStream
PRIVATE TYPE IWorkbook Workbook

PRIVATE TYPE TExcelFileMerge RECORD
   mergedFile STRING,
   filelist   DYNAMIC ARRAY OF STRING
END RECORD

PUBLIC DEFINE excelMerge TExcelFileMerge

MAIN
   DEFINE filename STRING

   CALL excelMerge.init()

   LET filename = SFMT("..%1files%1%2", os.Path.separator(), "fgl00040052.xlsx")
   IF NOT excelMerge.addFile(filename) THEN
      DISPLAY SFMT("File %1 does not exist", filename)
      EXIT PROGRAM -1
   END IF

   LET filename = SFMT("..%1files%1%2", os.Path.separator(), "fgl00140052.xlsx")
   IF NOT excelMerge.addFile(filename) THEN
      DISPLAY SFMT("File %1 does not exist", filename)
      EXIT PROGRAM -1
   END IF

   LET filename = SFMT("..%1files%1%2", os.Path.separator(), "fgl00240052.xlsx")
   IF NOT excelMerge.addFile(filename) THEN
      DISPLAY SFMT("File %1 does not exist", filename)
      EXIT PROGRAM -1
   END IF

   IF NOT excelMerge.mergeFiles() THEN
      DISPLAY "Error merging Excel files"
      EXIT PROGRAM -1
   END IF

   DISPLAY SFMT("Merged Excel file %1 was created successfully!", excelMerge.mergedFile)

END MAIN


PUBLIC FUNCTION (self TExcelFileMerge) init() RETURNS ()

   LET self.mergedFile = os.Path.makeTempName(), ".xlsx"
   CALL self.filelist.clear()

END FUNCTION

PUBLIC FUNCTION (self TExcelFileMerge) addFile(filename STRING) RETURNS BOOLEAN
   DEFINE idx INTEGER

   IF NOT os.Path.exists(filename) THEN
      RETURN FALSE
   END IF

   CALL self.filelist.appendElement()
   LET idx = self.filelist.getLength()
   LET self.filelist[idx] = filename

   RETURN TRUE

END FUNCTION #addFile

PUBLIC FUNCTION (self TExcelFileMerge) mergeFiles() RETURNS (BOOLEAN)
   DEFINE successFlag BOOLEAN = TRUE
   DEFINE fileStream FileInputStream
   DEFINE streamList InputStreamArray
   DEFINE excelWorkbook org.apache.poi.xssf.usermodel.XSSFWorkbook
   DEFINE idx INTEGER
   DEFINE merge CopySheets
   DEFINE workbook IWorkbook
   DEFINE outputStream FileOutputStream

   TRY
      LET excelWorkbook = XSSFWorkbook.create()
      LET merge = CopySheets.create()
      LET streamList = InputStreamArray.create(self.filelist.getLength())

      FOR idx = 1 TO self.filelist.getLength()

         LET fileStream = FileInputStream.create(self.filelist[idx])
         LET streamList[idx] = fileStream

      END FOR

      #Do the merge
      LET workbook = merge.mergeExcelFiles(excelWorkbook, streamList)

      #Write the file out
      LET outputStream = FileOutputStream.create(self.mergedFile)
      CALL workbook.write(outputStream)
      CALL outputStream.close() 

   CATCH

      DISPLAY SFMT("Status: %1", STATUS)
      DISPLAY SFMT("Message: %1", err_get(STATUS))
      DISPLAY base.Application.getStackTrace()
      LET successFlag = FALSE

   END TRY

   RETURN successFlag

END FUNCTION #mergeFiles

