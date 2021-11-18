import java.io.*;
import java.util.ArrayList;
import java.util.List;
import com.fourjs.spreadsheet.CopySheets;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelMerge {

   public static void main(String[] args) {
      Workbook wb = new XSSFWorkbook();
      CopySheets cs = new CopySheets();
      List<InputStream> inList = new ArrayList<InputStream>();

      try {
      
         inList.add(new FileInputStream("../files/fgl00040052.xlsx"));
         inList.add(new FileInputStream("../files/fgl00140052.xlsx"));
         inList.add(new FileInputStream("../files/fgl00240052.xlsx"));
   
         InputStream[] inArray = new InputStream[inList.size()];
         inList.toArray(inArray);
         wb = cs.mergeExcelFiles(wb, inArray);
         FileOutputStream outStream = new FileOutputStream("../files/merged.xlsx");
         wb.write(outStream);
         outStream.close();

         System.out.println("Excel file created: ../files/merged.xlsx");

      } catch (IOException e) {
         System.out.println("IO Exception occurred");
         System.out.println(e.toString());
      } catch (Exception e) {
         System.out.println("General Exception occurred");
         System.out.println(e.toString());
      }

   }


}
