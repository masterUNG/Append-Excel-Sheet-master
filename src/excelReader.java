import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.File;
import java.io.IOException;

/**
 * Created by Torab on 19-Jan-17.
 */
public class excelReader {
    public static void main(String[] args) {
        File file=new File("work2.xls");
        if(!file.exists())
        {
            try {
                create();
                excelWritingwriting();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            } catch (BiffException e) {
                e.printStackTrace();
            }
        }
        try {
            excelWritingwriting();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }

    }

    private static void excelWritingwriting() throws WriteException, IOException, BiffException {
        Workbook aWorkBook = Workbook.getWorkbook(new File("work2.xls"));
        WritableWorkbook aCopy = Workbook.createWorkbook(new File("work2.xls"), aWorkBook);

        WritableSheet aCopySheet = aCopy.getSheet(0);

        Label anotherWritableCel =  new Label(2,2,"wasd");
        Label anotherWritableCe2 =  new Label(23,2,"1234");
        Label anotherWritableCe3 =  new Label(23,3,"1234");
        Label anotherWritableCe4 =  new Label(8,4,"1234");
        Label anotherWritableCe5 =  new Label(18,4,"1234");
        Label anotherWritableCe6 =  new Label(24,4,"1234");
        Label anotherWritableCe7 =  new Label(34,4,"1234");
        Label anotherWritableCe8 =  new Label(2,5,"1234");
        Label anotherWritableCe9 =  new Label(13,5,"1234");
        Label anotherWritableCe10 =  new Label(24,5,"1234");
        Label anotherWritableCe11 =  new Label(2,6,"1234");
        Label anotherWritableCe12 =  new Label(13,6,"1234");
        Label anotherWritableCel3 =  new Label(24,6,"1234");
        Label anotherWritableCe14 =  new Label(35,6,"1234");
        Label anotherWritableCe15 =  new Label(29,8,"1234");
        Label anotherWritableCe16 =  new Label(16,10,"1234");
        Label anotherWritableCel7 =  new Label(21,10,"1234");
        Label anotherWritableCe18 =  new Label(34,10,"1234");
        Label anotherWritableCe19 =  new Label(12,11,"1234");
        Label anotherWritableCe20 =  new Label(18,11,"1234");
        Label anotherWritableCe2l =  new Label(24,11,"1234");
        Label anotherWritableCe22 =  new Label(33,11,"1234");
        Label anotherWritableCe23 =  new Label(23,12,"1234");
        Label anotherWritableCe24 =  new Label(29,12,"1234");
        Label anotherWritableCe25 =  new Label(36,12,"1234");
        Label anotherWritableCe26 =  new Label(12,13,"1234");
        Label anotherWritableCe27 =  new Label(13,14,"1234");
        Label anotherWritableCe28 =  new Label(23,14,"1234");
        Label anotherWritableCe29 =  new Label(35,16,"1234");
        Label anotherWritableCe30 =  new Label(35,17,"1234");
        Label anotherWritableCe31 =  new Label(19,18,"1234");
        Label anotherWritableCe32 =  new Label(25,19,"1234");
        Label anotherWritableCe34 =  new Label(18,20,"1234");
        Label anotherWritableCe35 =  new Label(36,20,"1234");
        Label anotherWritableCe36 =  new Label(34,21,"1234");
        Label anotherWritableCe37 =  new Label(2,24,"1234");
        Label anotherWritableCe38 =  new Label(18,29,"1234");
        Label anotherWritableCe39 =  new Label(7,31,"1234");
        Label anotherWritableCe40 =  new Label(18,31,"1234");
        Label anotherWritableCe41 =  new Label(24,31,"1234");
        Label anotherWritableCe42 =  new Label(34,31,"1234");
        Label anotherWritableCe43 =  new Label(12,32,"1234");
        Label anotherWritableCe44 =  new Label(34,32,"1234");

        aCopySheet.addCell(anotherWritableCel);
        aCopySheet.addCell(anotherWritableCe2);
        aCopySheet.addCell(anotherWritableCe3);
        aCopySheet.addCell(anotherWritableCe4);
        aCopySheet.addCell(anotherWritableCe5);
        aCopySheet.addCell(anotherWritableCe6);
        aCopySheet.addCell(anotherWritableCe7);
        aCopySheet.addCell(anotherWritableCe8);
        aCopySheet.addCell(anotherWritableCe9);
        aCopySheet.addCell(anotherWritableCe10);
        aCopySheet.addCell(anotherWritableCe11);
        aCopySheet.addCell(anotherWritableCe12);
        aCopySheet.addCell(anotherWritableCel3);
        aCopySheet.addCell(anotherWritableCe14);
        aCopySheet.addCell(anotherWritableCe15);
        aCopySheet.addCell(anotherWritableCe16);
        aCopySheet.addCell(anotherWritableCel7);
        aCopySheet.addCell(anotherWritableCe18);
        aCopySheet.addCell(anotherWritableCe19);
        aCopySheet.addCell(anotherWritableCe20);
        aCopySheet.addCell(anotherWritableCe2l);
        aCopySheet.addCell(anotherWritableCe22);
        aCopySheet.addCell(anotherWritableCe23);
        aCopySheet.addCell(anotherWritableCe24);
        aCopySheet.addCell(anotherWritableCe25);
        aCopySheet.addCell(anotherWritableCe26);
        aCopySheet.addCell(anotherWritableCe27);
        aCopySheet.addCell(anotherWritableCe28);
        aCopySheet.addCell(anotherWritableCe29);
        aCopySheet.addCell(anotherWritableCe30);
        aCopySheet.addCell(anotherWritableCe31);
        aCopySheet.addCell(anotherWritableCe32);
        aCopySheet.addCell(anotherWritableCe34);
        aCopySheet.addCell(anotherWritableCe35);
        aCopySheet.addCell(anotherWritableCe36);
        aCopySheet.addCell(anotherWritableCe37);
        aCopySheet.addCell(anotherWritableCe38);
        aCopySheet.addCell(anotherWritableCe39);
        aCopySheet.addCell(anotherWritableCe40);
        aCopySheet.addCell(anotherWritableCe41);
        aCopySheet.addCell(anotherWritableCe42);
        aCopySheet.addCell(anotherWritableCe43);
        aCopySheet.addCell(anotherWritableCe44);

        aCopy.write();
        aCopy.close();

    }

    private static void create() throws IOException, WriteException {

        WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File("work2.xls"));
        writableWorkbook.createSheet("land",0);

        writableWorkbook.write();
        writableWorkbook.close();
    }
}
