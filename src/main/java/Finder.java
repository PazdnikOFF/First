package first;

import first.writer.ExcelReader;
import first.writer.ExcelWriter;
import first.writer.conn;

import java.io.File;
import java.io.IOException;
import java.sql.SQLException;

public class Finder {

    public static void findFile() throws IOException {
        int fcount = 0;
        File folder = new File(".");
        //System.out.println(System.getProperty("user.dir"));
        File[] listOfFiles = folder.listFiles();
        try
        {
       for (File lofFiles: listOfFiles) {
            if (lofFiles.getName().indexOf("site")<0 && lofFiles.isFile() && lofFiles.getName().endsWith(".xlsx")||lofFiles.getName().endsWith(".xls")||lofFiles.getName().endsWith(".XLSX")||lofFiles.getName().endsWith(".XLS")) {
                if (lofFiles.getName().indexOf("~") != 0) {
                    var reader = new ExcelReader();
                    reader.read(lofFiles.getName());
                }
            }
        }
       }
        catch(IOException e) {
            throw new RuntimeException("Unknown action: " + fcount);
        }
     }
    public static void main(String[] args) throws IOException, SQLException {

        try {
         conn.Conn();
         conn.CreateDB();
        }
        catch(SQLException e)
        {
            System.err.println(e.getMessage());
        }
        catch(ClassNotFoundException e)
        {
            System.err.println(e.getMessage());
        }

        first.writer.ExcelWriter.create("site.xlsx");
        findFile();
        conn.CloseDB();
        first.writer.ExcelWriter.close("site.xlsx");
    }

}