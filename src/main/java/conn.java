package first.writer;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class conn {
    public static Connection conn;
    public static Statement statmt;
    public static ResultSet resSet;

    // --------ПОДКЛЮЧЕНИЕ К БАЗЕ ДАННЫХ--------
    public static void Conn() throws ClassNotFoundException, SQLException
    {
        conn = null;
        Class.forName("org.sqlite.JDBC");
        conn = DriverManager.getConnection("jdbc:sqlite:db");

    //    System.out.println("База Подключена!");
    }

    // --------Создание таблицы--------
    public static void CreateDB()  throws  SQLException
    {
        statmt = conn.createStatement();
        //statmt.execute("DROP TABLE 'products';");
        statmt.execute("CREATE TABLE if not exists 'products' ('id' INTEGER PRIMARY KEY AUTOINCREMENT, 'manufacture' text, 'articul' text, 'name' text, 'price' text, 'avaliable' text, 'supplierArticul' text, 'code' text, 'supplier' text, 'city' text);");
    //    System.out.println("Таблица создана или уже существует.");
    }

    // --------Заполнение таблицы--------
    public static void WriteDB(String manufacture, String articul, String name, String price, String avaliable, String supplierArticul, String code, String supplier, String city) throws SQLException
    {
            var id = ReadID(name,articul,supplier);

            if (id > 0) {
                statmt.execute("UPDATE products SET price=" + price + ", avaliable=" + avaliable + " WHERE id=" + id + ";");
            } else {
                statmt.execute("INSERT INTO products (manufacture, articul, name, price, avaliable, supplierArticul, code, supplier, city) " +
                        "VALUES ('" + manufacture + "', '" + articul + "', '" + name + "', " + price + ", " + avaliable + ", '" + supplierArticul + "', '" + String.format("%08d", ReadCode(name, articul, supplier)) + "', '" + supplier + "', '" + city + "'); ");
            }
    }

    // -------- Вывод таблицы--------
    public static int ReadID(String name, String articul, String supplier)  throws  SQLException
    {
        resSet = statmt.executeQuery("SELECT id FROM products WHERE name='"+name+"' AND articul='"+articul+"' AND supplier='"+supplier+"'");
        if (resSet.next() == false) {
            return 0;
        }else {
            return resSet.getInt("id");
        }
    }

    public static int ReadCode(String name, String articul, String supplier)  throws  SQLException
    {
        resSet = statmt.executeQuery("SELECT code FROM products WHERE name='"+name+"' AND articul='"+articul+"' AND supplier='"+supplier+"'");
        if (resSet.next() == false) {
            return GenerateCode();
        }else {
            return resSet.getInt("code");
        }
    }

    public static int GenerateCode()  throws  SQLException
    {   int art;
        resSet = statmt.executeQuery("SELECT code FROM products ORDER BY id DESC LIMIT 1");
        if (resSet.next() == false) {
            return 1;
        }else {
            return resSet.getInt("code")+1;
        }
    }

    // --------Закрытие--------
    public static void CloseDB()  throws  SQLException
    {
        resSet = statmt.executeQuery("SELECT * FROM products");

        while(resSet.next()) {
        first.writer.ExcelWriter.createCells(
                resSet.getString("manufacture"),
                resSet.getString("articul"),
                resSet.getString("name"),
                addPercent(resSet.getString("price")),
                resSet.getString("avaliable"),
                resSet.getString("supplierArticul"),
                resSet.getString("code"),
                resSet.getString("supplier"),
                resSet.getString("city"));
        }
        conn.close();
        statmt.close();
        resSet.close();
    }

    private static String addPercent(String input) {
        return String.valueOf(Math.round(Float.valueOf(input)+(20 / 100.0f)*Float.valueOf(input)));
    }
}
