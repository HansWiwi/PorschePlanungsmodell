import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.ListIterator;


public class Main {

    /* Main-Methode
     *
     */
        public static void main(String[] args) throws IOException {

        Projekt P  = new Projekt();
        P.DatenEinlesen();
        P.ParameterVergleichen();
        P.finalize();
        P.DatenAusgeben();
        P.AusgabeDateiErzeugen();


        }
}









