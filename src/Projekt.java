import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.util.*;
import java.io.FileOutputStream;

public class Projekt {
    //----------------------------------------------------------------------------------------------------------------
    // Konstaten für Name der TAB-Sheets in Exceldatei definieren
    //----------------------------------------------------------------------------------------------------------------

    public static final String TABNAME_Produkt = "Produkt";
    public static final String TABNAME_Produkteigenschaft = "Produkteigenschaft";
    public static final String TABNAME_Eigenschaft = "Eigenschaft";
    public static final String TABNAME_Ressource = "Ressource";
    public static final String TABNAME_Ressourcenparameter = "Ressourcenparameter";
    public static final String TABNAME_Parameter = "Parameter";

    //----------------------------------------------------------------------------------------------------------------
    // Konstanten für erlaubte Datentypen definieren
    //----------------------------------------------------------------------------------------------------------------

    public static final String DATENTYP_Double = "Double";
    public static final String DATENTYP_Integer = "Integer";
    public static final String DATENTYP_Boolean = "Boolean";
    public static final String DATENTYP_String = "String";

    //----------------------------------------------------------------------------------------------------------------
    // Konstanten für erlaubte Operatoren definieren
    //----------------------------------------------------------------------------------------------------------------

    public static final String OPERATOR_Gleich = "=";
    public static final String OPERATOR_UnGleich = "<>";
    public static final String OPERATOR_Groesser = ">";
    public static final String OPERATOR_GroesserGleich = ">=";
    public static final String OPERATOR_Kleiner = "<";
    public static final String OPERATOR_KleinerGleich = "<=";

    //----------------------------------------------------------------------------------------------------------------
    // Konstruktor für das Projekt zum Initialisieren und Öffnen der Protokoll-Datei
    //----------------------------------------------------------------------------------------------------------------

    PrintWriter PWriter = null;

    public Projekt(){

        try {
            PWriter = new PrintWriter(new BufferedWriter(new FileWriter("c:\\temp\\PorscheParameterVergleich.txt")));
            PWriter.println("Start");
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }

    //----------------------------------------------------------------------------------------------------------------
    // Destruktor für das Projekt zum Schließen der Protokoll-Datei
    //----------------------------------------------------------------------------------------------------------------

    protected void finalize() throws IOException {
        if (PWriter != null) {
            PWriter.println("Ende");
            PWriter.flush();
            PWriter.close();
        }
    }

    //----------------------------------------------------------------------------------------------------------------
    // Erzeugung je einer LinkedList für Parameter, Eigenschaften, Ressourcen und Produkte
    //----------------------------------------------------------------------------------------------------------------

    LinkedList <Parameter> ParameterListe = new LinkedList();
    LinkedList <Eigenschaft> EigenschaftenListe = new LinkedList();
    LinkedList <Ressource> RessourcenListe = new LinkedList();
    LinkedList <Produkt> ProduktListe = new LinkedList();

    //----------------------------------------------------------------------------------------------------------------
    // Prüft, ob ein Parameter mit der übergebenen ID in der Parameter-Liste vorhanden ist
    // Return TRUE, falls ja
    //        FALSE, sonst
    //----------------------------------------------------------------------------------------------------------------

    public boolean ParameterVorhanden (String ID){
        boolean retval = false;
        for(int num=0; num<ParameterListe.size(); num++)
        {
            Parameter P = (Parameter)ParameterListe.get(num);
            String S1 = P.GetID();
            if (S1.equals(ID)) {
                retval = true;
                break;
            }
        }
        return retval;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Sucht den Parameter mit der übergebenen ID in der Parameter-Liste und gibt diesen zurück
    // Return Parameter mit ID, wenn in Liste vorhanden
    //        NULL, wenn Parameter nicht gefunden
    //----------------------------------------------------------------------------------------------------------------

    public Parameter GetParameter (String ID){
        for(int num=0; num<ParameterListe.size(); num++)
        {
            Parameter P = (Parameter)ParameterListe.get(num);
            String S1 = P.GetID();
            if (S1.equals(ID)) {
                return P;
            }
        }
        return null;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Prüft, ob eine Ressource mit der übergebenen ID in der Ressourcen-Liste vorhanden ist
    // Return TRUE, falls ja
    //        FALSE, sonst
    //----------------------------------------------------------------------------------------------------------------

    public boolean RessourceVorhanden (String ID){
        boolean retval = false;
        for(int num=0; num<RessourcenListe.size(); num++)
        {
            Ressource R = (Ressource)RessourcenListe.get(num);
            String S1 = R.GetID();
            if (S1.equals(ID)) {
                retval = true;
                break;
            }
        }
        return retval;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Sucht die Ressource mit der übergebenen ID in der Ressourcen-Liste und gibt diese zurück
    // Return Ressource mit ID, wenn in Liste vorhanden
    //        NULL, wenn Ressource nicht gefunden
    //----------------------------------------------------------------------------------------------------------------

    public Ressource GetRessource (String ID){
        for(int num=0; num<RessourcenListe.size(); num++)
        {
            Ressource R = (Ressource)RessourcenListe.get(num);
            String S1 = R.GetID();
            if (S1.equals(ID)) {
                return R;
            }
        }
        return null;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Prüft, ob ein Produkt mit der übergebenen ID in der Produkt-Liste vorhanden ist
    // Return TRUE, falls ja
    //        FALSE, sonst
    //----------------------------------------------------------------------------------------------------------------

    public boolean ProduktVorhanden (String ID){
        boolean retval = false;
        for(int num=0; num<ProduktListe.size(); num++)
        {
            Produkt P = (Produkt)ProduktListe.get(num);
            String S1 = P.GetID();
            if (S1.equals(ID)) {
                retval = true;
                break;
            }
        }
        return retval;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Sucht das Produkt mit der übergebenen ID in der Produkt-Liste und gibt dieses zurück
    // Return Produkt mit ID, wenn in Liste vorhanden
    //        NULL, wenn Produkt nicht gefunden
    //----------------------------------------------------------------------------------------------------------------

    public Produkt GetProdukt (String ID){
        for(int num=0; num<ProduktListe.size(); num++)
        {
            Produkt P = (Produkt)ProduktListe.get(num);
            String S1 = P.GetID();
            if (S1.equals(ID)) {
                return P;
            }
        }
        return null;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Prüft, ob eine Eigenschaft mit der übergebenen ID in der Eigenschaften-Liste vorhanden ist
    // Return TRUE, falls ja
    //        FALSE, sonst
    //----------------------------------------------------------------------------------------------------------------

    public boolean EigenschaftVorhanden (String ID){
        boolean retval = false;
        for(int num=0; num<EigenschaftenListe.size(); num++)
        {
            Eigenschaft E = (Eigenschaft)EigenschaftenListe.get(num);
            String S1 = E.GetID();
            if (S1.equals(ID)) {
                retval = true;
                break;
            }
        }
        return retval;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Sucht die Eigenschaft mit der übergebenen ID in der Eigenschaften-Liste und gibt diese zurück
    // Return Eigenschaft mit ID, wenn in Liste vorhanden
    //        NULL, wenn Eigenschaft nicht gefunden
    //----------------------------------------------------------------------------------------------------------------

    public Eigenschaft GetEigenschaft (String ID){
        for(int num=0; num<EigenschaftenListe.size(); num++)
        {
            Eigenschaft E = (Eigenschaft)EigenschaftenListe.get(num);
            String S1 = E.GetID();
            if (S1.equals(ID)) {
                return E;
            }
        }
        return null;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Prüft, ob der Datentyp zulässig ist. Erlaubt sind Double, Integer, Boolean und String
    // Die Prüfung erfolgt gegen die erlaubten Datentypen, die als Konstanten definiert sind (vgl. oben)
    // Return TRUE, falls Datentyp zulässig
    //        FALSE, sonst
    //----------------------------------------------------------------------------------------------------------------

    public boolean DatentypZulaessig (String DT){
        if (DT.equalsIgnoreCase(DATENTYP_Double)) return true;
        if (DT.equalsIgnoreCase(DATENTYP_Integer)) return true;
        if (DT.equalsIgnoreCase(DATENTYP_Boolean)) return true;
        if (DT.equalsIgnoreCase(DATENTYP_String)) return true;
        return false;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Prüft, ob der Operator zulässig ist. Erlaubt sind =, <>, >, >=, <, <=
    // Die Prüfung erfolgt gegen die erlaubten Operatoren, die als Konstanten definiert sind (vgl. oben)
    // Return TRUE, falls Operator zulässig
    //        FALSE, sonst
    //----------------------------------------------------------------------------------------------------------------

    public boolean OperatorZulaessig (String OP){
        if (OP.equalsIgnoreCase(OPERATOR_Gleich)) return true;
        if (OP.equalsIgnoreCase(OPERATOR_UnGleich)) return true;
        if (OP.equalsIgnoreCase(OPERATOR_Groesser)) return true;
        if (OP.equalsIgnoreCase(OPERATOR_GroesserGleich)) return true;
        if (OP.equalsIgnoreCase(OPERATOR_Kleiner)) return true;
        if (OP.equalsIgnoreCase(OPERATOR_KleinerGleich)) return true;
        return false;
    }

    //----------------------------------------------------------------------------------------------------------------
    // Liest die Parameter aus dem übergebenen Excel-Tabsheet in die Parameter-Liste ein
    // Prüfung des Datentyps auf Zulässigkeit
    // Doppelte Parameter werden nicht der Liste hinzugefügt
    //----------------------------------------------------------------------------------------------------------------

    public void ParameterEinlesen (Sheet ParameterSheet) {

    // Deklaration und Initialisierung der Variablen eines Parameters (siehe Excel-Liste)
    String ID = "";
    String Name = "";
    String Datentyp = "";

    // Deklaration der Spalten-Indizes der Variablen eines Parameters in der Excel-Liste
    // Initialisierung mit -1 (unzulässiger Spalten-Index)
    // Nach Lesen der Zeile mit den Spaltenüberschriften in der Excel-Liste werden diese Indizes auf die tatsächlichen Spalten-Indizes gesetzt.
    // Anordnung der einzelnen Spalten in Excel-Liste ist daher beliebig.
    int idx_id = -1;
    int idx_name = -1;
    int idx_datentyp = -1;

    PWriter.println(" ");
    PWriter.println("  ------------------------");
    PWriter.println("  Start Parameter Einlesen");
    PWriter.println("  ------------------------");

    // Durchlaufen aller Zeilen und Spalten des übergebenen Parameter-Sheets
    for (Iterator<Row> rowIterator = ParameterSheet.rowIterator(); rowIterator.hasNext(); ) {
        Row row = rowIterator.next();
        if (row.getRowNum() == 0) {
            // Überschriften-Zeile
            // Setzen der Spalten-Indizes, Groß- und Kleinschreibung in der Überschriften-Zeile beliebig
            for (Iterator<Cell> cellIterator = row.cellIterator(); cellIterator.hasNext(); ) {
                Cell cell = cellIterator.next();
                if (cell.getCellType().equals(CellType.STRING)) {
                    // Zelle vom Typ String
                    if (cell.getStringCellValue().equalsIgnoreCase("ID")) {
                        idx_id = cell.getColumnIndex();
                    } else if (cell.getStringCellValue().equalsIgnoreCase("NAME")) {
                        idx_name = cell.getColumnIndex();
                    } else if (cell.getStringCellValue().equalsIgnoreCase("DATENTYP")) {
                        idx_datentyp = cell.getColumnIndex();
                    } else {
                        PWriter.println("  Falscher Spalten-Header (" + cell.getStringCellValue() + ") >> Abbruch");
                        return;
                    }
                } else if (cell.getCellType().equals(CellType.BLANK)) {
                    // Zelle vom Typ Blank (kein Inhalt) >> Tue nichts ...
                } else {
                    // Zelle von anderem Typ
                    PWriter.println("  Überschrift mit unzulässigem Typ (nicht String)");
                }
            }
        } else{
            // Ab zweiter Zeile: Parameter einlesen
            ID = "";
            Name = "";
            Datentyp = "";

            if (idx_id < 0) {
                // Spalte mit ID nicht vorhanden
                PWriter.println("  Spalten-Header (ID) nicht vorhanden >> Abbruch");
                return;
            }
            Cell cell = row.getCell(idx_id);
            if (cell != null){
                if (cell.getCellType().equals(CellType.STRING)) {
                    ID = cell.getStringCellValue();
                } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                    ID = String.valueOf(cell.getNumericCellValue());
                } else if (cell.getCellType().equals(CellType.BLANK)) {
                    ID = "";
                } else {
                    PWriter.println("  Falscher Zellenwert für ID (nicht numerisch und nicht String)");
                }

                if (idx_name < 0) {
                    // Spalte mit Name nicht vorhanden
                    PWriter.println("  Spalten-Header (NAME) nicht vorhanden >> Abbruch");
                }
                cell = row.getCell(idx_name);
                if (cell.getCellType().equals(CellType.STRING)) {
                    Name = cell.getStringCellValue();
                } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                    Name = String.valueOf(cell.getNumericCellValue());
                } else if (cell.getCellType().equals(CellType.BLANK)) {
                    Name = "";
                } else {
                    PWriter.println("  Falscher Zellenwert für NAME (nicht numerisch und nicht String)");
                }

                if (ID != "") {
                    if (Name != "") {
                        if (idx_datentyp < 0) {
                            // Spalte mit Datentyp nicht vorhanden
                            PWriter.println("  Spalten-Header (DATENTYP) nicht vorhanden >> Abbruch");
                        }
                        cell = row.getCell(idx_datentyp);
                        if (cell.getCellType().equals(CellType.STRING)) {
                            Datentyp = cell.getStringCellValue();
                        } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                            Datentyp = String.valueOf(cell.getNumericCellValue());
                        } else if (cell.getCellType().equals(CellType.BLANK)) {
                            Datentyp = "";
                        } else {
                            PWriter.println("  Falscher Zellenwert für DATENTYP (nicht numerisch und nicht String)");
                        }

                        if (ID != "") {
                            if (DatentypZulaessig(Datentyp)){
                                if (ParameterVorhanden(ID)) {
                                    PWriter.println("  ++FEHLER++ Parameter: " + ID + " bereits vorhanden");
                                } else {
                                    ParameterListe.add(new Parameter (ID, Name, Datentyp));
                                    PWriter.println("  - Neuer Parameter: " + ID + " - " + Name + " - " + Datentyp);
                                }
                            }
                            else {
                                PWriter.println("  ++FEHLER++ Falscher Wert für DATENTYP (" + Datentyp + ") >> Parameter " + ID + " nicht eingelesen!");
                            }
                        }
                    }
                    else {
                        PWriter.println("  ++FEHLER++ Spalte Name nicht gefüllt >> Parameter " + ID + " nicht eingelesen!");
                    }
                }
            }
        }
    }
    PWriter.println("  Ende  Parameter Einlesen");
    PWriter.println("  ------------------------");
}

    //----------------------------------------------------------------------------------------------------------------
    // Liest die Eigenschaften aus dem übergebenen Excel-Tabsheet in die Eigenschaften-Liste ein
    // Prüfung des Datentyps auf Zulässigkeit
    // Prüfung, ob der Parameter für den Vergleich vorhanden ist
    // Prüfung, ob Datentyp von Parameter und Datentyp von Eigenschaft gleich sind
    // Doppelte Eigenschaften werden nicht der Liste hinzugefügt
    //----------------------------------------------------------------------------------------------------------------

    public void EigenschaftenEinlesen (Sheet EigenschaftenSheet) {
        String ID = "";
        String Name = "";
        String Datentyp = "";
        String IDParameter = "";
        int idx_id = -1;
        int idx_name = -1;
        int idx_datentyp = -1;
        int idx_idparameter = -1;

        PWriter.println(" ");
        PWriter.println("  ----------------------------");
        PWriter.println("  Start Eigenschaften Einlesen");
        PWriter.println("  ----------------------------");

        for (Iterator<Row> rowIterator = EigenschaftenSheet.rowIterator(); rowIterator.hasNext(); ) {
            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                for (Iterator<Cell> cellIterator = row.cellIterator(); cellIterator.hasNext(); ) {
                    Cell cell = cellIterator.next();
                    if (cell.getCellType().equals(CellType.STRING)) {
                        if (cell.getStringCellValue().equalsIgnoreCase("ID")) {
                            idx_id = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("NAME")) {
                            idx_name = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("DATENTYP")) {
                            idx_datentyp = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("ID_PARAMETER")) {
                            idx_idparameter = cell.getColumnIndex();
                        } else {
                            PWriter.println("  Falscher Spalten-Header (" + cell.getStringCellValue() + ") >> Abbruch");
                            return;
                        }
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        // Nichts ...
                    } else {
                        PWriter.println("  Überschrift mit unzulässigerm Typ (nicht String)");
                    }
                }
            } else{
                ID = "";
                Name = "";
                Datentyp = "";
                IDParameter = "";

                if (idx_id < 0) {
                    PWriter.println("  Spalten-Header (ID) nicht vorhanden >> Abbruch");
                    return;
                }
                Cell cell = row.getCell(idx_id);
                if (cell != null){
                    if (cell.getCellType().equals(CellType.STRING)) {
                        ID = cell.getStringCellValue();
                    } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                        ID = String.valueOf(cell.getNumericCellValue());
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        ID = "";
                    } else {
                        PWriter.println("  Falscher Zellenwert für ID (nicht numerisch und nicht String)");
                    }

                    if (ID != "") {
                        if (idx_name < 0) {
                            PWriter.println("  Spalten-Header (NAME) nicht vorhanden >> Abbruch");
                        }
                        cell = row.getCell(idx_name);
                        if (cell.getCellType().equals(CellType.STRING)) {
                            Name = cell.getStringCellValue();
                        } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                            Name = String.valueOf(cell.getNumericCellValue());
                        } else if (cell.getCellType().equals(CellType.BLANK)) {
                            Name = "";
                        } else {
                            PWriter.println("  Falscher Zellenwert für NAME (nicht numerisch und nicht String)");
                        }

                        if (Name != "") {
                            if (idx_datentyp < 0) {
                                PWriter.println("  Spalten-Header (DATENTYP) nicht vorhanden >> Abbruch");
                            }
                            cell = row.getCell(idx_datentyp);
                            if (cell.getCellType().equals(CellType.STRING)) {
                                Datentyp = cell.getStringCellValue();
                            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                Datentyp = String.valueOf(cell.getNumericCellValue());
                            } else if (cell.getCellType().equals(CellType.BLANK)) {
                                Datentyp = "";
                            } else {
                                PWriter.println("  Falscher Zellenwert für DATENTYP (nicht numerisch und nicht String)");
                            }

                            if (DatentypZulaessig(Datentyp)){
                                if (idx_idparameter < 0) {
                                    PWriter.println("  Spalten-Header (ID_PARAMETER) nicht vorhanden >> Abbruch");
                                }
                                cell = row.getCell(idx_idparameter);
                                if (cell != null) {
                                    if (cell.getCellType().equals(CellType.STRING)) {
                                        IDParameter = cell.getStringCellValue();
                                    } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                        IDParameter = String.valueOf(cell.getNumericCellValue());
                                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                                        IDParameter = "";
                                    } else {
                                        PWriter.println("  Falscher Zellenwert für ID_PARAMETER (nicht numerisch und nicht String)");
                                    }

                                    if (ID != "") {
                                        if (EigenschaftVorhanden(ID)) {
                                            PWriter.println("  ++FEHLER++ Eigenschaft: " + ID + " bereits vorhanden");
                                        } else {
                                            // Test, ob Parameter vorhanden
                                            if (ParameterVorhanden(IDParameter)) {
                                                // Wenn Parameter vorhanden, dann  muss auch der Datentyp der Eigenschaft gleich dem Datentyp des Parameters sein
                                                Parameter P = GetParameter (IDParameter);
                                                String DT = P.GetDatentyp();
                                                if (DT.equals(Datentyp)) {
                                                    EigenschaftenListe.add(new Eigenschaft (ID, Name, Datentyp, IDParameter));
                                                    PWriter.println("  - Neue Eigenschaft: " + ID + " - " + Name + " - " + Datentyp + " - " + IDParameter);
                                                }
                                                else {
                                                    PWriter.println("  ++FEHLER++ Datentyp " + Datentyp + "ungleich Datentyp " + DT + "des Parameters >> Eigenschaft " + ID + " nicht eingelesen");
                                                }
                                            }
                                            else {
                                                PWriter.println("  ++FEHLER++ Parameter " + IDParameter + " nicht  vorhanden >>  Eigenschaft " + ID + " nicht eingelesen");
                                            }
                                        }
                                    }
                                }
                                else {
                                    PWriter.println("  ++FEHLER++ Falscher Zellenwert für ID_PARAMETER (nicht numerisch und nicht String) >> Eigenschaft " + ID + " nicht eingelesen");
                                }
                            }
                            else {
                                PWriter.println("  ++FEHLER++ Falscher Wert für DATENTYP (" + Datentyp + ") >> Eigenschaft " + ID + " nicht eingelesen");
                            }
                        }
                        else {
                            PWriter.println("  ++FEHLER++ Spalte Name nicht gefüllt >> Eigenschaft " + ID + " nicht eingelesen!");
                        }
                    }
                }
            }
        }
        PWriter.println("  Ende  Eigenschaften Einlesen");
        PWriter.println("  ----------------------------");
    }

    //----------------------------------------------------------------------------------------------------------------
    // Liest die Ressourcen aus dem übergebenen Excel-Tabsheet in die Ressourcen-Liste ein
    // Doppelte Ressourcen werden nicht der Liste hinzugefügt
    //----------------------------------------------------------------------------------------------------------------

    public void RessourcenEinlesen (Sheet RessourceSheet) {
        String ID = "";
        String Name = "";
        int idx_id = -1;
        int idx_name = -1;

        PWriter.println(" ");
        PWriter.println("  -------------------------");
        PWriter.println("  Start Ressourcen Einlesen");
        PWriter.println("  -------------------------");

        for (Iterator<Row> rowIterator = RessourceSheet.rowIterator(); rowIterator.hasNext(); ) {
            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                for (Iterator<Cell> cellIterator = row.cellIterator(); cellIterator.hasNext(); ) {
                    Cell cell = cellIterator.next();
                    if (cell.getCellType().equals(CellType.STRING)) {
                        if (cell.getStringCellValue().equalsIgnoreCase("ID")) {
                            idx_id = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("NAME")) {
                            idx_name = cell.getColumnIndex();
                        } else {
                            PWriter.println("  Falscher Spalten-Header (" + cell.getStringCellValue() + ") >> Abbruch");
                            return;
                        }
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        // Nichts ...
                    } else {
                        PWriter.println("  Überschrift mit unzulässigerm Typ (nicht String)");
                    }
                }
            } else{
                ID = "";
                Name = "";
                // ID lesen
                if (idx_id < 0) {
                    PWriter.println("  Spalten-Header (ID) nicht vorhanden >> Abbruch");
                }
                Cell cell = row.getCell(idx_id);
                if (cell != null){
                    if (cell.getCellType().equals(CellType.STRING)) {
                        ID = cell.getStringCellValue();
                    } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                        ID = String.valueOf(cell.getNumericCellValue());
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        ID = "";
                    } else {
                        PWriter.println("  Falscher Zellenwert für ID (nicht numerisch und nicht String)");
                    }

                    if (ID != "") {
                        if (idx_name < 0) {
                            PWriter.println("  Spalten-Header (NAME) nicht vorhanden >> Abbruch");
                            return;
                        }
                        cell = row.getCell(idx_name);
                        if (cell != null){
                            if (cell.getCellType().equals(CellType.STRING)) {
                                Name = cell.getStringCellValue();
                            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                Name = String.valueOf(cell.getNumericCellValue());
                            } else if (cell.getCellType().equals(CellType.BLANK)) {
                                Name = "";
                            } else {
                                PWriter.println("  ++FEHLER++ Falscher Zellenwert für NAME (nicht numerisch und nicht String) >> Ressource " + ID + " nicht eingelesen");
                            }

                            if (Name != ""){
                                if (RessourceVorhanden(ID)) {
                                    PWriter.println("  ++FEHLER++ Ressource: " + ID + " bereits vorhanden");
                                } else {
                                    RessourcenListe.add(new Ressource (ID, Name));
                                    PWriter.println("  - Neue Ressource: " + ID + " - " + Name);
                                }
                            }
                            else {
                                PWriter.println("  ++FEHLER++ Spalte Name nicht gefüllt >> Ressource " + ID + " nicht eingelesen!");
                            }
                        }
                        else {
                            PWriter.println("  ++FEHLER++ Spalte NAME nicht gefüllt >> Ressource " + ID + " nicht eingelesen");
                        }
                    }
                }
            }
        }
        PWriter.println("  Ende  Ressourcen Einlesen");
        PWriter.println("  -------------------------");
    }

    //----------------------------------------------------------------------------------------------------------------
    // Liest die Produkte aus dem übergebenen Excel-Tabsheet in die Produkt-Liste ein
    // Doppelte Produkte werden nicht der Liste hinzugefügt
    //----------------------------------------------------------------------------------------------------------------


    public void ProdukteEinlesen (Sheet ProdukteSheet) {
        String ID = "";
        String Name = "";
        int idx_id = -1;
        int idx_name = -1;

        PWriter.println(" ");
        PWriter.println("  -----------------------");
        PWriter.println("  Start Produkte Einlesen");
        PWriter.println("  -----------------------");

        for (Iterator<Row> rowIterator = ProdukteSheet.rowIterator(); rowIterator.hasNext(); ) {
            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                for (Iterator<Cell> cellIterator = row.cellIterator(); cellIterator.hasNext(); ) {
                    Cell cell = cellIterator.next();
                    if (cell.getCellType().equals(CellType.STRING)) {
                        if (cell.getStringCellValue().equalsIgnoreCase("ID")) {
                            idx_id = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("NAME")) {
                            idx_name = cell.getColumnIndex();
                        } else {
                            PWriter.println("  Falscher Spalten-Header (" + cell.getStringCellValue() + ") >> Abbruch");
                            return;
                        }
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        // Nichts ...
                    } else {
                        PWriter.println("  Überschrift mit unzulässigerm Typ (nicht String)");
                    }
                }
            } else{
                ID = "";
                Name = "";
                // ID lesen
                if (idx_id < 0) {
                    PWriter.println("  Spalten-Header (ID) nicht vorhanden >> Abbruch");
                    return;
                }
                Cell cell = row.getCell(idx_id);
                if (cell != null){
                    if (cell.getCellType().equals(CellType.STRING)) {
                        ID = cell.getStringCellValue();
                    } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                        ID = String.valueOf(cell.getNumericCellValue());
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        ID = "";
                    } else {
                        PWriter.println("  Falscher Zellenwert für ID (nicht numerisch und nicht String)");
                    }

                    if (ID != "") {
                        if (idx_name < 0) {
                            PWriter.println("  Spalten-Header (NAME) nicht vorhanden >> Abbruch");
                        }
                        cell = row.getCell(idx_name);
                        if (cell != null){
                            if (cell.getCellType().equals(CellType.STRING)) {
                                Name = cell.getStringCellValue();
                            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                Name = String.valueOf(cell.getNumericCellValue());
                            } else if (cell.getCellType().equals(CellType.BLANK)) {
                                Name = "";
                            } else {
                                PWriter.println("  ++FEHLER++ Falscher Zellenwert für NAME (nicht numerisch und nicht String) >> Produkt " + ID + " nicht eingelesen");
                            }

                            if (ProduktVorhanden(ID)) {
                                PWriter.println("  ++FEHLER++ Produkt: " + ID + " bereits vorhanden");
                            } else {
                                ProduktListe.add(new Produkt (ID, Name));
                                PWriter.println("  - Neue Produkt: " + ID + " - " + Name);
                            }
                        }
                        else {
                            PWriter.println("  ++FEHLER++ Spalte NAME nicht gefüllt >> Produkt " + ID + " nicht eingelesen");
                        }
                    }
                }
            }
        }
        PWriter.println("  Ende  Produkte Einlesen");
        PWriter.println("  -----------------------");
    }

    //----------------------------------------------------------------------------------------------------------------
    // Liest die Ressourcenparameter aus dem übergebenen Excel-Tabsheet ein und fügt sie in die Liste bei der zugehörigen Ressource ein.
    // Dafür müssen die zugehörige Ressource und der zugehörige Parameter vorhanden sein.
    // Bei Datentyp String und Boolean wird nur Operator 1 und Wert 1 verwendet.
    // Bei Datentyp Double und Integer kann auch Operator 2 und Wert 2 verwendet werden (->Intervall).
    //----------------------------------------------------------------------------------------------------------------

    public void RessourcenparameterEinlesen (Sheet RessourcenparameterSheet) {
        String IDRessource = "";
        String IDParameter = "";
        String Operator1 = "";
        String Operator2 = "";
        String Einheit = "";
        double DWert1 = 0;
        double DWert2 = 0;
        String SWert = "";
        Boolean BWert = false;
        int idx_idressource = -1;
        int idx_idparameter = -1;
        int idx_operator1 = -1;
        int idx_operator2 = -1;
        int idx_wert1 = -1;
        int idx_wert2 = -1;
        int idx_einheit = -1;

        PWriter.println(" ");
        PWriter.println("  ----------------------------------");
        PWriter.println("  Start Ressourcenparameter Einlesen");
        PWriter.println("  ----------------------------------");

        for (Iterator<Row> rowIterator = RessourcenparameterSheet.rowIterator(); rowIterator.hasNext(); ) {
            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                for (Iterator<Cell> cellIterator = row.cellIterator(); cellIterator.hasNext(); ) {
                    Cell cell = cellIterator.next();
                    if (cell.getCellType().equals(CellType.STRING)) {
                        if (cell.getStringCellValue().equalsIgnoreCase("ID_RESSOURCE")) {
                            idx_idressource = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("ID_PARAMETER")) {
                            idx_idparameter = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("OPERATOR1")) {
                            idx_operator1 = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("OPERATOR2")) {
                            idx_operator2 = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("WERT1")) {
                            idx_wert1 = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("WERT2")) {
                            idx_wert2 = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("EINHEIT")) {
                            idx_einheit = cell.getColumnIndex();
                        } else {
                            PWriter.println("  Falscher Spalten-Header (" + cell.getStringCellValue() + ") >> Abbruch");
                            return;
                        }
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        // Nichts ...
                    } else {
                        PWriter.println("  Überschrift mit unzulässigerm Typ (nicht String)");
                    }
                }
            } else{
                IDRessource = "";
                IDParameter = "";
                Operator1 = "";
                Operator2 = "";
                Einheit = "";
                DWert1 = 0;
                DWert2 = 0;
                SWert = "";
                BWert = false;

                // IDRessource lesen
                if (idx_idressource < 0) {
                    PWriter.println("  Spalten-Header (ID_RESSOURCE) nicht vorhanden >> Abbruch");
                    return;
                }
                Cell cell = row.getCell(idx_idressource);
                if (cell != null){
                    if (cell.getCellType().equals(CellType.STRING)) {
                        IDRessource = cell.getStringCellValue();
                    } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                        IDRessource = String.valueOf(cell.getNumericCellValue());
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        IDRessource = "";
                    } else {
                        PWriter.println("  Falscher Zellenwert für ID_RESSOURCE (nicht numerisch und nicht String)");
                    }

                    if (IDRessource != "") {
                        // Prüfen, ob Ressource vorhanden
                        Ressource R = GetRessource(IDRessource);
                        if (R != null){
                            if (idx_idparameter < 0) {
                                PWriter.println("  Spalten-Header (ID_PARAMETER) nicht vorhanden >> Abbruch");
                                return;
                            }
                            cell = row.getCell(idx_idparameter);
                            if (cell != null) {
                                if (cell.getCellType().equals(CellType.STRING)) {
                                    IDParameter = cell.getStringCellValue();
                                } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                    IDParameter = String.valueOf(cell.getNumericCellValue());
                                } else if (cell.getCellType().equals(CellType.BLANK)) {
                                    IDParameter = "";
                                } else {
                                    PWriter.println("  Falscher Zellenwert für ID_PARAMETER (nicht numerisch und nicht String)");
                                }

                                // Prüfen, ob der Parameter vorhanden ist
                                Parameter P = GetParameter(IDParameter);
                                if (P != null){
                                    // restliche Daten einlesen ....
                                    // Operator1 lesen
                                    if (idx_operator1 < 0) {
                                        PWriter.println("  Spalten-Header (OPERATOR1) nicht vorhanden >> Abbruch");
                                    }
                                    cell = row.getCell(idx_operator1);
                                    if (cell != null) {
                                        if (cell.getCellType().equals(CellType.STRING)) {
                                            Operator1 = cell.getStringCellValue();
                                        } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                            Operator1 = String.valueOf(cell.getNumericCellValue());
                                        } else if (cell.getCellType().equals(CellType.BLANK)) {
                                            Operator1 = "";
                                        } else {
                                            PWriter.println("  ++FEHLER++ Falscher Zellenwert für OPERATOR1 (nicht numerisch und nicht String) >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                            continue;
                                        }

                                        // Operator 1 muss immer gefüllt sein.
                                        if (Operator1 != ""){
                                            if (OperatorZulaessig(Operator1) == false){
                                                PWriter.println("  ++FEHLER++ Operator1 hat einen unzulässigen Wert: " + Operator1 + " >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                                continue;
                                            }
                                        }
                                        else {
                                            PWriter.println("  ++FEHLER++ Spalte Operator1 nicht gefüllt >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                            continue;
                                        }

                                        // Operator2 lesen
                                        if (idx_operator2 < 0) {
                                            PWriter.println("  Spalten-Header (OPERATOR2) nicht vorhanden >> Abbruch");
                                            return;
                                        }
                                        cell = row.getCell(idx_operator2);
                                        if (cell != null) {
                                            if (cell.getCellType().equals(CellType.STRING)) {
                                                Operator2 = cell.getStringCellValue();
                                            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                                Operator2 = String.valueOf(cell.getNumericCellValue());
                                            } else if (cell.getCellType().equals(CellType.BLANK)) {
                                                Operator2 = "";
                                            } else {
                                                PWriter.println("  Falscher Zellenwert für OPERATOR2 (nicht numerisch und nicht String)");
                                                return;
                                            }
                                        }
                                        else {
                                            Operator2 = "";
                                        }

                                        // Operator 2 kann auch leer sein. Wenn gefüllt, dann muss der Operator zulässig sein
                                        if (Operator2 != ""){
                                            if (OperatorZulaessig(Operator2) == false){
                                                PWriter.println("  ++FEHLER++ Operator2 hat einen unzulässigen Wert: " + Operator2 + " >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                                continue;
                                            }
                                        }

                                        // Einheit lesen
                                        if (idx_einheit < 0) {
                                            PWriter.println("  Spalten-Header (EINHEIT) nicht vorhanden >> Abbruch");
                                        }
                                        cell = row.getCell(idx_einheit);
                                        if (cell != null) {
                                            if (cell.getCellType().equals(CellType.STRING)) {
                                                Einheit = cell.getStringCellValue();
                                            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                                Einheit = String.valueOf(cell.getNumericCellValue());
                                            } else if (cell.getCellType().equals(CellType.BLANK)) {
                                                Einheit = "";
                                            } else {
                                                PWriter.println("  Falscher Zellenwert für EINHEIT (nicht numerisch und nicht String)");
                                            }
                                        }
                                        else {
                                            Einheit = "";
                                        }

                                        // Wert 1 einlesen
                                        // Abhängig vom Datentyp des Parameters unterschiedliche Varianten ...
                                        if (P.GetDatentyp().equals(DATENTYP_String)){
                                            if (idx_wert1 < 0) {
                                                PWriter.println("  Spalten-Header (WERT1) nicht vorhanden >> Abbruch");
                                                return;
                                            }
                                            cell = row.getCell(idx_wert1);
                                            if (cell.getCellType().equals(CellType.STRING)) {
                                                SWert = cell.getStringCellValue();
                                            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                                SWert = String.valueOf(cell.getNumericCellValue());
                                            } else if (cell.getCellType().equals(CellType.BLANK)) {
                                                SWert = "";
                                            } else {
                                                PWriter.println("  ++FEHLER++ Spalte Wert1 hat falschen Typ (nicht String) >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                                continue;
                                            }

                                            // Ressourcenparameter anlegen
                                            R.NeuerRessourcenParameterString(P, Operator1, SWert, Einheit, PWriter);
                                        }
                                        else if (P.GetDatentyp().equals(DATENTYP_Boolean)){
                                            if (idx_wert1 < 0) {
                                                PWriter.println("  Spalten-Header (WERT1) nicht vorhanden >> Abbruch");
                                                return;
                                            }
                                            cell = row.getCell(idx_wert1);
                                            if (cell.getCellType().equals(CellType.BOOLEAN)) {
                                                BWert = cell.getBooleanCellValue();
                                            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                                int inthelp = (int)cell.getNumericCellValue();
                                                if (inthelp == 0) {
                                                    BWert = false;
                                                }
                                                else {
                                                    BWert = true;
                                                }
                                            } else {
                                                PWriter.println("  ++FEHLER++ Spalte Wert1 hat falschen Typ (nicht Boolean) >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                                continue;
                                            }

                                            // Ressourcenparameter erstellen
                                            R.NeuerRessourcenParameterBoolean(P, Operator1, BWert, Einheit, PWriter);
                                        }
                                        else if (P.GetDatentyp().equals(DATENTYP_Double) || P.GetDatentyp().equals(DATENTYP_Integer)){
                                            if (idx_wert1 < 0) {
                                                PWriter.println("  Spalten-Header (WERT1) nicht vorhanden >> Abbruch");
                                                return;
                                            }
                                            cell = row.getCell(idx_wert1);
                                            if (cell == null) {
                                                PWriter.println("  ++FEHLER++ Spalte Wert1 hat falschen Typ (nicht numerisch) >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                                continue;
                                            }
                                            if (cell.getCellType().equals(CellType.NUMERIC)) {
                                                DWert1 = cell.getNumericCellValue();
                                            } else {
                                                PWriter.println("  ++FEHLER++ Spalte Wert1 hat falschen Typ (nicht numerisch) >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                                continue;
                                            }

                                            if (Operator2 != "") {
                                                // Für Intervalle
                                                // Wenn Operator 2 gefüllt ist, dann muss auch Wert 2 eingelesen werden.
                                                if (idx_wert2 < 0) {
                                                    PWriter.println("  Spalten-Header (WERT2) nicht vorhanden >> Abbruch");
                                                    return;
                                                }
                                                cell = row.getCell(idx_wert2);
                                                if (cell == null) {
                                                    PWriter.println("  ++FEHLER++ Spalte Wert2 hat falschen Typ (nicht numerisch) >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                                    continue;
                                                }
                                                if (cell.getCellType().equals(CellType.NUMERIC)) {
                                                    DWert2 = cell.getNumericCellValue();
                                                } else {
                                                    PWriter.println("  ++FEHLER++ Spalte Wert2 hat falschen Typ (nicht numerisch) >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                                    continue;
                                                }
                                            }

                                            // Ressourcenparameter erstellen
                                            R.NeuerRessourcenParameterDouble(P, Operator1, DWert1, Operator2, DWert2, Einheit, PWriter);
                                        }
                                    }
                                    else {
                                        PWriter.println("  ++FEHLER++ Spalte Operator1 nicht gefüllt >> Ressourcenparameter " + IDParameter + " für Ressource " + IDRessource + " nicht eingelesen");
                                    }
                                }
                                else {
                                    PWriter.println("  ++FEHLER++ Parameter " + IDParameter +  " nicht in Parameterliste vorhanden  >> Ressourcenparameter für Ressource " + IDRessource + " nicht eingelesen");
                                }
                            }
                            else {
                                PWriter.println("  ++FEHLER++ Spalte ID_PARAMETER nicht gefüllt >> Ressourcenparameter für Ressource " + IDRessource + " nicht eingelesen");
                            }
                        }
                        else {
                            PWriter.println("  ++FEHLER++ Ressource " + IDRessource + " nicht vorhanden. >> Ressourcenparameter nicht eingelesen");
                        }
                    }
                }
            }
        }
        PWriter.println("  Ende  Ressourcenparameter Einlesen");
        PWriter.println("  ----------------------------------");
    }

    //----------------------------------------------------------------------------------------------------------------
    // Liest die Produkteigenschaften aus dem übergebenen Excel-Tabsheet ein und fügt sie in die Liste bei dem zugehörigen Produkt ein.
    // Dafür müssen das zugehörige Produkt, die zugehörige Ressource und die zugehörige Eigenschaft vorhanden sein.
    //----------------------------------------------------------------------------------------------------------------

    public void ProdukteigenschaftEinlesen (Sheet ProdukteigenschaftenSheet) {
        String IDProdukt = "";
        String IDRessource = "";
        String IDEigenschaft = "";
        String Einheit = "";
        double DWert = 0;
        String SWert = "";
        Boolean BWert = false;
        int idx_idprodukt = -1;
        int idx_idressource = -1;
        int idx_ideigenschaft = -1;
        int idx_einheit = -1;
        int idx_wert = -1;

        PWriter.println(" ");
        PWriter.println("  ---------------------------------");
        PWriter.println("  Start Produkteigenschaft Einlesen");
        PWriter.println("  ---------------------------------");

        for (Iterator<Row> rowIterator = ProdukteigenschaftenSheet.rowIterator(); rowIterator.hasNext(); ) {
            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                for (Iterator<Cell> cellIterator = row.cellIterator(); cellIterator.hasNext(); ) {
                    Cell cell = cellIterator.next();
                    if (cell.getCellType().equals(CellType.STRING)) {
                        if (cell.getStringCellValue().equalsIgnoreCase("ID_PRODUKT")) {
                            idx_idprodukt = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("ID_RESSOURCE")) {
                            idx_idressource = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("ID_EIGENSCHAFT")) {
                            idx_ideigenschaft = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("WERT")) {
                            idx_wert = cell.getColumnIndex();
                        } else if (cell.getStringCellValue().equalsIgnoreCase("EINHEIT")) {
                            idx_einheit = cell.getColumnIndex();
                        } else {
                            PWriter.println("  Falscher Spalten-Header (" + cell.getStringCellValue() + ") >> Abbruch");
                            return;
                        }
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        // Nichts ...
                    } else {
                        PWriter.println("  Überschrift mit unzulässigerm Typ (nicht String)");
                    }
                }
            } else{
                IDProdukt = "";
                IDRessource = "";
                IDEigenschaft = "";
                Einheit = "";
                DWert = 0;
                SWert = "";
                BWert = false;

                // IDProdukt lesen
                if (idx_idprodukt < 0) {
                    PWriter.println("  Spalten-Header (ID_PRODUKT) nicht vorhanden >> Abbruch");
                    return;
                }
                Cell cell = row.getCell(idx_idprodukt);
                if (cell != null){
                    if (cell.getCellType().equals(CellType.STRING)) {
                        IDProdukt = cell.getStringCellValue();
                    } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                        IDProdukt = String.valueOf(cell.getNumericCellValue());
                    } else if (cell.getCellType().equals(CellType.BLANK)) {
                        IDProdukt = "";
                    } else {
                        PWriter.println("  Falscher Zellenwert für ID_PRODUKT (nicht numerisch und nicht String)");
                    }

                    if (IDProdukt != "") {
                        // Prüfen, ob Produkt vorhanden
                        Produkt P = GetProdukt(IDProdukt);
                        if (P != null){
                            // Ressource einlesen
                            if (idx_idressource < 0) {
                                PWriter.println("  Spalten-Header (ID_RESSOURCE) nicht vorhanden >> Abbruch");
                                return;
                            }
                            cell = row.getCell(idx_idressource);
                            if (cell == null) {
                                IDRessource = "";
                                PWriter.println("  ++FEHLER++ Ressource " + IDRessource + " nicht vorhanden. >> Produkteigenschaft für Produkt " + IDProdukt + " nicht eingelesen");
                                continue;
                            }
                            if (cell.getCellType().equals(CellType.STRING)) {
                                IDRessource = cell.getStringCellValue();
                            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                IDRessource = String.valueOf(cell.getNumericCellValue());
                            } else if (cell.getCellType().equals(CellType.BLANK)) {
                                IDRessource = "";
                            } else {
                                PWriter.println("  Falscher Zellenwert für ID_RESSOURCE (nicht numerisch und nicht String)");
                            }

                            // Prüfen, ob die Ressource vorhanden ist
                            Ressource R = GetRessource(IDRessource);
                            if (R != null){
                                // Eigenschaft einlesen ....
                                if (idx_ideigenschaft < 0) {
                                    PWriter.println("  Spalten-Header (ID_EIGENSCHAFT) nicht vorhanden >> Abbruch");
                                    return;
                                }
                                cell = row.getCell(idx_ideigenschaft);
                                if (cell.getCellType().equals(CellType.STRING)) {
                                    IDEigenschaft = cell.getStringCellValue();
                                } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                    IDEigenschaft = String.valueOf(cell.getNumericCellValue());
                                } else if (cell.getCellType().equals(CellType.BLANK)) {
                                    IDEigenschaft = "";
                                } else {
                                    PWriter.println("  Falscher Zellenwert für ID_EIGENSCHAFT (nicht numerisch und nicht String)");
                                }

                                // Prüfen, ob Eigenschaft vorhanden ist
                                Eigenschaft E = GetEigenschaft (IDEigenschaft);
                                if (E != null){
                                    // Einheit einlesen
                                    if (idx_einheit < 0) {
                                        PWriter.println("  Spalten-Header (EINHEIT) nicht vorhanden >> Abbruch");
                                    }
                                    cell = row.getCell(idx_einheit);
                                    if (cell != null) {
                                        if (cell.getCellType().equals(CellType.STRING)) {
                                            Einheit = cell.getStringCellValue();
                                        } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                            Einheit = String.valueOf(cell.getNumericCellValue());
                                        } else if (cell.getCellType().equals(CellType.BLANK)) {
                                            Einheit = "";
                                        } else {
                                            PWriter.println("  Falscher Zellenwert für EINHEIT (nicht numerisch und nicht String)");
                                        }
                                    }
                                    else {
                                        Einheit = "";
                                    }

                                    // Ist der benötigte Ressourcenparameter vorhanden? Dies muss so sein, da dieser zum Vergleich benötigt wird
                                    if (R.GetRessourcenParameter(E.GetIDParameter()) != null) {
                                        // Wert einlesen
                                        // Abhängig vom Datentyp des Parameters unterschiedliche Varianten ...
                                        if (E.GetDatentyp().equals(DATENTYP_String)){
                                            if (idx_wert < 0) {
                                                PWriter.println("  Spalten-Header (WERT) nicht vorhanden >> Abbruch");
                                                return;
                                            }
                                            cell = row.getCell(idx_wert);
                                            if (cell != null) {
                                                if (cell.getCellType().equals(CellType.STRING)) {
                                                    SWert = cell.getStringCellValue();
                                                } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                                    SWert = String.valueOf(cell.getNumericCellValue());
                                                } else if (cell.getCellType().equals(CellType.BLANK)) {
                                                    SWert = "";
                                                } else {
                                                    PWriter.println("  ++FEHLER++ Spalte Wert hat falschen Typ (nicht String) >> Produkteigenschaft für Produkt " + IDProdukt + " und Eigenschaft " + IDEigenschaft + " für Ressource " + IDRessource + " nicht eingelesen");
                                                    continue;
                                                }

                                                // Produkteigenschaft erstellen
                                                P.NeueProduktEigenschaftString(R, E, SWert, Einheit, PWriter);
                                            }
                                        }
                                        else if (E.GetDatentyp().equals(DATENTYP_Boolean)){
                                            if (idx_wert < 0) {
                                                PWriter.println("  Spalten-Header (WERT) nicht vorhanden >> Abbruch");
                                                return;
                                            }
                                            cell = row.getCell(idx_wert);
                                            if (cell != null) {
                                                if (cell.getCellType().equals(CellType.BOOLEAN)) {
                                                    BWert = cell.getBooleanCellValue();
                                                } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                                                    int inthelp = (int)cell.getNumericCellValue();
                                                    if (inthelp == 0) {
                                                        BWert = false;
                                                    }
                                                    else {
                                                        BWert = true;
                                                    }
                                                } else {
                                                    PWriter.println("  ++FEHLER++ Spalte Wert hat falschen Typ (nicht Boolean) >> Produkteigenschaft für Produkt " + IDProdukt + " und Eigenschaft " + IDEigenschaft + " für Ressource " + IDRessource + " nicht eingelesen");
                                                    continue;
                                                }

                                                // Produkteigenschaft erstellen
                                                P.NeueProduktEigenschaftBoolean(R, E, BWert, Einheit, PWriter);
                                            }
                                        }
                                        else if (E.GetDatentyp().equals(DATENTYP_Double) || E.GetDatentyp().equals(DATENTYP_Integer)){
                                            if (idx_wert < 0) {
                                                PWriter.println("  Spalten-Header (WERT) nicht vorhanden >> Abbruch");
                                                return;
                                            }
                                            cell = row.getCell(idx_wert);
                                            if (cell.getCellType().equals(CellType.NUMERIC)) {
                                                DWert = cell.getNumericCellValue();
                                            } else {
                                                PWriter.println("  ++FEHLER++ Spalte Wert hat falschen Typ (nicht numerisch) >> Produkteigenschaft für Produkt " + IDProdukt + " und Eigenschaft " + IDEigenschaft + " für Ressource " + IDRessource + " nicht eingelesen");
                                                continue;
                                            }

                                            // Produkteigenschaft erstellen
                                            P.NeueProduktEigenschaftDouble(R, E, DWert, Einheit, PWriter);
                                        }

                                    }
                                    else {
                                        PWriter.println("  ++FEHLER++ Eigenschaft: " + IDEigenschaft + "  Fehlender Parameter " + E.GetIDParameter() + " bei Ressource " + R.GetID() + " >> Produkteigenschaft für Produkt " + IDProdukt + " nicht eingelesen!");
                                        continue;
                                    }
                                }
                                else {
                                    PWriter.println("  ++FEHLER++ Eigenschaft " + IDEigenschaft + " nicht vorhanden. >> Produkteigenschaft für Produkt " + IDProdukt + " nicht eingelesen");
                                    continue;
                                }
                            }
                            else {
                                PWriter.println("  ++FEHLER++ Ressource " + IDRessource + " nicht vorhanden. >> Produkteigenschaft für Produkt " + IDProdukt + " nicht eingelesen");
                                continue;
                            }
                        }
                        else {
                            PWriter.println("  ++FEHLER++ Produkt " + IDProdukt + " nicht vorhanden. >> Produkteigenschaft nicht eingelesen");
                            continue;
                        }
                    }
                }
            }
        }
        PWriter.println("  Ende  Produkteigenschaft Einlesen");
        PWriter.println("  ---------------------------------");
    }

    //----------------------------------------------------------------------------------------------------------------
    // Excel-Datei zum Einlesen der Daten kann ausgewählt werden
    // Für die einzelnen Excel-Tabsheets werden die jeweiligen Einlesemethoden aufgerufen.
    //----------------------------------------------------------------------------------------------------------------

    public void DatenEinlesen() throws IOException{

        JFileChooser RessourcenDatei = new JFileChooser();
        int returnValue = RessourcenDatei.showOpenDialog(null);

        PWriter.println(" ");
        PWriter.println("---------------------------------------------------------------------------");
        PWriter.println("Start Daten einlesen");
        PWriter.println("---------------------------------------------------------------------------");

        if (returnValue == JFileChooser.APPROVE_OPTION) {
            try {
                Workbook workbook = new XSSFWorkbook(new FileInputStream(RessourcenDatei.getSelectedFile()));
                Sheet ParameterTabSheet = workbook.getSheet(TABNAME_Parameter);
                ParameterEinlesen (ParameterTabSheet);

                Sheet EigenschaftenTabSheet = workbook.getSheet(TABNAME_Eigenschaft);
                EigenschaftenEinlesen (EigenschaftenTabSheet);

                Sheet RessourcenTabSheet = workbook.getSheet(TABNAME_Ressource);
                RessourcenEinlesen(RessourcenTabSheet);

                Sheet ProduktTabSheet = workbook.getSheet(TABNAME_Produkt);
                ProdukteEinlesen(ProduktTabSheet);

                Sheet RessourcenparameterTabSheet = workbook.getSheet(TABNAME_Ressourcenparameter);
                RessourcenparameterEinlesen(RessourcenparameterTabSheet);

                Sheet ProdukteigenschaftenTabSheet = workbook.getSheet(TABNAME_Produkteigenschaft);
                ProdukteigenschaftEinlesen(ProdukteigenschaftenTabSheet);
            }

            catch (Exception e) {

            }
        }

        PWriter.println(" ");
        PWriter.println("Ende  Daten einlesen");
        PWriter.println("---------------------------------------------------------------------------");
    }

    //----------------------------------------------------------------------------------------------------------------
    // Liste der Produkte wird durchlaufen.
    // Pro Produkt wird die zugehörige Liste der Produkteigenschaften durchlaufen.
    // Pro Produkteigenschaft wird geprüft, ob der Wert der Produkteigenschaft vom entsprechenden Ressourcenparameter erfüllt wird.
    //----------------------------------------------------------------------------------------------------------------

    public void ParameterVergleichen() throws IOException{

        PWriter.println(" ");
        PWriter.println("---------------------------------------------------------------------------");
        PWriter.println("Start ParameterVergleichen");
        PWriter.println("---------------------------------------------------------------------------");

        // Alle eingelesenen Produkte betrachten
        for(int num=0; num<ProduktListe.size(); num++) {

            boolean testergebnis = true;

            Produkt P = (Produkt)ProduktListe.get(num);
            PWriter.println(" ");
            PWriter.println("Produkt: " + P.GetName() + " Parameter: " + P.GetID());

            // Jetzt für das Produkt alle Produkteigenschaften prüfen ...
            testergebnis = P.ProduktEigenschaftenPruefen(PWriter);

            if (testergebnis){
                PWriter.println("Produkt: " + P.GetName() + " Parameter: " + P.GetID() + "   OK");
            }
            else {
                PWriter.println("Produkt: " + P.GetName() + " Parameter: " + P.GetID() + "   NICHT OK");
            }
        }

        PWriter.println(" ");
        PWriter.println("---------------------------------------------------------------------------");
        PWriter.println("Ende  ParameterVergleichen");
        PWriter.println("---------------------------------------------------------------------------");
    }

    //Methode zum Ausgeben der gespeicherten Vergleichsdaten in der Konsole zum Testen

    public void DatenAusgeben() throws IOException{
        for(int num=0; num<ProduktListe.size(); num++) {
            Produkt P = (Produkt)ProduktListe.get(num);
            System.out.println("Das Produkt " + P.GetName() + " mit der ID " + P.GetID() + " ist " + P.getProduktAnforderungenErfuellt());

            for(int num2=0; num2<P.GetProdukteigenschaftenliste().size(); num2++) {
                Produkteigenschaft PE = (Produkteigenschaft) P.GetProdukteigenschaftenliste().get(num2);
                System.out.println("  - Die Produkteigenschaft " + P.GetProdukteigenschaftenliste().get(num2).GetEigenschaftName() + " des Produktes " + P.GetName() + " mit der ID " + P.GetProdukteigenschaftenliste().get(num2).GetEigenschaftID() + " ist " + P.GetProdukteigenschaftenliste().get(num2).getIstErfuellt());
                if (PE.GetReferenzAufRessourcenParameter().GetReferenzAufParameter().GetDatentyp().equals("Double")){
                    System.out.println("    - Die Differenz beträgt " + P.GetProdukteigenschaftenliste().get(num2).getDifferenz());
                }
            }
        }

    }

    public void AusgabeDateiErzeugen() throws IOException{
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Vergleichsergebnis");
        Row rowÜberschriften = sheet.createRow(0);
        rowÜberschriften.createCell(0).setCellValue("Produkt Name");                    //Spalte A
        rowÜberschriften.createCell(1).setCellValue("Produkt ID");                      //Spalte B
        rowÜberschriften.createCell(2).setCellValue("Eigenschaft Name");                //Spalte C
        rowÜberschriften.createCell(3).setCellValue("Eigenschaft ID");                  //Spalte D
        rowÜberschriften.createCell(4).setCellValue("Produkteigenschaft Wert");         //Spalte E
        rowÜberschriften.createCell(5).setCellValue("Produkteigenschaft Einheit");      //Spalte F
        rowÜberschriften.createCell(6).setCellValue("Ressource Name");                  //Spalte G
        rowÜberschriften.createCell(7).setCellValue("Ressource ID");                    //Spalte H
        rowÜberschriften.createCell(8).setCellValue("Parameter Name");                  //Spalte I
        rowÜberschriften.createCell(9).setCellValue("Parameter ID");                    //Spalte J
        rowÜberschriften.createCell(10).setCellValue("Ressourcenparameter Wert");       //Spalte K
        rowÜberschriften.createCell(11).setCellValue("Ressourcenparameter Einheit");    //Spalte L
        rowÜberschriften.createCell(12).setCellValue("Vergleichsergebnis");             //Spalte M
        rowÜberschriften.createCell(13).setCellValue("Differenz");                      //Spalte N


        for(int num=0; num<ProduktListe.size(); num++) {
            Produkt P = (Produkt)ProduktListe.get(num);

            for(int num2=0; num2<P.GetProdukteigenschaftenliste().size(); num2++) {
                Produkteigenschaft PE = (Produkteigenschaft) P.GetProdukteigenschaftenliste().get(num2);

                Row rowEinzelneZeilen = sheet.createRow(num + num2 + num + 1);

                rowEinzelneZeilen.createCell(0).setCellValue(P.GetName());                                                                  //Spalte A
                rowEinzelneZeilen.createCell(1).setCellValue(P.GetID());                                                                    //Spalte B
                rowEinzelneZeilen.createCell(2).setCellValue(PE.GetEigenschaftName());                                                      //Spalte C
                rowEinzelneZeilen.createCell(3).setCellValue(PE.GetEigenschaftID());                                                        //Spalte D
                rowEinzelneZeilen.createCell(4).setCellValue(PE.GetDWert1());                                      //Spalte E
                rowEinzelneZeilen.createCell(5).setCellValue(PE.GetEinheit());                                                              //Spalte F
                rowEinzelneZeilen.createCell(6).setCellValue(PE.GetReferenzAufRessource().GetName());                                       //Spalte G
                rowEinzelneZeilen.createCell(7).setCellValue(PE.GetReferenzAufRessource().GetID());                                         //Spalte H
                rowEinzelneZeilen.createCell(8).setCellValue(PE.GetReferenzAufRessourcenParameter().GetReferenzAufParameter().GetName());   //Spalte I
                rowEinzelneZeilen.createCell(9).setCellValue(PE.GetReferenzAufRessourcenParameter().GetReferenzAufParameter().GetID());     //Spalte J
                rowEinzelneZeilen.createCell(10).setCellValue(PE.GetReferenzAufRessourcenParameter().GetDWert1());                          //Spalte K
                rowEinzelneZeilen.createCell(11).setCellValue(PE.GetReferenzAufRessourcenParameter().GetEinheit());                         //Spalte L
                rowEinzelneZeilen.createCell(12).setCellValue(PE.getIstErfuellt());                                                         //Spalte M
                rowEinzelneZeilen.createCell(13).setCellValue(PE.getDifferenz());                                                           //Spalte N
            }
        }


        try {

            FileOutputStream output = new FileOutputStream("Ergebnis Datei.xlsx");
            workbook.write(output);
            output.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

}
