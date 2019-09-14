import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.PrintWriter;
import java.util.LinkedList;

public class Produkt {

    /* Deklaration der Variablen zur Beschreibung eines Produktes bzw. Prozesses
     *
     */

    private String id;
    private String name;
    private LinkedList<Produkteigenschaft> ProduktEigenschaftenListe;
    private boolean produktAnforderungenErfuellt;

    public void setProduktAnforderungenErfuellt(boolean istErfuellt){
        this.produktAnforderungenErfuellt = istErfuellt;
    }

    public boolean getProduktAnforderungenErfuellt(){
        return this.produktAnforderungenErfuellt;
    }

    /* Konstruktor
     *
     */

    public Produkt(String ID, String Name){
        this.id = ID;
        this.name = Name;
        this.ProduktEigenschaftenListe = new LinkedList();

    }

    public String GetID() {
        return this.id;
    }

    public String GetName() { return this.name;}

    //Test
    public LinkedList<Produkteigenschaft> GetProdukteigenschaftenliste(){ return this.ProduktEigenschaftenListe;}

    private boolean ProduktEigenschaftVorhanden(String IDRessource, String IDEigenschaft){
        boolean retval = false;
        for(int num=0; num<ProduktEigenschaftenListe.size(); num++)
        {
            Produkteigenschaft PE = (Produkteigenschaft)ProduktEigenschaftenListe.get(num);
            String S1 = PE.GetEigenschaftID();
            String S2 = PE.GetRessourceID();
            if (S1.equals(IDEigenschaft) && S2.equals(IDRessource)) {
                retval = true;
                break;
            }
        }
        return retval;
    }

    public void NeueProduktEigenschaftString (Ressource R, Eigenschaft E, String SWert, String Einheit, PrintWriter PW) {
        if (ProduktEigenschaftVorhanden(R.GetID(), E.GetID())){
            PW.println("  - Produkt: " + id + " Ressource: " + R.GetID() + " Eigenschaft: " + E.GetID() + " bereits vorhanden");
        }
        else {
            // Produkteigenschaft erstellen und in Liste einhängen
            this.ProduktEigenschaftenListe.add(new Produkteigenschaft(this, R, E, SWert, Einheit));
            PW.println("  - Produkt: " + id + " Ressource: " + R.GetID() + " Eigenschaft: " + E.GetID() + " neu erstellt");
        }
    }

    public void NeueProduktEigenschaftBoolean (Ressource R, Eigenschaft E, Boolean BWert, String Einheit, PrintWriter PW) {
        if (ProduktEigenschaftVorhanden(R.GetID(), E.GetID())){
            PW.println("  - Produkt: " + id + " Ressource: " + R.GetID() + " Eigenschaft: " + E.GetID() + " bereits vorhanden");
        }
        else {
            // Produkteigenschaft erstellen und in Liste einhängen
            this.ProduktEigenschaftenListe.add(new Produkteigenschaft(this, R, E, BWert, Einheit));
            PW.println("  - Produkt: " + id + " Ressource: " + R.GetID() + " Eigenschaft: " + E.GetID() + " neu erstellt");
        }
    }

    public void NeueProduktEigenschaftDouble (Ressource R, Eigenschaft E, Double DWert, String Einheit, PrintWriter PW) {
        if (ProduktEigenschaftVorhanden(R.GetID(), E.GetID())){
            PW.println("  - Produkt: " + id + " Ressource: " + R.GetID() + " Eigenschaft: " + E.GetID() + " bereits vorhanden");
        }
        else {
            // Produkteigenschaft erstellen und in Liste einhängen
            this.ProduktEigenschaftenListe.add(new Produkteigenschaft(this, R, E, DWert, Einheit));
            PW.println("  - Produkt: " + id + " Ressource: " + R.GetID() + " Eigenschaft: " + E.GetID() + " neu erstellt");
        }
    }

    public boolean ProduktEigenschaftenPruefen(PrintWriter PW){
        boolean retval = true;
        for(int num=0; num<ProduktEigenschaftenListe.size(); num++) {
            Produkteigenschaft PE = (Produkteigenschaft)ProduktEigenschaftenListe.get(num);

            if (PE.ProduktEigenschaftOK()){
                PW.println("  - Eigenschaft:  " + PE.GetEigenschaftID() + "   OK");
            }
            else {
                PW.println("  - Eigenschaft:  " + PE.GetEigenschaftID() + "   NICHT OK");
                retval = false;
            }

            PE.ProdukteigenschaftDifferenz();
        }
        this.produktAnforderungenErfuellt = retval;
        return retval;
    }
}
