import com.sun.xml.bind.v2.model.core.ID;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.util.Iterator;
import javax.swing.*;
import java.io.*;
import java.util.LinkedList;

public class Ressource {

    // Deklaration der Variablen zur Beschreibung einer Ressource

    private String id;
    private String name;
    private LinkedList<Ressourcenparameter> RessourcenParameterListe;

    // * Konstruktor für die Ressource

    public Ressource(String ID, String Name) {
        this.id = ID;
        this.name = Name;
        this.RessourcenParameterListe = new LinkedList();
    }

    public String GetID() {
        return this.id;
    }

    public String GetName() { return this.name;}

    private boolean RessourcenParameterVorhanden(String IDParameter){
        boolean retval = false;
        for(int num=0; num<RessourcenParameterListe.size(); num++)
        {
            Ressourcenparameter RP = (Ressourcenparameter)RessourcenParameterListe.get(num);
            String S1 = RP.GetParameterID();
            if (S1.equals(IDParameter)) {
                retval = true;
                break;
            }
        }
        return retval;
    }

    public Ressourcenparameter GetRessourcenParameter(String IDParameter){
        boolean retval = false;
        for(int num=0; num<RessourcenParameterListe.size(); num++)
        {
            Ressourcenparameter RP = (Ressourcenparameter)RessourcenParameterListe.get(num);
            String S1 = RP.GetParameterID();
            if (S1.equals(IDParameter)) {
                return RP;
            }
        }
        return null;
    }

    public void NeuerRessourcenParameterString (Parameter P, String Operator1, String SWert, String Einheit, PrintWriter PW) {
        if (RessourcenParameterVorhanden(P.GetID())){
            PW.println("  - Ressource: " + id + " Ressourcenparameter: " + P.GetID() + " bereits vorhanden");
        }
        else {
            // Ressourcenparameter erstellen und in Liste einhängen
            this.RessourcenParameterListe.add(new Ressourcenparameter(this, P, Operator1, SWert, Einheit));
            PW.println("  - Ressource: " + id + " Ressourcenparameter: " + P.GetID() + " neu erstellt");
        }
    }

    public void NeuerRessourcenParameterBoolean (Parameter P, String Operator1, boolean BWert, String Einheit, PrintWriter PW) {

        if (RessourcenParameterVorhanden(P.GetID())){
            PW.println("  - Ressource: " + id + " Ressourcenparameter: " + P.GetID() + " bereits vorhanden");
        }
        else {
            // Ressourcenparameter erstellen und in Liste einhängen
            this.RessourcenParameterListe.add(new Ressourcenparameter(this, P, Operator1, BWert, Einheit));
            PW.println("  - Ressource: " + id + " Ressourcenparameter: " + P.GetID() + " neu erstellt");
        }
    }


    public void NeuerRessourcenParameterDouble (Parameter P, String Operator1, double DWert1, String Operator2, double DWert2, String Einheit, PrintWriter PW) {

        if (RessourcenParameterVorhanden(P.GetID())){
            PW.println("  - Ressource: " + id + " Ressourcenparameter: " + P.GetID() + " bereits vorhanden");
        }
        else {
            // Ressourcenparameter erstellen und in Liste einhängen
            this.RessourcenParameterListe.add(new Ressourcenparameter(this, P, Operator1, DWert1, Operator2, DWert2, Einheit));
            PW.println("  - Ressource: " + id + " Ressourcenparameter: " + P.GetID() + " neu erstellt");
        }
    }


}
