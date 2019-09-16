public class Ressourcenparameter {

    /* Deklaration der Variablen zur Beschreibung eines Ressourcenparameter
     *
     */

    private Ressource referenzAufRessource;               // Referenz auf Klasse Ressource
    private Parameter referenzAufParameter;               // Referenz auf Klasse Paramneter
    private String operator1;
    private String operator2;
    private double dwert1;                                // dwert1 = Wert 1, wenn Datentyp Double
    private double dwert2;                                // dwert2 = Wert 2, wenn Datentyp Double
    private String swert;                                // swert1 = Wert 1, wenn Datentyp String
    private boolean bwert;
    private String einheit;


    // Konstruktor für den Fall: String-Vergleich

    public Ressourcenparameter(Ressource ReferenzAufRessource, Parameter ReferenzAufParameter, String Operator1,
                               String SWert1, String Einheit){

        this.referenzAufRessource = ReferenzAufRessource;
        this.referenzAufParameter = ReferenzAufParameter;
        this.operator1 = Operator1;
        this.operator2 = "";
        this.dwert1 = 0;
        this.dwert2 = 0;
        this.bwert = false;
        this.swert = SWert1;
        this.einheit = Einheit;
    }

    // Konstruktor für den Fall: Double-Vergleich mit einem oder mehreren Werten (Einfacher Vergleich oder Intervall-Vergleich)

    public Ressourcenparameter(Ressource ReferenzAufRessource, Parameter ReferenzAufParameter, String Operator1,
                               double DWert1, String Operator2, double DWert2, String Einheit){

        this.referenzAufRessource = ReferenzAufRessource;
        this.referenzAufParameter = ReferenzAufParameter;
        this.operator1 = Operator1;
        this.operator2 = "";
        this.dwert1 = DWert1;
        this.operator2 = Operator2;
        this.dwert2 = DWert2;
        this.swert = "";
        this.bwert = false;
        this.einheit = Einheit;
    }

    // Konstruktor für den Fall: Boolean-Vergleich

    public Ressourcenparameter(Ressource ReferenzAufRessource, Parameter ReferenzAufParameter, String Operator1,
                               boolean BWert, String Einheit){

        this.referenzAufRessource = ReferenzAufRessource;
        this.referenzAufParameter = ReferenzAufParameter;
        this.operator1 = Operator1;
        this.operator2 = "";
        this.dwert1 = 0;
        this.dwert2 = 0;
        this.swert = "";
        this.bwert = BWert;
        this.einheit = Einheit;
    }

    // Methode überprüft, ob der übergebene Wert der Produkteigenschaft durch den Ressourcenparameter abgedeckt wird

    public boolean IstWertZulaessig(double DWert, boolean BWert, String SWert) {
        boolean retval = false;

        // Double-Vergleich mit Intervall

        if (dwert2 != 0){
            if (operator1.equals("<") && operator2.equals("<")){
                if (dwert1 < DWert && DWert < dwert2){
                    retval = true;
                }

            }
            else if (operator1.equals("<") && operator2.equals("<=")){
                if (dwert1 < DWert && DWert <= dwert2){
                    retval = true;
                }

            }
            else if (operator1.equals("<=") && operator2.equals("<")){
                if (dwert1 <= DWert && DWert < dwert2){
                    retval = true;
                }

            }
            else if (operator1.equals("<=") && operator2.equals("<=")){
                if (dwert1 <= DWert && DWert <= dwert2){
                    retval = true;
                }

            }
            else {
                System.out.println("Operatoren falsche eingegeben!");
            }
        }


        // String-Vergleich

        else if (this.referenzAufParameter.GetDatentyp().equals("String")) {
            if (swert.contains(SWert)){
                retval = true;
            }
        }

        // Einfacher Double-Vergleich

        else if (this.referenzAufParameter.GetDatentyp().equals("Double")) {
            switch (operator1){
                case "<":
                    if (dwert1 < DWert){
                        retval = true;
                        }
                    break;

                case ">":
                    if (dwert1 > DWert){
                        retval = true;
                    }
                    break;

                case "<=":
                    if (dwert1 <= DWert){
                        retval = true;
                    }
                    break;

                case ">=":
                    if (dwert1 >= DWert){
                        retval = true;
                    }
                    break;

                case "=":
                    if (dwert1 == DWert){
                        retval = true;
                    }
                    break;

                default:
                    System.out.println("Operator nicht definiert!");
            }
        }

        // Boolean-Vergleich

        else if (this.referenzAufParameter.GetDatentyp().equals("Boolean")) {

            if (bwert == true){
                retval = true;
            }else if (BWert == false){
                retval = true;
            }
        }

        return retval;
    }

    // Methode gibt die Differenz zurück (Test)

    public double Differenz(double DWert){
        // Fallunterscheidungen je nach Operator errgänzen!
        double differenz = DWert - this.dwert1;
        return differenz;
    }

    public String GetParameterID() {
        return this.referenzAufParameter.GetID();
    }

    public String GetRessourcenID() { return this.referenzAufRessource.GetID();}

    public Parameter GetReferenzAufParameter() {
        return this.referenzAufParameter;
    }

    public double GetDWert1() { return this.dwert1; }

    public String GetEinheit() { return this.einheit; }
}
