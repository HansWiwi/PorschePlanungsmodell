public class Produkteigenschaft {

    /* Deklaration der Variablen zur Beschreibung einer ProduktProzessEigenschaft
     *
     */
    private Produkt referenzAufProdukt;                             // Referenz auf Klasse ProduktProzess
    private Ressource referenzAufRessource;                         // Referenz auf Klasse Ressource
    private Eigenschaft referenzAufEigenschaft;                     // Referenz auf Klasse Eigenschaft
    private Ressourcenparameter referenzAufRessourcenParameter;     // Referenz auf Klasse Ressourcenparameter (für Parametervergleich)
    private double dwert;                                           // dwert = Wert, wenn Datentyp Double
    private String swert;                                           // swert = Wert, wenn Datentyp String
    private Boolean bwert;
    private String einheit;
    private boolean istErfuellt;
    private double differenz;


    // Konstruktor, wenn Wert vom Datentyp Double

    public Produkteigenschaft(Produkt ReferenzAufProdukt, Ressource ReferenzAufRessource,
                              Eigenschaft ReferenzAufEigenschaft, double DWert, String Einheit){

        this.referenzAufProdukt = ReferenzAufProdukt;
        this.referenzAufRessource = ReferenzAufRessource;
        this.referenzAufEigenschaft = ReferenzAufEigenschaft;
        this.dwert = DWert;
        this.bwert = false;
        this.swert = "";
        this.einheit = Einheit;

        if (this.referenzAufRessource != null && this.referenzAufEigenschaft != null) {
            this.referenzAufRessourcenParameter = this.referenzAufRessource.GetRessourcenParameter(this.referenzAufEigenschaft.GetIDParameter());

        }
        else {
            this.referenzAufRessourcenParameter = null;
        }
    }

    // Konstruktor, wenn Wert vom Datentyp String

    public Produkteigenschaft(Produkt ReferenzAufProdukt, Ressource ReferenzAufRessource,
                              Eigenschaft ReferenzAufEigenschaft, String SWert, String Einheit){

        this.referenzAufProdukt = ReferenzAufProdukt;
        this.referenzAufRessource = ReferenzAufRessource;
        this.referenzAufEigenschaft = ReferenzAufEigenschaft;
        this.swert = SWert;
        this.dwert = 0;
        this.bwert = false;
        this.einheit = Einheit;

        if (this.referenzAufRessource != null && this.referenzAufEigenschaft != null) {
            this.referenzAufRessourcenParameter = this.referenzAufRessource.GetRessourcenParameter(this.referenzAufEigenschaft.GetIDParameter());

        }
        else {
            this.referenzAufRessourcenParameter = null;
        }
    }

    // Konstruktor, wenn Wert vom Datentyp Boolean

    public Produkteigenschaft(Produkt ReferenzAufProdukt, Ressource ReferenzAufRessource,
                              Eigenschaft ReferenzAufEigenschaft, Boolean  BWert, String Einheit){

        this.referenzAufProdukt = ReferenzAufProdukt;
        this.referenzAufRessource = ReferenzAufRessource;
        this.referenzAufEigenschaft = ReferenzAufEigenschaft;
        this.bwert = BWert;
        this.swert = "";
        this.dwert = 0;
        this.einheit = Einheit;

        if (this.referenzAufRessource != null && this.referenzAufEigenschaft != null) {
            this.referenzAufRessourcenParameter = this.referenzAufRessource.GetRessourcenParameter(this.referenzAufEigenschaft.GetIDParameter());

        }
        else {
            this.referenzAufRessourcenParameter = null;
        }
    }

    public String GetEigenschaftName() {
        return this.referenzAufEigenschaft.GetName();
    }

    public String GetEigenschaftID() {
        return this.referenzAufEigenschaft.GetID();
        }

    public String GetRessourceID() {
        return this.referenzAufRessource.GetID();
    }

    public boolean ProduktEigenschaftOK() {
        this.setIstErfuellt(this.referenzAufRessourcenParameter.IstWertZulaessig(this.dwert, this.bwert, this.swert));
        return this.referenzAufRessourcenParameter.IstWertZulaessig(this.dwert, this.bwert, this.swert);
    }

    public void ProdukteigenschaftDifferenz() {
        this.differenz = this.referenzAufRessourcenParameter.Differenz(this.dwert);
    }

    public void setIstErfuellt(boolean istErfuellt) {
        this.istErfuellt = istErfuellt;
    }

    public boolean getIstErfuellt(){
        return this.istErfuellt;
    }

    public void SetDifferenz(double Differenz) {
        this.differenz = Differenz;
    }

    public double getDifferenz() {
        return differenz;
    }

    public Ressourcenparameter GetReferenzAufRessourcenParameter() {
        return this.referenzAufRessourcenParameter;
    }
}
