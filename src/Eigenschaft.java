public class Eigenschaft {

    /* Deklaration der Variablen zur Beschreibung eines Parameters
     *
     */

    private String id;
    private String name;
    private String datentyp;
    private String id_parameter;

    /* Konstruktor
     *
     */

    public Eigenschaft(String ID, String Name, String Datentyp, String ID_Parameter){

        this.id = ID;
        this.name = Name;
        this.datentyp = Datentyp;
        this.id_parameter = ID_Parameter;
    }

    public String GetID() {
        return this.id;
    }

    public String GetDatentyp() {
        return this.datentyp;
    }

    public String GetIDParameter() {
        return this.id_parameter;
    }

    public String GetName() { return this.name;}

}
