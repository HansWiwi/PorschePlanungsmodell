public class Parameter {

    /* Deklaration der Variablen zur Beschreibung eines Parameters
     *
     */

    private String id;
    private String name;
    private String datentyp;

    /* Konstruktor
     *
     */

    public Parameter (String ID, String Name, String Datentyp) {
        this.id = ID;
        this.name = Name;
        this.datentyp = Datentyp;
    }

    public String GetID() { return this.id; }

    public String GetDatentyp() { return this.datentyp; }

    public String GetName() { return this.name; }

}
