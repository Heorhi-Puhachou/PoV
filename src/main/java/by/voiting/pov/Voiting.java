package by.voiting.pov;

public class Voiting {
    private String date;
    private String name;
    private String result;

    public Voiting(String date, String name, String result) {
        this.date = date;
        this.name = name;
        this.result = result;
    }

    public String getDate() {
        return date;
    }

    public String getName() {
        return name;
    }

    public String getResult() {
        return result;
    }
}
