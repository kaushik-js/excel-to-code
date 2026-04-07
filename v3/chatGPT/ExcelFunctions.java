
public class ExcelFunctions {

    public static double IF(boolean condition, double trueVal, double falseVal) {
        return condition ? trueVal : falseVal;
    }

    public static double ROUND(double value, int places) {
        double scale = Math.pow(10, places);
        return Math.round(value * scale) / scale;
    }

    public static double SUM(double... values) {
        double total = 0;
        for (double v : values) total += v;
        return total;
    }

    public static boolean AND(boolean... conditions) {
        for (boolean c : conditions) if (!c) return false;
        return true;
    }

    public static boolean OR(boolean... conditions) {
        for (boolean c : conditions) if (c) return true;
        return false;
    }
}
