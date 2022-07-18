package first.writer;

public class ProductInfo {

    private final String name;
    private final String articul;
    private final int price;
    private final String intArticul;

    public ProductInfo(String articul, String name, int price,  String intArticul) {
        this.name = name;
        this.price = price;
        this.articul = articul;
        this.intArticul = intArticul;
    }

    public String getName() {
        return name;
    }
    public int getPrice() {
        return price;
    }
    public String getArticul() {
        return articul;
    }
    public String getintArticul() {
        return intArticul;
    }

    @Override
    public String toString() {
        return "Item{" +
                "articul='" + articul + '\'' +
                ", name='" + name + '\''  +
                ", price=" + price +
                ", intArticul='" + intArticul + '\''  +
                '}';
    }
}