// The "PP1Car" class.
public class PP1Car
{
    public static void main (String[] args)
    {
	final double HST = 0.13;
	
	double finalPrice;
	double basePrice;
	
	System.out.print("Enter your final offer: ");
	finalPrice = ReadLib.readDouble();
	basePrice = finalPrice / (HST + 1);
	
	System.out.println("            Base Price: " + Math.round(basePrice));
	System.out.println("             Tax (HST): " + HST);
	System.out.println("           Final Price: " + finalPrice);
	
    } // main method
} // PP1Car class
