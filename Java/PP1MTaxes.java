// The "PP1MTaxes" class.
public class PP1MTaxes
{
    public static void main (String[] args)
    {
	final double HST = 0.13;
	
	double subtotal;
	double finalPrice;
	
	System.out.print("Enter a price: ");
	
	subtotal = ReadLib.readDouble();
	finalPrice = subtotal + subtotal * HST;
	
	System.out.println("Final Price: " + finalPrice);
	
    } // main method
} // PP1MTaxes class
