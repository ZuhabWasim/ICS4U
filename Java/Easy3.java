// The "Easy3" class.
public class Easy3
{
    public static void main (String[] args)
    {
	final double TAXRATE = 0.15;
	
	int items = 4;
	double price = 2.50;
	double subtotal;
	double tax;
	double total;
	
	subtotal = items * price;
	tax = subtotal * TAXRATE;
	total = subtotal + tax;
	
	System.out.println("Subtotal: " + subtotal);
	System.out.println("     Tax: " + tax);
	System.out.println("   Total: " + total); 
    } // main method
} // Easy3 class
