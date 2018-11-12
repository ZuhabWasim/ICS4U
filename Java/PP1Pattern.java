// The "PP1Pattern" class.
public class PP1Pattern
{
    public static void main (String[] args)
    {
	int n;
	
	System.out.print("Enter a #: ");
	n = ReadLib.readInt();
	
	for(int i = n; i > 0; i--) {
	    for(int k = i; k > 0; k--) {
		System.out.print((k) + " ");
	    }
	    System.out.println("");
	}
    } // main method
} // PP1Pattern class
