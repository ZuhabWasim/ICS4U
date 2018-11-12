// The "Easy6" class.
public class Easy6
{
    public static void main (String[] args)
    {
	int n;
	
	n = (int) (Math.random() * 100) + 1;
	
	System.out.println(n);
	if (n <= 50)
	    System.out.println("LOW");
	else
	    System.out.println("HIGH");
    } // main method
} // Easy6 class
