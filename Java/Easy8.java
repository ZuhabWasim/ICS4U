// The "Easy8" class.
public class Easy8
{
    public static void main (String[] args)
    {
	int x, y;
	
	System.out.print("Value for x? ");
	x = ReadLib.readInt();
	System.out.print("Value for y? ");
	y = ReadLib.readInt();
	
	if(x < 5 || y > 2)
	{
	    System.out.println("Yes");
	}
	else
	{
	    System.out.println("No");
	}
    
    } // main method
} // Easy8 class
