// The "Easy10" class.
public class Easy10
{
    public static void main (String[] args)
    {
	int number = 0;
	int sum = 0;
	
	while(number != -1)
	{
	    System.out.print("Enter #: ");
	    number = ReadLib.readInt();
	    sum += number;
	}
	
	System.out.println("Total = " + sum);
    } // main method
} // Easy10 class
