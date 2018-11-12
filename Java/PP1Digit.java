// The "PP1Digit" class.
public class PP1Digit
{
    public static void main (String[] args)
    {
	String st;
	
	System.out.print("Enter a 3-digit number: ");
	
	st = ReadLib.readString();
	
	System.out.println(st.substring(0, 1) + " " + st.substring(1,2) + " " + st.substring(2,3));
    } // main method
} // PP1Digit class
