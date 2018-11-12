// The "PP1Temp" class.
public class PP1Temp
{
    public static void main (String[] args)
    {
	double cels;
	double fahr;

	System.out.print ("Enter celsius: ");

	cels = ReadLib.readDouble ();
	fahr = 9.0 / 5.0 * cels + 32.0;

	System.out.println (cels + " degrees celsius is " + fahr + " degrees fahrenheit.");

    } // main method
} // PP1Temp class
