// The "Classes_Exercises" class.
import java.text.DecimalFormat;

public class Classes_Exercises
{
    public static void main (String[] args)
    {
	//(a)
	DecimalFormat fiveDec = new DecimalFormat ("0.00000");
	double number = getRandNum ();
	System.out.println ("The number is: " + fiveDec.format (number));

	//(b)
	DecimalFormat twoDec = new DecimalFormat ("0.00");
	double radius;
	System.out.print ("Radius: ");
	radius = ReadLib.readDouble ();
	System.out.println ("Area: " + twoDec.format (getArea (radius)));

	//(c)
	DecimalFormat fourDec = new DecimalFormat ("0.0000");
	System.out.println ("Num:" + "\t" + "sqrx" + "\t" + "x^3");
	for (int i = 5 ; i <= 25 ; i++)
	{
	    System.out.println (i + "\t" + fourDec.format (Math.sqrt (i)) + "\t" + fourDec.format (Math.pow (i, 3)));
	}
    } // main method


    //(a)
    public static double getRandNum ()
    {

	final int HIGH = 1;
	final int LOW = 0;

	double num = Math.random ();

	return num;

    }


    //(b)
    public static double getArea (double radius)
    {

	double area = Math.PI * Math.pow (radius, 2);

	return area;
    }
} // Classes_Exercises class
