// The "Cylinder_Exercise" class.
public class Cylinder_Exercise
{
    public static void main (String[] args)
    {

	double radius;
	double length;

	System.out.print ("Enter radius: ");
	radius = ReadLib.readDouble ();

	System.out.print ("Enter length: ");
	length = ReadLib.readDouble ();

	Cylinder cyln = new Cylinder (radius, length);

	System.out.println ("              Cylinder Information:");
	System.out.println ("      Radius: " + cyln.getRadius());
	System.out.println ("      Length: " + cyln.getLength());
	System.out.println ("Surface Area: " + cyln.getSurfaceArea());
	System.out.println ("      Volume: " + cyln.getVolume());

    } // main method
} // Cylinder_Exercise class
