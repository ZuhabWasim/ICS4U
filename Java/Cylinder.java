// The "Cylinder" class.
public class Cylinder
{
    private double radius;
    private double length;
    private double surfaceArea;
    private double volume;

    //constructor
    public Cylinder (double r, double l)
    {

	radius = r;
	length = l;

    }


    //getters
    public double getRadius ()
    {

	return radius;

    }


    public double getLength ()
    {

	return length;

    }


    public double getSurfaceArea ()
    {

	surfaceArea = (2 * Math.PI * Math.pow (radius, 2)) + (2 * Math.PI * radius * length);
	return surfaceArea;

    }


    public double getVolume ()
    {

	volume = Math.PI * Math.pow (radius, 2) * length;
	return volume;

    }
} // Cylinder class
