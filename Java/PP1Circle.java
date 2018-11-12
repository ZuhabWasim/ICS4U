// The "PP1Circle" class.
public class PP1Circle
{
    public static void main (String[] args)
    {
	
	double radius;
	double area;
	double circumference;
	
	System.out.print("Enter a radius: ");
	radius = ReadLib.readDouble();
	
	area = Math.PI * Math.pow(radius, 2);
	circumference = Math.PI * radius;
	
	System.out.println("        Radius: " + radius);
	System.out.println("          Area: " + area);
	System.out.println(" Circumference: " + circumference);
    } // main method
} // PP1Circle class
