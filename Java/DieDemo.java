// The "DieDemo" class.
public class DieDemo
{
    public static void main (String[] args)
    {
	Die d1 = new Die ();
	Die d2 = new Die ();

	d1.roll ();
	d2.roll ();

	System.out.println (d1.getValue () + "\t" + d2.getValue ());

    } // main method
} // DieDemo class
