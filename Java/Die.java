// The "Die" class.
public class Die
{
    private int value;

    public Die ()
    {
	roll ();
    }


    public void roll ()
    {
	value = (int) (Math.random () * 6 + 1);
    }


    public int getValue ()
    {
	return value;
    }
} // Die class
