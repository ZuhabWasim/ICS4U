// The "TestQuestionAlex" class.
public class TestQuestionAlex
{
    public static void main (String[] args)
    {
	String phrase;
	char target;
	int countCh = 0;
	String chTest;


	System.out.println ("Enter a string!");
	phrase = ReadLib.readString ();

	System.out.println ("Enter a character!");
	target = ReadLib.readChar ();

	chTest = target + "";

	for (int i = 0 ; i < phrase.length () ; i++)
	{
	    if ((((phrase.toUpperCase ()).charAt (i)).toString () == chTest.toUpperCase ()))
	    {
		countCh++;
	    }
	}


	// Place your code here
    } // main method
} // TestQuestionAlex class
