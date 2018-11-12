// The "TestQuestion" class.
public class TestQuestion
{
    public static void main (String[] args)
    {
	String phrase;
	char target;
	int countCh = 0;
	
	System.out.println("Enter a string!");
	phrase = ReadLib.readString();
	
	System.out.println("Enter a character!");
	target = ReadLib.readChar();
	
	for(int i = 0; i < phrase.length(); i++) {
	    if (((phrase.charAt(i) + "").toUpperCase()).equals((target + "").toUpperCase())) {
		countCh++;
	    }
	}
	
	System.out.println("Number of occurences: " + countCh);
    } // main method
} // TestQuestion class
