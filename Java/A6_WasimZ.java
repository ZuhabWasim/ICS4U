// The "A6_WasimZ" class.
public class A6_WasimZ
{
    public static void main (String[] args)
    {
	String words[] = {"APPLE", "MANGO", "WATERMELON", "HONEYDEW", "PINEAPPLE", "GRAPES", "STRAWBERRY", "PEACH", "POMEGRANETE", "BANANA"};
	
	final int HIGH = 9;
	final int LOW = 0;
	
	int numGuesses = 0;
	char guess;
	String targetWord = words [(int) (Math.random () * (HIGH - LOW + 1)) + LOW];
	String tempWord;
	
	char letters[] = new char [targetWord.length ()];
	boolean win = false;
	boolean playAgain = false;
	
	for (int i = 0 ; i < letters.length ; i++)
	{
	    letters [i] = '-';
	}

	System.out.println ("Word Guessing Game:");
	System.out.println (getStrChoices (letters));

	do
	{
	    System.out.print ("Enter a letter ($ for entire word): ");
	    guess = ReadLib.readChar ();
	    numGuesses++;
	    if (guess == '$')
	    {
		System.out.print ("What is your guess");
		if ((ReadLib.readString ().toUpperCase ()).equals (targetWord))
		{
		    win = true;
		}
		else
		{
		    break;
		}
	    }
	    else
	    {
		if (letterMatch (letters, targetWord, guess))
		{
		    if ((getStrChoices (letters)).equals (targetWord))
		    {
			win = true;
		    }
		    else
		    {
			System.out.println (getStrChoices (letters));
		    }

		}
	    }
	}
	while (!win);


	if (win)
	{
	    System.out.println ("You won! Secret word was " + targetWord);
	    System.out.println ("Total number of guesses: " + numGuesses);
	}
	else
	{
	    System.out.println ("You lose! Secret word was " + targetWord);
	    System.out.println ("Total number of guesses: " + numGuesses);
	}
	
	System.out.println ("Do you want to play again? [Yes or No]");
    } // main method


    public static String getStrChoices (char letters[])
    {

	String s = "";

	for (int i = 0 ; i < letters.length ; i++)
	{
	    s = s + letters [i];
	}

	return s;

    }


    public static boolean letterMatch (char letters[], String target, char guess)
    {

	boolean match = false;

	for (int i = 0 ; i < letters.length ; i++)
	{
	    if (target.charAt (i) == Character.toUpperCase (guess))
	    {
		letters [i] = Character.toUpperCase (guess);
		match = true;
	    }
	}

	return match;
    }
} // A6_WasimZ class
