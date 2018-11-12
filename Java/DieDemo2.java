// The "DieDemo2" class.
public class DieDemo2
{
    public static void main (String[] args)
    {
	Die whiteDie = new Die();
	Die greenDie = new Die();
	
	int whiteDieRolls[] = new int[10];
	int greenDieRolls[] = new int[10];
	boolean isMatch[] = new boolean[10];

	System.out.println ("Die :  Green   White   Match?");

	for (int i = 0 ; i < 10 ; i++)
	{
	    whiteDieRolls[i] = whiteDie.roll ();
	    greenDieRolls[i] = greenDie.roll ();
	    isMatch[i] = (whiteDieRolls[i] == greenDieRolls[i]) ? true:
	    false;
	    System.out.println ("\t" + greenDieRolls[i] + "\t" + greenDieRolls[i] + "\t" + ((isMatch[i]) ? "Yes" : "No"));
	}
    } // main method
} // DieDemo2 class
