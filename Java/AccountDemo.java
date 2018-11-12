// The "AccountDemo" class.
import java.text.DecimalFormat;

public class AccountDemo
{
    public static void main (String[] args)
    {

	System.out.println ("Welcome to your account!");
	System.out.print ("Please enter an account number: ");

	String accountNum = new String (ReadLib.readString ());
	DecimalFormat currency = new DecimalFormat ("$#,###,##0.00");

	Account account = new Account (accountNum, 0);

	String nextInput = "";
	double amount;

	while (!nextInput.equals ("*"))
	{
	    System.out.println ("-----------------------------");
	    System.out.println ("What would you like to do?");
	    System.out.println (" ! : Show Account Information");
	    System.out.println (" # : Show Account Number");
	    System.out.println (" $ : Show Account Balance");
	    System.out.println (" + : Deposit");
	    System.out.println (" - : Withdraw");
	    System.out.println (" * : Exit");
	    System.out.println ("-----------------------------");
	    System.out.print ("Enter: ");

	    nextInput = ReadLib.readString ();

	    if (nextInput.equals ("!"))
	    {
		System.out.println ("Account Number : " + account.getAccNum ());
		System.out.println ("Account Balance: " + account.getBalance ());
	    }
	    else if (nextInput.equals ("#"))
	    {
		System.out.println ("Account Number : " + account.getAccNum ());
	    }
	    else if (nextInput.equals ("$"))
	    {
		System.out.println ("Account Balance: " + account.getBalance ());
	    }
	    else if (nextInput.equals ("+"))
	    {
		System.out.print ("How much would you like to deposit? ");
		amount = ReadLib.readDouble ();
		account.deposit (amount);
		System.out.println ("You have deposited " + currency.format (amount) + ", your balance is now " + account.getBalance ());
	    }
	    else if (nextInput.equals ("-"))
	    {
		System.out.print ("How much would you like to withdraw? ");
		amount = ReadLib.readDouble ();
		account.withdraw (amount);
		System.out.println ("You have withdrawed " + currency.format (amount) + ", your balance is now " + account.getBalance ());
	    }
	    else if (nextInput.equals ("*"))
	    {
		System.out.println ("Thank you! Please come again!");
	    }
	    else
	    {
		System.out.println ("Invalid command: " + nextInput);
	    }

	}
    } // main method


    public boolean validAccNum (String acc)
    {

	if (acc.length () == 6)
	{
	    //for(int i = 0; i < acc.length() - 1; i++) {
	    //if(!(acc.substring(i, i + 1) >= "0" && acc.substring(i, i + 1) <= "9")) {
	    //    return false;
	    //    break;
	    //}
	    //}
	    return true;
	}
	else
	{
	    return false;
	}
    }
} // AccountDemo class


