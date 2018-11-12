// The "Easy7" class.
public class Easy7
{
    public static void main (String[] args)
    {
	int age;
	String name;
	
	System.out.print("Enter your age: ");
	age = ReadLib.readInt();
	
	System.out.print("Enter your name: ");
	name = ReadLib.readString();
	
	if(age > 18)
	{
	    System.out.println(name + ", you are");
	    System.out.println("too old.");
	}
	
	System.out.println("You are " + age + ".");
    } // main method
} // Easy7 class
