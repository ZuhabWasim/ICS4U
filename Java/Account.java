// The "Account" class.
import java.text.DecimalFormat;

public class Account
{
    private DecimalFormat currency = new DecimalFormat("$#,###,##0.00");
    private String accountNum;
    private double balance;
    
    public Account(String accNum, double bal) {
	
	accountNum = new String(accNum);
	balance = bal;
    } 
    
    public void deposit(double val) {
	
	balance += val;
    }
    
    public void withdraw(double val) {
	
	balance -= val;
    }
    
    public String getAccNum() {
	
	return accountNum;
    }
    
    public String getBalance() {
	
	return currency.format(balance);
    }

} // Account class
