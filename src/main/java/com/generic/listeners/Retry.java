package com.generic.listeners;

import java.util.Properties;

import org.apache.log4j.Logger;
import org.testng.IRetryAnalyzer;
import org.testng.ITestResult;

import com.generic.property.PropertyManager;
import com.generic.utilities.Logg;
import com.generic.utilities.Utilities;

public class Retry implements IRetryAnalyzer {
	
	private static final Properties APPLICATIONPROPERTY = PropertyManager.loadApplicationPropertyFile();
    private int retryCount = 0;
    private int maxRetryCount = Integer.parseInt(APPLICATIONPROPERTY.getProperty("retry"));//change this as required

	protected static final Logger LOGG = Logg.createLogger();

// Below method returns 'true' if the test method has to be retried else 'false' 
//and it takes the 'Result' as parameter of the test method that just ran
	@Override
	public boolean retry(ITestResult result) {
    	System.out.println("Enter Retry");
        if (retryCount < maxRetryCount) {

        	
        	LOGG.info(Utilities.getCurrentThreadId() + "Retrying test " + result.getName() + " with status "
                    + getResultStatusName(result.getStatus()) + " for the " + (retryCount+1) + " time(s).");
            retryCount++;
            return true;
        }
        return false;
    }
    
    public String getResultStatusName(int status) {
    	String resultName = null;
    	if(status==1)
    		resultName = "SUCCESS";
    	if(status==2)
    		resultName = "FAILURE";
    	if(status==3)
    		resultName = "SKIP";
		return resultName;
    }
}