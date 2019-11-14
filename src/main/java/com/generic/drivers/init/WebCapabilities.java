package com.generic.drivers.init;

import org.apache.log4j.Logger;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.remote.DesiredCapabilities;

import com.generic.enums.BrowserType;
import com.generic.utilities.Logg;
import com.generic.utilities.Utilities;

public class WebCapabilities {

    private static DesiredCapabilities capabilities;
	protected static final Logger LOGG = Logg.createLogger();

    
    private static String language;

    private WebCapabilities() {
    }
        
    public static DesiredCapabilities getDesiredCapability(Browser browser){
    	
    	setBrowserLanguage(browser.getLanguage());
    	capabilities = getCapability(browser.getName());    	
    	capabilities.setBrowserName(browser.getName());
        capabilities.setPlatform(browser.getPlatform());
        capabilities.setVersion(browser.getVersion());     
        capabilities.setJavascriptEnabled(true);
        return capabilities;
    }

    private static DesiredCapabilities getCapability(String browserName) {

    	if(BrowserType.CHROME.getBrowserValue().contentEquals(browserName)){

    		ChromeOptions options = new ChromeOptions();
    		options.addArguments("start-maximized", "forced-maximize-mode",
    				"no-default-browser-check", "always-authorize-plugins","test-type");
    		
    		options.addArguments("--lang=" + WebCapabilities.language);  
    		capabilities = DesiredCapabilities.chrome();

    		capabilities.setCapability(ChromeOptions.CAPABILITY, options);

    	} else if(BrowserType.FIREFOX.getBrowserValue().contentEquals(browserName)){   		

    		capabilities = DesiredCapabilities.firefox();   		
    		
    		FirefoxProfile firefoxProfile = new FirefoxProfile();
        	firefoxProfile.setPreference("intl.accept_languages", WebCapabilities.language);       	
        	
        	capabilities.setCapability(FirefoxDriver.PROFILE, firefoxProfile);
        	
    	}	   else if(BrowserType.INTERNETEXPLORER.getBrowserValue().contentEquals(browserName)){
    		capabilities = DesiredCapabilities.internetExplorer();

    	}else {
    		capabilities = DesiredCapabilities.chrome();
    	}
    	return capabilities;
    }

	private static String setBrowserLanguage(String lang) {
		
		LOGG.debug(Utilities.getCurrentThreadId() + "Browser locale : " + lang);

		language= lang;	
		return language;
	}
}

