package com.qa.para;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class LoginPage {

	@FindBy(name = "username")
	private WebElement username;
	
	@FindBy(name = "password")
	private WebElement password;
	
	@FindBy(name = "FormsButton2")
	private WebElement button;
	
	@FindBy(xpath = "/html/body/table/tbody/tr/td[1]/big/blockquote/blockquote/font/center")
	private WebElement loginAttempt;
	
	public void loginUser(String user, String pass) {
		username.sendKeys(user);
		password.sendKeys(pass);
		button.click();
	}
	
	public String loginAttemptText() {
		return loginAttempt.getText();
	}
	
	
}
