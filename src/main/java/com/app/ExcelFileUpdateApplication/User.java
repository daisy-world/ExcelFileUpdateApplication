package com.app.ExcelFileUpdateApplication;
public class User {

	
	private String userId ;
	private String userName ;
	public String getUserId() {
		return userId;
	}
	public void setUserId(String userId) {
		this.userId = userId;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public User() {
		super();
	}
	public User(String userId, String userName) {
		super();
		this.userId = userId;
		this.userName = userName;
	}
	@Override
	public String toString() {
		return "User [userId=" + userId + ", userName=" + userName + "]";
	}
	
	
	
}
