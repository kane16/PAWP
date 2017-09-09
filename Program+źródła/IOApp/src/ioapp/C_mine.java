package ioapp;

import java.util.ArrayList;

public class C_mine implements Comparable<C_mine>{
	
	String lp;
	String mine;
	String error;
	ArrayList<Double> list=new ArrayList<>();
	
	public C_mine(String lp, String mine, ArrayList<Double> list, String error){
		this.list=list;
		this.lp=lp;
		this.mine=mine;
		this.error=error;
	}

	@Override
	public int compareTo(C_mine other) {
		return other.list.get(list.size()-1).compareTo(list.get(list.size()-1));
	}
}
