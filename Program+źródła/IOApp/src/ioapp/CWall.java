package ioapp;
import java.util.*;
public class CWall implements Comparable<CWall>{
	
	String lp;
	String wall;
	String mine;
	String error;
	ArrayList<Double> list=new ArrayList<>();
	
	ArrayList<CWall> elem;
	
	public CWall(String lp, String wall, String mine, ArrayList<Double> list, String error){
		this.list=list;
		this.lp=lp;
		this.wall=wall;
		this.mine=mine;
		this.error=error;
	}

	void getList(ArrayList<CWall> elemlist){
		elem=elemlist;
	}
	
	@Override
	public int compareTo(CWall other) {
		CWall secelem=null;
		CWall secelemoth=null;
		for(CWall c: elem){
			if(c.wall.equals(other.wall)){
				secelemoth=c;
			}
		}
		for(CWall c: elem){
			if(c.wall.equals(wall)){
				secelem=c;
			}
		}
		double elem1=secelemoth.list.get(13)*other.list.get(list.size()-1);
		double elem2=secelem.list.get(13)*list.get(list.size()-1);
		if(elem1>elem2){
			return 1;
		}else if(elem1==elem2){
			return 0;
		}else return -1;
	}
	
}
