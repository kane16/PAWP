package ioapp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;


public class CToExcel {
	
	ArrayList<CWall> list=new ArrayList<>();
	ArrayList<CWall> wlist=new ArrayList<>();
	ArrayList<C_mine> mlist=new ArrayList<>();
	ArrayList<C_mine> resmlist=new ArrayList<>();
	ArrayList<Double> param=new ArrayList<>();
	String path;
	int NUMBER_OF_WORKING_DAYS;
	boolean SUNDAYS_AND_SATURDAYS;
	boolean SHORT_WORKING_TIME;
	int ADDITIONAL_WEEKENDS;
	double LEVEL_OF_EMPLOYMENT_BOTTOM;
	double LEVEL_OF_EMPLOYMENT_CEILING;
	double ABSENCE_BOTTOM;
	double ABSENCE_CEILING;
	double WDb;
	double WDn;
	
	public CToExcel(String path, int NUMBER_OF_WORKING_DAYS, boolean SUNDAYS_AND_SATURDAYS, 
			boolean SHORT_WORKING_TIME, int ADDITIONAL_WEEKENDS){
		this.path=path;
		this.ADDITIONAL_WEEKENDS=ADDITIONAL_WEEKENDS;
		this.NUMBER_OF_WORKING_DAYS=NUMBER_OF_WORKING_DAYS;
		this.SUNDAYS_AND_SATURDAYS=SUNDAYS_AND_SATURDAYS;
		this.SHORT_WORKING_TIME=SHORT_WORKING_TIME;
	}
	
	void dread_from() throws IOException{
		FileInputStream file = new FileInputStream(new File(path));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		Iterator<Row> rowIterator = sheet.iterator();
		
		String txtpath="wall.txt";
		File txtfile=new File(txtpath);
		PrintWriter write=new PrintWriter(txtfile);
		int i=0;
		while(rowIterator.hasNext()){
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while(cellIterator.hasNext()){
				Cell cell = cellIterator.next();
				switch(cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						write.print((int)(cell.getNumericCellValue()*1000) + "\t\t");
						break;
					case Cell.CELL_TYPE_STRING:
						if(i>0){
							write.print(cell.getStringCellValue() + "\t\t");
							write.println();
						}
						break;
				}
			}
			i++;
			write.println();
		}
		write.close();
		
		Scanner sc=new Scanner(txtfile);
		File respath=new File("data.txt");
		PrintWriter writedata=new PrintWriter("data.txt");
		wlist=new ArrayList<>();
		String lp=null;
		boolean ismounted=false;
		while(ismounted==false){
			lp=sc.nextLine();
			if(lp.startsWith("1.")){
				ismounted=true;
			}
		}
		ismounted=false;
		int j=1;
		while(ismounted==false){
			String wall=sc.nextLine();
			String mine=sc.nextLine();
			ArrayList<Double> list=new ArrayList<>();
			while(sc.hasNextInt()){
				list.add(sc.nextInt()/1000.0);
			}
			wlist.add(new CWall(lp,wall,mine,list,""));
			sc.nextLine();
			lp=sc.nextLine();
			j++;
			if(!lp.startsWith(j+".")){
				ismounted=true;
			}
		}
		
		String type=null;
		ismounted=false;
		while(ismounted==false){
			type=sc.nextLine();
			if(type.startsWith("Typ i zakres danych")){
				ismounted=true;
			}
		}
		
		String build=sc.nextLine();
		int k=0;
		
		while(build.startsWith("<") || build.startsWith("{")){
			if(build.startsWith("<")){
				build=build.replaceAll("<", " ");
				build=build.replaceAll(";", " ");
				build=build.replaceAll(">", " ");
				Scanner str=new Scanner(build);
				double a=str.nextDouble();
				double b=str.nextDouble();
				for(CWall c: wlist){
					if(c.list.get(k)<a || c.list.get(k)>b){
						c.error=c.error+"  Poza zakresem danych parametr "+(k+1);
					}
				}
			}
			if(build.startsWith("{")){
				build=build.replaceAll("\\{", " ");
				build=build.replaceAll(";", " ");
				build=build.replaceAll("\\}", " ");
				Scanner str=new Scanner(build);
				ArrayList<Integer> arr=new ArrayList<>();
				while(str.hasNextInt()){
					arr.add(str.nextInt());
				}
				for(CWall c: wlist){
					boolean present=false;
					for(Integer in: arr){
						if((double)in==c.list.get(k)){
							present=true;
							break;
						}
					}
					if(present==false){
						c.error=c.error+"  Parametr "+(k+1)+" nie mieści się w zbiorze danych";
					}
				}
				str.close();
			}
			k++;
			build=sc.nextLine();
		}
		
		
		for(CWall c:wlist){
			writedata.print(c.lp+" "+c.wall+" "+c.mine+" ");
			for(Double d: c.list){
				writedata.print(" "+ d +" ");
			}
			writedata.print(c.error);
			writedata.println();
		}
		
		ismounted=false;
		while(ismounted==false){
			type=sc.nextLine();
			if(type.startsWith("Poziom zatrudnienia na dole")){
				ismounted=true;
			}
		}
		
		sc.nextLine();
		sc.nextLine();
		LEVEL_OF_EMPLOYMENT_BOTTOM=sc.nextInt()/1000.0;
		
		for(int s=0;s<5;s++){
			sc.nextLine();
		}
		
		LEVEL_OF_EMPLOYMENT_CEILING=sc.nextInt()/1000.0;
		
		for(int s=0;s<5;s++){
			sc.nextLine();
		}
		
		ABSENCE_BOTTOM=sc.nextInt()/1000.0;
		
		for(int s=0;s<5;s++){
			sc.nextLine();
		}
		
		ABSENCE_CEILING=sc.nextInt()/1000.0;
		
		for(int s=0;s<5;s++){
			sc.nextLine();
		}
		
		WDb=sc.nextInt()/1000.0;
		
		for(int s=0;s<5;s++){
			sc.nextLine();
		}
		
		WDn=sc.nextInt()/1000.0;
		
		ismounted=false;
		while(ismounted==false){
			type=sc.nextLine();
			if(type.startsWith("1.")){
				ismounted=true;
			}
		}
		
		ismounted=false;
		j=1;
		while(ismounted==false){
			String mine=sc.nextLine();
			ArrayList<Double> list=new ArrayList<>();
			while(sc.hasNextInt()){
				list.add(sc.nextInt()/1000.0);
			}
			mlist.add(new C_mine(lp,mine,list,""));
			sc.nextLine();
			lp=sc.nextLine();
			j++;
			if(!lp.startsWith(j+".")){
				ismounted=true;
			}
		}
		
		type=null;
		ismounted=false;
		while(ismounted==false){
			type=sc.nextLine();
			if(type.startsWith("Typ i zakres danych")){
				ismounted=true;
			}
		}
		
		build=sc.nextLine();
		k=0;
		
		while(build.startsWith("<") || build.startsWith("{")){
			if(build.startsWith("<")){
				build=build.replaceAll("<", " ");
				build=build.replaceAll(";", " ");
				build=build.replaceAll(">", " ");
				Scanner str=new Scanner(build);
				double a=str.nextDouble();
				double b=str.nextDouble();
				for(C_mine c: mlist){
					if(c.list.get(k)<a || c.list.get(k)>b){
						c.error=c.error+"  Poza zakresem danych parametr "+(k+1);
					}
				}
			}
			if(build.startsWith("{")){
				build=build.replaceAll("\\{", " ");
				build=build.replaceAll(";", " ");
				build=build.replaceAll("\\}", " ");
				Scanner str=new Scanner(build);
				ArrayList<Integer> arr=new ArrayList<>();
				while(str.hasNextInt()){
					arr.add(str.nextInt());
				}
				for(C_mine c: mlist){
					boolean present=false;
					for(Integer in: arr){
						if((double)in==c.list.get(k)){
							present=true;
							break;
						}
					}
					if(present==false){
						c.error=c.error+"  Parametr "+(k+1)+" nie mieści się w zbiorze danych";
					}
				}
				str.close();
			}
			k++;
			build=sc.nextLine();
		}
		
		writedata.close();
		file.close();
		sc.close();
		FileOutputStream out = new FileOutputStream(new File(path));
		workbook.write(out);
		out.close();
		workbook.close();
		respath.delete();
		txtfile.delete();
	}
	
	void make_computing() throws FileNotFoundException{
		File file=new File("result.txt");
		double Eef=0;
		double LZPDRmax=0;
		double LZKDRmax=0;
		double TPKdmax=0;
		double TPKzmmax=0;
		double WDBmax_dr=0;
		double WDBmax_sn=0;
		double WDBmax=0;
		double RWmax=0;
		double WDBmin=0;
		double LZPDR=0;
		double LZKDR=0;
		double SWDB=0;
		double WDBZNmax_dr=0;
		double LZPDRZNmax=0;
		for(CWall c: wlist){
			if(SUNDAYS_AND_SATURDAYS==false){
				c.list.set(12, 0.0);
			}
			if(SHORT_WORKING_TIME==false){
				c.list.set(1, 0.0);
			}
			CWall newc=new CWall(c.lp, c.wall, c.mine, new ArrayList<Double>(), c.error);
			list.add(newc);
			for(int i=0;i<20;i++){
				newc.list.add(0.0);
			}
			if(c.list.get(1)==0){
				Eef=450-2.0*c.list.get(0);
			}
			if(c.list.get(1)==1){
				Eef=360-2.0*c.list.get(0);
			}
			TPKdmax=1440*(c.list.get(3)/100)*(1-c.list.get(2)/100);
			LZPDRmax=1440/Eef*c.list.get(7)/c.list.get(8)/(1+c.list.get(7)/c.list.get(8));
			WDBZNmax_dr=c.list.get(5)/c.list.get(6)*c.list.get(16);
			LZPDRZNmax=LZPDRmax*WDBZNmax_dr/TPKdmax*60/c.list.get(4);
			if(LZPDRZNmax<LZPDRmax){
				LZPDRmax=LZPDRZNmax;
			}
			LZKDRmax=LZPDRmax/c.list.get(7)*c.list.get(8);
			double LZPDR0=c.list.get(7);
			TPKzmmax=TPKdmax/LZPDR0;
			if(LZPDRmax<LZPDR0){
				TPKzmmax=TPKdmax/LZPDRmax;
			}
			if(TPKzmmax>Eef*(c.list.get(3)/100)){
				TPKzmmax=Eef*(c.list.get(3)/100);
			}
			WDBmax_dr=LZPDRmax*TPKzmmax/60*c.list.get(4);
			newc.list.set(13, Eef);
			newc.list.set(14, LZPDRmax);
			newc.list.set(15, LZKDRmax);
			newc.list.set(16, TPKdmax);
			newc.list.set(17, TPKzmmax);
			if(c.list.get(12)<2*LZPDRmax){
				WDBmax_sn=c.list.get(12)*TPKzmmax/60*c.list.get(4);
			}else WDBmax_sn=2*LZPDRmax*TPKzmmax/60*c.list.get(4);
			WDBmax=WDBmax_dr+WDBmax_sn*ADDITIONAL_WEEKENDS/NUMBER_OF_WORKING_DAYS;
			newc.list.set(18, WDBmax);
			RWmax=WDBmax-c.list.get(16);
			newc.list.set(0, RWmax);
			WDBmin=c.list.get(11)*c.list.get(4)*TPKzmmax/60;
			LZPDR=c.list.get(11);
			LZKDR=LZPDR*c.list.get(8)/c.list.get(7);
			newc.list.set(1, LZPDR);
			newc.list.set(2, LZKDR);
			SWDB=SWDB+WDBmin;
			newc.list.set(19, c.list.get(4)*TPKzmmax/60);
			newc.list.set(11, WDBmin);
			newc.list.set(12, LZPDRZNmax);
			}
		for(CWall cwall: list){
			cwall.getList(wlist);
		}
		compare_Walls(list);
		File resfile=new File("iterresult.txt");
		PrintWriter write=new PrintWriter(resfile);
		write.printf("%-16s%-24s%-16s%-10s%-10s%-10s%-10s%-10s%-10s%-10s%-10s%-10s%-10s","LP", "Nazwa ściany", "Nazwa kopalni",
				"Eef", "LZPDRmax","LZKDRmax","TPKdmax","TPKzmmax","WDBmax","RWmax","WDBzm", "WDBmin", "LZPDRZNmax");
		write.println();
		for(CWall c: list){
			write.printf("%s%s%s%-10.2f%-10.2f%-10.2f%-10.2f%-10.2f%-10.2f%-10.2f%-10.2f%-10.2f%-10.2f",c.lp, c.wall,c.mine, c.list.get(13),c.list.get(14),c.list.get(15),c.list.get(16),
					c.list.get(17),c.list.get(18),c.list.get(0),c.list.get(19),c.list.get(11),c.list.get(12));
			write.println();
		}
		int p=0;
		for(CWall c: list){
			for(CWall wc: wlist){
				if(c.wall.equals(wc.wall)){
					while(c.list.get(1)<c.list.get(14)){
						write.print("Iteracja "+p);
						write.println();
						p++;
						write.printf("%-16s%-24s%-22s%-10s%-10s%-10s%-10s", "LP", "Nazwa ściany", "Nazwa kopalni", "LZPDR", "LZKDR",
								"LZPSN","LZKSN");
						write.println();
						for(CWall wall: list){
							write.printf("%s%s%s",wall.lp,wall.wall,wall.mine);
							write.printf("%10.2f",wall.list.get(1));
							write.printf("%10.2f",wall.list.get(2));
							write.printf("%10.2f",wall.list.get(3));
							write.printf("%10.2f",wall.list.get(4));
							write.println();
						}
						write.println("SWDB:" +SWDB);
						if(SWDB<WDb){
							c.list.set(1, c.list.get(1)+1);
							c.list.set(2, c.list.get(1)*wc.list.get(8)/wc.list.get(7));
							SWDB=SWDB+wc.list.get(4)*c.list.get(17)/60;
						}else break;
					}
					if(c.list.get(1)>c.list.get(14)){
						SWDB=SWDB-(c.list.get(1)-c.list.get(14))*wc.list.get(4)*c.list.get(17)/60;
						c.list.set(1, c.list.get(14));
						c.list.set(2, c.list.get(1)*wc.list.get(8)/wc.list.get(7));
					}
					if(SWDB>WDb){
						c.list.set(1, c.list.get(1)-(SWDB-WDb)/wc.list.get(4)/c.list.get(17)*60);
						SWDB=WDb;
						c.list.set(2, c.list.get(1)*wc.list.get(8)/wc.list.get(7));
						break;
					}
					if(wc.list.get(12)>(2*c.list.get(14))){
						wc.list.set(12, 2*c.list.get(14));
					}
					if(ADDITIONAL_WEEKENDS!=0){
						while(c.list.get(3)<wc.list.get(12)){
							write.print("Iteracja "+p);
							write.println();
							p++;
							write.printf("%-16s%-24s%-22s%-10s%-10s%-10s%-10s", "LP", "Nazwa ściany", "Nazwa kopalni", "LZPDR", "LZKDR",
									"LZPSN","LZKSN");
							write.println();
							for(CWall wall: list){
								write.printf("%s%s%s",wall.lp,wall.wall,wall.mine);
								write.printf("%10.2f",wall.list.get(1));
								write.printf("%10.2f",wall.list.get(2));
								write.printf("%10.2f",wall.list.get(3));
								write.printf("%10.2f",wall.list.get(4));
								write.println();
							}
							write.println("SWDB: "+SWDB);
							if(SWDB<WDb){
								c.list.set(3, c.list.get(3)+1);
								c.list.set(4, c.list.get(3)*wc.list.get(8)/wc.list.get(7));
								SWDB=SWDB+wc.list.get(4)*c.list.get(17)/60*ADDITIONAL_WEEKENDS/NUMBER_OF_WORKING_DAYS;
							}else break;
						}
						if(c.list.get(3)>wc.list.get(12)){
							SWDB=SWDB-(c.list.get(3)-wc.list.get(12))*wc.list.get(4)*c.list.get(17)/60*ADDITIONAL_WEEKENDS/NUMBER_OF_WORKING_DAYS;
							c.list.set(3, wc.list.get(12));
							c.list.set(4, c.list.get(3)*wc.list.get(8)/wc.list.get(7));
						}
						if(SWDB>WDb){
							c.list.set(3, c.list.get(3)-(SWDB-WDb)/wc.list.get(4)/c.list.get(17)*60/ADDITIONAL_WEEKENDS*NUMBER_OF_WORKING_DAYS);
							SWDB=WDb;
							c.list.set(4, c.list.get(3)*wc.list.get(8)/wc.list.get(7));
							break;
						}
					}
				}
			}
			if(SWDB==WDb){
				break;
			}
		}
		write.close();
		for(CWall c: list){
			for(CWall wc: wlist){
				if(c.wall.equals(wc.wall)){
					c.list.set(5, (c.list.get(1)+c.list.get(3)*ADDITIONAL_WEEKENDS/NUMBER_OF_WORKING_DAYS)*wc.list.get(4)*c.list.get(17)/60);
					c.list.set(6, c.list.get(5)-wc.list.get(16));
					c.list.set(7, (c.list.get(1)-wc.list.get(7))*wc.list.get(13)+(c.list.get(2)-wc.list.get(8))*wc.list.get(14));
					c.list.set(8, (c.list.get(3)-wc.list.get(9))*wc.list.get(13)+(c.list.get(4)-wc.list.get(10))*wc.list.get(14));
					c.list.set(9, c.list.get(7)/(1-(ABSENCE_BOTTOM/100))+c.list.get(8)*ADDITIONAL_WEEKENDS/((1-(ABSENCE_BOTTOM/100))*NUMBER_OF_WORKING_DAYS+150/7.5));
				}
			}
		}
		
		
		for(C_mine m: mlist){
			ArrayList<Double> arr=new ArrayList<>();
			for(int i=0;i<40;i++){
				arr.add(0.0);
			}
			resmlist.add(new C_mine(m.lp,m.mine,arr,m.error));
		}
		
		for(C_mine m: resmlist){
			double MAXDREef=0;
			double MAXSNEef=0;
			double MAXSNLZP=0;
			double LZOWDR=0;
			double LZOPWDR=0;
			double LZOWSN=0;
			double LZOPWSN=0;
			double dOPWDR=0;
			double dOPPWDR=0;
			double dOPWSN=0;
			double dOPPWSN=0;
			double LZDIVIDEDR=0;
			double LZDIVIDESN=0;
			for(CWall c: wlist){
				if(m.mine.equals(c.mine)){
					m.list.set(0, m.list.get(0)+c.list.get(16));
				}
			}
			
			for(CWall c: list){
				if(m.mine.equals(c.mine)){
					m.list.set(1, m.list.get(1)+c.list.get(5));
				}
			}
			
			m.list.set(2, m.list.get(1)-m.list.get(0));
			double LZPDRMAX=0;
			for(CWall c: list){
				if(m.mine.equals(c.mine)){
					if(m.list.get(3)<c.list.get(1)){
						m.list.set(3, c.list.get(1));
						MAXDREef=c.list.get(13);
						for(CWall cd: wlist){
							if(cd.wall.equals(c.wall)){
								LZPDRMAX=cd.list.get(7);
								LZDIVIDEDR=cd.list.get(8)/cd.list.get(7);
							}
						}
					}
				}
			}
			
			for(CWall c: list){
				if(m.mine.equals(c.mine)){
					if(m.list.get(4)<=c.list.get(3)){
						m.list.set(4, c.list.get(3));
						MAXSNEef=c.list.get(13);
						for(CWall cd: wlist){
							if(cd.wall.equals(c.wall)){
								LZDIVIDESN=cd.list.get(8)/cd.list.get(7);
							}
						}
					}
				}
			}
			for(CWall c: list){
				if(m.mine.equals(c.mine)){
					m.list.set(5, m.list.get(5)+c.list.get(9));
				}
			}
			if(LZPDRMAX==0){
				LZOWDR=0;
			}else{
				LZOWDR=4*m.list.get(3)/LZPDRMAX;
			}
			//Do wprowadzenia warunku sprawdzenia obłożenia wydobycia i przeróbki
			LZOWDR = 4;
			if(LZOWDR>4){
				LZOWDR=4;
			}
			LZOPWDR=3/4.0*LZOWDR;
			LZOWSN=4*m.list.get(4)*(1+LZDIVIDESN)*MAXSNEef/1440;
			if(LZOWSN>8){
				LZOWSN=8;
			}
			LZOPWSN=3/4.0*LZOWSN;
			for(C_mine c: mlist){
				if(c.mine.equals(m.mine)){
					if(LZOWDR==0){
						dOPWDR=0;
						dOPPWDR=0;
					}else{
					dOPWDR=c.list.get(0)*c.list.get(5)/100*(1-(c.list.get(2)/100))/4*(LZOWDR-4);
					dOPPWDR=c.list.get(1)*c.list.get(4)/100*(1-(c.list.get(3)/100))/3*(LZOPWDR-3);
					}
					if(LZOWSN==0){
						dOPWSN=0;
						dOPPWSN=0;
					}else{
					dOPWSN=c.list.get(0)*c.list.get(5)/100*(1-(c.list.get(2)/100))/4*LZOWSN;
					dOPPWSN=c.list.get(1)*c.list.get(4)/100*(1-(c.list.get(3)/100))/3*LZOPWSN;
					}
					m.list.set(6, dOPWDR+dOPPWDR);
					m.list.set(7, dOPWSN+dOPPWSN);
					m.list.set(8,dOPWDR/(1-(c.list.get(2)/100))
							+dOPPWDR/(1-(c.list.get(3)/100))
							+dOPWSN*ADDITIONAL_WEEKENDS/((1-(c.list.get(2)/100))*NUMBER_OF_WORKING_DAYS+150/7.5)
							+dOPPWSN*ADDITIONAL_WEEKENDS/((1-(c.list.get(3)/100))*NUMBER_OF_WORKING_DAYS+150/8.0));
					m.list.set(9, m.list.get(5)+m.list.get(8));
					m.list.set(10, m.list.get(0)*NUMBER_OF_WORKING_DAYS/(c.list.get(0)+c.list.get(1)));
					m.list.set(11, m.list.get(1)*NUMBER_OF_WORKING_DAYS/(c.list.get(0)+c.list.get(1)+m.list.get(9)));
					m.list.set(12, m.list.get(11)-m.list.get(10));
					m.list.set(13, m.list.get(12)/m.list.get(10)*100);
				}
			}
		}
		
		
		double WDBGK=0;
		double ZGK=0;
		double WGK=0;
		double WDBGKk=0;
		double ZGKk=0;
		double WGKk=0;
		double dWGK=0;
		double dWGKpr=0;
		
		for(C_mine m: resmlist){
			WDBGK=WDBGK+m.list.get(0);
			for(C_mine c: mlist){
				if(c.mine.equals(m.mine)){
					ZGK=ZGK+c.list.get(0)+c.list.get(1);
				}
			}
		}
		WGK=WDBGK*NUMBER_OF_WORKING_DAYS/ZGK;
		ZGKk=ZGK;
		for(C_mine m: resmlist){
			WDBGKk=WDBGKk+m.list.get(1);
			ZGKk=ZGKk+m.list.get(9);
		}
		WGKk=WDBGKk*NUMBER_OF_WORKING_DAYS/ZGKk;
		dWGK=WGKk-WGK;
		dWGKpr=dWGK/WGK*100;
		param.add(WGK);
		param.add(WGKk);
		param.add(dWGK);
		param.add(dWGKpr);
		file.delete();
	}
	
	void compare_Walls(ArrayList<CWall> list){
		Collections.sort(list);
	}
	
	void create_result(){
		
	}
	
	void modify_sheet(XSSFSheet sheet, ArrayList<CWall> list) throws FileNotFoundException{
		for(int i=5;i<list.size()+5;i++){
			Cell cell1=sheet.getRow(i).getCell(1);
			cell1.setCellValue(list.get(i-5).wall);
			Cell cell2=sheet.getRow(i).getCell(2);
			cell2.setCellValue(list.get(i-5).mine);
			for(int j=3;j<13;j++){
				Cell cell=sheet.getRow(i).getCell(j);
				cell.setCellValue(list.get(i-5).list.get(j-3));
			}
			Cell cell=sheet.getRow(i).getCell(13);
			cell.setCellValue(list.get(i-5).error);
		}
	}
	
	void modify_sheet_with_sorting(XSSFSheet unsorted, XSSFSheet sorted) throws FileNotFoundException{
		modify_sheet(sorted,list);
		for(CWall c: list){
			for(CWall wc: wlist){
				if(c.wall.equals(wc.wall)){
					wc.list=c.list;
				}
			}
		}
		modify_sheet(unsorted,wlist);
		for(int i=28;i<resmlist.size()+28;i++){
			Cell cell1=unsorted.getRow(i).getCell(1);
			cell1.setCellValue(resmlist.get(i-28).mine);
			for(int j=2;j<16;j++){
				Cell cell=unsorted.getRow(i).getCell(j);
				cell.setCellValue(resmlist.get(i-28).list.get(j-2));
			}
			Cell cell=unsorted.getRow(i).getCell(16);
			cell.setCellValue(resmlist.get(i-28).error);
		}
		for(int i=12;i<param.size()+12;i++){
			Cell cell=unsorted.getRow(32).getCell(i);
			cell.setCellValue(param.get(i-12));
		}
	}
	
	void soverride_worksheet_in_known_xlsx() throws IOException, EncryptedDocumentException, InvalidFormatException {
		FileInputStream fileinp = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fileinp);
		XSSFSheet unsorted=workbook.getSheetAt(1);
		XSSFSheet sorted=workbook.getSheetAt(2);
		modify_sheet_with_sorting(unsorted,sorted);
		FileOutputStream fileOut = new FileOutputStream(path);
		workbook.write(fileOut);
		fileOut.close(); 
    }
	
	String getPath(){
		File currentdirfile=new File("");
		String currentdir=currentdirfile.getAbsolutePath();
		return currentdir;
	}
	
}
