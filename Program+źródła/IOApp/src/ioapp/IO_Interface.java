package ioapp;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.NoSuchElementException;
import java.util.ArrayList;
import java.util.Scanner;

import javax.swing.*;
import org.apache.*;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class IO_Interface extends JFrame implements ActionListener{
	
	String pathfrom;
	static int NUMBER_OF_WORKING_DAYS;
	static boolean SUNDAYS_AND_SATURDAYS;
	static boolean SHORT_WORKING_TIME;
	static int ADDITIONAL_WEEKENDS;
	static int NUMBER_OF_WALLS;
	JLabel jl;
	
	
	public IO_Interface(){
		
	}
	
	void display(String message){
		setSize(600,100);
        setLayout(new FlowLayout());
        add(new JLabel(message));
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setVisible(true);
	}

	public static void main(String[] args){
		IO_Interface inf=new IO_Interface();
		inf.pathfrom=inf.getPath()+"\\config.txt";
		File file=new File(inf.pathfrom);
		Scanner sc=null;
		BufferedReader br = null;
		String path=null;
		try {
			sc = new Scanner(file);
		} catch (FileNotFoundException e) {
			inf.display("Nie znaleziono pliku config");
		}
		try{
			sc.nextLine();
			path=inf.getPath()+"\\"+sc.nextLine()+".xlsx";
			sc.nextLine();
			NUMBER_OF_WORKING_DAYS=sc.nextInt();
			sc.nextLine();
			sc.nextLine();
			String area=sc.nextLine();
			area=area.replaceAll("<", " ");
			area=area.replaceAll(";" , " ");
			area=area.replaceAll(">", " ");
			Scanner st=new Scanner(area);
			if(NUMBER_OF_WORKING_DAYS<st.nextInt() || NUMBER_OF_WORKING_DAYS>st.nextInt()){
				throw new ArithmeticException();
			}
			sc.nextLine();
			if(sc.nextLine().startsWith("tak")){
				SUNDAYS_AND_SATURDAYS=true;
			}else SUNDAYS_AND_SATURDAYS=false;
			sc.nextLine();
			if(sc.nextLine().startsWith("tak")){
				SHORT_WORKING_TIME=true;
			}else SHORT_WORKING_TIME=false;
			sc.nextLine();
			ADDITIONAL_WEEKENDS=sc.nextInt();
			st.close();
		}catch(ArithmeticException ex){
			inf.display("Liczba dni roboczych w roku poza zakresem");
		}catch(NoSuchElementException ex){
			ex.printStackTrace();
			inf.display("Coś nie tak z plikiem config.txt");
		}
		sc.close();
		CToExcel el=new CToExcel(path, NUMBER_OF_WORKING_DAYS, SUNDAYS_AND_SATURDAYS, SHORT_WORKING_TIME, ADDITIONAL_WEEKENDS);
		try {
			el.dread_from();
		} catch (IOException e1) {
			inf.display("Proces nie może uzyskać dostępu do pliku, ponieważ jest on używany przez inny proces");
		}
		try {
			el.make_computing();
		} catch (FileNotFoundException e1) {
			inf.display("Błąd I/O");
		}
		try {
			el.soverride_worksheet_in_known_xlsx();
		} catch (EncryptedDocumentException e) {
			inf.display("Błąd przy odczycie");
		} catch (InvalidFormatException e) {
			inf.display("Niewłaściwy format");
		} catch (IOException e) {
			inf.display("Błąd I/O");
		}
		System.out.println("Sukces !!!");
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		
	}
	
	String getPath(){
		File currentdirfile=new File("");
		String currentdir=currentdirfile.getAbsolutePath();
		return currentdir;
	}

}
