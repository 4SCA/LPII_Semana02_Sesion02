package org.cibertec;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EscribirExcel {

	public static void crearAPartirDeArrayList() 
	{
		ArrayList<Persona> personas = new ArrayList<Persona>();
		personas.add(new Persona("Oliver Jara", "https://parzibyte.me", 60));
		personas.add(new Persona("Rasmus Lerdorf", "https://toys.lerdorf.com/", 53));
		personas.add(new Persona("Brian Kernighan", "https://www.cs.princeton.edu/~bwk/", 80));
		Workbook libro = new XSSFWorkbook();
		final String nombreArchivo = "Personas.xlsx";
		Sheet hoja = libro.createSheet("Personas");
		
		String[] encabezados = {"Nombre","Web","Edad"};
		int indiceFila= 0;
		
		Row fila = hoja.createRow(indiceFila);
		for(int i =0; i <encabezados.length; i++) {	//Creacion de los encabezados
			String encabezado = encabezados[i]; //Encapsulos los encabezados
			Cell celda = fila.createCell(i);	//Se crea la celda 
			celda.setCellValue(encabezado);		//Se coloca el valor encapsulado en la celda creada
		}
		indiceFila++;
		for(int i = 0; i<personas.size(); i++) {
			fila = hoja.createRow(indiceFila);
			Persona persona = personas.get(i);	//Obtiene los datos del ArrayList 1 por 1
			fila.createCell(0).setCellValue(persona.getNombre()); //Crea la celda y coloca el respectivo dato
			fila.createCell(1).setCellValue(persona.getWeb());
			fila.createCell(2).setCellValue(persona.getEdad());
			indiceFila++; //Aumenta la fila para bajar y colocar los datos de los siguientes objetos
		}
		
		//Guardamos
		File directorioActual = new File(".");
		String ubicacion = directorioActual.getAbsolutePath(); //
		String ubicacionArchivoSalida = ubicacion.substring(0,ubicacion.length()-1)+nombreArchivo; //Donde se ubicara el archivo
	
		FileOutputStream outputStream;
		try {
			outputStream = new FileOutputStream(ubicacionArchivoSalida);
			libro.write(outputStream);
			libro.close();
			System.out.println("Libro de personas guardado correctamente");
		} catch (FileNotFoundException ex) {
			System.out.println("Error de filenotfound");
		}catch(IOException ex) {
			System.out.println("Error de IOException");
		}
	}
	
	public static void main(String[] args) {
		crearAPartirDeArrayList();
	}
}
