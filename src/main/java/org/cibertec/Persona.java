package org.cibertec;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class Persona {

	private String nombre, web; //modificar private es opcional
	private int edad;			//debido a la clase lombok
	
	public Persona(String nombre, String web, int edad) {
		super();
		this.nombre = nombre;
		this.web = web;
		this.edad = edad;
	}
	
	
}
