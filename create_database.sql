DROP DATABASE IF EXISTS test_altos_ejecutivos;
CREATE DATABASE test_altos_ejecutivos;

USE test_altos_ejecutivos;

CREATE TABLE Cargos (
	`id` INT AUTO_INCREMENT NOT NULL,
	`nombre` VARCHAR(255) NOT NULL,
	`grado` VARCHAR(5) NOT NULL,
	`genero` CHAR(1) NOT NULL,
	`nacionalidad` VARCHAR(26) NOT NULL,
	CONSTRAINT pk_cargos_id PRIMARY KEY(id)
);


CREATE TABLE Rentas (
	`id` INT AUTO_INCREMENT NOT NULL,
	`cargo_id` INT NOT NULL,
	`renta_bruta` INT NOT NULL,
	CONSTRAINT pk_rentas_id PRIMARY KEY(id),
	FOREIGN KEY (cargo_id) REFERENCES Cargos(id)
);



