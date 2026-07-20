--UNIVERSIDAD UTC
--INTRODUCCION A LA PROGRAMACION
--ALUMNO: CRISTIAN PADILLA S.
--PROYECTO BASES DE DATOS

-- =============================================
-- FASE 2: CREACIÓN DE LA ESTRUCTURA (DDL)
-- SISTEMA DE GESTIÓN DE REPARACIONES TECNOLÓGICAS
-- =============================================

/* Comandos mas utilizados */
--DDL Data definition language: Create, Drop, Truncate, Alter

-- Crear la base de datos
CREATE DATABASE SistemaReparaciones -- Crea el espacio de almacenamiento para el sistema
GO

-- Seleccionar la base de datos para trabajar
USE SistemaReparaciones -- Indica que todas las siguientes acciones se harán en esta base
GO

-- =============================================
-- Tabla: Usuarios
-- =============================================
CREATE TABLE Usuarios
(
    UsuarioID        INT PRIMARY KEY IDENTITY(1,1), -- Identificador único autoincremental
    Nombre           VARCHAR(100) NOT NULL,          -- Campo obligatorio
    CorreoElectronico VARCHAR(150) NOT NULL UNIQUE,   -- Obligatorio y único para evitar duplicados
    Telefono         VARCHAR(20) NULL                 -- Campo opcional
)
GO

-- =============================================
-- Tabla: Equipos
-- =============================================
CREATE TABLE Equipos
(
    EquipoID         INT PRIMARY KEY IDENTITY(1,1), -- Identificador único autoincremental
    TipoEquipo       VARCHAR(50) NOT NULL,          -- Tipo de equipo (ej: Laptop, Impresora) - obligatorio
    Modelo           VARCHAR(100) NOT NULL,         -- Marca y modelo del equipo - obligatorio
    UsuarioID        INT NOT NULL,                  -- Enlace al dueño del equipo - obligatorio
    FOREIGN KEY (UsuarioID) REFERENCES Usuarios(UsuarioID) -- Relación con la tabla Usuarios
)
GO

-- =============================================
-- Tabla: Tecnicos
-- =============================================
CREATE TABLE Tecnicos
(
    TecnicoID        INT PRIMARY KEY IDENTITY(1,1), -- Identificador único autoincremental
    Nombre           VARCHAR(100) NOT NULL,          -- Nombre completo del técnico - obligatorio
    Especialidad     VARCHAR(80) NOT NULL             -- Área de conocimiento - obligatorio
)
GO

-- =============================================
-- Tabla: Reparaciones
-- =============================================
CREATE TABLE Reparaciones
(
    ReparacionID     INT PRIMARY KEY IDENTITY(1,1), -- Identificador único autoincremental
    EquipoID         INT NOT NULL,                  -- Enlace al equipo a reparar - obligatorio
    FechaSolicitud   DATETIME NOT NULL,              -- Fecha de solicitud, se ingresa manualmente
    Estado           VARCHAR(30) NOT NULL,           -- Estado de la reparación, se ingresa manualmente
    FOREIGN KEY (EquipoID) REFERENCES Equipos(EquipoID) -- Relación con la tabla Equipos
)
GO

-- =============================================
-- Tabla: Asignaciones
-- =============================================
CREATE TABLE Asignaciones
(
    AsignacionID     INT PRIMARY KEY IDENTITY(1,1), -- Identificador único autoincremental
    ReparacionID     INT NOT NULL,                  -- Enlace a la reparación - obligatorio
    TecnicoID        INT NOT NULL,                  -- Enlace al técnico encargado - obligatorio
    FechaAsignacion  DATETIME NOT NULL,              -- Fecha de asignación, se ingresa manualmente
    FOREIGN KEY (ReparacionID) REFERENCES Reparaciones(ReparacionID), -- Relación con Reparaciones
    FOREIGN KEY (TecnicoID) REFERENCES Tecnicos(TecnicoID) -- Relación con Tecnicos
)
GO

-- =============================================
-- Tabla: DetallesReparacion
-- =============================================
CREATE TABLE DetallesReparacion
(
    DetalleID        INT PRIMARY KEY IDENTITY(1,1), -- Identificador único autoincremental
    ReparacionID     INT NOT NULL,                  -- Enlace a la reparación - obligatorio
    Descripcion      VARCHAR(500) NOT NULL,          -- Detalle de la falla o trabajo - obligatorio
    FechaInicio      DATETIME NULL,                  -- Fecha de inicio, se ingresa manualmente
    FechaFin         DATETIME NULL,                  -- Fecha de finalización, se ingresa manualmente
    FOREIGN KEY (ReparacionID) REFERENCES Reparaciones(ReparacionID) -- Relación con Reparaciones
)
GO

--_______________________________________________________________________________________________________________________________________

-- =============================================
-- FASE 3: MANIPULACIÓN Y LÓGICA DE NEGOCIO
-- SISTEMA DE GESTIÓN DE REPARACIONES TECNOLÓGICAS
-- =============================================



-- =============================================
-- 1. POBLAMIENTO DE TABLAS
-- Insertar al menos 5 registros en cada tabla
-- =============================================

-- Insertar en Usuarios
INSERT INTO Usuarios (Nombre, CorreoElectronico, Telefono)
VALUES 
('Carlos Ruiz', 'carlos.ruiz@correo.com', '8765-1122'),
('María Gómez', 'maria.gomez@correo.com', '8901-3344'),
('Luis Fernández', 'luis.fernandez@correo.com', '8654-5566'),
('Ana Castro', 'ana.castro@correo.com', '8876-7788'),
('Jorge Mora', 'jorge.mora@correo.com', '8543-9900')
GO

-- Insertar en Equipos
INSERT INTO Equipos (TipoEquipo, Modelo, UsuarioID)
VALUES 
('Laptop', 'Lenovo IdeaPad 3', 1),
('Impresora', 'Epson L3250', 2),
('Celular', 'iPhone 13', 3),
('Tablet', 'Samsung Tab S6', 4),
('PC Escritorio', 'HP Pavilion', 5)
GO

-- Insertar en Tecnicos
INSERT INTO Tecnicos (Nombre, Especialidad)
VALUES 
('Roberto Díaz', 'Reparación de Hardware'),
('Sofía Herrera', 'Mantenimiento de Software'),
('Manuel Acosta', 'Redes y Conectividad'),
('Patricia Soto', 'Periféricos y Accesorios'),
('David Rojas', 'Diagnóstico General')
GO

-- Insertar en Reparaciones
INSERT INTO Reparaciones (EquipoID, FechaSolicitud, Estado)
VALUES 
(1, '05/07/2026', 'Pendiente'),
(2, '06/07/2026', 'Pendiente'),
(3, '07/07/2026', 'En proceso'),
(4, '08/07/2026', 'Pendiente'),
(5, '09/07/2026', 'En proceso')
GO

-- Insertar en Asignaciones
INSERT INTO Asignaciones (ReparacionID, TecnicoID, FechaAsignacion)
VALUES 
(1, 1, '05/07/2026'),
(2, 2, '06/07/2026'),
(3, 3, '07/07/2026'),
(4, 4, '08/07/2026'),
(5, 5, '09/07/2026')
GO

-- Insertar en DetallesReparacion
INSERT INTO DetallesReparacion (ReparacionID, Descripcion, FechaInicio, FechaFin)
VALUES 
(1, 'Equipo no enciende, revisión de fuente de poder', '05/07/2026', NULL),
(2, 'Atasco de papel y limpieza de cabezales', '06/07/2026', NULL),
(3, 'Problema con batería y bajo rendimiento', '07/07/2026', NULL),
(4, 'Pantalla táctil no responde correctamente', '08/07/2026', NULL),
(5, 'Sistema operativo lento y con errores', '09/07/2026', NULL)
GO

-- =============================================
-- 2. OPERACIONES CRUD
-- =============================================

-- CREAR: Insertar un nuevo registro
INSERT INTO Usuarios (Nombre, CorreoElectronico, Telefono)
VALUES ('Pedro López', 'pedro.lopez@correo.com', '8234-5678') -- Agrega un nuevo usuario
GO

-- LEER: Consultar registros
SELECT * FROM Usuarios -- Muestra todos los usuarios registrados
GO

-- ACTUALIZAR: Modificar un dato existente
UPDATE Usuarios SET Telefono = '8234-9999' WHERE UsuarioID = 6 -- Cambia el teléfono del usuario 6
GO

-- BORRAR: Eliminar un registro
DELETE FROM Usuarios WHERE UsuarioID = 6 -- Elimina el usuario número 6
GO