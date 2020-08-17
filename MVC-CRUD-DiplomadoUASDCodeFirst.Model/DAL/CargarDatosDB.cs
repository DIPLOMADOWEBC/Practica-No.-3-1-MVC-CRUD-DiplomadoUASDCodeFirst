using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MVC_CRUD_DiplomadoUASDCodeFirst.Model.Models;

namespace MVC_CRUD_DiplomadoUASDCodeFirst.Model.DAL
{
    public class CargarDatosDB : System.Data.Entity.DropCreateDatabaseIfModelChanges<EstudianteContext>
    {
        protected override void Seed(EstudianteContext context)
        {

            var Cursos = new List<Curso>
            {
                new Curso { CursoID = 1, Descripcion = "Matematica", },
                new Curso { CursoID = 2, Descripcion = "Sociales", },
                new Curso { CursoID = 3, Descripcion = "Naturales", },
                new Curso { CursoID = 3, Descripcion = "Fisica", }
            };
            Cursos.ForEach(s => context.Cursos.Add(s));
            context.SaveChanges();

            var Estudiantes = new List<Estudiante>
            {
                new Estudiante{Nombres="Juan ",Apellidos="Carlos Reyes .",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="JHOJAN LEONARDO ",Apellidos="ABREU MONTAS",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="ROBINSON SANTIAGO ",Apellidos="ACOSTA REGALADO",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="JUAN CARLOS ",Apellidos="ALCANTARA ENCARNACION",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="SANTIAGO ",Apellidos="ANTUANE POLO",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="LUIS MIGUEL ",Apellidos="BALLENILLA RAMIREZ",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="ANGEL DANIEL ",Apellidos="BAUTISTA BRITO",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="DALWIN ",Apellidos="CUELLO BUENO",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="Silverio ",Apellidos="Del Orbe",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="MANUEL ARIEL ",Apellidos="GALVA FERNÁNDEZ",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="Juan Esteban ",Apellidos="González Merán",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="Abnel ",Apellidos="Lluberes",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="ALFRED JODE ",Apellidos="LOPEZ JIMENEZ",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="JERLIS ANTONIO ",Apellidos="MEDINA GOMEZ",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="CARLOS RAUL ",Apellidos="PEÑA EVERTSZ",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="ISIDRO ",Apellidos="PEREZ ALCÁNTARA",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="MIGUEL ",Apellidos="RODRIGUEZ PERDOMO",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="YOHELKIS ESTERLIN ",Apellidos="SANTANA GONZALEZ",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="JEURY ",Apellidos="SUAREZ HERNANDEZ",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="ANA MERCEDES ",Apellidos="TEJEDA TEJEDA",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="DARWIN DAVID ",Apellidos="VALDEZ OVALLE",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
                new Estudiante{Nombres="FREDDY JOEL ",Apellidos="VIRGIL DE LA CRUZ",Fecha_Inscripcion=DateTime.Parse("07-08-2020")},
            };
            Estudiantes.ForEach(s => context.Estudiantes.Add(s));
            context.SaveChanges();

            var Inscripcion = new List<Inscripcion>
            {
                new Inscripcion{InscripcionID =1,CursoID=1,EstudianteID=1,Semestre = 1},
                new Inscripcion{InscripcionID =2,CursoID=1,EstudianteID=2,Semestre = 6},
                new Inscripcion{InscripcionID =3,CursoID=2,EstudianteID=3,Semestre = 7},
                new Inscripcion{InscripcionID =4,CursoID=2,EstudianteID=4,Semestre = 3},
            };
            Inscripcion.ForEach(s => context.Inscripciones.Add(s));
            context.SaveChanges();
        }
    }
}
