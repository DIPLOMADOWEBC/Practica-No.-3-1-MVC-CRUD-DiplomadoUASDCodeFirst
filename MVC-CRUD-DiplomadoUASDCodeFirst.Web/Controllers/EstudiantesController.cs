using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using MVC_CRUD_DiplomadoUASDCodeFirst.Model.DAL;
using MVC_CRUD_DiplomadoUASDCodeFirst.Model.Models;

namespace MVC_CRUD_DiplomadoUASDCodeFirst.Web.Controllers
{
    public class EstudiantesController : Controller
    {
        private EstudianteContext db = new EstudianteContext();

        // GET: Estudiantes
        public ActionResult Index()
        {
            return View(db.Estudiantes.ToList());
        }

        // GET: Estudiantes/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Estudiante estudiante = db.Estudiantes.Find(id);
            if (estudiante == null)
            {
                return HttpNotFound();
            }
            return View(estudiante);
        }

        // GET: Estudiantes/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Estudiantes/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "EstudianteID,Nombres,Apellidos,Fecha_Inscripcion")] Estudiante estudiante)
        {
            if (ModelState.IsValid)
            {
                db.Estudiantes.Add(estudiante);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(estudiante);
        }

        // GET: Estudiantes/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Estudiante estudiante = db.Estudiantes.Find(id);
            if (estudiante == null)
            {
                return HttpNotFound();
            }
            return View(estudiante);
        }

        // POST: Estudiantes/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "EstudianteID,Nombres,Apellidos,Fecha_Inscripcion")] Estudiante estudiante)
        {
            if (ModelState.IsValid)
            {
                db.Entry(estudiante).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(estudiante);
        }

        // GET: Estudiantes/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Estudiante estudiante = db.Estudiantes.Find(id);
            if (estudiante == null)
            {
                return HttpNotFound();
            }
            return View(estudiante);
        }
        public FileResult ExportarEstudiantesCSV()
        {
            EstudianteContext objEstudiante = new EstudianteContext();

            //Indicamos las columnas que tendra el archivo generado por la accion FileResult
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[4]
            {
                new DataColumn("EstudianteID"),
                new DataColumn("Nombres"),
                new DataColumn("Apellidos"),
                new DataColumn("Fecha_Inscripcion")
            });

            //Solo 20 primeros registros 
           // var estudantes = from Estudiante in objEstudiante.Estudiantes.Take(20)
                             //select Estudiante;

            //Que en nombre empiece con letra "J"
            //var estudantes = from Estudiante in objEstudiante.Estudiantes.Take(20)
                             //where Estudiante.Nombres.ToLower().StartsWith("j")
                             //select Estudiante;

            //Que en el apellido contenga la letra "a"
            var estudantes = from Estudiante in objEstudiante.Estudiantes.Take(20)
                            where Estudiante.Apellidos.ToLower().Contains("a")
                            select Estudiante;



            //Recoger y agregar los datos al archivo
            foreach (var Estudiante in estudantes)
            {
                dt.Rows.Add(Estudiante.EstudianteID, Estudiante.Nombres, Estudiante.Apellidos, Estudiante.Fecha_Inscripcion);
            }

            //Generar el tipo de objeto, en este caso un archivo excel
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (System.IO.MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocumment.spreadsheetmml.sheet", "Grid.xlsx");
                }
            }

        }

        // POST: Estudiantes/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Estudiante estudiante = db.Estudiantes.Find(id);
            db.Estudiantes.Remove(estudiante);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
