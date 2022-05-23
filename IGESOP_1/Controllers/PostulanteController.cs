using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using General.Librerias.AccesoDatos;

namespace IGESOP_1.Controllers
{
    public class PostulanteController : Controller
    {
        // GET: Postulante
        public ActionResult Inicio()
        {
            return View();
        }

        public string cargar()
        {
            string rpta = "";
            daSQL odaSQL = new daSQL("conIGESOP");
            rpta = odaSQL.EjecutarComando("dbo.usp_PostulantesManteGenCsv_o");
            return rpta;
        }
    }
}