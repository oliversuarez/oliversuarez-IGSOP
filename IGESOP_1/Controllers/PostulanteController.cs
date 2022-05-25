using System.Web.Mvc;
using io = System.IO;
using General.Librerias.AccesoDatos;

namespace IGESOP_1.Controllers
{
    public class PostulanteController : Controller
    {
        // GET: Postulante
        public ActionResult Inicio()
        {
            string archivoPopup = Server.MapPath("~/Popud/postulante.txt");
            if (io.File.Exists(archivoPopup))
            {
                ViewBag.Popup = io.File.ReadAllText(archivoPopup);
            }
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
