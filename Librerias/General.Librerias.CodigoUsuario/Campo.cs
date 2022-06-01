using System.Collections.Generic;

namespace General.Librerias.CodigoUsuario
{
    public class Campo
    {
        public static List<beCampo> CrearLista(string data, char sepReg = ';', char sepCampo='|')
        {
            List<beCampo> lbeCampo = new List<beCampo>();
            string[] lista = data.Split(sepReg);
            int n = lista.Length;
            string[] campos;
            beCampo obeCampo;
            for (int i=0;i<n;i++)
            {
                obeCampo = new beCampo();
                campos = lista[i].Split(sepCampo);
                obeCampo.Codigo = campos[0];
                obeCampo.Nombre = campos[1];
                lbeCampo.Add(obeCampo);
            }
            return lbeCampo;
        }
    }
}
