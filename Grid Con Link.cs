using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Entidades;
using Conexion;


namespace Daniel_Delgado_prueba
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Iniciar();
        }

        private void Iniciar()
        {
            List<ListaGiros> girosOficina = Conexion.Conexion.ListarOficinas();

            if (girosOficina != null)
            {
                GridView1.DataSource = girosOficina;
                GridView1.DataBind();
            }
        }


        protected void lnkView_Click(object sender, EventArgs e)
        {
            GridViewRow grdrow = (GridViewRow)((LinkButton)sender).NamingContainer;

            string idOficina = grdrow.Cells[0].Text;
            Response.Redirect(string.Format("~/GirosOficina.aspx?IdOficina={0}", idOficina));       
        }
    }
}
