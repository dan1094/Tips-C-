
		//Evaluar Querystring C#
		private void EvaluarQueryString() 
        {
            var valor = Request.QueryString.GetValues("IdOficina");
            int idOficina = 0;
            if (valor != null )
            {
                idOficina = Convert.ToInt32(valor[0]);
            }

            GridView1.DataSource = Conexion.Conexion.ListarGiros(idOficina);
            GridView1.DataBind();
        }
		
		
		
		
		//Llenar dropdownlist
		  <asp:DropDownList ID="ddlPaises"
            runat="server">
        </asp:DropDownList>
			ddlPaises.DataSource = WSControlBox.ObtenerPaises();
            ddlPaises.DataTextField = "Nombre";
            ddlPaises.DataValueField = "Id";
            ddlPaises.DataBind();
		
		
		