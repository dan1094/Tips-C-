using System;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
//using Microsoft.Reporting.WebForms;


public class Tools
{
	public Tools()
	{
	}

	private static string ByteArrayToString(byte[] arrInput)
	{
		StringBuilder stringBuilder = new StringBuilder((int)arrInput.Length);
		for (int i = 0; i <= (int)arrInput.Length - 1; i++)
		{
			stringBuilder.Append(arrInput[i].ToString("X2"));
		}
		return stringBuilder.ToString().ToLower();
	}

	public static string GenerateHash(string strData)
	{
		MD5CryptoServiceProvider mD5CryptoServiceProvider = new MD5CryptoServiceProvider();
		byte[] bytes = Encoding.UTF8.GetBytes(strData);
		byte[] numArray = mD5CryptoServiceProvider.ComputeHash(bytes);
		mD5CryptoServiceProvider.Clear();
		return Tools.ByteArrayToString(numArray);
	}

    public static string GenerateSignaturePayU(string strOrden, float amount)
    {
        //ApiKey~merchantId~referenceCode~amount~currency
        string strCurrency = System.Configuration.ConfigurationManager.AppSettings["Currency"];
        string strMerchantId = System.Configuration.ConfigurationManager.AppSettings["MerchantId"];
        string strApiKey = System.Configuration.ConfigurationManager.AppSettings["ApiKey"];

        return GenerateHash(strApiKey + "~" + strMerchantId + "~" + strOrden + "~" + amount.ToString()
            .Replace(",", ".") + "~" + strCurrency);
    }

    public static string GenerateSignaturePayUTest(string strOrden, float amount)
    {
        string str = ConfigurationManager.AppSettings["com.PayU.Currency"].ToString();
        string str1 = "508029";
        string str2 = "4Vj8eK4rloUd272L48hsrarnUA";
        string[] strArrays = new string[] { str2, "~", str1, "~", strOrden, "~", amount.ToString().Replace(",", "."), "~", str };
        return Tools.GenerateHash(string.Concat(strArrays));
    }

    public static string GetSessionId()
    {
        return System.Web.HttpContext.Current.Session.SessionID;
    }

    public static bool isPriceInPesos()
	{
		if (ConfigurationManager.AppSettings["tipo_de_moneda"].ToString() == "Pesos")
		{
			return true;
		}
		return false;
	}

	public static string ReturnTypeCoin()
	{
		if (Tools.isPriceInPesos())
		{
			return "$";
		}
		return "USD $";
	}

	public static bool TryBool(object dato)
	{
		
		try
		{
			return Convert.ToBoolean(dato);
		}
		catch (Exception)
		{
		    return false;
		}
	}

	public static int TryInt(object dato)
	{
		int num;
		try
		{
			num = Convert.ToInt32(dato);
		}
		catch (Exception)
		{
			num = 0;
		}
		return num;
	}

    public static double TryDouble(object dato)
    {
        try
        {
            return Convert.ToDouble(dato);
        }
        catch (Exception)
        {
        }
        return 0;
    }

    public static float TryFloat(object dato)
    {
        try
        {
            return float.Parse(dato.ToString());
        }
        catch (Exception)
        {
        }
        return 0;
    }

    public static string GetTextByTag(string text, string tag, string tagEnd)
    {
        try
        {
            int startIndex = text.IndexOf(tag, StringComparison.Ordinal);
            tag = tagEnd + tag.Substring(1);
            int endIndex = text.IndexOf(tag, StringComparison.Ordinal);
            return text.Substring(startIndex, endIndex + tag.Length - startIndex);
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    //public static Model.PlantillasEmailGenerica GetTemplate(string strTemplateName)
    //{
    //    Model.PlantillasEmail plantillasEmail = Model.PlantillasEmail.SelectByName(strTemplateName);
    //    Model.PlantillasEmailGenerica genericTemplate = new Model.PlantillasEmailGenerica();
    //    if (Resources.Language.Idioma == "En")
    //    {
    //        genericTemplate.Body = plantillasEmail.BodyEn;
    //        genericTemplate.Asunto = plantillasEmail.AsuntoEn;
    //    }
    //    else
    //    {
    //        genericTemplate.Body = plantillasEmail.BodyEs;
    //        genericTemplate.Asunto = plantillasEmail.AsuntoEs;
    //    }
    //    return genericTemplate;
    //}

    public static DateTime TryDateTime(object value)
    {
        try
        {
            return DateTime.ParseExact(value.ToString(), "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
        }
        catch (Exception)
        {
            return DateTime.MinValue;
        }

    }

    public static DateTime TryDateTimeHtml(object value)
    {
        try
        {
            return Convert.ToDateTime(value);
        }
        catch (Exception)
        {

            return DateTime.Now;
        }

    }

    public static DateTime TryDateTimeTwo(object value)
    {
        try
        {
            DateTime dtDate;
            DateTime.TryParse(value.ToString(), out dtDate);
            return dtDate;
        }
        catch (Exception)
        {
            return DateTime.MinValue;
        }

    }

    public static byte[] GenerateInvoicePdf(int ordenId, bool isAffiliate)
    {
        //float flIva = System.Configuration.ConfigurationManager.AppSettings["Gnl_Iva_Colombia"];

        //ReportViewer rptInvoice = new ReportViewer();

        //string companyName =
        //    Configuracion.GetConfiguration(null, null, "Gnl_Code_Company_Invoice").FirstOrDefault().Valor;

        //DataTable dtCompany = Compañia.GetInfomationCompanyByCodeDataTable(companyName);

        //DataTable dtReport = Model.Orden.GetInvoiceReport(ordenId);
        //rptInvoice.LocalReport.ReportPath = "Reportes/rptInvoice.rdlc";
        //rptInvoice.LocalReport.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource("dsOrdenes", dtReport));
        //rptInvoice.LocalReport.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource("dsOrdenesProductos", dtReport));
        //rptInvoice.LocalReport.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource("dsFacturas", dtReport));
        //rptInvoice.LocalReport.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource("dsCompañia", dtCompany));
        //rptInvoice.LocalReport.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource("dsProductos", dtReport));

        ////  this.rptInvoice.LocalReport.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource("DataSet1_FFW_COMPANIA", dtCompany));

        //Numalet numalet = new Numalet(true, "Pesos", "Con", true);
        //string strPrice = numalet.ToCustomCardinal(dtReport.Rows[0]["ord_total"].ToString());
        //rptInvoice.ProcessingMode = ProcessingMode.Local;
        //Microsoft.Reporting.WebForms.ReportParameter[] p = new Microsoft.Reporting.WebForms.ReportParameter[4];

        //p[0] = new Microsoft.Reporting.WebForms.ReportParameter("Loc_Title", Resources.Language.Accept);
        //p[1] = new Microsoft.Reporting.WebForms.ReportParameter("Loc_Description", strPrice.ToUpper());
        //p[2] = new Microsoft.Reporting.WebForms.ReportParameter("Loc_IVA", flIva.ToString(CultureInfo.InvariantCulture));
        //p[3] = new Microsoft.Reporting.WebForms.ReportParameter("Loc_Order", ordenId.ToString());


        //rptInvoice.LocalReport.SetParameters(p);


        //String cad1 = String.Empty;
        //String cad2 = String.Empty;
        //String cad3 = String.Empty;
        //String[] cads;
        //Microsoft.Reporting.WebForms.Warning[] cads1;

        //return rptInvoice.LocalReport.Render("pdf", cad1, out cad1, out cad2, out cad3, out cads, out cads1);
        return new byte[0];
    }
}