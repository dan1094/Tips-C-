using System.Net;
using System.Web;
using System.Web.Script.Serialization;



			//JavaScriptSerializer serializer = new JavaScriptSerializer();
            //string request = serializer.Serialize(requestEfective).Replace("[", "").Replace("]", "");
            //string responseCashPayment = PagosRULE.Request(strUrlApi, request);
            //ResponseTransaction responseTransaction = JsonConvert.DeserializeObject<ResponseTransaction>(responseCashPayment);

//Request in C#
  public static string Request(string url, string jsonContent)
        {
            long length = 0;
            try
            {

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                //request.KeepAlive = false;

                System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
                Byte[] byteArray = encoding.GetBytes(jsonContent);

                request.ContentLength = byteArray.Length;
                request.ContentType = @"application/json";
                request.Accept = @"application/json";

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                Stream dataStream = request.GetRequestStream();

                dataStream.Write(byteArray, 0, byteArray.Length);


                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    length = response.ContentLength;

                    //Response.Write(((HttpWebResponse)response).StatusDescription);
                    // Get the stream containing content returned by the server.
                    dataStream = response.GetResponseStream();
                    // Open the stream using a StreamReader for easy access.
                    StreamReader reader = new StreamReader(dataStream);
                    // Read the content.
                    string responseFromServer = reader.ReadToEnd();
                    // Display the content.
                    // Response.Write(responseFromServer);
                    return responseFromServer;
                }
            }
            catch (WebException ex)
            {
                return ex.Message + " " + ex.StackTrace;// Log exception and throw as for GET example above
            }
            return null;
        }
		
		

		public static IList<PaymentMethods> RequestPaymentMethods()
        {
            InitializeValues();
            List<Peticion> paymentMethods = new List<Peticion>{
                   new Peticion{
                       test = strTest,
                       language = strLanguage,
                       command = "GET_PAYMENT_METHODS",
                       //  command = "PING",
                       merchant = new Merchant(){
                           apiLogin = strApiLogin,
                           apiKey = strApiKey
                       } }
                   };

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            string paymentMet = serializer.Serialize(paymentMethods).Replace("[", "").Replace("]", "");
            string respPayment = Request(strUrlApi, paymentMet);

            var jss = new JavaScriptSerializer();
            var dict = jss.Deserialize<dynamic>(respPayment);
            ResponsePayments responsePayments = JsonConvert.DeserializeObject<ResponsePayments>(respPayment);

            if (responsePayments.code == "ERROR")
            {
                return new List<PaymentMethods>();
            }
            if (responsePayments.paymentMethods != null && responsePayments.paymentMethods.Count > 0)
            {
                IList<PaymentMethods> re = (from d in responsePayments.paymentMethods
                                            where d.country == strCountry && d.enabled == "True"
                                            && (d.description.Contains("VISA") || d.description.Contains("MASTERCARD")
                                                || d.description.Contains("AMEX") || d.description.Contains("DINER")
                                                 || d.description.Contains("VISA") || d.description.Contains("CODENSA")
                                                 || d.description.Contains("CMR"))
                                            select d).ToList().Distinct().ToList();

                IList<PaymentMethods> result = new List<PaymentMethods>();

                foreach (var item in re)
                {
                    if (!result.Any(a => a.description.Equals(item.description)))
                    {
                        result.Add(item);
                    }
                }
                return result;
            }
            else
                return new List<PaymentMethods>();
        }
