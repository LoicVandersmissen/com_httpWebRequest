using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Security;

namespace ws_declaration_sociale
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual), Guid("9CB95D4C-7956-45bd-96A2-31A166FB9667")]
    public interface DSCOM_Interface
    {
        //send http request to the url and return http status code
        [DispId(1)]
        Int32 reqHttp(string url, string httpVerb, string[] header, string data);

        //get the http request response's data
        [DispId(2)]
        String GetRequestResponse();

        //get the http request response's data
        [DispId(3)]
        String GetRequestResponse(string encodingName);

        //add a client certificat 
        [DispId(4)]
        Int32 SetClientCertificate(string clientCertificate, string passwordClientCertificate);

        //add a server certificat trust policy
        [DispId(5)]
        Int32 SetTrusAllCertificatePolicy(string value);

    }

    // Events interface DeclarationSociale_COMObjectEvents 
    [Guid("EF0808AD-786B-4145-AA1A-7FF8ACA0E7C7"),
    InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface DSCOM_Events
    {
    }





    [ComVisible(true)]

    [Guid("DEB604DA-0FE1-47e0-B5AB-1E433A912640"),
    ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(DSCOM_Events))]
    public class main : DSCOM_Interface
    {

        private String _requestResponse;

        private Int32 _codeHttp;

        private String _clientCertificate, _passwordClientCertificate, _clientCertificateType;

        private Boolean _postGZipData;

        private bool _trusAllCertificatePolicy;


        public Int32 reqHttp(string url, string httpVerb, string[] headers, string data)
        {

            /*
            _clientCertificate = @"G:\Projet PB126\declaration_sociale\certificats DSN pro-BTP\guichetprobrz.p12";

            _passwordClientCertificate = "Pr0Brz14";
            */


            string contentType;

            contentType = "";

            byte[] bytes;


            _codeHttp = 0;


            /***
             * 
             *Create connection to the Uri by HttpWebRequest 
             * 
             ***/
            Uri address = new Uri(url);
            HttpWebRequest request = WebRequest.Create(address) as HttpWebRequest;
            request.Method = httpVerb;


            /****
             * Attach header to the HTTP request             * 
             ****/
            foreach (string header in headers)
            {
                Int32 colonPosition;
                string keyHeader, valueHeader;


                colonPosition = header.IndexOf(":");

                keyHeader = header.Substring(0, colonPosition - 1).Trim();

                valueHeader = header.Substring(colonPosition + 1).Trim();
                try
                {
                    if (keyHeader != string.Empty && valueHeader != string.Empty)
                    {
                        switch (keyHeader)
                        {
                            ///////////////////////////////
                            //standard header in attributs 
                            //
                            case "Content-Type":

                                contentType = valueHeader;
                                request.ContentType = valueHeader;
                                break;

                            case "Content-Length":
                                request.ContentLength = int.Parse(valueHeader);
                                break;
                            //
                            ///////////////////////////////


                            ///////////////////////////////
                            //headers add by method
                            //

                            default:

                                if (keyHeader == "Content-Encoding" && valueHeader == "gzip")
                                {
                                    _postGZipData = true;
                                }


                                request.Headers.Add(keyHeader, valueHeader);
                                break;

                            //
                            /////////////////////////////// 
                        }

                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }

            }


            /***
             * 
             *Adding a client certificate. Work only with file.
             * 
             ***/
            if (!string.IsNullOrEmpty(_clientCertificate))
            {
                try
                {
                    X509Certificate2 Cert;
                    Cert = null;


                    switch (_clientCertificateType)
                    {

                        case "P12":
                            Cert = new X509Certificate2(_clientCertificate, _passwordClientCertificate);

                            break;
                        case "CER":
                            Cert = new X509Certificate2(_clientCertificate, "");

                            break;
                        default:
                            break;
                    }
                    if (Cert != null)
                    {
                        MessageBox.Show("enter in add cert");

                        request.ClientCertificates.Add(Cert);

                        MessageBox.Show("after add");

                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Erreur");
                }

            }
            else
            {
                if (_trusAllCertificatePolicy)
                {
                    //legacy code, suppress code
                    //System.Net.ServicePointManager.CertificatePolicy = new TrustAllCertificatePolicy();

                    //replace by this
                    System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                }
            }


            /***
             * 
             *1- Sending Data if data isn't null or empty string        
             *
             *2- Get the http status code and data (when send data or not)
             * 
            ***/
            try
            {
                if (!string.IsNullOrEmpty(data))
                {

                    request.Timeout = 1000000;
                    request.SendChunked = true;

                    Stream requestStream = request.GetRequestStream();



                    bytes = UTF8Encoding.UTF8.GetBytes(data);

                    if (_postGZipData == false)
                    {
                        requestStream.Write(bytes, 0, bytes.Length);
                    }
                    else
                    {

                        GZipStream gzipStream = new GZipStream(requestStream, CompressionMode.Compress);

                        gzipStream.Write(bytes, 0, bytes.Length);

                        gzipStream.Close();
                    }

                    requestStream.Close();

                }


                HttpWebResponse response = request.GetResponse() as HttpWebResponse;

                StreamReader reader = new StreamReader(response.GetResponseStream());
                _requestResponse = reader.ReadToEnd();

                _codeHttp = (int)response.StatusCode;

                reader.Close();
                response.Close();


            }


            /****
             * 
             * catch error
             * 
             ***/

            catch (WebException we)
            {
                MessageBox.Show(we.Message, "WE");
                _codeHttp = (int)((HttpWebResponse)we.Response).StatusCode;
                MessageBox.Show(_codeHttp.ToString());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
            }



            return _codeHttp;
        }


        //Get the data return by the response object stored in private attribut of main class
        public String GetRequestResponse()
        {
            return _requestResponse;
        }

        public String GetRequestResponse(string encodingName)
        {
            switch (encodingName)
            {
                case "UTF8":
                    convertToUtf8(ref _requestResponse);
                    break;
                case "ASCII":
                    convertToAscii(ref _requestResponse);
                    break;
                default:
                    break;
            }




            return _requestResponse;
        }

        //set values to use client certificate (not using for now)
        public Int32 SetClientCertificate(string clientCertificate, string passwordClientCertificate)
        {
            //return -1 if certicate path/name is empty, -2 is password is empty with a p12 certificate.

            Int32 lastIndexOfDot;

            lastIndexOfDot = 0;

            _clientCertificateType = "";


            if (clientCertificate == string.Empty)
            {
                return -1;
            }


            lastIndexOfDot = clientCertificate.LastIndexOf('.');

            _clientCertificateType = clientCertificate.Substring(lastIndexOfDot + 1).ToUpper();


            //   MessageBox.Show(clientCertificate, "clientCertificate");
            //   MessageBox.Show(lastIndexOfDot.ToString(), "lastIndexOfDot");


            //Password is required for P12 certificate type.
            if (passwordClientCertificate == string.Empty)
            {
                if (_clientCertificateType == "P12")
                {
                    return -2;
                }
            }


            convertToUtf8(ref clientCertificate);
            _clientCertificate = clientCertificate;
            //  MessageBox.Show(_clientCertificate, "_clientCertificate");

            convertToUtf8(ref  passwordClientCertificate);
            _passwordClientCertificate = passwordClientCertificate;

            // MessageBox.Show(passwordClientCertificate, "passwordClientCertificate");

            return 1;


        }

        /// <summary>
        /// convert source string to UTF8 encoding
        /// </summary>
        /// <param name="sourceString"></param>
        /// <returns></returns>
        private Boolean convertToUtf8(ref String sourceString)
        {
            byte[] utfBytes = Encoding.Default.GetBytes(sourceString);

            sourceString = Encoding.UTF8.GetString(utfBytes, 0, utfBytes.Length);


            return true;
        }

        /// <summary>
        /// convert source string to ASCII encoding
        /// </summary>
        /// <param name="sourceString"></param>
        /// <returns></returns>
        private Boolean convertToAscii(ref String sourceString)
        {
            byte[] utfBytes = Encoding.Default.GetBytes(sourceString);

            sourceString = Encoding.ASCII.GetString(utfBytes, 0, utfBytes.Length);


            return true;
        }


        /// <summary>
        /// Set the trusAllCertificatePolicy to true if value parameter == "O"
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public Int32 SetTrusAllCertificatePolicy(string value)
        {
            if (value == "O")
            {
                _trusAllCertificatePolicy = true;
            }

            return 1;
        }

        //legacy code, will be suppress
        ///// <summary>
        ///// Class used for server certificate problem
        ///// </summary>
        //public class TrustAllCertificatePolicy : System.Net.ICertificatePolicy
        //{
        //    public TrustAllCertificatePolicy() { }
        //    public bool CheckValidationResult(ServicePoint sp,
        //        X509Certificate cert,
        //        WebRequest req,
        //        int problem)
        //    {
        //        return true;
        //    }
        //}

    }
}
