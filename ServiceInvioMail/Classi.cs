using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceInvioMail
{
   
        public class MailDto
        {
            public string Destinatario { set; get; }//ToRiceved destinatario
            public string DestinatarioLst { set; get; }//ToRiceved destinatario
            public string Cc { set; get; }
            public string Ccn { set; get; }
            public string Oggetto { get; set; }
            public string Messaggio { get; set; }

            public string UtenteFiguraInvio { get; set; }//IdirizzoDiInvio
            public string NominativoInvio { get; set; }//DenominazioneInvio

            public string UserMail { get; set; } //MailConfig
            public string Password { get; set; }//Pwd
            public bool? SsLAuto { get; set; }//SsLAuto
            public bool? UseDefaultCredential { get; set; }//UseDefaultCredential
            public string SmtPostaUscita { get; set; }
            public int SmtpPort { get; set; }
        }

        public partial class MailImpoDto
        {
            public int IdInvio { get; set; }
            public string Descrizione { get; set; }
            public string ToRiceved { get; set; }
            public string Cc { get; set; }
            public string CCn { get; set; }
            public string IdirizzoDiInvio { get; set; }
            public string DenominazioneInvio { get; set; }
            public string MailConfig { get; set; }
            public string Pwd { get; set; }
            public string Oggetto { get; set; }
            public Nullable<bool> SsLAuto { get; set; }
            public Nullable<bool> UseDefaultCredential { get; set; }
            public string SmtPostaUscita { get; set; }
            public Nullable<int> SmtpPort { get; set; }
        }

        public partial class MailLogDto
        {
            public int IdConta { get; set; }
            public Nullable<System.DateTime> Data { get; set; }

            public string Commenti { get; set; }
            public Nullable<int> Esito { get; set; }
            public Nullable<int> Tipo { get; set; }
        }

        public class Articoli
        {
            public string IdArticolo { get; set; }
            public string DescriArticolo { get; set; }
            public int? QuantiInMagazzino { get; set; }
            public int? MinMagazzino { get; set; }

        }

        //public class CampiPdf
        //{
        //    public string Campo1 { get; set; }
        //    public string Campo2 { get; set; }
        //    public string Campo3 { get; set; }
        //    public string Campo4 { get; set; }
        //}

}
