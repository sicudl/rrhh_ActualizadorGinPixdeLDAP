//#define SERVEISWEB_noORACLECLIENTINSTALL_NEEDED

using System;
using System.Collections;
using System.Collections.Generic;

using System.Text;

using RecursosProgramacio;
using MultiTools.General;
using ConectorLDAP;

namespace ActualizadorGinPixdeLDAP
{
    class GinPixLDAPFixer
    {
        private ConectorLDAP.Conector LDAP = null;

        public void DoFix()
        {

            /*            string SaltMD5 = "01XdfGgjFhc319.3pQXcnEU";
                        Secreter miCypher = new Secreter();*/

            this.LDAP = new ConectorLDAP.Conector();
            MultiTools.General.Email miEmail = new MultiTools.General.Email();

            string GlobalHTMLLogText = "";
            ConectorOracle.Conector bdGinPIX = null;

            string strMail = "";
            string rsComproba = "";
            string strConsultaEmpleadosSinEmailenGinPix = "";

            MultiTools.General.AppLog miLog = new AppLog("", DateTime.Now);
            List<Hashtable> Dades;

            long RegistresSenseEMAIL = 0;
            long RegistresSenseUID = 0;
            strConsultaEmpleadosSinEmailenGinPix =
                "select apellid1 as ape1,apellid2 as ape2,ndniempl || ddniempl as completo,ndniempl as numero from " +
"personal,XP_USUPORTAL,eq_personal where eq_personal.empleado_cod=personal.codiempl AND activo_hoy='S' AND personal.codienti = xp_usuportal.codienti and personal.codiempl = xp_usuportal.codiempl and email is null";
            //AND ndniempl || ddniempl='72489575S'";

#if SERVEISWEB_noORACLECLIENTINSTALL_NEEDED
            /*ServeiWeb = new global::ActualizadorGinPixdeLDAP.serveiWeb_eADMIN.Service();
            if(UrlSWConector!="")
                ServeiWeb.Url = UrlSWConector;*/
            bdGinPIX = RunScriptSQL(strConsultaEmpleadosSinEmailenGinPix);

#else
          bdGinPIX = new ConectorOracle.Conector("ginpixdb.udl.cat", "1521", "gxpower", "gxpower", "gxpower");
            bdGinPIX.GetDades(strConsultaEmpleadosSinEmailenGinPix);
            Dades=bdGinPIX.Dades;
#endif

            RegistresSenseEMAIL = 0;
            if (bdGinPIX != null)
                RegistresSenseEMAIL = bdGinPIX.Dades.Count;

            GlobalHTMLLogText = GlobalHTMLLogText + "<hr>SQL<p> select apellid1 as ape1,apellid2 as ape2,ndniempl || ddniempl as completo,ndniempl as numero from " +
"personal,XP_USUPORTAL,eq_personal where eq_personal.empleado_cod=personal.codiempl AND activo_hoy='S' AND personal.codienti = xp_usuportal.codienti and personal.codiempl = xp_usuportal.codiempl and email is null</p><p>Activo a data " + System.DateTime.Now + ", transpassat a portal web, i sense email a taula personal</p><br>";

            miLog.WriteEntry("empleats sense email : " + RegistresSenseEMAIL.ToString());

            System.Console.WriteLine("SQL Ginpix persones sense email, executada ok");

            string strComproba = "";
            if (RegistresSenseEMAIL > 0)
            {
                GlobalHTMLLogText = GlobalHTMLLogText + "s'han trobat " + bdGinPIX.Dades.Count + " persones, ara les anem a arreglar, posant-hi el email de LDAP, tot cercant al LDAP per digits DNI i digits DNI+ lletra  a camps LDAP employeeNumber i carLicense<br><br><p>";

                foreach (Hashtable Registre in bdGinPIX.Dades)
                {
                    bool Actualizado = false;

                    Actualizado = CercaLDAP((string)Registre["NUMERO"], (string)Registre["COMPLETO"]);

                    if (Actualizado)
                    {
                        Hashtable RegistreLDAP = (Hashtable)LDAP.Dades[0];
                        strMail = (string)RegistreLDAP["mail"];

                        strConsultaEmpleadosSinEmailenGinPix = "UPDATE personal SET EMAIL='" + strMail + "' WHERE ndniempl='" + (string)Registre["NUMERO"] + "'";
                        //                        bdGinPIX.Begin(strConsultaEmpleadosSinEmailenGinPix);
                        bdGinPIX = RunScriptSQL(strConsultaEmpleadosSinEmailenGinPix);
                        GlobalHTMLLogText = GlobalHTMLLogText + "<b>OK</b>" + (string)Registre["COMPLETO"] + "  " + (string)Registre["NUMERO"] + "  " + (string)Registre["APE1"] + "  " + (string)Registre["APE2"] + "<br>";
                        System.Console.WriteLine("##" + strConsultaEmpleadosSinEmailenGinPix);
                    }
                    else
                    {
                        miLog.WriteEntry("sense EMAIL, i problemes al LDAP, no hi ha manera de localitzar a : " + (string)Registre["NUMERO"] + (string)Registre["APE1"] + (string)Registre["APE2"]);
                        GlobalHTMLLogText = GlobalHTMLLogText + "<b>epic FAIL!</b>" + (string)Registre["COMPLETO"] + "  " + (string)Registre["NUMERO"] + "  " + (string)Registre["APE1"] + "  " + (string)Registre["APE2"] + "<br>";
                    }
                }
                GlobalHTMLLogText = GlobalHTMLLogText + "</p>";
            }


            strConsultaEmpleadosSinEmailenGinPix = "select apellid1 as ape1,apellid2 as ape2,ndniempl as numero, xp_usuportal.codiempl as codi,ndniempl || ddniempl as completo,idenldap FROM XP_USUPORTAL,personal,eq_personal where eq_personal.empleado_cod=personal.codiempl AND activo_hoy='S' AND personal.codienti = xp_usuportal.codienti and personal.codiempl = xp_usuportal.codiempl and idenldap is null";
            //AND ndniempl || ddniempl='72489575S'";



            GlobalHTMLLogText = GlobalHTMLLogText + "<hr>SQL<p> select apellid1 as ape1,apellid2 as ape2,ndniempl as numero, xp_usuportal.codiempl as codi,ndniempl || ddniempl as completo,idenldap FROM XP_USUPORTAL,personal,eq_personal where eq_personal.empleado_cod=personal.codiempl AND activo_hoy='S' AND personal.codienti = xp_usuportal.codienti and personal.codiempl = xp_usuportal.codiempl and idenldap is null</p><p>Activo a data " + System.DateTime.Now + ", transpassat a portal web, i el camp login de LDAP sense informar a la taula de portal</p><br>";
            

            //            bdGinPIX.GetDades(strConsultaEmpleadosSinEmailenGinPix);
            bdGinPIX = RunScriptSQL(strConsultaEmpleadosSinEmailenGinPix);

            RegistresSenseUID = 0;
            if (bdGinPIX != null)
                RegistresSenseUID = bdGinPIX.Dades.Count;

            miLog.WriteEntry("empleats sense uid : " + RegistresSenseUID.ToString());

            System.Console.WriteLine("SQL Ginpix persones sense UID del LDAP, executada ok");

            if (RegistresSenseUID > 0)
            {
                GlobalHTMLLogText = GlobalHTMLLogText + "s'han trobat " + bdGinPIX.Dades.Count + " persones, ara les anem a arreglar, posant-hi el uid del LDAP tot cercant a LDAP per digits DNI i digits DNI+ lletra  a camps LDAP employeeNumber i carLicense<br><br><p>";
                foreach (Hashtable Registre in bdGinPIX.Dades)
                {
                    bool Actualizado = false;

                    Actualizado = CercaLDAP((string)Registre["NUMERO"], (string)Registre["COMPLETO"]);

                    if (Actualizado)
                    {
                        Hashtable RegistreLDAP = (Hashtable)LDAP.Dades[0];
                        string strUID = (string)RegistreLDAP["uid"];

                        strConsultaEmpleadosSinEmailenGinPix = "UPDATE XP_USUPORTAL SET IDENLDAP='" + strUID + "' WHERE codiempl='" + (string)Registre["CODI"] + "'";
                        //                        bdGinPIX.Begin(strConsultaEmpleadosSinEmailenGinPix);
                        bdGinPIX = RunScriptSQL(strConsultaEmpleadosSinEmailenGinPix);
                        GlobalHTMLLogText = GlobalHTMLLogText + "<b>OK</b>" + (string)Registre["COMPLETO"] + "  " + (string)Registre["NUMERO"] + "  " + (string)Registre["APE1"] + "  " + (string)Registre["APE2"] + "<br>";
                        System.Console.WriteLine("##" + strConsultaEmpleadosSinEmailenGinPix);
                    }
                    else
                    {
                        miLog.WriteEntry("Sense UID i problemes al LDAP, no hi ha manera de localitzar a : " + (string)Registre["NUMERO"] + (string)Registre["APE1"] + (string)Registre["APE2"]);
                        GlobalHTMLLogText = GlobalHTMLLogText + "<b>epic FAIL!</b>"+(string)Registre["COMPLETO"] + "  " + (string)Registre["NUMERO"] + "  " + (string)Registre["APE1"] + "  " + (string)Registre["APE2"] + "<br>";
                    }
                }
                GlobalHTMLLogText = GlobalHTMLLogText + "</p>";
            }


            //Actualitzar a Ginpix la gent que no té correu...a ficha personal o a portal? als dos?

            //RECOMPTES FINALS
            strConsultaEmpleadosSinEmailenGinPix = "select apellid1 as ape1,apellid2 as ape2,ndniempl || ddniempl as completo,ndniempl as numero,codiempl as CODI from personal,eq_personal where eq_personal.empleado_cod=personal.codiempl AND activo_hoy='S' AND email is null order by ape1";

            GlobalHTMLLogText = GlobalHTMLLogText + "<hr>SQL<p> select apellid1 as ape1,apellid2 as ape2,ndniempl || ddniempl as completo,ndniempl as numero from personal,eq_personal where eq_personal.empleado_cod=personal.codiempl AND activo_hoy='S' AND email is null order by ape1</p><p>Activo a data " + System.DateTime.Now + ", i el camp email buit a  la taula personal :</p><br>";

            //            bdGinPIX.GetDades(strConsultaEmpleadosSinEmailenGinPix);
            bdGinPIX = RunScriptSQL(strConsultaEmpleadosSinEmailenGinPix);
            if (bdGinPIX != null)
            {
                RegistresSenseEMAIL = bdGinPIX.Dades.Count;
                GlobalHTMLLogText = GlobalHTMLLogText + "s'han trobat " + bdGinPIX.Dades.Count + " persones, ara les anem a arreglar posant-hi el email de LDAP i, tot cercant al LDAP pels digits DNI i digits DNI+ lletra a camps LDAP employeeNumber i carLicense<br><br><p>";

                foreach (Hashtable Registre in bdGinPIX.Dades)
                {
                    bool esTrobat = false;

                    esTrobat = CercaLDAP((string)Registre["NUMERO"], (string)Registre["COMPLETO"]);

                    if (esTrobat)
                    {
                        Hashtable RegistreLDAP = (Hashtable)LDAP.Dades[0];
                        string strCorreu = (string)RegistreLDAP["mail"];

                        strConsultaEmpleadosSinEmailenGinPix = "UPDATE PERSONAL SET EMAIL='" + strCorreu + "' WHERE codiempl='" + (string)Registre["CODI"] + "'";
                        //                        bdGinPIX.Begin(strConsultaEmpleadosSinEmailenGinPix);
                        bdGinPIX = RunScriptSQL(strConsultaEmpleadosSinEmailenGinPix);
                        GlobalHTMLLogText = GlobalHTMLLogText + "<b>OK</b>" + (string)Registre["COMPLETO"] + "  " + (string)Registre["NUMERO"] + "  " + (string)Registre["APE1"] + "  " + (string)Registre["APE2"] + "<br>";
                    }
                    else
                    {
                        GlobalHTMLLogText = GlobalHTMLLogText + "<b>epic FAIL!</b>" + (string)Registre["COMPLETO"] + "  " + (string)Registre["NUMERO"] + "  " + (string)Registre["APE1"] + "  " + (string)Registre["APE2"] + "<br>";
                        miLog.WriteSimpleTextLine((string)Registre["COMPLETO"] +" "+(string)Registre["NUMERO"] + (string)Registre["APE1"] + "      "+(string)Registre["APE2"] + "<br>");
                    }
                }
                GlobalHTMLLogText = GlobalHTMLLogText + "</p>";
            }

            bool test = false;
            test = miEmail.SendUdLEmail("Errors GinPix i LDAP - " + System.DateTime.Now.ToString(), GlobalHTMLLogText, "k4371605", "nitro1911", "k4371605@alumnes.udl.cat", "isaac.munoz@udl.cat");
            test = miEmail.SendUdLEmail("Errors GinPix i LDAP - " + System.DateTime.Now.ToString(), GlobalHTMLLogText, "k4371605", "nitro1911", "k4371605@alumnes.udl.cat", "jaume.esteban@udl.cat");
        }

        private bool CercaLDAP(string DNINumero, string DNIComplet)
        {

            bool Actualizado = false;

            //Conque tenim tanta creativitat alhora de informar els camp(s)(!!!!) de DNI al LDAP 
            //Calen varios intentos per asegurar que la persona NO esta al LDAP.
            //Primer intento...
            if (!Actualizado)
            {
                //Testme!
                //LDAP.GetDades("(&(employee//Number=437160))", true);
                LDAP.GetDades("(&(employeeNumber=" + DNINumero + "))", true);

                if (LDAP.Dades.Count > 0)
                    Actualizado = true;
                //Sino com a estudiant
                LDAP.GetDades("(&(employeeNumber=" + DNINumero + "))", false);

                if (LDAP.Dades.Count > 0)
                    Actualizado = true;
            }

            if (!Actualizado)
            {
                LDAP.GetDades("(&(employeeNumber=" + DNIComplet + "))", true);

                if (LDAP.Dades.Count > 0)
                    Actualizado = true;
                //Sino com a estudiant
                LDAP.GetDades("(&(employeeNumber=" + DNIComplet + "))", false);

                if (LDAP.Dades.Count > 0)
                    Actualizado = true;
            }

            if (!Actualizado)
            {
                LDAP.GetDades("(&(carLicense=" + DNINumero + "))", true);

                if (LDAP.Dades.Count > 0)
                    Actualizado = true;

                //Sino com a estudiant
                LDAP.GetDades("(&(carLicense=" + DNINumero + "))", false);

                if (LDAP.Dades.Count > 0)
                    Actualizado = true;
            }

            if (!Actualizado)
            {
                LDAP.GetDades("(&(carLicense=" + DNIComplet + "))", true);

                if (LDAP.Dades.Count > 0)
                    Actualizado = true;

                //Sino com a estudiant
                LDAP.GetDades("(&(carLicense=" + DNIComplet + "))", false);

                if (LDAP.Dades.Count > 0)
                    Actualizado = true;
            }
            return Actualizado;
        }

        private ConectorOracle.Conector RunScriptSQL(string SQLQuery)
        {
            //string AppServer = "http://localhost:1070/ERdpServeiWeb/Service.asmx";
            //string AppServer = "http://ercd.udl.net:8080/ercd/service.asmx";

            // OK pre canvi a nouercd <2025
            string AppServer = "http://ercd.udl.net:4646/service.asmx";
            
            // >1!/2025
            //string AppServer = "http://nouercd.udl.cat:8888/Service.asmx";
               

            ClientOracle.Conector DadesUXXIOracle = null;
            ConectorOracle.Conector mbdGinPIX = null;

            DadesUXXIOracle = new ClientOracle.Conector(AppServer, "gxpower", "ginpixdb.udl.cat", "gxpower");            
            DadesUXXIOracle.GetDades(SQLQuery);
            if (DadesUXXIOracle.HasRegisters)
            {
                System.Data.DataTable mDades = DadesUXXIOracle.GetThinDataTable("DADES");
                //DadesUXXIOracle.GetDataTable("DADES");
                mbdGinPIX = new ConectorOracle.Conector(mDades);
            }
            return mbdGinPIX;
        }
    }
    class Program
    {              
        static void Main(string[] args)
        {
            GinPixLDAPFixer MiFixer = new GinPixLDAPFixer();
            MiFixer.DoFix();
        }                     
    }
}
