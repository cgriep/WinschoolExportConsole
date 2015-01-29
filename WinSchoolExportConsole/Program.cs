using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Data.Odbc;

namespace WinSchoolExportConsole
{
    class Program
    {
        private static string CONFIGURATIONFILE = "winschoolexport.cfg";
        private static string WINSCHOOLEXPORT = "WinschoolExport";
        private string WWWUSER = "user";
        private string WWWPASSWORD = "";
        private string DBSERVER = "ws_data";
        private string DATABASE = "wsdata";
        private string DBUSER = "sa";
        private string DBPASSWORD = "oszimt";
        private string UPLOADPAGE = "https://url/winschool/uploadimport.php";
        private string MAILTO = "mailaddresse";
        private string MAILFROM = "WinSchoolExport <noreply@oszimt.de>";
        private string MAILSERVER = "";
        private string MAILUSER = "";
        private string MAILPASSWORD = "";
        private string DATAFILE = "winschool";
        private string SQLSCHULJAHR = "SELECT LfdSchuljahr, Halbjahr FROM Globals";
        private string SQLSCHUELER = "SELECT SchuelerTable.NR, Vorname, Name, Klasse, Tutor, Betrieb.Name1 AS Betrieb, Geburtsdatum, Geschlecht, Abteilung "+
            "FROM SchuelerTable INNER JOIN KlassenTable ON KurzForm=Klasse LEFT JOIN dbo.Betrieb ON Kürzel=SchuelerTable.Ausbildungsbetrieb;";
        private string SQLKURSE = "SELECT SchuelerNr, AnzahlFächer, Kurs0, Fach0, Art0, Kurs1, Fach1, Art1, Kurs2, Fach2, Art2, Kurs3, Fach3, Art3, "+
            "Kurs4, Fach4, Art4, Kurs5, Fach5, Art5, Kurs6, Fach6, Art6, Kurs7, Fach7, Art7, Kurs8, Fach8, Art8, Kurs9, Fach9, Art9, Kurs10, Fach10, Art10, "+
            "Kurs11, Fach11, Art11, Kurs12, Fach12, Art12, Kurs13, Fach13, Art13, Kurs14, Fach14, Art14, Kurs15, Fach15, Art15 "+ 
            "FROM Schulhalbjahre WHERE Kurs0 IS NOT NULL AND Schulhalbjahr LIKE '%[[SCHULJAHR]]'"; // Platzhalter [[Schuljahr]]

        static void Main(string[] args)
        {
            Program wc = new Program();
            wc.Run();
        }
        private void Run()
        {
            // load configuration
            if (!EventLog.SourceExists(WINSCHOOLEXPORT))
            {
                EventLog.CreateEventSource(WINSCHOOLEXPORT, "WinschoolExport Log");
            }
            if (System.IO.File.Exists(CONFIGURATIONFILE))
            {
                System.Xml.XmlDocument xdoc = new System.Xml.XmlDocument();
                try
                {
                    xdoc.Load(CONFIGURATIONFILE);
                    try
                    {
                        WWWUSER = xdoc.SelectSingleNode("//wwwuser").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        DBSERVER = xdoc.SelectSingleNode("//dbserver").InnerText;
                    }
                    catch (Exception) { };

                    try
                    {
                        DBUSER = xdoc.SelectSingleNode("//dbuser").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        WWWPASSWORD = xdoc.SelectSingleNode("//wwwpassword").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        DBPASSWORD = xdoc.SelectSingleNode("//dbpassword").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        UPLOADPAGE = xdoc.SelectSingleNode("//uploadpage").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        MAILTO = xdoc.SelectSingleNode("//mailto").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        MAILFROM = xdoc.SelectSingleNode("//mailfrom").InnerText;
                    }
                    catch (Exception) { };
                    /*
                    try
                    {
                        WWWUSER = xdoc.SelectSingleNode("//wwwserver").InnerText;
                    }
                    catch (Exception) { };
                     */
                    try
                    {
                        MAILSERVER = xdoc.SelectSingleNode("//mailserver").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        MAILUSER = xdoc.SelectSingleNode("//mailuser").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        DATABASE = xdoc.SelectSingleNode("//database").InnerText;
                    }
                    catch (Exception) { };

                    try
                    {
                        MAILPASSWORD = xdoc.SelectSingleNode("//mailpassword").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        DATAFILE = xdoc.SelectSingleNode("//datafile").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        SQLSCHUELER = xdoc.SelectSingleNode("//sqlschueler").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        SQLSCHULJAHR = xdoc.SelectSingleNode("//sqlschuljahr").InnerText;
                    }
                    catch (Exception) { };
                    try
                    {
                        SQLKURSE = xdoc.SelectSingleNode("//sqlkurse").InnerText;
                    }
                    catch (Exception) { };
                }
                catch (Exception err)
                {
                    // wrong configuration file
                    EventLog.WriteEntry(WINSCHOOLEXPORT, "Falsche Konfiguration: " + err.Message, EventLogEntryType.Error);
                }
            }
            exportiereDaten();
        }
        private string sendeDaten(string datafile, int anzahl)
        {
            // Daten senden
            string message = "";
            try
            {
                if (System.IO.File.Exists(datafile))
                {
                    System.Net.WebClient client = new System.Net.WebClient();
                    client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                    //Creating an instance of a credential cache, 
                    //and passing the username and password to it
                    System.Net.CredentialCache mycache = new System.Net.CredentialCache();
                    mycache.Add(new Uri(UPLOADPAGE),"Basic", new System.Net.NetworkCredential(WWWUSER, WWWPASSWORD));
                    client.Credentials = mycache;
                    try
                    {
                        
                        Byte[] antwort = client.UploadFile(UPLOADPAGE, "POST", datafile);
                        
                        EventLog.WriteEntry(WINSCHOOLEXPORT, anzahl + " Datensätze in "+datafile+" erfolgreich hochgeladen.", EventLogEntryType.Information);
                        message = "Export von "+datafile+" ist erfolgreich verlaufen ("+System.Text.Encoding.Default.GetString(antwort)+").\nEs wurden " +
                            anzahl + " Datensätze in "+datafile+" exportiert.";
                    }
                    catch (Exception err)
                    {
                        EventLog.WriteEntry(WINSCHOOLEXPORT, "Fehler beim Hochladen: " + err.Message + "\n\n" + err.StackTrace, EventLogEntryType.Error);
                        message = "Fehler beim Hochladen von "+datafile+"\n\n"+err.Message + "\n\n" + err.StackTrace;
                    }
                }
            }
            catch (Exception err)
            {
                EventLog.WriteEntry(WINSCHOOLEXPORT, "Fehler beim Senden: " + err.Message + "\n\n" + err.StackTrace, EventLogEntryType.Error);
                message= "Fehler beim Senden von "+datafile+"\n\n"+ err.Message + "\n\n" + err.StackTrace;

            }
            return message;
        }
        private void sendMail(string subject, string text)
        {
            if (MAILSERVER != "" && MAILUSER != "" && MAILPASSWORD != "" )
            {
                int versuch = 0;
                while (versuch < 10)
                    try
                    {
                        System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(MAILFROM, MAILTO);
                        message.Subject = WINSCHOOLEXPORT + " " + subject;
                        message.Body = text + "\n\nDiese Nachricht wurde automatisch generiert " + DateTime.Now.ToString();
                        System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                        // Credentials are necessary if the server requires the client 
                        // to authenticate before it will send e-mail on the client's behalf.
                        /*
                        System.Net.CredentialCache mycache = new System.Net.CredentialCache();
                        mycache.Add(new Uri(WWWSERVER),
                            "Basic", new System.Net.NetworkCredential(MAILUSER, MAILPASSWORD));
                        client.Credentials = mycache;
                         */
                        client.Host = MAILSERVER;
                        client.UseDefaultCredentials = false;
                        client.Credentials = new System.Net.NetworkCredential(MAILUSER, MAILPASSWORD);
                        client.Send(message);
                        versuch = 100;
                    }
                    catch (Exception err)
                    {
                        EventLog.WriteEntry(WINSCHOOLEXPORT, "Mailen Versuch " + versuch + ":\n" + err.Message + "\n\n" + err.StackTrace, EventLogEntryType.Error);
                        versuch++;
                        if (versuch >= 10)
                        {
                            EventLog.WriteEntry(WINSCHOOLEXPORT, "Mailen Abbruch: zu viele Fehler. Keine Mail versendet.", EventLogEntryType.Error);
                        }
                        else
                        {
                            // Sekunde warten
                            System.Threading.Thread.Sleep(1000);
                        }
                    }
            }
        }
        private void exportiereDaten()
        {
            string connectionString = "DSN=" + DBSERVER + ";UID=" + DBUSER + ";PWD=" + DBPASSWORD + ";Database=" + DATABASE;
            EventLog.WriteEntry(WINSCHOOLEXPORT, "Exportiere Dateien von " + DBSERVER + " als User " + DBUSER, EventLogEntryType.Information);
            string queryString = SQLSCHULJAHR;
            string Schuljahr = "";
            int anzahl = 0;
            string message = "";
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    OdbcCommand command = new OdbcCommand(queryString, connection);
                    connection.Open();
                    OdbcDataReader reader = command.ExecuteReader();
                    try
                    {
                        System.Xml.XmlDocument xdoc = new System.Xml.XmlDocument();
                        System.Xml.XmlNode node = xdoc.CreateNode(System.Xml.XmlNodeType.Element,"winschool","");
                        xdoc.AppendChild(node);
                        while (reader.Read())
                        {
                            System.Xml.XmlNode pupil = xdoc.CreateNode(System.Xml.XmlNodeType.Element,"schuljahr", "");
                            node.AppendChild(pupil);
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                System.Xml.XmlNode entry = xdoc.CreateNode(System.Xml.XmlNodeType.Element, reader.GetName(i),"");
                                System.Xml.XmlText text = xdoc.CreateTextNode(reader[i].ToString());
                                entry.AppendChild(text);
                                pupil.AppendChild(entry);                                
                            }
                            Schuljahr = reader.GetString(1) + " " + reader.GetString(0);
                        }
                        xdoc.Save(DATAFILE+".schuljahr.xml");
                        message += sendeDaten(DATAFILE+".schuljahr.xml", 1);
                    }
                    catch (Exception err)
                    {
                        EventLog.WriteEntry(WINSCHOOLEXPORT, "Schuljahr: " + err.Message + "\n" + err.StackTrace, EventLogEntryType.Error);
                        sendMail("Fehler beim Schuljahr holen",
                            "ACHTUNG:\n\nBeim Schuljahr-Exportieren ist ein Fehler aufgetreten:\n\n" +
                            err.Message + "\n\n" + err.StackTrace + "\n\nBenutzte Einstellungen:\nBenutzer: " +
                            DBUSER + "\nServer: " + DBSERVER);
                    }
                    finally
                    {
                        // Always call Close when done reading.
                        reader.Close();
                    }               
                    queryString = SQLSCHUELER;
                    anzahl = 0;
                    command = new OdbcCommand(queryString, connection);
                    reader = command.ExecuteReader();
                    try
                    {
                        System.Xml.XmlDocument xdoc = new System.Xml.XmlDocument();
                        System.Xml.XmlNode node = xdoc.CreateNode(System.Xml.XmlNodeType.Element,"winschool","");
                        xdoc.AppendChild(node);
                        while (reader.Read())
                        {
                            System.Xml.XmlNode pupil = xdoc.CreateNode(System.Xml.XmlNodeType.Element,"schueler", "");
                            node.AppendChild(pupil);
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                System.Xml.XmlNode entry = xdoc.CreateNode(System.Xml.XmlNodeType.Element, reader.GetName(i),"");
                                System.Xml.XmlText text = xdoc.CreateTextNode(reader[i].ToString());
                                entry.AppendChild(text);
                                pupil.AppendChild(entry);                                
                            }
                            anzahl++;
                        }
                        xdoc.Save(DATAFILE+".schueler.xml");
                        message += "\n\n"+sendeDaten(DATAFILE+".schueler.xml", anzahl);
                    }
                    catch (Exception err)
                    {
                        EventLog.WriteEntry(WINSCHOOLEXPORT, "Datenschreiben: " + err.Message + "\n" + err.StackTrace, EventLogEntryType.Error);
                        sendMail("Fehler beim Daten schreiben",
                            "ACHTUNG:\n\nBeim Schüler-Exportieren ist ein Fehler aufgetreten:\n\n" +
                            err.Message + "\n\n" + err.StackTrace + "\n\nBenutzte Einstellungen:\nBenutzer: " +
                            DBUSER + "\nServer: " + DBSERVER);
                    }
                    finally
                    {
                        // Always call Close when done reading.
                        reader.Close();
                    }
                    EventLog.WriteEntry(WINSCHOOLEXPORT, "Lese Kurse für Schuljahr " + Schuljahr, EventLogEntryType.Information);   
                    queryString = SQLKURSE.Replace("[[SCHULJAHR]]",Schuljahr);
                    anzahl = 0;
                    command = new OdbcCommand(queryString, connection);
                    reader = command.ExecuteReader();
                    try
                    {
                        System.Xml.XmlDocument xdoc = new System.Xml.XmlDocument();
                        System.Xml.XmlNode node = xdoc.CreateNode(System.Xml.XmlNodeType.Element,"winschool","");
                        xdoc.AppendChild(node);
                        System.Xml.XmlNode entry = xdoc.CreateNode(System.Xml.XmlNodeType.Element, "Schuljahr", ""); // Schulhalbjahr
                        System.Xml.XmlText text = xdoc.CreateTextNode(Schuljahr);
                        entry.AppendChild(text);
                        node.AppendChild(entry);
                        while (reader.Read())
                        {
                            System.Xml.XmlNode pupil = xdoc.CreateNode(System.Xml.XmlNodeType.Element,"Schueler", "");                                                
                            entry = xdoc.CreateNode(System.Xml.XmlNodeType.Element, reader.GetName(0),"");
                            text = xdoc.CreateTextNode(reader[0].ToString());
                            entry.AppendChild(text);
                            pupil.AppendChild(entry);
                            node.AppendChild(pupil);
                            int anzahlFaecher = reader.GetInt16(1);
                            int nr = 2; // offset der vorangegangenen Felder
                            for (int i = 0; i < anzahlFaecher; i++)
                            {
                                System.Xml.XmlNode kurs = xdoc.CreateNode(System.Xml.XmlNodeType.Element, "Kurs", "");
                                entry = xdoc.CreateNode(System.Xml.XmlNodeType.Element, "Kurs", "");
                                text = xdoc.CreateTextNode(reader[nr].ToString());
                                entry.AppendChild(text);
                                kurs.AppendChild(entry);
                                nr++;
                                entry = xdoc.CreateNode(System.Xml.XmlNodeType.Element, "Fach", "");
                                text = xdoc.CreateTextNode(reader[nr].ToString());
                                entry.AppendChild(text);
                                kurs.AppendChild(entry);
                                nr++;
                                entry = xdoc.CreateNode(System.Xml.XmlNodeType.Element, "Art", "");
                                if (reader[nr].ToString() == "3")
                                    text = xdoc.CreateTextNode("gk");
                                else
                                    text = xdoc.CreateTextNode("LK");
                                entry.AppendChild(text);
                                kurs.AppendChild(entry);
                                nr++;
                                anzahl++;
                                pupil.AppendChild(kurs);
                            }                            
                        }
                        xdoc.Save(DATAFILE+".kurse.xml");
                        message += "\n\n" +sendeDaten(DATAFILE+".kurse.xml", anzahl);
                        sendMail("Export vom " + DateTime.Now.ToString(), message);
                    }
                    catch (Exception err)
                    {
                        EventLog.WriteEntry(WINSCHOOLEXPORT, "Datenschreiben: " + err.Message + "\n" + err.StackTrace, EventLogEntryType.Error);
                        sendMail("Fehler beim Daten schreiben",
                            "ACHTUNG:\n\nBeim Exportieren ist ein Fehler aufgetreten:\n\n" +
                            err.Message + "\n\n" + err.StackTrace + "\n\nBenutzte Einstellungen:\nBenutzer: " +
                            DBUSER + "\nServer: " + DBSERVER);
                    }
                    finally
                    {
                        // Always call Close when done reading.
                        reader.Close();
                    }
                    connection.Close();
                }
            }
            catch (System.Data.Odbc.OdbcException err)
            {
                EventLog.WriteEntry(WINSCHOOLEXPORT, "Export: " + err.Message + "\n" + err.StackTrace, EventLogEntryType.Error);
                /*
                try
                {
                    System.IO.File.Delete(DATAFILE);
                }
                catch (Exception)
                {
                }
                 */
                sendMail("Fehler beim Datenbankabruf",
                    "ACHTUNG:\n\nBeim Exportieren ist ein Fehler aufgetreten:\n\n" +
                    err.Message + "\n\n" + err.StackTrace + "\n\nBenutzte Einstellungen:\nBenutzer: " +
                    DBUSER + "\nServer: " + DBSERVER);
            }
        }

    }
}
